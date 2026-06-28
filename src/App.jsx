import React, { useState, useEffect } from 'react';
import { Calendar, Upload, MapPin, Clock, Link as LinkIcon, BookOpen, Loader2, ChevronLeft, ChevronRight, Sun, FileText, Video, Bell, Image as ImageIcon, Settings, Check, X } from 'lucide-react';

export default function App() {
  const [events, setEvents] = useState([]);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [currentMonth, setCurrentMonth] = useState(new Date());
  const [selectedDate, setSelectedDate] = useState(null);
  const [errorMsg, setErrorMsg] = useState('');
  const [successMsg, setSuccessMsg] = useState('');

  // ใหม่: State สำหรับเก็บรูปภาพและการแจ้งเตือน
  const [uploadedImage, setUploadedImage] = useState(null);
  const [lineToken, setLineToken] = useState('');
  const [notifyDays, setNotifyDays] = useState(1);
  const [notifiedLog, setNotifiedLog] = useState(new Set()); 
  const [showSettings, setShowSettings] = useState(false);
  const [notifyStatus, setNotifyStatus] = useState(null);

  // ฟังก์ชันสำหรับเรียกใช้ Gemini API
  const analyzeImage = async (base64Image) => {
    setIsAnalyzing(true);
    setErrorMsg('');
    setSuccessMsg('');
    
    // Read API key from environment variable or check if provided
    const apiKey = import.meta.env.VITE_GEMINI_API_KEY || "";
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;

    const payload = {
      contents: [
        {
          role: "user",
          parts: [
            { text: "Extract training or event details from this poster. If any field is missing or cannot be found, leave it as an empty string. The dates should be strictly in YYYY-MM-DD format." },
            { inlineData: { mimeType: "image/jpeg", data: base64Image.split(',')[1] } }
          ]
        }
      ],
      generationConfig: {
        responseMimeType: "application/json",
        responseSchema: {
          type: "OBJECT",
          properties: {
            events: {
              type: "ARRAY",
              items: {
                type: "OBJECT",
                properties: {
                  eventName: { type: "STRING", description: "ชื่องาน" },
                  courseName: { type: "STRING", description: "ชื่อคอร์สอบรม" },
                  topic: { type: "STRING", description: "ชื่อเรื่อง" },
                  date: { type: "STRING", description: "วันอบรม (Format YYYY-MM-DD, e.g., 2026-02-28)" },
                  time: { type: "STRING", description: "เวลาอบรม" },
                  location: { type: "STRING", description: "สถานที่อบรม" },
                  registerLink: { type: "STRING", description: "ลิ้งลงทะเบียน" },
                  joinLink: { type: "STRING", description: "ลิ้งเข้าอบรม" }
                },
                required: ["eventName", "date"]
              }
            }
          }
        }
      }
    };

    let retries = 5;
    let delay = 1000;

    while (retries > 0) {
      try {
        const response = await fetch(url, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload)
        });

        if (!response.ok) throw new Error(`API Error: ${response.status}`);
        
        const data = await response.json();
        const extractedText = data.candidates?.[0]?.content?.parts?.[0]?.text;
        
        if (extractedText) {
          const parsedData = JSON.parse(extractedText);
          if (parsedData.events && parsedData.events.length > 0) {
            // นำ base64Image แนบเข้าไปในแต่ละ event ด้วย เพื่อนำไปแสดงผล
            const eventsWithImage = parsedData.events.map(e => ({ ...e, sourceImage: base64Image }));
            setEvents(prev => [...prev, ...eventsWithImage]);
            setSuccessMsg('ดึงข้อมูลและเพิ่มลงในปฏิทินเรียบร้อยแล้ว!');
            
            // Set the selected date to the newly found event's date
            const newDate = new Date(parsedData.events[0].date);
            if (!isNaN(newDate)) {
               setCurrentMonth(newDate);
               setSelectedDate(parsedData.events[0].date);
            }
          } else {
             setErrorMsg('ไม่พบข้อมูลการอบรมในรูปภาพนี้');
          }
          break; // Success, exit retry loop
        } else {
           throw new Error("No text extracted");
        }
      } catch (err) {
        retries--;
        if (retries === 0) {
          setErrorMsg('เกิดข้อผิดพลาดในการวิเคราะห์รูปภาพ กรุณาลองใหม่อีกครั้ง');
        } else {
          await new Promise(resolve => setTimeout(resolve, delay));
          delay *= 2; // Exponential backoff
        }
      }
    }
    setIsAnalyzing(false);
  };

  const handleImageUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    if (!file.type.startsWith('image/')) {
        setErrorMsg('กรุณาอัพโหลดไฟล์รูปภาพเท่านั้น');
        return;
    }

    const reader = new FileReader();
    reader.onloadend = () => {
      setUploadedImage(reader.result);
      analyzeImage(reader.result);
    };
    reader.readAsDataURL(file);
    // Reset file input
    e.target.value = null;
  };

  // --- ระบบแจ้งเตือน LINE Notify ---
  const sendLineNotify = async (message, eventId, type) => {
    if (!lineToken) return;
    
    try {
      // Note: การยิง LINE Notify ตรงๆ จาก Browser อาจจะติด CORS 
      // โค้ดนี้เป็นการเตรียมฟังก์ชันไว้ให้ หรือใช้โหมด no-cors เพื่อให้ยิงออกไปได้แม้ไม่มี Backend
      await fetch('https://notify-api.line.me/api/notify', {
        method: 'POST',
        mode: 'no-cors', // เลี่ยง error จาก Browser (จะอ่าน Response ขากลับไม่ได้ แต่ส่งผ่าน)
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
          'Authorization': `Bearer ${lineToken}`
        },
        body: new URLSearchParams({ message: message })
      });

      // แจ้งสถานะ UI ในหน้าระบบ
      setNotifyStatus(`แจ้งเตือนเข้า LINE แล้ว: ${message.substring(0, 30)}...`);
      setTimeout(() => setNotifyStatus(null), 5000);
      
      // จดจำว่าแจ้งเตือนอีเวนต์นี้ไปแล้ว จะได้ไม่สแปมซ้ำ
      setNotifiedLog(prev => {
        const newLog = new Set(prev);
        newLog.add(`${eventId}-${type}`);
        return newLog;
      });
    } catch (error) {
      console.error("LINE Notify Error:", error);
    }
  };

  // ตรวจสอบเวลาทุกๆ 1 นาที สำหรับการแจ้งเตือน
  useEffect(() => {
    if (events.length === 0 || !lineToken) return;

    const checkSchedule = setInterval(() => {
      const now = new Date();
      const todayStr = formatDateString(now.getFullYear(), now.getMonth(), now.getDate());
      
      events.forEach((event, idx) => {
        if (!event.date) return;
        
        const eventDate = new Date(event.date);
        const diffTime = eventDate.getTime() - now.getTime();
        // คำนวณความต่างของวัน
        const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
        const eventId = `event-${idx}-${event.date}`;

        // 1. แจ้งเตือนล่วงหน้า x วัน
        if (diffDays === parseInt(notifyDays) && !notifiedLog.has(`${eventId}-advance`)) {
          sendLineNotify(`\n🔔 แจ้งเตือนล่วงหน้า ${notifyDays} วัน!\nคอร์ส: ${event.courseName || '-'}\nวันที่: ${event.date}`, eventId, 'advance');
        }

        // 2. แจ้งเตือนเมื่อถึงวันและเวลา
        if (event.date === todayStr && !notifiedLog.has(`${eventId}-now`)) {
          sendLineNotify(`\n🚨 วันนี้มีคิวอบรม!\nคอร์ส: ${event.courseName || '-'}\nเวลา: ${event.time || '-'}\nลิงก์: ${event.joinLink || '-'}`, eventId, 'now');
        }
      });
    }, 60000); // เช็คทุก 1 นาที

    return () => clearInterval(checkSchedule);
  }, [events, lineToken, notifyDays, notifiedLog]);

  // Calendar Logic
  const getDaysInMonth = (year, month) => new Date(year, month + 1, 0).getDate();
  const getFirstDayOfMonth = (year, month) => new Date(year, month, 1).getDay();

  const prevMonth = () => setCurrentMonth(new Date(currentMonth.getFullYear(), currentMonth.getMonth() - 1, 1));
  const nextMonth = () => setCurrentMonth(new Date(currentMonth.getFullYear(), currentMonth.getMonth() + 1, 1));

  const year = currentMonth.getFullYear();
  const month = currentMonth.getMonth();
  const daysInMonth = getDaysInMonth(year, month);
  const firstDay = getFirstDayOfMonth(year, month);

  const days = [];
  for (let i = 0; i < firstDay; i++) {
    days.push(null); // Empty slots
  }
  for (let i = 1; i <= daysInMonth; i++) {
    days.push(i);
  }

  const getEventsForDate = (dateStr) => {
    return events.filter(e => e.date === dateStr);
  };

  const formatDateString = (year, month, day) => {
    return `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
  };

  const selectedDateEvents = selectedDate ? getEventsForDate(selectedDate) : [];

  // Parse event date and time into start/end format for Outlook/ICS
  const parseEventDateTime = (dateStr, timeStr) => {
    if (!dateStr) return { start: '', end: '' };
    
    let startHour = "09";
    let startMin = "00";
    let endHour = "17";
    let endMin = "00";
    
    if (timeStr) {
      const times = timeStr.match(/(\d{1,2})[:.](\d{2})/g);
      if (times && times.length >= 1) {
        const startParts = times[0].split(/[:.]/);
        startHour = startParts[0].padStart(2, '0');
        startMin = startParts[1].padStart(2, '0');
        
        if (times.length >= 2) {
          const endParts = times[1].split(/[:.]/);
          endHour = endParts[0].padStart(2, '0');
          endMin = endParts[1].padStart(2, '0');
        } else {
          const hourNum = parseInt(startHour);
          endHour = String(hourNum >= 23 ? 23 : hourNum + 1).padStart(2, '0');
          endMin = startMin;
        }
      }
    }
    
    return {
      start: `${dateStr}T${startHour}:${startMin}:00`,
      end: `${dateStr}T${endHour}:${endMin}:00`
    };
  };

  const getOutlookCalendarLink = (event, platform) => {
    const { start, end } = parseEventDateTime(event.date, event.time);
    
    const baseUrl = platform === 'office365' 
      ? 'https://outlook.office.com/calendar/deeplink/compose' 
      : 'https://outlook.live.com/calendar/deeplink/compose';
      
    const subject = encodeURIComponent(event.eventName || event.courseName || 'อบรม SolarSync');
    const bodyText = `คอร์ส: ${event.courseName || '-'}\nหัวข้อ: ${event.topic || '-'}\nเวลา: ${event.time || '-'}\nสถานที่: ${event.location || '-'}\nลิงก์เข้าอบรม: ${event.joinLink || '-'}\nลิงก์ลงทะเบียน: ${event.registerLink || '-'}\n\nบันทึกจาก SolarSync Training Booking`;
    const body = encodeURIComponent(bodyText);
    const location = encodeURIComponent(event.location || '');
    
    return `${baseUrl}?path=/calendar/action/compose&rru=addevent&subject=${subject}&startdt=${start}&enddt=${end}&body=${body}&location=${location}`;
  };

  const formatToICSDateTime = (dateTimeStr) => {
    return dateTimeStr.replace(/[-:]/g, '');
  };

  const downloadICSFile = (event) => {
    const { start, end } = parseEventDateTime(event.date, event.time);
    const icsStart = formatToICSDateTime(start);
    const icsEnd = formatToICSDateTime(end);
    const timestamp = formatToICSDateTime(new Date().toISOString().split('.')[0].replace('Z', ''));
    
    const subject = event.eventName || event.courseName || 'อบรม SolarSync';
    const description = `คอร์ส: ${event.courseName || '-'}\\nหัวข้อ: ${event.topic || '-'}\\nเวลา: ${event.time || '-'}\\nสถานที่: ${event.location || '-'}\\nลิงก์เข้าอบรม: ${event.joinLink || '-'}\\nลิงก์ลงทะเบียน: ${event.registerLink || '-'}`;
    const location = event.location || '';
    
    const icsContent = [
      'BEGIN:VCALENDAR',
      'VERSION:2.0',
      'PRODID:-//SolarSync//Calendar Event//EN',
      'CALSCALE:GREGORIAN',
      'BEGIN:VEVENT',
      `UID:${Date.now()}@solarsync`,
      `DTSTAMP:${timestamp}`,
      `DTSTART:${icsStart}`,
      `DTEND:${icsEnd}`,
      `SUMMARY:${subject}`,
      `DESCRIPTION:${description}`,
      `LOCATION:${location}`,
      'END:VEVENT',
      'END:VCALENDAR'
    ].join('\n');
    
    const blob = new Blob([icsContent], { type: 'text/calendar;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.setAttribute('download', `${subject.replace(/[^a-zA-Z0-9ก-๙]/g, '_')}.ics`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800 font-sans selection:bg-orange-200">
      {/* Header */}
      <header className="bg-white shadow-sm border-b border-slate-200 sticky top-0 z-10">
        <div className="max-w-6xl mx-auto px-4 py-4 flex items-center justify-between">
          <div className="flex items-center gap-2 text-orange-600">
            <Sun className="w-7 h-7" />
            <h1 className="text-xl font-bold tracking-tight text-slate-800">Solar<span className="text-orange-600">Sync</span> Training Booking</h1>
          </div>
          <div className="flex items-center gap-4">
            <button 
              onClick={() => setShowSettings(!showSettings)}
              className={`p-2 rounded-full transition-colors ${showSettings ? 'bg-green-100 text-green-600' : 'text-slate-500 hover:text-green-600 hover:bg-green-50'}`}
              title="ตั้งค่าแจ้งเตือน LINE"
            >
              <Settings className="w-5 h-5" />
            </button>
            <div className="text-sm font-medium text-slate-500 hidden sm:block">
              วิศวกรไฟฟ้า: คุณกั้ง
            </div>
          </div>
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-4 py-8 grid grid-cols-1 lg:grid-cols-3 gap-8">
        
        {/* Left Column: Upload & Calendar */}
        <div className="lg:col-span-2 space-y-6">
          
          {/* Notifications Status Bubble */}
          {notifyStatus && (
            <div className="bg-green-600 text-white p-3 rounded-xl shadow-lg flex items-center gap-2 animate-in fade-in slide-in-from-top-4">
              <Bell className="w-5 h-5" />
              <span className="text-sm font-medium">{notifyStatus}</span>
            </div>
          )}

          {/* LINE Settings Panel */}
          {showSettings && (
            <section className="bg-slate-800 text-white p-6 rounded-2xl shadow-md border border-slate-700 animate-in fade-in slide-in-from-top-4">
              <div className="flex justify-between items-center mb-4">
                <h2 className="text-lg font-semibold flex items-center gap-2">
                  <Bell className="w-5 h-5 text-green-400" /> ตั้งค่าการแจ้งเตือน LINE Notify
                </h2>
                <button onClick={() => setShowSettings(false)} className="text-slate-400 hover:text-white"><X className="w-5 h-5" /></button>
              </div>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <label className="block text-sm text-slate-300 mb-1">LINE Notify Token</label>
                  <input 
                    type="password" 
                    value={lineToken}
                    onChange={(e) => setLineToken(e.target.value)}
                    placeholder="วาง Token ของคุณที่นี่..."
                    className="w-full bg-slate-900 border border-slate-600 rounded-lg px-4 py-2 text-sm text-white focus:outline-none focus:border-green-400"
                  />
                </div>
                <div>
                  <label className="block text-sm text-slate-300 mb-1">แจ้งเตือนล่วงหน้า (วัน)</label>
                  <input 
                    type="number" 
                    min="1"
                    value={notifyDays}
                    onChange={(e) => setNotifyDays(e.target.value)}
                    className="w-full bg-slate-900 border border-slate-600 rounded-lg px-4 py-2 text-sm text-white focus:outline-none focus:border-green-400"
                  />
                </div>
              </div>
              <p className="text-xs text-slate-400 mt-3">*ระบบจะรันการตรวจสอบและยิงแจ้งเตือนอัตโนมัติ เมื่อเปิดหน้านี้ทิ้งไว้ในเบราว์เซอร์</p>
            </section>
          )}

          {/* Upload Section */}
          <section className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100">
            <div className="flex items-center justify-between mb-4">
              <h2 className="text-lg font-semibold flex items-center gap-2">
                <Upload className="w-5 h-5 text-orange-500" /> อัพโหลดตารางอบรม / โปสเตอร์
              </h2>
              {uploadedImage && (
                <button onClick={() => setUploadedImage(null)} className="text-sm text-slate-500 hover:text-red-500 flex items-center gap-1 transition-colors">
                  <X className="w-4 h-4" /> ลบรูปภาพ
                </button>
              )}
            </div>
            
            {uploadedImage ? (
              <div className="relative rounded-xl overflow-hidden border border-slate-200 bg-slate-50 flex justify-center p-2 h-64">
                <img src={uploadedImage} alt="Uploaded Poster" className="max-h-full object-contain rounded-lg shadow-sm" />
              </div>
            ) : (
              <label className={`flex flex-col items-center justify-center w-full h-40 border-2 border-dashed rounded-xl cursor-pointer transition-colors ${isAnalyzing ? 'bg-orange-50 border-orange-200' : 'bg-slate-50 border-slate-300 hover:bg-slate-100'}`}>
                <div className="flex flex-col items-center justify-center pt-5 pb-6">
                  {isAnalyzing ? (
                    <>
                      <Loader2 className="w-10 h-10 text-orange-500 animate-spin mb-3" />
                      <p className="text-sm text-orange-600 font-medium">AI กำลังอ่านข้อมูลจากรูปภาพ...</p>
                    </>
                  ) : (
                    <>
                      <Upload className="w-10 h-10 text-slate-400 mb-3" />
                      <p className="mb-2 text-sm text-slate-600"><span className="font-semibold">คลิกเพื่ออัพโหลด</span> หรือลากไฟล์มาวางที่นี่</p>
                      <p className="text-xs text-slate-500">รองรับไฟล์ JPG, PNG</p>
                    </>
                  )}
                </div>
                <input type="file" className="hidden" accept="image/*" onChange={handleImageUpload} disabled={isAnalyzing} />
              </label>
            )}

            {/* Messages */}
            {errorMsg && <div className="mt-4 p-3 bg-red-50 text-red-600 text-sm rounded-lg border border-red-100">{errorMsg}</div>}
            {successMsg && <div className="mt-4 p-3 bg-green-50 text-green-700 text-sm rounded-lg border border-green-100">{successMsg}</div>}
          </section>

          {/* Calendar Section */}
          <section className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100">
             <div className="flex items-center justify-between mb-6">
              <h2 className="text-lg font-semibold flex items-center gap-2">
                <Calendar className="w-5 h-5 text-orange-500" /> ปฏิทินอบรม
              </h2>
              <div className="flex items-center gap-4">
                <button onClick={prevMonth} className="p-2 hover:bg-slate-100 rounded-full transition-colors"><ChevronLeft className="w-5 h-5" /></button>
                <span className="font-medium min-w-[120px] text-center">
                  {currentMonth.toLocaleString('th-TH', { month: 'long', year: 'numeric' })}
                </span>
                <button onClick={nextMonth} className="p-2 hover:bg-slate-100 rounded-full transition-colors"><ChevronRight className="w-5 h-5" /></button>
              </div>
            </div>

            <div className="grid grid-cols-7 gap-1 mb-2">
              {['อา', 'จ', 'อ', 'พ', 'พฤ', 'ศ', 'ส'].map(day => (
                <div key={day} className="text-center text-xs font-medium text-slate-500 py-2">{day}</div>
              ))}
            </div>

            <div className="grid grid-cols-7 gap-2">
              {days.map((day, index) => {
                if (day === null) return <div key={`empty-${index}`} className="h-14"></div>;
                
                const dateStr = formatDateString(year, month, day);
                const hasEvents = getEventsForDate(dateStr).length > 0;
                const isSelected = selectedDate === dateStr;
                const isToday = formatDateString(new Date().getFullYear(), new Date().getMonth(), new Date().getDate()) === dateStr;

                return (
                  <button
                    key={day}
                    onClick={() => setSelectedDate(dateStr)}
                    className={`
                      relative h-14 rounded-xl flex flex-col items-center justify-center text-sm transition-all
                      ${isSelected ? 'bg-orange-600 text-white shadow-md shadow-orange-200 ring-2 ring-orange-600 ring-offset-2' : 
                        hasEvents ? 'bg-orange-50 text-orange-900 hover:bg-orange-100 border border-orange-100' : 
                        isToday ? 'bg-slate-100 font-bold text-slate-900' : 'hover:bg-slate-50 text-slate-700'}
                    `}
                  >
                    <span>{day}</span>
                    {hasEvents && (
                      <span className={`mt-1 block w-1.5 h-1.5 rounded-full ${isSelected ? 'bg-white' : 'bg-orange-500'}`}></span>
                    )}
                  </button>
                );
              })}
            </div>
          </section>

        </div>

        {/* Right Column: Event Details */}
        <div className="space-y-6">
          <section className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100 h-full min-h-[500px]">
            <h2 className="text-lg font-semibold mb-6 flex items-center gap-2 border-b border-slate-100 pb-4">
               <FileText className="w-5 h-5 text-orange-500" /> 
               รายละเอียดการอบรม
            </h2>
            
            {!selectedDate ? (
              <div className="flex flex-col items-center justify-center h-64 text-slate-400 text-center">
                <Calendar className="w-12 h-12 mb-3 opacity-20" />
                <p>เลือกวันที่บนปฏิทิน<br/>เพื่อดูรายละเอียดการอบรม</p>
              </div>
            ) : selectedDateEvents.length === 0 ? (
              <div className="flex flex-col items-center justify-center h-64 text-slate-400 text-center">
                <p>ไม่มีคิวอบรมในวันที่<br/>{new Date(selectedDate).toLocaleDateString('th-TH', { year: 'numeric', month: 'long', day: 'numeric' })}</p>
              </div>
            ) : (
              <div className="space-y-6">
                <div className="text-sm text-slate-500 mb-2 font-medium bg-slate-50 inline-block px-3 py-1 rounded-full">
                  วันที่: {new Date(selectedDate).toLocaleDateString('th-TH', { year: 'numeric', month: 'long', day: 'numeric' })}
                </div>
                
                {selectedDateEvents.map((event, idx) => (
                  <div key={idx} className="bg-slate-50 p-5 rounded-xl border border-slate-100 relative overflow-hidden">
                    {/* Decorative bar */}
                    <div className="absolute left-0 top-0 bottom-0 w-1 bg-orange-500"></div>
                    
                    {/* โชว์รูปภาพที่อัพโหลดในช่องรายละเอียด */}
                    {event.sourceImage && (
                      <div className="mb-4 rounded-lg overflow-hidden border border-slate-200 bg-white flex justify-center">
                        <img src={event.sourceImage} alt="Training Poster" className="max-h-64 w-full object-contain" />
                      </div>
                    )}

                    <h3 className="text-xl font-bold text-slate-800 mb-1">{event.eventName || 'ไม่ได้ระบุชื่องาน'}</h3>
                    <p className="text-orange-600 font-medium mb-4 flex items-center gap-2">
                      <BookOpen className="w-4 h-4" /> {event.courseName || 'ไม่ได้ระบุชื่อคอร์ส'}
                    </p>
                    
                    <div className="space-y-3 text-sm">
                      {event.topic && (
                        <div>
                          <span className="text-slate-500 block text-xs">หัวข้อ/เรื่อง</span>
                          <span className="font-medium text-slate-700">{event.topic}</span>
                        </div>
                      )}
                      
                      <div className="grid grid-cols-2 gap-4">
                        {event.time && (
                          <div className="flex items-start gap-2">
                            <Clock className="w-4 h-4 text-slate-400 mt-0.5" />
                            <div>
                              <span className="text-slate-500 block text-xs">เวลา</span>
                              <span className="font-medium text-slate-700">{event.time}</span>
                            </div>
                          </div>
                        )}
                        {event.location && (
                          <div className="flex items-start gap-2">
                            <MapPin className="w-4 h-4 text-slate-400 mt-0.5" />
                            <div>
                              <span className="text-slate-500 block text-xs">สถานที่</span>
                              <span className="font-medium text-slate-700">{event.location}</span>
                            </div>
                          </div>
                        )}
                      </div>

                      <div className="pt-4 mt-2 border-t border-slate-200 space-y-2">
                        {event.registerLink && (
                          <a href={event.registerLink.startsWith('http') ? event.registerLink : `https://${event.registerLink}`} target="_blank" rel="noreferrer" 
                             className="flex items-center justify-center gap-2 w-full py-2 bg-white border border-slate-300 text-slate-700 rounded-lg hover:bg-slate-50 hover:text-orange-600 transition-colors font-medium">
                            <LinkIcon className="w-4 h-4" /> ลิ้งค์ลงทะเบียน
                          </a>
                        )}
                        {event.joinLink && (
                          <a href={event.joinLink.startsWith('http') ? event.joinLink : `https://${event.joinLink}`} target="_blank" rel="noreferrer"
                             className="flex items-center justify-center gap-2 w-full py-2 bg-orange-600 text-white rounded-lg hover:bg-orange-700 shadow-sm shadow-orange-200 transition-colors font-medium">
                            <Video className="w-4 h-4" /> ลิ้งค์เข้าอบรม
                          </a>
                        )}
                      </div>

                      <div className="pt-4 mt-2 border-t border-slate-200 space-y-2">
                        <span className="text-slate-500 block text-xs mb-1 font-semibold">ซิงค์ไปยัง Microsoft Calendar</span>
                        <div className="grid grid-cols-2 gap-2">
                          <a href={getOutlookCalendarLink(event, 'live')} target="_blank" rel="noreferrer"
                             className="flex items-center justify-center gap-1.5 py-1.5 px-3 bg-blue-50 text-blue-700 rounded-lg hover:bg-blue-100 transition-colors text-xs font-medium border border-blue-100">
                            <Calendar className="w-3.5 h-3.5 text-blue-500" /> Outlook (ส่วนตัว)
                          </a>
                          <a href={getOutlookCalendarLink(event, 'office365')} target="_blank" rel="noreferrer"
                             className="flex items-center justify-center gap-1.5 py-1.5 px-3 bg-indigo-50 text-indigo-700 rounded-lg hover:bg-indigo-100 transition-colors text-xs font-medium border border-indigo-100">
                            <Calendar className="w-3.5 h-3.5 text-indigo-500" /> Microsoft 365
                          </a>
                        </div>
                        <button onClick={() => downloadICSFile(event)}
                                className="flex items-center justify-center gap-2 w-full py-2 bg-slate-100 text-slate-700 rounded-lg hover:bg-slate-200 transition-colors text-sm font-medium border border-slate-200">
                          <FileText className="w-4 h-4 text-slate-500" /> ดาวน์โหลดไฟล์ปฏิทิน (.ics)
                        </button>
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </section>
        </div>
        
      </main>
    </div>
  );
}
