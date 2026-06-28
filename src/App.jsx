import React, { useState, useEffect } from 'react';
import { Calendar, Upload, MapPin, Clock, Link as LinkIcon, BookOpen, Loader2, ChevronLeft, ChevronRight, Sun, FileText, Video, Bell, Image as ImageIcon, Settings, Check, X, Plus } from 'lucide-react';
import { PublicClientApplication } from "@azure/msal-browser";

// MSAL Initialization helper
const initializeMsal = async (clientId, tenantId = 'c8445030-e602-4993-a805-7e41f70338e8') => {
  const finalClientId = clientId || '5073bf5c-1947-460d-8dd0-de9b883343d9';
  const finalTenantId = (!tenantId || tenantId === 'common') ? 'c8445030-e602-4993-a805-7e41f70338e8' : tenantId;

  const msalConfig = {
    auth: {
      clientId: finalClientId,
      authority: `https://login.microsoftonline.com/${finalTenantId}`,
      redirectUri: window.location.origin + window.location.pathname,
    },
    cache: {
      cacheLocation: "sessionStorage",
      storeAuthStateInCookie: false,
    }
  };
  const pca = new PublicClientApplication(msalConfig);
  await pca.initialize();
  return pca;
};

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

  // ใหม่: MSAL/Microsoft states
  const [msClientId, setMsClientId] = useState(localStorage.getItem('ms_client_id') || '5073bf5c-1947-460d-8dd0-de9b883343d9');
  const [msTenantId, setMsTenantId] = useState(localStorage.getItem('ms_tenant_id') || 'c8445030-e602-4993-a805-7e41f70338e8');
  const [msAccount, setMsAccount] = useState(null);
  const [msAccessToken, setMsAccessToken] = useState('');
  const [isMsSyncing, setIsMsSyncing] = useState(false);

  // ใหม่: State สำหรับกรอกข้อมูลเอง (Manual Entry)
  const [showManualForm, setShowManualForm] = useState(false);
  const [manualEvent, setManualEvent] = useState({
    eventName: '',
    courseName: '',
    topic: '',
    date: '',
    time: '',
    location: '',
    registerLink: '',
    joinLink: '',
  });
  const [manualImageFile, setManualImageFile] = useState(null);

  // Set default MS configuration on first mount if not already set
  useEffect(() => {
    if (!localStorage.getItem('ms_client_id')) {
      localStorage.setItem('ms_client_id', '5073bf5c-1947-460d-8dd0-de9b883343d9');
      setMsClientId('5073bf5c-1947-460d-8dd0-de9b883343d9');
    }
    if (!localStorage.getItem('ms_tenant_id') || localStorage.getItem('ms_tenant_id') === 'common') {
      localStorage.setItem('ms_tenant_id', 'c8445030-e602-4993-a805-7e41f70338e8');
      setMsTenantId('c8445030-e602-4993-a805-7e41f70338e8');
    }
  }, []);

  // Check active Microsoft session on startup
  useEffect(() => {
    const checkActiveSession = async () => {
      if (!msClientId) return;
      try {
        const pca = await initializeMsal(msClientId, msTenantId);
        const accounts = pca.getAllAccounts();
        if (accounts.length > 0) {
          pca.setActiveAccount(accounts[0]);
          const response = await pca.acquireTokenSilent({
            scopes: ["User.Read", "Calendars.ReadWrite"]
          });
          setMsAccount(accounts[0]);
          setMsAccessToken(response.accessToken);
        }
      } catch (err) {
        console.warn("Silent token acquisition failed, session expired:", err);
      }
    };
    checkActiveSession();
  }, [msClientId, msTenantId]);

  // Microsoft Login / Logout handlers
  const handleMicrosoftLogin = async () => {
    if (!msClientId) {
      setErrorMsg('กรุณากรอก Microsoft Client ID ในเมนูตั้งค่าก่อนเข้าสู่ระบบ');
      setShowSettings(true);
      return;
    }
    
    setErrorMsg('');
    setSuccessMsg('');
    
    try {
      const pca = await initializeMsal(msClientId, msTenantId);
      const loginRequest = {
        scopes: ["User.Read", "Calendars.ReadWrite"],
        prompt: "select_account"
      };
      
      const response = await pca.loginPopup(loginRequest);
      if (response && response.account) {
        setMsAccount(response.account);
        setMsAccessToken(response.accessToken);
        setSuccessMsg(`เชื่อมต่อ Microsoft Account สำเร็จ: ${response.account.username}`);
      }
    } catch (err) {
      console.error("Microsoft Login Error:", err);
      setErrorMsg(`เข้าสู่ระบบ Microsoft ล้มเหลว: ${err.message || err}`);
    }
  };

  const handleMicrosoftLogout = () => {
    setMsAccount(null);
    setMsAccessToken('');
    setSuccessMsg('ออกจากระบบ Microsoft เรียบร้อยแล้ว');
  };

  // Direct Calendar Booking via Graph API
  const bookDirectOutlook = async (event) => {
    if (!msAccount || !msAccessToken) {
      setErrorMsg('กรุณาเข้าสู่ระบบ Microsoft ก่อนดำเนินการซิงค์ตรง');
      return;
    }
    
    setIsMsSyncing(true);
    setErrorMsg('');
    setSuccessMsg('');
    
    try {
      const { start, end } = parseEventDateTime(event.date, event.time);
      const subject = event.eventName || event.courseName || 'อบรม SolarSync';
      
      const bodyContent = `คอร์ส: ${event.courseName || '-'}<br/>
หัวข้อ: ${event.topic || '-'}<br/>
เวลา: ${event.time || '-'}<br/>
สถานที่: ${event.location || '-'}<br/>
ลิงก์ลงทะเบียน: <a href="${event.registerLink}">${event.registerLink}</a><br/>
ลิงก์เข้าอบรม: <a href="${event.joinLink}">${event.joinLink}</a><br/><br/>
<i>บันทึกจาก SolarSync Training Booking</i>`;

      const requestBody = {
        subject: subject,
        body: {
          contentType: "HTML",
          content: bodyContent
        },
        start: {
          dateTime: start,
          timeZone: "SE Asia Standard Time"
        },
        end: {
          dateTime: end,
          timeZone: "SE Asia Standard Time"
        },
        location: {
          displayName: event.location || 'ไม่ได้ระบุสถานที่'
        }
      };
      
      const response = await fetch('https://graph.microsoft.com/v1.0/me/calendar/events', {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${msAccessToken}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(requestBody)
      });
      
      if (!response.ok) {
        const errData = await response.json();
        throw new Error(errData.error?.message || `HTTP ${response.status}`);
      }
      
      setSuccessMsg(`บันทึกชื่องาน "${subject}" ลงในปฏิทิน Microsoft ของคุณเรียบร้อยแล้ว!`);
    } catch (err) {
      console.error("Microsoft Graph API Error:", err);
      setErrorMsg(`บันทึกปฏิทินล้มเหลว: ${err.message || err}`);
    } finally {
      setIsMsSyncing(false);
    }
  };

  // ฟังก์ชันสำหรับเรียกใช้ Gemini API
  const analyzeImage = async (base64Image) => {
    setIsAnalyzing(true);
    setErrorMsg('');
    setSuccessMsg('');
    
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
            const eventsWithImage = parsedData.events.map(e => ({ ...e, sourceImage: base64Image }));
            setEvents(prev => [...prev, ...eventsWithImage]);
            setSuccessMsg('ดึงข้อมูลและเพิ่มลงในปฏิทินเรียบร้อยแล้ว!');
            
            const newDate = new Date(parsedData.events[0].date);
            if (!isNaN(newDate)) {
               setCurrentMonth(newDate);
               setSelectedDate(parsedData.events[0].date);
            }
          } else {
             setErrorMsg('ไม่พบข้อมูลการอบรมในรูปภาพนี้');
          }
          break;
        } else {
           throw new Error("No text extracted");
        }
      } catch (err) {
        retries--;
        if (retries === 0) {
          setErrorMsg('เกิดข้อผิดพลาดในการวิเคราะห์รูปภาพ กรุณาลองใหม่อีกครั้ง');
        } else {
          await new Promise(resolve => setTimeout(resolve, delay));
          delay *= 2;
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
    e.target.value = null;
  };

  // --- ระบบแจ้งเตือน LINE Notify ---
  const sendLineNotify = async (message, eventId, type) => {
    if (!lineToken) return;
    
    try {
      await fetch('https://notify-api.line.me/api/notify', {
        method: 'POST',
        mode: 'no-cors',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
          'Authorization': `Bearer ${lineToken}`
        },
        body: new URLSearchParams({ message: message })
      });

      setNotifyStatus(`แจ้งเตือนเข้า LINE แล้ว: ${message.substring(0, 30)}...`);
      setTimeout(() => setNotifyStatus(null), 5000);
      
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
        const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
        const eventId = `event-${idx}-${event.date}`;

        if (diffDays === parseInt(notifyDays) && !notifiedLog.has(`${eventId}-advance`)) {
          sendLineNotify(`\n🔔 แจ้งเตือนล่วงหน้า ${notifyDays} วัน!\nคอร์ส: ${event.courseName || '-'}\nวันที่: ${event.date}`, eventId, 'advance');
        }

        if (event.date === todayStr && !notifiedLog.has(`${eventId}-now`)) {
          sendLineNotify(`\n🚨 วันนี้มีคิวอบรม!\nคอร์ส: ${event.courseName || '-'}\nเวลา: ${event.time || '-'}\nลิงก์: ${event.joinLink || '-'}`, eventId, 'now');
        }
      });
    }, 60000);

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
    days.push(null);
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
            {msAccount ? (
              <div className="flex items-center gap-2 bg-blue-50 text-blue-700 px-3 py-1.5 rounded-lg border border-blue-100 text-xs font-semibold">
                <div className="w-1.5 h-1.5 rounded-full bg-blue-500 animate-pulse"></div>
                <span className="hidden md:inline">{msAccount.username}</span>
                <span className="md:hidden">MS Account</span>
                <button onClick={handleMicrosoftLogout} className="text-slate-400 hover:text-red-500 font-bold ml-1" title="ออกจากระบบ Microsoft">
                  <X className="w-3.5 h-3.5" />
                </button>
              </div>
            ) : (
              <button 
                onClick={handleMicrosoftLogin}
                className="bg-blue-600 hover:bg-blue-700 text-white text-xs px-3 py-1.5 rounded-lg font-semibold transition-colors flex items-center gap-1.5 shadow-sm"
              >
                <Sun className="w-3.5 h-3.5" /> เชื่อมต่อ MS Calendar
              </button>
            )}

            <button 
              onClick={() => setShowSettings(!showSettings)}
              className={`p-2 rounded-full transition-colors ${showSettings ? 'bg-slate-200 text-slate-700' : 'text-slate-500 hover:text-slate-700 hover:bg-slate-100'}`}
              title="ตั้งค่าเชื่อมต่อภายนอก"
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

          {/* Settings Panel */}
          {showSettings && (
            <section className="bg-slate-800 text-white p-6 rounded-2xl shadow-md border border-slate-700 animate-in fade-in slide-in-from-top-4 space-y-6">
              <div>
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
                <p className="text-xs text-slate-400 mt-2">*ระบบจะรันการตรวจสอบและยิงแจ้งเตือนอัตโนมัติ เมื่อเปิดหน้านี้ทิ้งไว้ในเบราว์เซอร์</p>
              </div>

              <div className="border-t border-slate-700 pt-4">
                <h2 className="text-lg font-semibold flex items-center gap-2 mb-4">
                  <Calendar className="w-5 h-5 text-blue-400" /> ตั้งค่าเชื่อมต่อ Microsoft Calendar
                </h2>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-3">
                  <div>
                    <label className="block text-sm text-slate-300 mb-1">Application (client) ID</label>
                    <input 
                      type="text" 
                      value={msClientId}
                      onChange={(e) => {
                        setMsClientId(e.target.value);
                        localStorage.setItem('ms_client_id', e.target.value);
                      }}
                      placeholder="คัดลอก Client ID จาก Azure Portal..."
                      className="w-full bg-slate-900 border border-slate-600 rounded-lg px-4 py-2 text-sm text-white focus:outline-none focus:border-blue-400"
                    />
                  </div>
                  <div>
                    <label className="block text-sm text-slate-300 mb-1">Directory (tenant) ID</label>
                    <input 
                      type="text" 
                      value={msTenantId}
                      onChange={(e) => {
                        setMsTenantId(e.target.value);
                        localStorage.setItem('ms_tenant_id', e.target.value);
                      }}
                      placeholder="ค่าเริ่มต้นคือ common (หรือระบุ Tenant ID สำหรับ Single Tenant)"
                      className="w-full bg-slate-900 border border-slate-600 rounded-lg px-4 py-2 text-sm text-white focus:outline-none focus:border-blue-400"
                    />
                  </div>
                </div>
                <p className="text-xs text-slate-400">
                  *สิทธิ์ (Permissions): แอปพลิเคชันต้องการสิทธิ์แบบ Delegated `Calendars.ReadWrite` and `User.Read`<br/>
                  *ตั้งค่า Redirect URIs ใน Azure Portal: <code className="bg-slate-900 px-1 py-0.5 rounded text-blue-300">{window.location.origin + window.location.pathname}</code>
                </p>
              </div>
            </section>
          )}

          {/* Upload & Manual Entry Section */}
          <section className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100">
            <div className="flex items-center justify-between mb-4 flex-wrap gap-2">
              <h2 className="text-lg font-semibold flex items-center gap-2">
                <Upload className="w-5 h-5 text-orange-500" /> จัดการตารางอบรม / โปสเตอร์
              </h2>
              <div className="flex items-center gap-3">
                <button 
                  onClick={() => setShowManualForm(!showManualForm)}
                  className="text-sm bg-orange-50 hover:bg-orange-100 text-orange-700 px-3 py-1.5 rounded-lg border border-orange-200 transition-colors font-semibold flex items-center gap-1.5"
                >
                  <Plus className="w-4 h-4" /> กรอกข้อมูลเอง
                </button>
                {uploadedImage && (
                  <button onClick={() => setUploadedImage(null)} className="text-sm text-slate-500 hover:text-red-500 flex items-center gap-1 transition-colors font-semibold">
                    <X className="w-4 h-4" /> ลบรูปภาพ
                  </button>
                )}
              </div>
            </div>
            
            {showManualForm && (
              <div className="mt-4 p-5 bg-slate-50 rounded-xl border border-slate-200 space-y-4 mb-4">
                <div className="flex justify-between items-center border-b border-slate-200 pb-2">
                  <h3 className="text-sm font-bold text-slate-700">กรอกข้อมูลการอบรมด้วยตนเอง</h3>
                  <button onClick={() => setShowManualForm(false)} className="text-slate-400 hover:text-slate-600"><X className="w-4 h-4" /></button>
                </div>
                
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div>
                    <label className="block text-xs font-semibold text-slate-500 mb-1">ชื่องาน *</label>
                    <input 
                      type="text" 
                      required
                      value={manualEvent.eventName}
                      onChange={(e) => setManualEvent(prev => ({ ...prev, eventName: e.target.value }))}
                      placeholder="เช่น สัมมนา C&I PV and BESS"
                      className="w-full bg-white border border-slate-300 rounded-lg px-3 py-1.5 text-sm text-slate-700 focus:outline-none focus:border-orange-500"
                    />
                  </div>
                  
                  <div>
                    <label className="block text-xs font-semibold text-slate-500 mb-1">ชื่อคอร์สอบรม *</label>
                    <input 
                      type="text" 
                      required
                      value={manualEvent.courseName}
                      onChange={(e) => setManualEvent(prev => ({ ...prev, courseName: e.target.value }))}
                      placeholder="เช่น C&I PV and BESS Course"
                      className="w-full bg-white border border-slate-300 rounded-lg px-3 py-1.5 text-sm text-slate-700 focus:outline-none focus:border-orange-500"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-slate-500 mb-1">หัวข้อ/เรื่อง</label>
                    <input 
                      type="text" 
                      value={manualEvent.topic}
                      onChange={(e) => setManualEvent(prev => ({ ...prev, topic: e.target.value }))}
                      placeholder="เช่น การออกแบบระบบโซล่าร์เซลล์และแบตเตอรี่"
                      className="w-full bg-white border border-slate-300 rounded-lg px-3 py-1.5 text-sm text-slate-700 focus:outline-none focus:border-orange-500"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-slate-500 mb-1">วันที่อบรม *</label>
                    <input 
                      type="date" 
                      required
                      value={manualEvent.date}
                      onChange={(e) => setManualEvent(prev => ({ ...prev, date: e.target.value }))}
                      className="w-full bg-white border border-slate-300 rounded-lg px-3 py-1.5 text-sm text-slate-700 focus:outline-none focus:border-orange-500"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-slate-500 mb-1">เวลาอบรม</label>
                    <input 
                      type="text" 
                      value={manualEvent.time}
                      onChange={(e) => setManualEvent(prev => ({ ...prev, time: e.target.value }))}
                      placeholder="เช่น 13:00 - 15:30 น."
                      className="w-full bg-white border border-slate-300 rounded-lg px-3 py-1.5 text-sm text-slate-700 focus:outline-none focus:border-orange-500"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-slate-500 mb-1">สถานที่อบรม</label>
                    <input 
                      type="text" 
                      value={manualEvent.location}
                      onChange={(e) => setManualEvent(prev => ({ ...prev, location: e.target.value }))}
                      placeholder="เช่น โรงแรมคราวน์ พลาซ่า หรือ Online Zoom"
                      className="w-full bg-white border border-slate-300 rounded-lg px-3 py-1.5 text-sm text-slate-700 focus:outline-none focus:border-orange-500"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-slate-500 mb-1">ลิงก์ลงทะเบียน</label>
                    <input 
                      type="text" 
                      value={manualEvent.registerLink}
                      onChange={(e) => setManualEvent(prev => ({ ...prev, registerLink: e.target.value }))}
                      placeholder="วางลิงก์ลงทะเบียน..."
                      className="w-full bg-white border border-slate-300 rounded-lg px-3 py-1.5 text-sm text-slate-700 focus:outline-none focus:border-orange-500"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-slate-500 mb-1">ลิงก์เข้าอบรม</label>
                    <input 
                      type="text" 
                      value={manualEvent.joinLink}
                      onChange={(e) => setManualEvent(prev => ({ ...prev, joinLink: e.target.value }))}
                      placeholder="วางลิงก์เข้า Zoom / Teams..."
                      className="w-full bg-white border border-slate-300 rounded-lg px-3 py-1.5 text-sm text-slate-700 focus:outline-none focus:border-orange-500"
                    />
                  </div>

                  <div className="md:col-span-2">
                    <label className="block text-xs font-semibold text-slate-500 mb-1">แนบรูปภาพ (ถ้ามี)</label>
                    <input 
                      type="file" 
                      accept="image/*"
                      onChange={(e) => {
                        const file = e.target.files[0];
                        if (file) {
                          const reader = new FileReader();
                          reader.onloadend = () => {
                            setManualImageFile(reader.result);
                          };
                          reader.readAsDataURL(file);
                        }
                      }}
                      className="w-full text-xs text-slate-500 file:mr-4 file:py-1.5 file:px-3 file:rounded-lg file:border-0 file:text-xs file:font-semibold file:bg-orange-50 file:text-orange-700 hover:file:bg-orange-100"
                    />
                  </div>
                </div>

                <div className="flex justify-end gap-3 pt-2">
                  <button 
                    type="button"
                    onClick={() => {
                      setShowManualForm(false);
                      setManualEvent({ eventName: '', courseName: '', topic: '', date: '', time: '', location: '', registerLink: '', joinLink: '' });
                      setManualImageFile(null);
                    }}
                    className="px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-lg text-sm font-semibold transition-colors"
                  >
                    ยกเลิก
                  </button>
                  <button 
                    type="button"
                    onClick={() => {
                      if (!manualEvent.eventName || !manualEvent.date || !manualEvent.courseName) {
                        setErrorMsg('กรุณากรอกข้อมูลที่จำเป็น (* ชื่องาน, ชื่อคอร์ส, วันที่)');
                        return;
                      }
                      const newEvent = { 
                        ...manualEvent, 
                        sourceImage: manualImageFile 
                      };
                      setEvents(prev => [...prev, newEvent]);
                      setSuccessMsg('เพิ่มงานอบรมเข้าสู่ปฏิทินเรียบร้อยแล้ว!');
                      setCurrentMonth(new Date(manualEvent.date));
                      setSelectedDate(manualEvent.date);
                      setShowManualForm(false);
                      setManualEvent({ eventName: '', courseName: '', topic: '', date: '', time: '', location: '', registerLink: '', joinLink: '' });
                      setManualImageFile(null);
                    }}
                    className="px-4 py-2 bg-orange-600 hover:bg-orange-700 text-white rounded-lg text-sm font-semibold transition-colors shadow-sm"
                  >
                    เพิ่มงานอบรม
                  </button>
                </div>
              </div>
            )}
            
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
                        
                        {msAccount ? (
                          <button 
                            onClick={() => bookDirectOutlook(event)}
                            disabled={isMsSyncing}
                            className="flex items-center justify-center gap-2 w-full py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 disabled:bg-blue-300 transition-colors text-sm font-semibold shadow-sm"
                          >
                            {isMsSyncing ? (
                              <>
                                <Loader2 className="w-4 h-4 animate-spin" /> กำลังบันทึกลงปฏิทิน...
                              </>
                            ) : (
                              <>
                                <Check className="w-4 h-4" /> บันทึกลง MS Calendar ทันที
                              </>
                            )}
                          </button>
                        ) : (
                          <button 
                            onClick={handleMicrosoftLogin}
                            className="flex items-center justify-center gap-2 w-full py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg transition-colors text-sm font-semibold border border-slate-200"
                          >
                            <Sun className="w-4 h-4 text-blue-500" /> ล็อกอิน Microsoft เพื่อบันทึกด่วน
                          </button>
                        )}

                        <div className="grid grid-cols-2 gap-2 pt-1">
                          <a href={getOutlookCalendarLink(event, 'live')} target="_blank" rel="noreferrer"
                             className="flex items-center justify-center gap-1.5 py-1.5 px-3 bg-blue-50 text-blue-700 rounded-lg hover:bg-blue-100 transition-colors text-xs font-medium border border-blue-100">
                            <Calendar className="w-3.5 h-3.5 text-blue-500" /> Link Outlook (ส่วนตัว)
                          </a>
                          <a href={getOutlookCalendarLink(event, 'office365')} target="_blank" rel="noreferrer"
                             className="flex items-center justify-center gap-1.5 py-1.5 px-3 bg-indigo-50 text-indigo-700 rounded-lg hover:bg-indigo-100 transition-colors text-xs font-medium border border-indigo-100">
                            <Calendar className="w-3.5 h-3.5 text-indigo-500" /> Link Microsoft 365
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
