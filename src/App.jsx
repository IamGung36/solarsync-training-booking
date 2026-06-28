import React, { useState, useEffect } from 'react';
import { Calendar, Upload, MapPin, Clock, Link as LinkIcon, BookOpen, Loader2, ChevronLeft, ChevronRight, Sun, FileText, Video, Bell, Image as ImageIcon, Settings, Check, X, Plus } from 'lucide-react';
import { PublicClientApplication } from "@azure/msal-browser";

// MSAL Initialization helper
const initializeMsal = async (clientId, tenantId = 'cb445030-e602-4993-a605-7e41f70338e8') => {
  const finalClientId = clientId || '5073bf5c-1947-460d-8dd0-de9b883343d9';
  const finalTenantId = (!tenantId || tenantId === 'common') ? 'cb445030-e602-4993-a605-7e41f70338e8' : tenantId;

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

// Formatting date to Thai locale but using Gregorian Calendar (ค.ศ.)
const formatDisplayDate = (dateStr) => {
  if (!dateStr) return '';
  const dateObj = new Date(dateStr);
  if (isNaN(dateObj)) return dateStr;
  return dateObj.toLocaleDateString('th-TH-u-ca-gregory', {
    day: 'numeric',
    month: 'long',
    year: 'numeric'
  });
};

export default function App() {
  const [events, setEvents] = useState([]);
  const [msEvents, setMsEvents] = useState([]); // ดึงกิจกรรมจาก Microsoft Calendar
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [currentMonth, setCurrentMonth] = useState(new Date());
  const [selectedDate, setSelectedDate] = useState(null);
  const [errorMsg, setErrorMsg] = useState('');
  const [successMsg, setSuccessMsg] = useState('');

  // LINE Settings States
  const [uploadedImage, setUploadedImage] = useState(null);
  const [lineAccessToken, setLineAccessToken] = useState(localStorage.getItem('line_access_token') || localStorage.getItem('line_token') || 'VE6pRAQO6w7nKGvGPrpUmSuFZZb5n8+FciujImJMAIEqPbjHGMBF4aI6sOY+A3XME0trehIxYBNswHObwUFjIVK0MV0E7b6BL9NaqkdzR0JHnrYAjgpD7GZ147O4CujXQ9270Onm/aL05XKFapXl1AdB04t89/1O/w1cDnyilFU=');
  const [lineUserId, setLineUserId] = useState(localStorage.getItem('line_user_id') || 'gung1125');
  const [notifyDays, setNotifyDays] = useState(1);
  const [notifiedLog, setNotifiedLog] = useState(new Set()); 
  const [showSettings, setShowSettings] = useState(false);
  const [notifyStatus, setNotifyStatus] = useState(null);

  // Microsoft OAuth States
  const [msClientId, setMsClientId] = useState(localStorage.getItem('ms_client_id') || '5073bf5c-1947-460d-8dd0-de9b883343d9');
  const [msTenantId, setMsTenantId] = useState(localStorage.getItem('ms_tenant_id') || 'cb445030-e602-4993-a605-7e41f70338e8');
  const [msAccount, setMsAccount] = useState(null);
  const [msAccessToken, setMsAccessToken] = useState('');
  const [isMsSyncing, setIsMsSyncing] = useState(false);

  // Manual Event Form States
  const [showManualForm, setShowManualForm] = useState(false);
  const [manualEvent, setManualEvent] = useState({
    eventName: '',
    courseName: '',
    topic: '',
    date: '',
    startTime: '',
    endTime: '',
    location: '',
    registerLink: '',
    joinLink: '',
    bookingType: 'work',
  });
  const [manualImageFile, setManualImageFile] = useState(null);

  // Inline Editing States for Card
  const [editingEvent, setEditingEvent] = useState(null);
  const [editDate, setEditDate] = useState('');
  const [editTime, setEditTime] = useState('');

  // Set default MS and LINE configuration on first mount if not already set
  useEffect(() => {
    if (!localStorage.getItem('ms_client_id')) {
      localStorage.setItem('ms_client_id', '5073bf5c-1947-460d-8dd0-de9b883343d9');
      setMsClientId('5073bf5c-1947-460d-8dd0-de9b883343d9');
    }
    const curTenant = localStorage.getItem('ms_tenant_id');
    if (!curTenant || curTenant === 'common' || curTenant === 'c8445030-e602-4993-a805-7e41f70338e8' || curTenant === '') {
      localStorage.setItem('ms_tenant_id', 'cb445030-e602-4993-a605-7e41f70338e8');
      setMsTenantId('cb445030-e602-4993-a605-7e41f70338e8');
    }
    if (!localStorage.getItem('line_access_token')) {
      localStorage.setItem('line_access_token', 'VE6pRAQO6w7nKGvGPrpUmSuFZZb5n8+FciujImJMAIEqPbjHGMBF4aI6sOY+A3XME0trehIxYBNswHObwUFjIVK0MV0E7b6BL9NaqkdzR0JHnrYAjgpD7GZ147O4CujXQ9270Onm/aL05XKFapXl1AdB04t89/1O/w1cDnyilFU=');
      setLineAccessToken('VE6pRAQO6w7nKGvGPrpUmSuFZZb5n8+FciujImJMAIEqPbjHGMBF4aI6sOY+A3XME0trehIxYBNswHObwUFjIVK0MV0E7b6BL9NaqkdzR0JHnrYAjgpD7GZ147O4CujXQ9270Onm/aL05XKFapXl1AdB04t89/1O/w1cDnyilFU=');
    }
    if (!localStorage.getItem('line_user_id')) {
      localStorage.setItem('line_user_id', 'gung1125');
      setLineUserId('gung1125');
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

  // Fetch Microsoft Calendar Events
  const fetchMicrosoftEvents = async () => {
    if (!msAccessToken) return;
    try {
      // ดึงนัดหมาย 100 รายการล่าสุด
      const response = await fetch('https://graph.microsoft.com/v1.0/me/calendar/events?$top=100&$select=subject,start,end,location,bodyPreview', {
        headers: {
          'Authorization': `Bearer ${msAccessToken}`,
          'Prefer': 'outlook.timezone="SE Asia Standard Time"'
        }
      });
      
      if (!response.ok) {
        throw new Error(`HTTP ${response.status}`);
      }
      
      const data = await response.json();
      if (data.value) {
        const mapped = data.value.map(item => {
          const startDateStr = item.start.dateTime.split('T')[0];
          const startTimeStr = item.start.dateTime.split('T')[1].substring(0, 5);
          const endTimeStr = item.end.dateTime.split('T')[1].substring(0, 5);
          
          return {
            eventName: item.subject,
            topic: item.bodyPreview || '',
            date: startDateStr,
            time: `${startTimeStr} - ${endTimeStr} น.`,
            location: item.location?.displayName || '',
            bookingType: 'work', // ดึงมาเป็นงานจัดประเภทที่ทำงาน
            isMsEvent: true,     // ทำเครื่องหมายว่าเป็นกิจกรรมของระบบ Microsoft
            id: item.id
          };
        });
        setMsEvents(mapped);
      }
    } catch (err) {
      console.error("Error fetching MS events:", err);
    }
  };

  // Sync MS Calendar events whenever msAccessToken is set/updated
  useEffect(() => {
    if (msAccessToken) {
      fetchMicrosoftEvents();
    } else {
      setMsEvents([]);
    }
  }, [msAccessToken]);

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
    setMsEvents([]);
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
      const subject = event.eventName || event.courseName || 'นัดหมาย My Booking Calendar';
      
      const bodyContent = `ชื่องาน: ${event.eventName || '-'}<br/>
คอร์ส: ${event.courseName || '-'}<br/>
หัวข้อ: ${event.topic || '-'}<br/>
เวลา: ${event.time || '-'}<br/>
สถานที่: ${event.location || '-'}<br/>
ประเภทนัดหมาย: ${event.bookingType === 'personal' ? 'Booking ส่วนตัว' : 'Booking ที่ทำงาน'}<br/>
ลิงก์ลงทะเบียน: <a href="${event.registerLink}">${event.registerLink}</a><br/>
ลิงก์เข้าอบรม: <a href="${event.joinLink}">${event.joinLink}</a><br/><br/>
<i>บันทึกจาก My Booking Calendar</i>`;

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
      // ดึงข้อมูลกิจกรรมล่าสุดจาก MS Calendar อีกรอบเพื่ออัปเดตบนหน้าปฏิทิน
      fetchMicrosoftEvents();
    } catch (err) {
      console.error("Microsoft Graph API Error:", err);
      setErrorMsg(`บันทึกปฏิทินล้มเหลว: ${err.message || err}`);
    } finally {
      setIsMsSyncing(false);
    }
  };

  // Gemini API Poster Analysis
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
            const eventsWithImage = parsedData.events.map(e => ({ 
              ...e, 
              bookingType: 'work', // default to work
              sourceImage: base64Image 
            }));
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

  // --- LINE Messaging API Push Function ---
  const sendLineMessage = async (message, eventId, type) => {
    if (!lineAccessToken || !lineUserId) return;
    
    try {
      const response = await fetch('https://corsproxy.io/?https://api.line.me/v2/bot/message/push', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${lineAccessToken}`
        },
        body: JSON.stringify({
          to: lineUserId,
          messages: [
            {
              type: 'text',
              text: message
            }
          ]
        })
      });

      if (!response.ok) {
        throw new Error(`HTTP ${response.status}`);
      }

      setNotifyStatus(`แจ้งเตือนเข้า LINE บอทสำเร็จ: ${message.substring(0, 30)}...`);
      setTimeout(() => setNotifyStatus(null), 5000);
      
      setNotifiedLog(prev => {
        const newLog = new Set(prev);
        newLog.add(`${eventId}-${type}`);
        return newLog;
      });
    } catch (error) {
      console.error("LINE Messaging API Error:", error);
      setNotifyStatus(`การส่งแจ้งเตือน LINE ล้มเหลว: ${error.message || error}`);
      setTimeout(() => setNotifyStatus(null), 5000);
    }
  };

  // ตรวจสอบเวลาทุกๆ 1 นาที สำหรับการแจ้งเตือน
  useEffect(() => {
    if (events.length === 0 || !lineAccessToken || !lineUserId) return;

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
          sendLineMessage(`[แจ้งเตือนนัดหมายล่วงหน้า ${notifyDays} วัน]\n📌 ชื่องาน: ${event.eventName || '-'}\n📅 วันที่: ${formatDisplayDate(event.date)}\n🕒 เวลา: ${event.time || '-'}\nประเภท: ${event.bookingType === 'personal' ? 'ส่วนตัว' : 'ที่ทำงาน'}`, eventId, 'advance');
        }

        if (event.date === todayStr && !notifiedLog.has(`${eventId}-now`)) {
          sendLineMessage(`[แจ้งเตือนนัดหมายวันนี้!]\n🚨 ชื่องาน: ${event.eventName || '-'}\n🕒 เวลา: ${event.time || '-'}\n📍 สถานที่: ${event.location || '-'}\n🔗 ลิงก์เข้า: ${event.joinLink || '-'}`, eventId, 'now');
        }
      });
    }, 60000);

    return () => clearInterval(checkSchedule);
  }, [events, lineAccessToken, lineUserId, notifyDays, notifiedLog]);

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

  // รวมกิจกรรมทั้ง Local และ Microsoft Calendar
  const getEventsForDate = (dateStr) => {
    const local = events.filter(e => e.date === dateStr);
    const ms = msEvents.filter(e => e.date === dateStr);
    return [...local, ...ms];
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
      
    const subject = encodeURIComponent(event.eventName || event.courseName || 'นัดหมาย My Booking Calendar');
    const bodyText = `ชื่องาน: ${event.eventName || '-'}\nคอร์ส: ${event.courseName || '-'}\nหัวข้อ: ${event.topic || '-'}\nเวลา: ${event.time || '-'}\nสถานที่: ${event.location || '-'}\nประเภทนัดหมาย: ${event.bookingType === 'personal' ? 'ส่วนตัว' : 'ที่ทำงาน'}\nลิงก์เข้าอบรม: ${event.joinLink || '-'}\nลิงก์ลงทะเบียน: ${event.registerLink || '-'}\n\nบันทึกจาก My Booking Calendar`;
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
    
    const subject = event.eventName || event.courseName || 'นัดหมาย My Booking Calendar';
    const description = `ชื่องาน: ${event.eventName || '-'}\\nคอร์ส: ${event.courseName || '-'}\\nหัวข้อ: ${event.topic || '-'}\\nเวลา: ${event.time || '-'}\\nสถานที่: ${event.location || '-'}\\nประเภทนัดหมาย: ${event.bookingType === 'personal' ? 'ส่วนตัว' : 'ที่ทำงาน'}\\nลิงก์เข้าอบรม: ${event.joinLink || '-'}\\nลิงก์ลงทะเบียน: ${event.registerLink || '-'}`;
    const location = event.location || '';
    
    const icsContent = [
      'BEGIN:VCALENDAR',
      'VERSION:2.0',
      'PRODID:-//MyBookingCalendar//Calendar Event//EN',
      'CALSCALE:GREGORIAN',
      'BEGIN:VEVENT',
      `UID:${Date.now()}@mybookingcalendar`,
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
    <div className="min-h-screen bg-slate-50 text-slate-800 font-sans selection:bg-blue-200">
      {/* Header */}
      <header className="bg-white shadow-sm border-b border-slate-200 sticky top-0 z-10">
        <div className="max-w-6xl mx-auto px-4 py-4 flex items-center justify-between">
          <div className="flex items-center gap-2 text-blue-600">
            <Calendar className="w-7 h-7 text-blue-600" />
            <h1 className="text-xl font-bold tracking-tight text-slate-800">My <span className="text-blue-600">Booking Calendar</span></h1>
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
            <div className="bg-blue-600 text-white p-3 rounded-xl shadow-lg flex items-center gap-2 animate-in fade-in slide-in-from-top-4">
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
                    <Bell className="w-5 h-5 text-green-400" /> ตั้งค่าการแจ้งเตือน LINE Messaging API
                  </h2>
                  <button onClick={() => setShowSettings(false)} className="text-slate-400 hover:text-white"><X className="w-5 h-5" /></button>
                </div>
                <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                  <div className="md:col-span-2">
                    <label className="block text-sm text-slate-300 mb-1">LINE Channel Access Token</label>
                    <input 
                      type="password" 
                      value={lineAccessToken}
                      onChange={(e) => {
                        setLineAccessToken(e.target.value);
                        localStorage.setItem('line_access_token', e.target.value);
                      }}
                      placeholder="วาง Channel Access Token ของคุณ..."
                      className="w-full bg-slate-900 border border-slate-600 rounded-lg px-4 py-2 text-sm text-white focus:outline-none focus:border-blue-400"
                    />
                  </div>
                  <div>
                    <label className="block text-sm text-slate-300 mb-1">LINE User ID</label>
                    <input 
                      type="text" 
                      value={lineUserId}
                      onChange={(e) => {
                        setLineUserId(e.target.value);
                        localStorage.setItem('line_user_id', e.target.value);
                      }}
                      placeholder="เช่น U12345..."
                      className="w-full bg-slate-900 border border-slate-600 rounded-lg px-4 py-2 text-sm text-white focus:outline-none focus:border-blue-400"
                    />
                  </div>
                </div>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mt-3">
                  <div>
                    <label className="block text-sm text-slate-300 mb-1">แจ้งเตือนล่วงหน้า (วัน)</label>
                    <input 
                      type="number" 
                      min="1"
                      value={notifyDays}
                      onChange={(e) => setNotifyDays(e.target.value)}
                      className="w-full bg-slate-900 border border-slate-600 rounded-lg px-4 py-2 text-sm text-white focus:outline-none focus:border-blue-400"
                    />
                  </div>
                </div>
                <p className="text-xs text-slate-400 mt-2">*ส่งผ่านระบบ Proxy เพื่อเลี่ยงข้อจำกัด CORS ฝั่งเบราว์เซอร์ และแจ้งเตือนตรงถึง User ID ของคุณ</p>
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
                      placeholder="ระบบระบุค่าเริ่มต้นขององค์กรคุณไว้แล้ว..."
                      className="w-full bg-slate-900 border border-slate-600 rounded-lg px-4 py-2 text-sm text-white focus:outline-none focus:border-blue-400"
                    />
                  </div>
                </div>
                <p className="text-xs text-slate-400">
                  *สิทธิ์ (Permissions): แอปพลิเคชันต้องการสิทธิ์แบบ Delegated `Calendars.ReadWrite` และ `User.Read`<br/>
                  *ตั้งค่า Redirect URIs ใน Azure Portal: <code className="bg-slate-900 px-1 py-0.5 rounded text-blue-300">{window.location.origin + window.location.pathname}</code>
                </p>
              </div>
            </section>
          )}

          {/* Upload & Manual Entry Section */}
          <section className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100">
            <div className="flex items-center justify-between mb-4 flex-wrap gap-2">
              <h2 className="text-lg font-semibold flex items-center gap-2">
                <Upload className="w-5 h-5 text-blue-500" /> จัดการตารางนัดหมาย / โปสเตอร์
              </h2>
              <div className="flex items-center gap-3">
                <button 
                  onClick={() => setShowManualForm(!showManualForm)}
                  className="text-sm bg-blue-50 hover:bg-blue-100 text-blue-700 px-3 py-1.5 rounded-lg border border-blue-200 transition-colors font-semibold flex items-center gap-1.5"
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
              <div className="mt-4 p-5 bg-slate-50 rounded-xl border border-slate-200 space-y-4 mb-4 animate-in fade-in duration-200">
                <div className="flex justify-between items-center border-b border-slate-200 pb-2">
                  <h3 className="text-sm font-bold text-slate-700 font-sans">กรอกข้อมูลการนัดหมายด้วยตนเอง</h3>
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
                      placeholder="เช่น ประชุมแผนงานประจำไตรมาส"
                      className="w-full bg-white border border-slate-300 rounded-lg px-3 py-1.5 text-sm text-slate-700 focus:outline-none focus:border-blue-500"
                    />
                  </div>
                  
                  <div>
                    <label className="block text-xs font-semibold text-slate-500 mb-1">ชื่อคอร์สอบรม (ไม่บังคับ)</label>
                    <input 
                      type="text" 
                      value={manualEvent.courseName}
                      onChange={(e) => setManualEvent(prev => ({ ...prev, courseName: e.target.value }))}
                      placeholder="เช่น C&I PV Course (ถ้ามี)"
                      className="w-full bg-white border border-slate-300 rounded-lg px-3 py-1.5 text-sm text-slate-700 focus:outline-none focus:border-blue-500"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-slate-500 mb-1">หัวข้อ/เรื่อง</label>
                    <input 
                      type="text" 
                      value={manualEvent.topic}
                      onChange={(e) => setManualEvent(prev => ({ ...prev, topic: e.target.value }))}
                      placeholder="เช่น อัปเดตงานออกแบบโซล่าร์รูฟท็อป"
                      className="w-full bg-white border border-slate-300 rounded-lg px-3 py-1.5 text-sm text-slate-700 focus:outline-none focus:border-blue-500"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-slate-500 mb-1">วันที่จัดงาน *</label>
                    <input 
                      type="date" 
                      required
                      value={manualEvent.date}
                      onChange={(e) => setManualEvent(prev => ({ ...prev, date: e.target.value }))}
                      className="w-full bg-white border border-slate-300 rounded-lg px-3 py-1.5 text-sm text-slate-700 focus:outline-none focus:border-blue-500"
                    />
                    <span className="text-[11px] text-slate-500 mt-1 block">
                      {manualEvent.date ? `แสดงผลเป็น: ${formatDisplayDate(manualEvent.date)}` : 'ตัวอย่าง: 16 มิถุนายน 2026'}
                    </span>
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-slate-500 mb-1">เวลาเริ่ม (ไม่บังคับ)</label>
                    <input 
                      type="time" 
                      value={manualEvent.startTime}
                      onChange={(e) => setManualEvent(prev => ({ ...prev, startTime: e.target.value }))}
                      className="w-full bg-white border border-slate-300 rounded-lg px-3 py-1.5 text-sm text-slate-700 focus:outline-none focus:border-blue-500"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-slate-500 mb-1">เวลาสิ้นสุด</label>
                    <input 
                      type="time" 
                      value={manualEvent.endTime}
                      onChange={(e) => setManualEvent(prev => ({ ...prev, endTime: e.target.value }))}
                      className="w-full bg-white border border-slate-300 rounded-lg px-3 py-1.5 text-sm text-slate-700 focus:outline-none focus:border-blue-500"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-slate-500 mb-1">สถานที่จัดงาน</label>
                    <input 
                      type="text" 
                      value={manualEvent.location}
                      onChange={(e) => setManualEvent(prev => ({ ...prev, location: e.target.value }))}
                      placeholder="เช่น ห้องประชุมใหญ่ หรือ Microsoft Teams"
                      className="w-full bg-white border border-slate-300 rounded-lg px-3 py-1.5 text-sm text-slate-700 focus:outline-none focus:border-blue-500"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-slate-500 mb-1">ประเภทการนัดหมาย *</label>
                    <select
                      value={manualEvent.bookingType}
                      onChange={(e) => setManualEvent(prev => ({ ...prev, bookingType: e.target.value }))}
                      className="w-full bg-white border border-slate-300 rounded-lg px-3 py-1.5 text-sm text-slate-700 focus:outline-none focus:border-blue-500"
                    >
                      <option value="work">💼 Booking ที่ทำงาน (Work)</option>
                      <option value="personal">🏠 Booking ส่วนตัว (Personal)</option>
                    </select>
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-slate-500 mb-1">ลิงก์ลงทะเบียน</label>
                    <input 
                      type="text" 
                      value={manualEvent.registerLink}
                      onChange={(e) => setManualEvent(prev => ({ ...prev, registerLink: e.target.value }))}
                      placeholder="วางลิงก์ลงทะเบียน..."
                      className="w-full bg-white border border-slate-300 rounded-lg px-3 py-1.5 text-sm text-slate-700 focus:outline-none focus:border-blue-500"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-slate-500 mb-1">ลิงก์เข้าจัดงาน (Zoom/Teams)</label>
                    <input 
                      type="text" 
                      value={manualEvent.joinLink}
                      onChange={(e) => setManualEvent(prev => ({ ...prev, joinLink: e.target.value }))}
                      placeholder="วางลิงก์เข้าประชุม..."
                      className="w-full bg-white border border-slate-300 rounded-lg px-3 py-1.5 text-sm text-slate-700 focus:outline-none focus:border-blue-500"
                    />
                  </div>

                  <div className="md:col-span-2">
                    <label className="block text-xs font-semibold text-slate-500 mb-1">แนบรูปภาพโปสเตอร์ (ถ้ามี)</label>
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
                      className="w-full text-xs text-slate-500 file:mr-4 file:py-1.5 file:px-3 file:rounded-lg file:border-0 file:text-xs file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100"
                    />
                  </div>
                </div>

                <div className="flex justify-end gap-3 pt-2">
                  <button 
                    type="button"
                    onClick={() => {
                      setShowManualForm(false);
                      setManualEvent({ eventName: '', courseName: '', topic: '', date: '', startTime: '', endTime: '', location: '', registerLink: '', joinLink: '', bookingType: 'work' });
                      setManualImageFile(null);
                    }}
                    className="px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-lg text-sm font-semibold transition-colors"
                  >
                    ยกเลิก
                  </button>
                  <button 
                    type="button"
                    onClick={() => {
                      if (!manualEvent.eventName || !manualEvent.date) {
                        setErrorMsg('กรุณากรอกข้อมูลที่จำเป็น (* ชื่องาน, วันที่)');
                        return;
                      }
                      
                      let formattedTime = '';
                      if (manualEvent.startTime) {
                        formattedTime = manualEvent.startTime;
                        if (manualEvent.endTime) {
                          formattedTime += ` - ${manualEvent.endTime}`;
                        }
                        formattedTime += ' น.';
                      }

                      const newEvent = { 
                        eventName: manualEvent.eventName,
                        courseName: manualEvent.courseName,
                        topic: manualEvent.topic,
                        date: manualEvent.date,
                        time: formattedTime,
                        location: manualEvent.location,
                        registerLink: manualEvent.registerLink,
                        joinLink: manualEvent.joinLink,
                        bookingType: manualEvent.bookingType,
                        sourceImage: manualImageFile 
                      };

                      setEvents(prev => [...prev, newEvent]);
                      setSuccessMsg('เพิ่มนัดหมายเข้าสู่ปฏิทินเรียบร้อยแล้ว!');
                      setCurrentMonth(new Date(manualEvent.date));
                      setSelectedDate(manualEvent.date);
                      setShowManualForm(false);
                      setManualEvent({ eventName: '', courseName: '', topic: '', date: '', startTime: '', endTime: '', location: '', registerLink: '', joinLink: '', bookingType: 'work' });
                      setManualImageFile(null);
                    }}
                    className="px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-sm font-semibold transition-colors shadow-sm"
                  >
                    เพิ่มรายการนัดหมาย
                  </button>
                </div>
              </div>
            )}
            
            {uploadedImage ? (
              <div className="relative rounded-xl overflow-hidden border border-slate-200 bg-slate-50 flex justify-center p-2 h-64">
                <img src={uploadedImage} alt="Uploaded Poster" className="max-h-full object-contain rounded-lg shadow-sm" />
              </div>
            ) : (
              <label className={`flex flex-col items-center justify-center w-full h-40 border-2 border-dashed rounded-xl cursor-pointer transition-colors ${isAnalyzing ? 'bg-blue-50 border-blue-200' : 'bg-slate-50 border-slate-300 hover:bg-slate-100'}`}>
                <div className="flex flex-col items-center justify-center pt-5 pb-6">
                  {isAnalyzing ? (
                    <>
                      <Loader2 className="w-10 h-10 text-blue-500 animate-spin mb-3" />
                      <p className="text-sm text-blue-600 font-medium">AI กำลังวิเคราะห์ข้อมูลจากภาพตาราง...</p>
                    </>
                  ) : (
                    <>
                      <Upload className="w-10 h-10 text-slate-400 mb-3" />
                      <p className="mb-2 text-sm text-slate-600"><span className="font-semibold">คลิกเพื่ออัพโหลด</span> หรือลากไฟล์รูปภาพนัดหมายมาวางที่นี่</p>
                      <p className="text-xs text-slate-500">รองรับไฟล์ JPG, PNG (จะประมวลผลเป็น Booking ที่ทำงานโดยอัตโนมัติ)</p>
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
                <Calendar className="w-5 h-5 text-blue-500" /> ปฏิทินนัดหมาย
              </h2>
              <div className="flex items-center gap-4">
                <button onClick={prevMonth} className="p-2 hover:bg-slate-100 rounded-full transition-colors"><ChevronLeft className="w-5 h-5" /></button>
                <span className="font-medium min-w-[120px] text-center">
                  {currentMonth.toLocaleString('th-TH-u-ca-gregory', { month: 'long', year: 'numeric' })}
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
                const dayEvents = getEventsForDate(dateStr);
                const hasEvents = dayEvents.length > 0;
                const isSelected = selectedDate === dateStr;
                const isToday = formatDateString(new Date().getFullYear(), new Date().getMonth(), new Date().getDate()) === dateStr;

                const hasWork = dayEvents.some(e => e.bookingType !== 'personal' && !e.isMsEvent);
                const hasPersonal = dayEvents.some(e => e.bookingType === 'personal' && !e.isMsEvent);
                const hasMsCalendarEvent = dayEvents.some(e => e.isMsEvent);

                let dayBgClass = '';
                if (isSelected) {
                  dayBgClass = 'ring-2 ring-slate-800 ring-offset-2 font-bold z-10';
                }
                
                // Color coding for day background based on event types present
                let itemBg = '';
                if (hasWork && hasPersonal) {
                  itemBg = 'bg-purple-50 hover:bg-purple-100 border border-purple-100 text-purple-900';
                } else if (hasWork) {
                  itemBg = 'bg-blue-50 hover:bg-blue-100 border border-blue-100 text-blue-900';
                } else if (hasPersonal) {
                  itemBg = 'bg-orange-50 hover:bg-orange-100 border border-orange-100 text-orange-900';
                } else if (hasMsCalendarEvent) {
                  // Microsoft Calendar events color code (light sky blue)
                  itemBg = 'bg-sky-50 hover:bg-sky-100 border border-sky-100 text-sky-900';
                } else if (isToday) {
                  itemBg = 'bg-slate-100 font-bold text-slate-900';
                } else {
                  itemBg = 'hover:bg-slate-50 text-slate-700';
                }

                return (
                  <button
                    key={day}
                    onClick={() => setSelectedDate(dateStr)}
                    className={`
                      relative h-14 rounded-xl flex flex-col items-center justify-center text-sm transition-all
                      ${dayBgClass} ${itemBg}
                    `}
                  >
                    <span>{day}</span>
                    <div className="flex gap-0.5 mt-1">
                      {hasWork && (
                        <span className="block w-1.5 h-1.5 rounded-full bg-blue-500"></span>
                      )}
                      {hasPersonal && (
                        <span className="block w-1.5 h-1.5 rounded-full bg-orange-500"></span>
                      )}
                      {hasMsCalendarEvent && (
                        <span className="block w-1.5 h-1.5 rounded-full bg-sky-500"></span>
                      )}
                    </div>
                  </button>
                );
              })}
            </div>
            
            {/* Color Legend */}
            <div className="flex flex-wrap gap-4 mt-6 pt-4 border-t border-slate-100 text-xs text-slate-500 justify-center">
              <div className="flex items-center gap-1.5">
                <span className="w-3 h-3 bg-blue-50 border border-blue-200 rounded-md block"></span>
                <span>💼 ที่ทำงาน (Work)</span>
              </div>
              <div className="flex items-center gap-1.5">
                <span className="w-3 h-3 bg-orange-50 border border-orange-200 rounded-md block"></span>
                <span>🏠 ส่วนตัว (Personal)</span>
              </div>
              <div className="flex items-center gap-1.5">
                <span className="w-3 h-3 bg-sky-50 border border-sky-200 rounded-md block"></span>
                <span>🌐 Microsoft Calendar</span>
              </div>
              <div className="flex items-center gap-1.5">
                <span className="w-3 h-3 bg-purple-50 border border-purple-200 rounded-md block"></span>
                <span>✨ มีทั้งคู่ (Both)</span>
              </div>
            </div>
          </section>

        </div>

        {/* Right Column: Event Details */}
        <div className="space-y-6">
          <section className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100 h-full min-h-[500px]">
            <div className="flex items-center justify-between mb-6 border-b border-slate-100 pb-4">
              <h2 className="text-lg font-semibold flex items-center gap-2">
                 <FileText className="w-5 h-5 text-blue-500" /> 
                 รายละเอียดนัดหมาย
              </h2>
              {msAccount && (
                <button 
                  onClick={fetchMicrosoftEvents}
                  className="text-xs text-blue-600 hover:text-blue-800 font-bold flex items-center gap-1 bg-blue-50 px-2 py-1 rounded border border-blue-100"
                  title="รีเฟรช MS Calendar"
                >
                  🔄 โหลด MS Calendar ใหม่
                </button>
              )}
            </div>
            
            {!selectedDate ? (
              <div className="flex flex-col items-center justify-center h-64 text-slate-400 text-center">
                <Calendar className="w-12 h-12 mb-3 opacity-20" />
                <p>เลือกวันที่บนปฏิทิน<br/>เพื่อดูรายละเอียดนัดหมาย</p>
              </div>
            ) : selectedDateEvents.length === 0 ? (
              <div className="flex flex-col items-center justify-center h-64 text-slate-400 text-center">
                <p>ไม่มีนัดหมายในวันที่<br/>{formatDisplayDate(selectedDate)}</p>
              </div>
            ) : (
              <div className="space-y-6">
                <div className="text-sm text-slate-500 mb-2 font-semibold bg-slate-50 inline-block px-3 py-1 rounded-full border border-slate-200">
                  วันที่: {formatDisplayDate(selectedDate)}
                </div>
                
                {selectedDateEvents.map((event, idx) => {
                  const isPersonal = event.bookingType === 'personal';
                  const isEditing = editingEvent === event;
                  
                  return (
                    <div key={idx} className={`p-5 rounded-xl border relative overflow-hidden transition-all duration-200 ${event.isMsEvent ? 'bg-sky-50/50 border-sky-100 text-sky-950' : isPersonal ? 'bg-orange-50/50 border-orange-100 text-orange-950' : 'bg-blue-50/50 border-blue-100 text-blue-950'}`}>
                      {/* Decorative bar */}
                      <div className={`absolute left-0 top-0 bottom-0 w-1.5 ${event.isMsEvent ? 'bg-sky-500' : isPersonal ? 'bg-orange-500' : 'bg-blue-500'}`}></div>
                      
                      {/* Booking Type Tag and Toggler */}
                      <div className="flex justify-between items-center mb-3 flex-wrap gap-2">
                        <span className={`text-xs px-2.5 py-0.5 rounded-full font-bold tracking-wide ${event.isMsEvent ? 'bg-sky-100 text-sky-800 border border-sky-200' : isPersonal ? 'bg-orange-100 text-orange-800 border border-orange-200' : 'bg-blue-100 text-blue-800 border border-blue-200'}`}>
                          {event.isMsEvent ? '🌐 Microsoft Calendar' : isPersonal ? '🏠 Booking ส่วนตัว' : '💼 Booking ที่ทำงาน'}
                        </span>
                        
                        {!event.isMsEvent && (
                          <div className="flex gap-2 text-xs">
                            <button 
                              onClick={() => {
                                setEvents(prev => prev.map((ev) => {
                                  if (ev === event) {
                                    return { ...ev, bookingType: isPersonal ? 'work' : 'personal' };
                                  }
                                  return ev;
                                }));
                              }}
                              className="text-slate-500 hover:text-slate-700 underline font-semibold transition-colors"
                            >
                              สลับประเภท
                            </button>
                            <span className="text-slate-300">|</span>
                            <button 
                              onClick={() => {
                                setEditingEvent(event);
                                setEditDate(event.date);
                                setEditTime(event.time || '');
                              }}
                              className="text-slate-500 hover:text-slate-700 underline font-semibold transition-colors"
                            >
                              แก้ไขวัน/เวลา
                            </button>
                            <span className="text-slate-300">|</span>
                            <button 
                              onClick={() => {
                                if (window.confirm(`คุณต้องการลบกิจกรรม "${event.eventName || 'ไม่มีชื่องาน'}" ใช่หรือไม่?`)) {
                                  setEvents(prev => prev.filter(ev => ev !== event));
                                  setSuccessMsg('ลบการนัดหมายเรียบร้อยแล้ว');
                                }
                              }}
                              className="text-red-500 hover:text-red-700 underline font-semibold transition-colors"
                            >
                              ลบ
                            </button>
                          </div>
                        )}
                      </div>

                      {/* Poster Image if available */}
                      {event.sourceImage && (
                        <div className="mb-4 rounded-lg overflow-hidden border border-slate-200 bg-white flex justify-center shadow-inner">
                          <img src={event.sourceImage} alt="Training Poster" className="max-h-64 w-full object-contain" />
                        </div>
                      )}

                      {/* Editing View */}
                      {isEditing ? (
                        <div className="space-y-3 bg-white p-4 rounded-lg border border-slate-200 mt-2 text-sm text-slate-800">
                          <div>
                            <label className="block text-xs font-semibold text-slate-500 mb-1">แก้ไขวันที่ *</label>
                            <input 
                              type="date" 
                              value={editDate}
                              onChange={(e) => setEditDate(e.target.value)}
                              className="w-full bg-slate-50 border border-slate-300 rounded px-2.5 py-1.5 text-xs text-slate-700 focus:outline-none focus:border-blue-500"
                            />
                            <span className="text-[10px] text-slate-400 mt-0.5 block">{formatDisplayDate(editDate)}</span>
                          </div>
                          <div>
                            <label className="block text-xs font-semibold text-slate-500 mb-1">แก้ไขเวลา</label>
                            <input 
                              type="text" 
                              value={editTime}
                              onChange={(e) => setEditTime(e.target.value)}
                              placeholder="เช่น 13:00 - 15:30 น."
                              className="w-full bg-slate-50 border border-slate-300 rounded px-2.5 py-1.5 text-xs text-slate-700 focus:outline-none focus:border-blue-500"
                            />
                          </div>
                          <div className="flex justify-end gap-2 pt-1">
                            <button 
                              onClick={() => setEditingEvent(null)}
                              className="px-3 py-1 text-xs bg-slate-100 hover:bg-slate-200 text-slate-700 rounded font-semibold"
                            >
                              ยกเลิก
                            </button>
                            <button 
                              onClick={() => {
                                if (!editDate) return;
                                setEvents(prev => prev.map(ev => {
                                  if (ev === event) {
                                    return { ...ev, date: editDate, time: editTime };
                                  }
                                  return ev;
                                }));
                                setSelectedDate(editDate);
                                setEditingEvent(null);
                                setSuccessMsg('อัปเดตวันและเวลาเรียบร้อยแล้ว');
                              }}
                              className="px-3 py-1 text-xs bg-blue-600 hover:bg-blue-700 text-white rounded font-semibold"
                            >
                              บันทึก
                            </button>
                          </div>
                        </div>
                      ) : (
                        <>
                          <h3 className="text-xl font-bold text-slate-800 mb-1">{event.eventName || 'ไม่ได้ระบุชื่องาน'}</h3>
                          
                          {event.courseName && (
                            <p className={`font-medium mb-4 flex items-center gap-2 text-sm ${event.isMsEvent ? 'text-sky-700' : isPersonal ? 'text-orange-700' : 'text-blue-700'}`}>
                              <BookOpen className="w-4 h-4" /> {event.courseName}
                            </p>
                          )}
                          
                          <div className="space-y-3 text-sm mt-2">
                            {event.topic && (
                              <div>
                                <span className="text-slate-500 block text-xs">หัวข้อ/เรื่อง/รายละเอียด</span>
                                <span className="font-semibold text-slate-700 block whitespace-pre-wrap">{event.topic}</span>
                              </div>
                            )}
                            
                            <div className="grid grid-cols-2 gap-4">
                              {event.time && (
                                <div className="flex items-start gap-2">
                                  <Clock className="w-4 h-4 text-slate-400 mt-0.5" />
                                  <div>
                                    <span className="text-slate-500 block text-xs">เวลา</span>
                                    <span className="font-semibold text-slate-700">{event.time}</span>
                                  </div>
                                </div>
                              )}
                              {event.location && (
                                <div className="flex items-start gap-2">
                                  <MapPin className="w-4 h-4 text-slate-400 mt-0.5" />
                                  <div>
                                    <span className="text-slate-500 block text-xs">สถานที่</span>
                                    <span className="font-semibold text-slate-700">{event.location}</span>
                                  </div>
                                </div>
                              )}
                            </div>

                            {/* Options specific to local events */}
                            {!event.isMsEvent && (
                              <div className="pt-4 mt-2 border-t border-slate-200 space-y-2">
                                {event.registerLink && (
                                  <a href={event.registerLink.startsWith('http') ? event.registerLink : `https://${event.registerLink}`} target="_blank" rel="noreferrer" 
                                     className="flex items-center justify-center gap-2 w-full py-2 bg-white border border-slate-300 text-slate-700 rounded-lg hover:bg-slate-50 hover:text-blue-600 transition-colors font-medium">
                                    <LinkIcon className="w-4 h-4" /> ลิ้งค์ลงทะเบียน
                                  </a>
                                )}
                                {event.joinLink && (
                                  <a href={event.joinLink.startsWith('http') ? event.joinLink : `https://${event.joinLink}`} target="_blank" rel="noreferrer"
                                     className="flex items-center justify-center gap-2 w-full py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 shadow-sm shadow-blue-200 transition-colors font-medium">
                                    <Video className="w-4 h-4" /> ลิ้งค์เข้าประชุม / ห้องจัดงาน
                                  </a>
                                )}
                              </div>
                            )}

                            <div className="pt-4 mt-2 border-t border-slate-200 space-y-2">
                              {!event.isMsEvent ? (
                                <>
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
                                </>
                              ) : (
                                <div className="p-3 bg-sky-100/50 rounded-lg border border-sky-200 text-sky-800 text-xs font-semibold flex items-center gap-1.5">
                                  <Check className="w-4 h-4 text-sky-600" /> กิจกรรมนี้อยู่ใน Microsoft Calendar แล้ว
                                </div>
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
                        </>
                      )}
                    </div>
                  );
                })}
              </div>
            )}
          </section>
        </div>
        
      </main>
    </div>
  );
}
