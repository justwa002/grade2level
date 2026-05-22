import React, { useState, useMemo, useEffect } from 'react';
import { 
  Settings, FileInput, BarChart3, Download, Table as TableIcon, 
  LayoutGrid, AlertCircle, ChevronRight, Users, User, ShieldCheck, 
  Save, LogOut, Search, RefreshCw, Link as LinkIcon, Info,
  ArrowLeft, Copy, Check
} from 'lucide-react';

// === 系統常數與使用者提供之設定 ===
const ADMIN_PASSWORD = 'admin'; // 管理員預設密碼可更改
const GRADES = ['7', '8', '9'];
const SUBJECTS = ['國文', '英文', '數學', '社會', '自然'];
const LEVELS = [
  { id: 'A++', color: 'bg-emerald-100 text-emerald-800' },
  { id: 'A+', color: 'bg-green-100 text-green-800' },
  { id: 'A', color: 'bg-lime-100 text-lime-800' },
  { id: 'B++', color: 'bg-blue-100 text-blue-800' },
  { id: 'B+', color: 'bg-indigo-100 text-indigo-800' },
  { id: 'B', color: 'bg-purple-100 text-purple-800' },
  { id: 'C', color: 'bg-red-100 text-red-800' }
];

const CLOUD_URLS = {
  '7': { grade: "https://docs.google.com/spreadsheets/d/e/2PACX-1vR9LhxgNWTLkGftNnMkHQTR449Y_7M0NDr_IR_Oi5lTYZvCF9s01onsLaBWrxuA69DPntEwv0hFNU72/pub?gid=2077033678&single=true&output=csv", dist: "https://docs.google.com/spreadsheets/d/e/2PACX-1vR9LhxgNWTLkGftNnMkHQTR449Y_7M0NDr_IR_Oi5lTYZvCF9s01onsLaBWrxuA69DPntEwv0hFNU72/pub?gid=859457249&single=true&output=csv" },
  '8': { grade: "https://docs.google.com/spreadsheets/d/e/2PACX-1vR9LhxgNWTLkGftNnMkHQTR449Y_7M0NDr_IR_Oi5lTYZvCF9s01onsLaBWrxuA69DPntEwv0hFNU72/pub?gid=0&single=true&output=csv", dist: "https://docs.google.com/spreadsheets/d/e/2PACX-1vR9LhxgNWTLkGftNnMkHQTR449Y_7M0NDr_IR_Oi5lTYZvCF9s01onsLaBWrxuA69DPntEwv0hFNU72/pub?gid=853170505&single=true&output=csv" },
  '9': { grade: "https://docs.google.com/spreadsheets/d/e/2PACX-1vR9LhxgNWTLkGftNnMkHQTR449Y_7M0NDr_IR_Oi5lTYZvCF9s01onsLaBWrxuA69DPntEwv0hFNU72/pub?gid=1634530372&single=true&output=csv", dist: "https://docs.google.com/spreadsheets/d/e/2PACX-1vR9LhxgNWTLkGftNnMkHQTR449Y_7M0NDr_IR_Oi5lTYZvCF9s01onsLaBWrxuA69DPntEwv0hFNU72/pub?gid=1683092563&single=true&output=csv" }
};

const SUBJECT_WEIGHTS = { '國文': 5, '英文': 3, '數學': 4, '社會': 3, '自然': 3 };
const TOTAL_WEIGHT = Object.values(SUBJECT_WEIGHTS).reduce((a, b) => a + b, 0);

const defaultSettings = `,等級,國文,英文,數學,社會,自然\n,A++,92,100,94,94,96\n,A+,89,98,89,88,92\n,A,84,95,79,80,86\n,B++,80,92,70,72,78\n,B+,74,88,62,64,68\n,B,52,50,28,34,36`;
const defaultDistribution = `分數組距,全校人數,累計人數\n100,0,0\n98-99.99,0,0\n96-97.99,23,23\n94-95.99,52,75\n92-93.99,76,151\n90-91.99,73,224\n87-90.99,102,326\n84-86.99,75,401\n80-83.99,117,518\n70-79.99,174,692\n60-69.99,127,819\n0-59.99,25,844`;

// === 免責說明元件 ===
const Disclaimer = () => (
  <div className="mt-8 px-4 pb-8 text-center animate-in fade-in">
    <div className="text-xs text-gray-500 space-y-1 p-4 bg-gray-100/80 rounded-xl border border-gray-200 inline-block text-left shadow-sm w-full max-w-sm">
      <p className="font-bold text-gray-700 mb-2 flex items-center gap-1 justify-center">
        <Info size={16} /> 免責說明
      </p>
      <p>• 本程式由 <strong className="text-indigo-600 text-sm">望子成龍工作室</strong> 開發。</p>
      <p>• 成績分數組距以學校成績單為主，預估排名結果僅供參考，不作為實際成績依據。</p>
    </div>
  </div>
);

// === 核心解析邏輯 ===
const getInitialAppData = () => {
  const saved = localStorage.getItem('gradeAppData');
  if (saved) return JSON.parse(saved);
  return {
    '7': { grade: defaultSettings, dist: defaultDistribution },
    '8': { grade: defaultSettings, dist: defaultDistribution },
    '9': { grade: defaultSettings, dist: defaultDistribution }
  };
};

const parseThresholds = (csv) => {
  const lines = csv.trim().split('\n').map(l => l.split(','));
  const headers = lines[0].map(h => h.trim());
  const thresholds = { 國文:{}, 英文:{}, 數學:{}, 社會:{}, 自然:{} };
  const levelIdx = headers.indexOf('等級');
  
  if (levelIdx === -1) return thresholds;

  for(let i = 1; i < lines.length; i++) {
    const row = lines[i];
    const level = row[levelIdx]?.trim();
    if(!level) continue;
    SUBJECTS.forEach(sub => {
      const idx = headers.indexOf(sub);
      if(idx !== -1 && row[idx]) thresholds[sub][level] = parseFloat(row[idx]);
    });
  }
  SUBJECTS.forEach(sub => thresholds[sub]['C'] = 0); // 預設 C
  return thresholds;
};

const processDistribution = (csv) => {
  const lines = csv.trim().split('\n').map(l => l.split(','));
  const headers = lines[0].map(h => h.trim());
  
  const distData = [];
  for(let i = 1; i < lines.length; i++) {
    const obj = {};
    headers.forEach((h, j) => { obj[h] = lines[i][j]; });
    distData.push(obj);
  }

  let previousCumulative = 0;
  return distData.map(row => {
    const rangeStr = row['分數組距'];
    if (!rangeStr) return null;
    let min, max;
    if (rangeStr.includes('-')) {
      const parts = rangeStr.split('-');
      min = parseFloat(parts[0]); max = parseFloat(parts[1]);
    } else {
      min = parseFloat(rangeStr); max = parseFloat(rangeStr);
    }
    const count = parseInt(row['全校人數'] || '0', 10);
    const cumulative = parseInt(row['累計人數'] || '0', 10);
    const result = { min, max, count, cumulative, startRank: previousCumulative + 1 };
    previousCumulative = cumulative;
    return result;
  }).filter(Boolean);
};

const calculateLevel = (score, subjectThresholds) => {
  if (score === null || isNaN(score) || score === '') return '-';
  for (const level of LEVELS) {
    if (Number(score) >= subjectThresholds[level.id]) return level.id;
  }
  return 'C';
};

const getSchoolRank = (average, distMap) => {
  if (isNaN(average) || !distMap || distMap.length === 0) return '-';
  for (let i = 0; i < distMap.length; i++) {
     if (average >= distMap[i].min - 0.001 && average <= distMap[i].max + 0.001) {
        if (distMap[i].count === 0) return '-';
        const range = distMap[i].max - distMap[i].min;
        let exactRank = distMap[i].startRank;
        if (range > 0) {
           const offsetRatio = (distMap[i].max - average) / range;
           exactRank = Math.round(distMap[i].startRank + offsetRatio * (distMap[i].count - 1));
        }
        return `${exactRank} (區間 ${distMap[i].startRank}~${distMap[i].cumulative})`;
     }
  }
  return '-';
};


// ==========================================
// 主應用程式元件
// ==========================================
export default function App() {
  const [role, setRole] = useState('');
  const [appSettings, setAppSettings] = useState(getInitialAppData);

  useEffect(() => {
    localStorage.setItem('gradeAppData', JSON.stringify(appSettings));
  }, [appSettings]);

  const parsedData = useMemo(() => {
    const data = {};
    GRADES.forEach(g => {
      data[g] = {
        thresholds: parseThresholds(appSettings[g].grade),
        distMap: processDistribution(appSettings[g].dist)
      };
    });
    return data;
  }, [appSettings]);

  if (!role) {
    return (
      <div className="min-h-screen bg-gray-50 flex flex-col items-center justify-center p-4">
        <div className="max-w-md w-full bg-white rounded-2xl shadow-xl p-8 space-y-8 animate-in zoom-in-95 duration-300">
          <div className="text-center">
            <div className="bg-indigo-600 text-white p-3 rounded-2xl inline-block mb-4 shadow-md">
              <BarChart3 size={32} />
            </div>
            <h1 className="text-2xl font-bold text-gray-900">成績落點與校排精算系統</h1>
            <p className="text-sm text-gray-500 mt-2">支援加權平均與組距內插法排名</p>
          </div>
          
          <div className="space-y-4">
            <button onClick={() => setRole('teacher')} className="w-full flex items-center p-4 border border-gray-200 hover:border-indigo-500 hover:bg-indigo-50 rounded-xl transition-all group">
              <div className="bg-indigo-100 p-3 rounded-lg text-indigo-600 group-hover:bg-indigo-600 group-hover:text-white transition-colors"><Users size={24} /></div>
              <div className="ml-4 text-left">
                <h3 className="font-bold text-gray-900">我是教師</h3>
                <p className="text-sm text-gray-500">輸入班級成績，計算班排與精準校排</p>
              </div>
            </button>
            <button onClick={() => setRole('parent')} className="w-full flex items-center p-4 border border-gray-200 hover:border-emerald-500 hover:bg-emerald-50 rounded-xl transition-all group">
              <div className="bg-emerald-100 p-3 rounded-lg text-emerald-600 group-hover:bg-emerald-600 group-hover:text-white transition-colors"><User size={24} /></div>
              <div className="ml-4 text-left">
                <h3 className="font-bold text-gray-900">我是家長/學生</h3>
                <p className="text-sm text-gray-500">查詢個人成績等級與預估校排</p>
              </div>
            </button>
            <button onClick={() => {
              const pwd = prompt("請輸入管理員密碼：");
              if (pwd === ADMIN_PASSWORD) setRole('admin');
              else if (pwd !== null) alert("密碼錯誤");
            }} className="w-full flex items-center p-4 border border-gray-200 hover:border-purple-500 hover:bg-purple-50 rounded-xl transition-all group">
              <div className="bg-purple-100 p-3 rounded-lg text-purple-600 group-hover:bg-purple-600 group-hover:text-white transition-colors"><ShieldCheck size={24} /></div>
              <div className="ml-4 text-left">
                <h3 className="font-bold text-gray-900">管理員中心</h3>
                <p className="text-sm text-gray-500">一鍵同步 Google Sheets 組距資料</p>
              </div>
            </button>
          </div>
        </div>
        <Disclaimer />
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-gray-50 text-gray-800 font-sans flex flex-col">
      <header className="bg-white shadow-sm sticky top-0 z-30">
        <div className="max-w-6xl mx-auto px-4 h-16 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <div className={`p-2 rounded-lg text-white ${role === 'admin' ? 'bg-purple-600' : role === 'teacher' ? 'bg-indigo-600' : 'bg-emerald-600'}`}>
              {role === 'admin' ? <ShieldCheck size={20} /> : role === 'teacher' ? <Users size={20} /> : <User size={20} />}
            </div>
            <h1 className="font-bold text-lg text-gray-900 hidden sm:block">
              {role === 'admin' ? '系統設定與同步中心' : role === 'teacher' ? '班級成績分析工具' : '個人落點精算系統'}
            </h1>
          </div>
          <button onClick={() => setRole('')} className="flex items-center gap-1.5 px-3 py-1.5 rounded-md text-sm font-bold text-gray-600 hover:text-gray-900 bg-gray-100 hover:bg-gray-200 transition-colors">
            <ArrowLeft size={16} /> 返回首頁
          </button>
        </div>
      </header>
      <main className="max-w-6xl mx-auto px-4 py-6 flex-1 w-full">
        {role === 'admin' && <AdminView appSettings={appSettings} setAppSettings={setAppSettings} parsedData={parsedData} />}
        {role === 'teacher' && <TeacherView parsedData={parsedData} />}
        {role === 'parent' && <ParentView parsedData={parsedData} />}
      </main>
      {(role === 'parent' || role === 'teacher' || role === 'admin') && <Disclaimer />}
    </div>
  );
}


// ==========================================
// Admin View Component
// ==========================================
function AdminView({ appSettings, setAppSettings, parsedData }) {
  const [isLoading, setIsLoading] = useState(false);
  const [msg, setMsg] = useState('');
  const [activeGrade, setActiveGrade] = useState('7');

  const handleSyncAllGrades = async () => {
    setIsLoading(true);
    setMsg('🔄 正在批次同步所有年級(7,8,9)雲端資料，請稍候...');
    const newSettings = { ...appSettings };
    let successCount = 0;
    
    try {
      for (const grade of ['7', '8', '9']) {
        const gradeUrl = CLOUD_URLS[grade]?.grade;
        const distUrl = CLOUD_URLS[grade]?.dist;
        
        if (gradeUrl) {
          const res = await fetch(gradeUrl);
          if (res.ok) { newSettings[grade].grade = await res.text(); successCount++; }
        }
        if (distUrl) {
          const res = await fetch(distUrl);
          if (res.ok) { newSettings[grade].dist = await res.text(); successCount++; }
        }
      }
      setAppSettings(newSettings);
      if(successCount > 0) setMsg(`✅ 同步完成！成功獲取 ${successCount} 份檔案。`);
      else setMsg('⚠️ 找不到有效的雲端連結，同步失敗。');
    } catch (err) {
      setMsg(`❌ 同步發生錯誤: ${err.message}`);
    } finally {
      setIsLoading(false);
    }
  };

  const distCount = parsedData[activeGrade].distMap.reduce((sum, d) => sum + d.count, 0);

  return (
    <div className="space-y-6 animate-in fade-in">
      <div className="bg-purple-50 border border-purple-200 rounded-xl p-5 flex flex-col md:flex-row justify-between items-center gap-4 sticky top-16 z-20 shadow-sm">
        <div>
          <h2 className="font-bold text-purple-900 text-lg flex items-center gap-2"><LinkIcon size={20}/> 雲端發佈同步中心</h2>
          <p className="text-sm text-purple-700 mt-1">點擊右方按鈕，系統將自動從您設定的 Google Sheets CSV 連結抓取 7-9 年級的門檻與組距。</p>
        </div>
        <div className="flex flex-col items-end gap-2">
          <button 
            onClick={handleSyncAllGrades} disabled={isLoading}
            className="flex items-center gap-2 bg-purple-600 hover:bg-purple-700 text-white px-6 py-2.5 rounded-lg font-bold transition-colors shadow-sm disabled:opacity-70"
          >
            <RefreshCw size={18} className={isLoading ? 'animate-spin' : ''} /> {isLoading ? '資料擷取中...' : '一鍵同步所有年級資料'}
          </button>
          {msg && <span className="text-sm font-bold text-purple-800 bg-purple-100 px-2 py-1 rounded">{msg}</span>}
        </div>
      </div>

      <div className="mb-6">
        <label className="block text-sm font-bold text-gray-700 mb-3 text-center">👇 選擇設定年級 👇</label>
        <div className="grid grid-cols-3 gap-3 sm:gap-6">
          {GRADES.map(g => (
            <button 
              key={g} 
              onClick={() => setActiveGrade(g)}
              className={`py-3 rounded-xl font-black text-lg transition-all duration-200 border-2 ${
                activeGrade === g 
                  ? 'bg-purple-600 text-white border-purple-600 shadow-lg transform scale-105' 
                  : 'bg-white text-gray-400 border-gray-200 hover:border-purple-300 hover:text-purple-500'
              }`}
            >
              {g} 年級
            </button>
          ))}
        </div>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
        <div className="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">
           <div className="bg-gray-50 p-3 border-b font-bold text-gray-700">各科等級門檻 (解析預覽)</div>
           <div className="p-4 space-y-4">
             {SUBJECTS.map(sub => (
                <div key={sub} className="flex flex-wrap items-center gap-2">
                  <span className="font-bold w-12 text-gray-600">{sub}</span>
                  {LEVELS.filter(l=>l.id!=='C').map(lvl => (
                    <span key={lvl.id} className={`text-xs px-2 py-1 rounded ${lvl.color}`}>
                      {lvl.id} ≥ {parsedData[activeGrade].thresholds[sub]?.[lvl.id] || '-'}
                    </span>
                  ))}
                </div>
             ))}
           </div>
        </div>

        <div className="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">
           <div className="bg-gray-50 p-3 border-b font-bold text-gray-700 flex justify-between">
              <span>全校分數組距 (解析預覽)</span>
              <span className="text-purple-600 text-sm">樣本數: {distCount} 人</span>
           </div>
           <div className="p-4 h-64 overflow-y-auto">
             <table className="w-full text-sm text-left">
                <thead className="bg-gray-100 sticky top-0">
                   <tr><th className="p-2">組距</th><th className="p-2">人數</th><th className="p-2">累計</th><th className="p-2">排名起點</th></tr>
                </thead>
                <tbody>
                   {parsedData[activeGrade].distMap.map((d, i) => (
                     <tr key={i} className="border-b border-gray-50 hover:bg-gray-50">
                        <td className="p-2 font-mono">{d.min} - {d.max}</td>
                        <td className="p-2">{d.count}</td><td className="p-2">{d.cumulative}</td><td className="p-2 text-purple-600">{d.startRank}</td>
                     </tr>
                   ))}
                </tbody>
             </table>
           </div>
        </div>
      </div>
    </div>
  );
}

// ==========================================
// Teacher View Component
// ==========================================
function TeacherView({ parsedData }) {
  const [activeGrade, setActiveGrade] = useState('7');
  const [rawData, setRawData] = useState('');
  const [results, setResults] = useState(null);
  const [copyOk, setCopyOk] = useState(false);

  const handleProcessData = () => {
    if (!rawData.trim()) return;
    const lines = rawData.trim().split('\n');
    const headers = lines[0].split('\t').map(h => h.trim());
    
    const subjectIndices = {};
    SUBJECTS.forEach(sub => {
      const idx = headers.findIndex(h => h.includes(sub));
      if (idx !== -1) subjectIndices[sub] = idx;
    });

    const gradeData = parsedData[activeGrade];
    const students = [];

    lines.slice(1).forEach(line => {
      const values = line.split('\t').map(v => v.trim());
      if (values.length < 2) return;
      
      const student = { id1: values[0], id2: values[1], scores: {}, levels: {} };
      let weightedSum = 0;

      SUBJECTS.forEach(sub => {
        const idx = subjectIndices[sub];
        if (idx !== undefined && values[idx] && !isNaN(values[idx])) {
          const score = Number(values[idx]);
          student.scores[sub] = score;
          student.levels[sub] = calculateLevel(score, gradeData.thresholds[sub]);
          weightedSum += (score * SUBJECT_WEIGHTS[sub]);
        } else {
          student.scores[sub] = 0; student.levels[sub] = '-';
        }
      });
      student.weightedAverage = (weightedSum / TOTAL_WEIGHT).toFixed(2);
      students.push(student);
    });

    const sortedByAvg = [...students].sort((a, b) => b.weightedAverage - a.weightedAverage);
    const finalData = students.map(s => {
      const classRank = sortedByAvg.findIndex(sorted => sorted.weightedAverage <= s.weightedAverage) + 1;
      const schoolRankStr = getSchoolRank(Number(s.weightedAverage), gradeData.distMap);
      return { ...s, classRank, schoolRankStr };
    });

    const stats = {};
    SUBJECTS.forEach(sub => {
      stats[sub] = {};
      LEVELS.forEach(l => stats[sub][l.id] = 0);
      finalData.forEach(s => {
        if (s.levels[sub] !== '-') stats[sub][s.levels[sub]]++;
      });
    });

    setResults({ data: finalData, stats });
  };

  const generateReportString = (isCsv = false) => {
    const sep = isCsv ? ',' : '\t';
    let content = "";
    
    const headers = ['座號', '姓名', '加權平均', '班排', '預估校排'];
    SUBJECTS.forEach(sub => { headers.push(`${sub}等級`); headers.push(`${sub}分數`); });
    content += headers.join(sep) + '\n';

    results.data.forEach(s => {
      const row = [s.id1, s.id2, s.weightedAverage, s.classRank, s.schoolRankStr];
      SUBJECTS.forEach(sub => { row.push(s.levels[sub]); row.push(s.scores[sub]); });
      content += row.join(sep) + '\n';
    });

    content += '\n';
    content += `[各科等級統計]\n`;
    const statHeaders = ['科目', ...LEVELS.map(l => l.id)];
    content += statHeaders.join(sep) + '\n';

    SUBJECTS.forEach(sub => {
      const row = [sub];
      LEVELS.forEach(l => {
        row.push(results.stats[sub][l.id]);
      });
      content += row.join(sep) + '\n';
    });

    return content;
  };

  const handleExportCSV = () => {
    const content = "\uFEFF" + generateReportString(true);
    const blob = new Blob([content], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", `${activeGrade}年級_班級成績與統計報表.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const handleCopyTable = () => {
    const content = generateReportString(false);
    const fallbackCopy = (text) => {
      const textArea = document.createElement("textarea");
      textArea.value = text;
      textArea.style.position = "fixed";
      document.body.appendChild(textArea);
      textArea.focus();
      textArea.select();
      try {
        document.execCommand('copy');
        setCopyOk(true);
        setTimeout(() => setCopyOk(false), 2000);
      } catch (err) {
        console.error('Copy failed', err);
      }
      document.body.removeChild(textArea);
    };

    if (navigator.clipboard && window.isSecureContext) {
      navigator.clipboard.writeText(content).then(() => {
        setCopyOk(true);
        setTimeout(() => setCopyOk(false), 2000);
      }).catch(() => fallbackCopy(content));
    } else {
      fallbackCopy(content);
    }
  };

  return (
    <div className="space-y-6 animate-in fade-in">
      <div className="mb-2">
        <label className="block text-sm font-bold text-gray-700 mb-3 text-center">👇 請先選擇要分析的年級 👇</label>
        <div className="grid grid-cols-3 gap-3 sm:gap-6">
          {GRADES.map(g => (
            <button 
              key={g} 
              onClick={() => { setActiveGrade(g); setResults(null); }}
              className={`py-4 rounded-xl font-black text-xl transition-all duration-200 border-2 ${
                activeGrade === g 
                  ? 'bg-indigo-600 text-white border-indigo-600 shadow-lg transform scale-105' 
                  : 'bg-white text-gray-400 border-gray-200 hover:border-indigo-300 hover:text-indigo-500'
              }`}
            >
              {g} 年級
            </button>
          ))}
        </div>
      </div>

      <div className="bg-white p-5 rounded-xl shadow-sm border border-gray-200">
        <div className="mb-4">
          <h2 className="font-bold text-gray-800 text-lg">設定分析條件</h2>
          <p className="text-sm text-gray-500">已選擇：<strong className="text-indigo-600">{activeGrade} 年級</strong>。請貼上班級成績表 (包含: 座號, 姓名, 國文, 英文, 數學, 社會, 自然)</p>
        </div>

        {!results && (
          <div className="space-y-4">
            <div className="bg-indigo-50 text-indigo-700 p-3 rounded-lg text-sm flex items-start gap-2">
               <Info size={18} className="mt-0.5 shrink-0"/>
               <span>加權計分比例：國文(5)、英文(3)、數學(4)、社會(3)、自然(3)。<br/>格式需求：座號、姓名，以及各科欄位(Tab分隔，例如從Excel複製)。</span>
            </div>
            <textarea
              className="w-full h-40 p-4 border border-gray-200 rounded-xl font-mono text-sm"
              placeholder="座號&#9;姓名&#9;國文&#9;英文&#9;數學&#9;社會&#9;自然..."
              value={rawData} onChange={e => setRawData(e.target.value)}
            />
            <button onClick={handleProcessData} className="w-full bg-indigo-600 hover:bg-indigo-700 text-white py-3 rounded-xl font-bold transition-colors">開始分析班級成績</button>
          </div>
        )}
      </div>

      {results && (
        <div className="space-y-6 animate-in slide-in-from-bottom-4">
          <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
            <h2 className="font-bold text-xl text-gray-800">分析結果 ({activeGrade}年級)</h2>
            <div className="flex flex-wrap items-center gap-2">
              <button onClick={handleCopyTable} className="flex items-center gap-2 bg-white border border-gray-300 hover:bg-gray-50 text-gray-700 px-4 py-2 rounded-lg font-bold text-sm transition-colors shadow-sm">
                {copyOk ? <Check size={16} className="text-emerald-500" /> : <Copy size={16} />} 
                {copyOk ? '已複製！' : '複製報表 (含統計)'}
              </button>
              <button onClick={handleExportCSV} className="flex items-center gap-2 bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-2 rounded-lg font-bold text-sm transition-colors shadow-sm">
                <Download size={16} /> 匯出 CSV 報表
              </button>
              <button onClick={() => setResults(null)} className="text-indigo-600 text-sm font-bold hover:underline ml-2">重新輸入</button>
            </div>
          </div>

          <div className="bg-white rounded-xl shadow-sm border border-indigo-100 overflow-hidden">
            <div className="bg-indigo-50 px-4 py-3 border-b border-indigo-100">
              <h3 className="font-bold text-indigo-900 flex items-center gap-2"><BarChart3 size={18}/> 各科等級人數總計</h3>
            </div>
            <div className="p-4 overflow-x-auto">
              <table className="w-full text-sm text-center min-w-[600px]">
                <thead>
                  <tr className="border-b-2 border-gray-200">
                    <th className="py-2 text-left text-gray-600 font-bold w-24">科目</th>
                    {LEVELS.map(l => <th key={l.id} className="py-2 font-bold"><span className={`px-2 py-1 rounded-md ${l.color}`}>{l.id}</span></th>)}
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-100">
                  {SUBJECTS.map(sub => (
                    <tr key={sub} className="hover:bg-gray-50">
                      <td className="py-3 text-left font-bold text-gray-800">{sub}</td>
                      {LEVELS.map(l => (
                        <td key={l.id} className="py-3 font-bold text-gray-700">{results.stats[sub][l.id]}</td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          <div className="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">
            <div className="p-4 flex justify-between items-center border-b">
              <h3 className="font-bold text-gray-800">成績報表 (依座號排序)</h3>
            </div>
            <div className="overflow-x-auto max-h-[60vh]">
              <table className="min-w-full divide-y divide-gray-200 text-sm">
                <thead className="bg-gray-50 sticky top-0 z-10 shadow-sm">
                  <tr>
                    <th className="px-4 py-3 text-left font-bold text-gray-900 border-r min-w-[60px]">座號/姓名</th>
                    <th className="px-4 py-3 text-center font-bold text-indigo-700 bg-indigo-50 border-r">加權平均</th>
                    <th className="px-4 py-3 text-center font-bold text-indigo-700 bg-indigo-50 border-r">班排</th>
                    <th className="px-4 py-3 text-center font-bold text-indigo-700 bg-indigo-50 border-r min-w-[160px]">內插法校排預估</th>
                    {SUBJECTS.map(sub => <th key={sub} className="px-3 py-3 text-center font-bold text-gray-900">{sub}</th>)}
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-200">
                  {results.data.map((s, i) => (
                    <tr key={i} className="hover:bg-gray-50">
                      <td className="px-4 py-3 font-bold border-r">{s.id1} {s.id2}</td>
                      <td className="px-4 py-3 text-center font-bold bg-indigo-50/30 border-r">{s.weightedAverage}</td>
                      <td className="px-4 py-3 text-center font-bold text-indigo-600 bg-indigo-50/30 border-r">{s.classRank}</td>
                      <td className="px-4 py-3 text-center font-bold text-purple-600 bg-indigo-50/30 border-r">{s.schoolRankStr}</td>
                      {SUBJECTS.map(sub => (
                        <td key={sub} className="px-2 py-3 text-center">
                           <div className="flex flex-col items-center">
                              <span className={`px-2 rounded text-xs font-bold ${LEVELS.find(l=>l.id===s.levels[sub])?.color}`}>{s.levels[sub]}</span>
                              <span className="text-[10px] text-gray-400 mt-1">{s.scores[sub]}</span>
                           </div>
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ==========================================
// Parent View Component
// ==========================================
function ParentView({ parsedData }) {
  const [activeGrade, setActiveGrade] = useState('7');
  const [scores, setScores] = useState({ 國文: '', 英文: '', 數學: '', 社會: '', 自然: '' });
  const [result, setResult] = useState(null);

  const handleCalculate = () => {
    const gradeData = parsedData[activeGrade];
    let weightedSum = 0;
    const levels = {};
    let hasEmpty = false;

    SUBJECTS.forEach(sub => {
      if (scores[sub] === '') hasEmpty = true;
      const num = Number(scores[sub] || 0);
      weightedSum += (num * SUBJECT_WEIGHTS[sub]);
      levels[sub] = calculateLevel(num, gradeData.thresholds[sub]);
    });

    if (hasEmpty) { alert("請填寫所有科目的成績"); return; }

    const average = (weightedSum / TOTAL_WEIGHT).toFixed(2);
    const schoolRankStr = getSchoolRank(Number(average), gradeData.distMap);
    setResult({ average, levels, schoolRankStr });
  };

  return (
    <div className="max-w-2xl mx-auto space-y-6 animate-in fade-in">
      <div className="bg-white rounded-2xl shadow-sm border border-emerald-100 overflow-hidden">
        <div className="bg-emerald-50 px-6 py-4 border-b border-emerald-100">
          <h2 className="font-bold text-emerald-900 text-lg flex items-center gap-2"><Search size={20}/> 個人成績落點精算</h2>
          <p className="text-sm text-emerald-700 mt-1">透過加權平均與校排組距內插法，計算最精準的預估排名。</p>
        </div>
        <div className="p-6 space-y-6">
          <div className="mb-2">
            <label className="block text-sm font-bold text-gray-700 mb-3 text-center">👇 請先選擇就讀年級 👇</label>
            <div className="grid grid-cols-3 gap-3 sm:gap-6">
              {GRADES.map(g => (
                <button 
                  key={g} 
                  onClick={() => { setActiveGrade(g); setResult(null); }}
                  className={`py-3 rounded-xl font-black text-lg transition-all duration-200 border-2 ${
                    activeGrade === g 
                      ? 'bg-emerald-600 text-white border-emerald-600 shadow-lg transform scale-105' 
                      : 'bg-white text-gray-400 border-gray-200 hover:border-emerald-300 hover:text-emerald-500'
                  }`}
                >
                  {g} 年級
                </button>
              ))}
            </div>
          </div>
          <div className="grid grid-cols-2 sm:grid-cols-3 gap-4">
            {SUBJECTS.map(sub => (
              <div key={sub}>
                <label className="block text-sm font-bold text-gray-700 mb-1">{sub} <span className="text-xs text-gray-400">(權重 x{SUBJECT_WEIGHTS[sub]})</span></label>
                <input type="number" value={scores[sub]} onChange={e => setScores(p => ({...p, [sub]: e.target.value}))} className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-emerald-500" placeholder="分數" />
              </div>
            ))}
          </div>
          <button onClick={handleCalculate} className="w-full bg-emerald-600 hover:bg-emerald-700 text-white py-3 rounded-xl font-bold text-lg transition-colors shadow-md">計算加權平均與落點</button>
        </div>
      </div>

      {result && (
        <div className="bg-white rounded-2xl shadow-xl border border-gray-200 overflow-hidden animate-in zoom-in-95 duration-300">
          <div className="p-6 bg-gradient-to-r from-emerald-500 to-teal-500 text-white text-center">
            <h3 className="font-bold text-emerald-50 mb-2">落點分析結果</h3>
            <div className="flex justify-center items-end gap-6">
              <div>
                <div className="text-sm text-emerald-100 mb-1">加權平均分數</div>
                <div className="text-4xl font-black">{result.average}</div>
              </div>
              <div className="w-px h-12 bg-emerald-400/50"></div>
              <div>
                <div className="text-sm text-emerald-100 mb-1">精算預估校排 (內插法)</div>
                <div className="text-3xl font-black text-yellow-300">{result.schoolRankStr}</div>
              </div>
            </div>
          </div>
          <div className="p-6 space-y-3">
            {SUBJECTS.map(sub => (
              <div key={sub} className="flex items-center justify-between p-3 bg-gray-50 rounded-xl border border-gray-100">
                <span className="font-bold text-gray-700">{sub}</span>
                <div className="flex items-center gap-4">
                  <span className="text-gray-500 font-medium">{scores[sub]} 分</span>
                  <span className={`px-3 py-1 rounded-lg font-bold text-sm w-12 text-center shadow-sm ${LEVELS.find(l=>l.id===result.levels[sub])?.color}`}>{result.levels[sub]}</span>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

// 隱藏捲軸樣式注入
const style = document.createElement('style');
style.textContent = `.scrollbar-hide::-webkit-scrollbar { display: none; } .scrollbar-hide { -ms-overflow-style: none; scrollbar-width: none; }`;
document.head.appendChild(style);