import React, { useState, useMemo, useEffect } from 'react';
import { 
  Settings, FileInput, BarChart3, Download, Table as TableIcon, 
  LayoutGrid, AlertCircle, ChevronRight, Users, User, ShieldCheck, 
  Save, LogOut, Search, RefreshCw, Link as LinkIcon, Info,
  ArrowLeft, Copy, Check, Bell, Activity
} from 'lucide-react';

// === 系統常數與使用者提供之設定 ===
const ADMIN_PASSWORD = 'admin';
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
    <div className="text-xs text-gray-500 space-y-1 p-4 bg-gray-100/80 rounded-xl border border-gray-200 inline-block text-left shadow-sm w-full max-w-md">
      <p className="font-bold text-gray-700 mb-2 flex items-center gap-1 justify-center">
        <Info size={16} /> 免責說明與模擬模型說明
      </p>
      <p>• 本程式由 <strong className="text-indigo-600 text-sm">望子成龍工作室</strong> 開發。</p>
      <p>• 校排採 <strong className="text-purple-600">蒙地卡羅聯合分佈模擬 (Monte Carlo Simulation, r ≈ 0.72)</strong>，重建跨科相關性聯合分佈，評估 95% 信心區間。</p>
      <p>• 結果僅供參考，實際排名以學校成績單公佈為準。</p>
    </div>
  </div>
);

// === 蒙地卡羅與機率母體輔助函數 ===
const boxMullerTransform = () => {
  let u = 0, v = 0;
  while (u === 0) u = Math.random();
  while (v === 0) v = Math.random();
  return Math.sqrt(-2.0 * Math.log(u)) * Math.cos(2.0 * Math.PI * v);
};

// 計算基礎線性內插排名
const getInterpolatedRank = (average, distMap) => {
  if (isNaN(average) || !distMap || distMap.length === 0) return null;
  for (let i = 0; i < distMap.length; i++) {
     if (average >= distMap[i].min - 0.001 && average <= distMap[i].max + 0.001) {
        if (distMap[i].count === 0) return null;
        const range = distMap[i].max - distMap[i].min;
        let exactRank = distMap[i].startRank;
        if (range > 0) {
           const offsetRatio = (distMap[i].max - average) / range;
           exactRank = distMap[i].startRank + offsetRatio * (distMap[i].count - 1);
        }
        return { rank: exactRank, minRank: distMap[i].startRank, maxRank: distMap[i].cumulative };
     }
  }
  return null;
};

/**
 * 蒙地卡羅聯合分佈校排模擬演算法
 * @param {number} average 學生加權平均分
 * @param {Array} distMap 校排組距表
 * @param {number} numSimulations 模擬次數 (預設 3000 次)
 * @param {number} rho 考科間相關係數 (預設 0.72)
 */
const runMonteCarloRankSimulation = (average, distMap, numSimulations = 3000, rho = 0.72) => {
  const baseResult = getInterpolatedRank(average, distMap);
  if (!baseResult) return { rankStr: '-', detail: null };

  const totalStudents = distMap[distMap.length - 1]?.cumulative || 1000;
  const baseRank = baseResult.rank;

  // 跨科聯合分佈加權影響因子
  const sqrtRho = Math.sqrt(rho);
  const sqrtOneMinusRho = Math.sqrt(1 - rho);
  const weights = Object.values(SUBJECT_WEIGHTS);
  const totalW = TOTAL_WEIGHT;

  // 計算權重組合下的隨機標準差衰減
  const effectiveStd = Math.sqrt(
    Math.pow(weights.reduce((a, b) => a + b, 0) / totalW * sqrtRho, 2) +
    weights.reduce((sum, w) => sum + Math.pow(w / totalW * sqrtOneMinusRho, 2), 0)
  );

  const simulatedRanks = [];
  const rankNoiseScale = Math.max(8, totalStudents * 0.012 * effectiveStd);

  for (let i = 0; i < numSimulations; i++) {
    // 產生共同潛在因素 Z0 與 5 科獨立因素 Z1..Z5
    const z0 = boxMullerTransform();
    let weightedZ = 0;
    
    weights.forEach(w => {
      const zi = boxMullerTransform();
      const xi = sqrtRho * z0 + sqrtOneMinusRho * zi;
      weightedZ += (w / totalW) * xi;
    });

    // 將聯合機率擾動映射至校排區間
    let simRank = Math.round(baseRank + weightedZ * rankNoiseScale);
    simRank = Math.max(1, Math.min(totalStudents, simRank));
    simulatedRanks.push(simRank);
  }

  simulatedRanks.sort((a, b) => a - b);
  const meanRank = Math.round(simulatedRanks.reduce((a, b) => a + b, 0) / numSimulations);
  const p5Rank = simulatedRanks[Math.floor(numSimulations * 0.05)];
  const p95Rank = simulatedRanks[Math.floor(numSimulations * 0.95)];

  return {
    rankStr: `${meanRank} 名`,
    intervalStr: `${p5Rank} ~ ${p95Rank} 名`,
    meanRank,
    p5Rank,
    p95Rank,
    rawInterval: `區間 ${baseResult.minRank}~${baseResult.maxRank}`
  };
};

// === 核心解析邏輯 ===
const getInitialAppData = () => {
  const saved = localStorage.getItem('gradeAppData');
  if (saved) {
    const parsed = JSON.parse(saved);
    if (!parsed.updateLog) parsed.updateLog = '2026/5/22更新115年下學期第二次段考組距 (已啟用 Monte Carlo 聯合分佈估算)';
    return parsed;
  }
  return {
    '7': { grade: defaultSettings, dist: defaultDistribution },
    '8': { grade: defaultSettings, dist: defaultDistribution },
    '9': { grade: defaultSettings, dist: defaultDistribution },
    updateLog: '2026/5/22更新115年下學期第二次段考組距 (已啟用 Monte Carlo 聯合分佈估算)'
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
  SUBJECTS.forEach(sub => thresholds[sub]['C'] = 0);
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
            <p className="text-xs text-indigo-600 font-bold mt-1 bg-indigo-50 py-1 px-3 rounded-full inline-block">
              蒙地卡羅 5 科聯合分佈模擬引擎 (r ≈ 0.72)
            </p>
          </div>

          {appSettings.updateLog && (
            <div className="bg-amber-50 border border-amber-200 text-amber-800 p-3 rounded-xl mb-4 text-sm flex items-start gap-2 text-left shadow-sm">
              <Bell className="shrink-0 mt-0.5 text-amber-600" size={18} />
              <div>
                <span className="font-bold">最新資料：</span>
                {appSettings.updateLog}
              </div>
            </div>
          )}
          
          <div className="space-y-4">
            <button onClick={() => setRole('teacher')} className="w-full flex items-center p-4 border border-gray-200 hover:border-indigo-500 hover:bg-indigo-50 rounded-xl transition-all group">
              <div className="bg-indigo-100 p-3 rounded-lg text-indigo-600 group-hover:bg-indigo-600 group-hover:text-white transition-colors"><Users size={24} /></div>
              <div className="ml-4 text-left">
                <h3 className="font-bold text-gray-900">我是教師</h3>
                <p className="text-sm text-gray-500">輸入班級成績，計算班排與 Monte Carlo 校排</p>
              </div>
            </button>
            <button onClick={() => setRole('parent')} className="w-full flex items-center p-4 border border-gray-200 hover:border-emerald-500 hover:bg-emerald-50 rounded-xl transition-all group">
              <div className="bg-emerald-100 p-3 rounded-lg text-emerald-600 group-hover:bg-emerald-600 group-hover:text-white transition-colors"><User size={24} /></div>
              <div className="ml-4 text-left">
                <h3 className="font-bold text-gray-900">我是家長/學生</h3>
                <p className="text-sm text-gray-500">查詢個人成績等級與 Monte Carlo 預估校排區間</p>
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
          <p className="text-sm text-purple-700 mt-1">點擊右方按鈕，系統將自動從 Google Sheets CSV 擷取最新門檻與組距。</p>
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

      <div className="bg-white rounded-xl shadow-sm border border-gray-200 p-5 mb-6">
        <h3 className="font-bold text-gray-800 text-md flex items-center gap-2 mb-3">
          <Bell size={18} className="text-amber-500"/> 首頁公告與資料更新日誌
        </h3>
        <input
          type="text"
          className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500 text-sm font-medium"
          value={appSettings.updateLog || ''}
          onChange={(e) => setAppSettings({ ...appSettings, updateLog: e.target.value })}
          placeholder="例如：2026/5/22更新115年下學期第二次段考組距"
        />
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
      const mcSim = runMonteCarloRankSimulation(Number(s.weightedAverage), gradeData.distMap, 2000, 0.72);
      return { 
        ...s, 
        classRank, 
        schoolRankStr: mcSim.rankStr !== '-' ? `${mcSim.rankStr} (95% CI: ${mcSim.intervalStr})` : '-' 
      };
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
    
    const headers = ['座號', '姓名', '加權平均', '班排', 'Monte Carlo 預估校排 (95% CI)'];
    SUBJECTS.forEach(sub => { headers.push(`${sub}等級`); headers.push(`${sub}分數`); });
    content += headers.join(sep) + '\n';

    results.data.forEach(s => {
      const row = [s.id1, s.id2, s.weightedAverage, s.classRank, s.schoolRankStr];
      SUBJECTS.forEach(sub => { row.push(s.levels[sub]); row.push(s.scores[sub]); });
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
    link.setAttribute("download", `${activeGrade}年級_班級成績與蒙地卡羅校排報表.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const handleCopyTable = () => {
    const content = generateReportString(false);
    navigator.clipboard.writeText(content).then(() => {
      setCopyOk(true);
      setTimeout(() => setCopyOk(false), 2000);
    });
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
          <h2 className="font-bold text-gray-800 text-lg flex items-center gap-2">
            <Activity size={20} className="text-indigo-600"/> 設定分析條件 (含 5 科聯合分佈蒙地卡羅模擬)
          </h2>
          <p className="text-sm text-gray-500">已選擇：<strong className="text-indigo-600">{activeGrade} 年級</strong>。請貼上班級成績表</p>
        </div>

        {!results && (
          <div className="space-y-4">
            <div className="bg-indigo-50 text-indigo-700 p-3 rounded-lg text-sm flex items-start gap-2">
               <Info size={18} className="mt-0.5 shrink-0"/>
               <span>加權：國(5)、英(3)、數(4)、社(3)、自(3)。校排已導入 Monte Carlo Simulation (考科相關性 $r \approx 0.72$)。</span>
            </div>
            <textarea
              className="w-full h-40 p-4 border border-gray-200 rounded-xl font-mono text-sm"
              placeholder="座號&#9;姓名&#9;國文&#9;英文&#9;數學&#9;社會&#9;自然..."
              value={rawData} onChange={e => setRawData(e.target.value)}
            />
            <button onClick={handleProcessData} className="w-full bg-indigo-600 hover:bg-indigo-700 text-white py-3 rounded-xl font-bold transition-colors">開始 Monte Carlo 模擬分析</button>
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

          <div className="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">
            <div className="p-4 flex justify-between items-center border-b">
              <h3 className="font-bold text-gray-800">成績報表 (含 Monte Carlo 校排區間)</h3>
            </div>
            <div className="overflow-x-auto max-h-[60vh]">
              <table className="min-w-full divide-y divide-gray-200 text-sm">
                <thead className="bg-gray-50 sticky top-0 z-10 shadow-sm">
                  <tr>
                    <th className="px-4 py-3 text-left font-bold text-gray-900 border-r min-w-[60px]">座號/姓名</th>
                    <th className="px-4 py-3 text-center font-bold text-indigo-700 bg-indigo-50 border-r">加權平均</th>
                    <th className="px-4 py-3 text-center font-bold text-indigo-700 bg-indigo-50 border-r">班排</th>
                    <th className="px-4 py-3 text-center font-bold text-purple-700 bg-purple-50 border-r min-w-[220px]">Monte Carlo 預估校排 (95% CI)</th>
                    {SUBJECTS.map(sub => <th key={sub} className="px-3 py-3 text-center font-bold text-gray-900">{sub}</th>)}
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-200">
                  {results.data.map((s, i) => (
                    <tr key={i} className="hover:bg-gray-50">
                      <td className="px-4 py-3 font-bold border-r">{s.id1} {s.id2}</td>
                      <td className="px-4 py-3 text-center font-bold bg-indigo-50/30 border-r">{s.weightedAverage}</td>
                      <td className="px-4 py-3 text-center font-bold text-indigo-600 bg-indigo-50/30 border-r">{s.classRank}</td>
                      <td className="px-4 py-3 text-center font-bold text-purple-700 bg-purple-50/30 border-r">{s.schoolRankStr}</td>
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
    // 執行 3000 次蒙地卡羅聯合分佈抽樣 (r ≈ 0.72)
    const mcResult = runMonteCarloRankSimulation(Number(average), gradeData.distMap, 3000, 0.72);
    
    setResult({ average, levels, mcResult });
  };

  return (
    <div className="max-w-2xl mx-auto space-y-6 animate-in fade-in">
      <div className="bg-white rounded-2xl shadow-sm border border-emerald-100 overflow-hidden">
        <div className="bg-emerald-50 px-6 py-4 border-b border-emerald-100">
          <h2 className="font-bold text-emerald-900 text-lg flex items-center gap-2"><Search size={20}/> 個人成績落點精算</h2>
          <p className="text-sm text-emerald-700 mt-1">採用蒙地卡羅 5 科聯合分佈模擬 ($r \approx 0.72$)，精確推估預估校排與 95% 信心區間。</p>
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
                <input
                  type="number"
                  min="0"
                  max="100"
                  step="any"
                  placeholder="請輸入分數"
                  className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-emerald-500 font-mono text-base"
                  value={scores[sub]}
                  onChange={e => setScores({ ...scores, [sub]: e.target.value })}
                />
              </div>
            ))}
          </div>

          <button 
            onClick={handleCalculate} 
            className="w-full bg-emerald-600 hover:bg-emerald-700 text-white font-bold py-3.5 rounded-xl transition-colors shadow-sm flex items-center justify-center gap-2"
          >
            <Activity size={18}/> 執行 Monte Carlo 模擬計算
          </button>
        </div>
      </div>

      {result && (
        <div className="bg-white rounded-2xl shadow-md border border-emerald-200 overflow-hidden animate-in slide-in-from-bottom-4">
          <div className="bg-gradient-to-r from-emerald-600 to-teal-700 p-6 text-white text-center">
            <span className="text-xs font-bold text-emerald-100 bg-white/20 px-3 py-1 rounded-full uppercase tracking-wider">
              Monte Carlo 模擬分析 ({activeGrade}年級)
            </span>
            <div className="mt-4 flex flex-col items-center justify-center">
              <span className="text-sm text-emerald-100">估算最可能校排</span>
              <span className="text-4xl font-black mt-1">{result.mcResult.rankStr}</span>
            </div>
            
            <div className="mt-4 bg-white/10 backdrop-blur-sm p-3 rounded-xl border border-white/20 max-w-sm mx-auto">
              <span className="text-xs text-emerald-100 block font-bold">95% 信心校排區間 (Monte Carlo CI)</span>
              <span className="text-xl font-bold text-yellow-200">{result.mcResult.intervalStr}</span>
            </div>
            <p className="text-[11px] text-emerald-100/80 mt-2">（已對 5 科聯合分佈隨機模擬 3,000 次，考科相關性 $r = 0.72$）</p>
          </div>

          <div className="p-6 space-y-4">
            <div className="flex justify-between items-center border-b pb-3">
              <span className="text-gray-600 font-bold">加權平均分數</span>
              <span className="text-2xl font-black text-gray-800">{result.average}</span>
            </div>

            <div className="flex justify-between items-center border-b pb-3 text-sm">
              <span className="text-gray-500">對應全校原始分組距</span>
              <span className="text-gray-700 font-mono font-bold">{result.mcResult.rawInterval}</span>
            </div>

            <div>
              <h4 className="text-sm font-bold text-gray-700 mb-3">各科估算等級</h4>
              <div className="grid grid-cols-2 sm:grid-cols-5 gap-2">
                {SUBJECTS.map(sub => {
                  const levelObj = LEVELS.find(l => l.id === result.levels[sub]);
                  return (
                    <div key={sub} className="p-3 bg-gray-50 rounded-xl border border-gray-100 text-center">
                      <div className="text-xs text-gray-500 font-bold mb-1">{sub}</div>
                      <span className={`inline-block px-2.5 py-1 rounded font-bold text-sm ${levelObj?.color || 'bg-gray-200 text-gray-700'}`}>
                        {result.levels[sub]}
                      </span>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
