import React, { useState, useMemo, useRef, useEffect } from 'react';
import { 
  Upload, FileDown, ClipboardCopy, Calculator, CheckCircle2, 
  Settings, ArrowLeft, Lock, Unlock, AlertTriangle, Users, 
  BookOpen, ChevronRight, FileSpreadsheet, Trash2, Home, Info,
  Cloud, RefreshCw
} from 'lucide-react';

// --- 雲端試算表固定連結設定 (請將此處替換為您的真實發佈連結) ---
const CLOUD_URLS = {
  '7': { grade: "https://docs.google.com/spreadsheets/d/e/2PACX-1vR9LhxgNWTLkGftNnMkHQTR449Y_7M0NDr_IR_Oi5lTYZvCF9s01onsLaBWrxuA69DPntEwv0hFNU72/pub?gid=2077033678&single=true&output=csv", dist: "https://docs.google.com/spreadsheets/d/e/2PACX-1vR9LhxgNWTLkGftNnMkHQTR449Y_7M0NDr_IR_Oi5lTYZvCF9s01onsLaBWrxuA69DPntEwv0hFNU72/pub?gid=859457249&single=true&output=csv" },
  '8': { grade: "https://docs.google.com/spreadsheets/d/e/2PACX-1vR9LhxgNWTLkGftNnMkHQTR449Y_7M0NDr_IR_Oi5lTYZvCF9s01onsLaBWrxuA69DPntEwv0hFNU72/pub?gid=0&single=true&output=csv", dist: "https://docs.google.com/spreadsheets/d/e/2PACX-1vR9LhxgNWTLkGftNnMkHQTR449Y_7M0NDr_IR_Oi5lTYZvCF9s01onsLaBWrxuA69DPntEwv0hFNU72/pub?gid=853170505&single=true&output=csv" },
  '9': { grade: "https://docs.google.com/spreadsheets/d/e/2PACX-1vR9LhxgNWTLkGftNnMkHQTR449Y_7M0NDr_IR_Oi5lTYZvCF9s01onsLaBWrxuA69DPntEwv0hFNU72/pub?gid=1634530372&single=true&output=csv", dist: "https://docs.google.com/spreadsheets/d/e/2PACX-1vR9LhxgNWTLkGftNnMkHQTR449Y_7M0NDr_IR_Oi5lTYZvCF9s01onsLaBWrxuA69DPntEwv0hFNU72/pub?gid=1683092563&single=true&output=csv" }
};

// --- 科目加權設定 ---
const SUBJECT_WEIGHTS = {
  '國文': 5,
  '英文': 3,
  '數學': 4,
  '社會': 3,
  '自然': 3
};

// --- 1. 核心計算邏輯 ---
const parseCSV = (csvText) => {
  if (!csvText) return [];
  const lines = csvText.trim().split(/\r?\n/).filter(line => line.trim() !== '');
  if (lines.length === 0) return [];
  const delimiter = lines[0].includes('\t') ? '\t' : ',';
  const headers = lines[0].split(delimiter).map(h => h.trim());
  return lines.slice(1).map(line => {
    const values = line.split(delimiter).map(v => v.trim());
    const obj = {};
    headers.forEach((header, i) => { obj[header] = values[i]; });
    return obj;
  });
};

const processSettings = (settingsData) => {
  const settings = {};
  if (!settingsData || settingsData.length === 0) return settings;
  const subjects = Object.keys(settingsData[0]).filter(k => k !== '等級' && k !== '');
  subjects.forEach(subject => {
    settings[subject] = settingsData
      .map(row => ({ level: row['等級'], minScore: parseFloat(row[subject]) }))
      .filter(item => !isNaN(item.minScore))
      .sort((a, b) => b.minScore - a.minScore);
  });
  return { settings, subjects };
};

const getGradeLevel = (score, subjectSettings) => {
  if (isNaN(score) || !subjectSettings) return '';
  for (let i = 0; i < subjectSettings.length; i++) {
    if (score >= subjectSettings[i].minScore) return subjectSettings[i].level;
  }
  return 'C'; 
};

const processDistribution = (distData) => {
  if (!distData || distData.length === 0) return [];
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

const getSchoolRank = (average, distMap) => {
  if (isNaN(average) || !distMap || distMap.length === 0) return '';
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
  return '';
};

const exportToCSV = (data, filename) => {
  if (!data || data.length === 0) return;
  const headers = Object.keys(data[0]);
  const csvContent = [
    headers.join(','),
    ...data.map(row => headers.map(header => `"${row[header] || ''}"`).join(','))
  ].join('\n');
  const blob = new Blob(['\uFEFF' + csvContent], { type: 'text/csv;charset=utf-8;' });
  const link = document.createElement('a');
  const url = URL.createObjectURL(blob);
  link.href = url; link.download = filename;
  document.body.appendChild(link); link.click(); document.body.removeChild(link); URL.revokeObjectURL(url);
};

// --- 2. 預設資料與 UI 配置 ---
const defaultSettings = `,等級,國文,英文,數學,社會,自然\n,A++,92,100,94,94,96\n,A+,89,98,89,88,92\n,A,84,95,79,80,86\n,B++,80,92,70,72,78\n,B+,74,88,62,64,68\n,B,52,50,28,34,36`;
const defaultDistribution = `分數組距,全校人數,累計人數\n100,0,0\n98-99.99,0,0\n96-97.99,23,23\n94-95.99,52,75\n92-93.99,76,151\n90-91.99,73,224\n87-90.99,102,326\n84-86.99,75,401\n80-83.99,117,518\n70-79.99,174,692\n60-69.99,127,819\n0-59.99,25,844`;

const FIXED_HEADERS = ['座號', '姓名', '國文', '英文', '數學', '社會', '自然'];
const INITIAL_GRID_ROWS = 45;

const COL_STYLES = {
  '#': { width: '45px' },
  '座號': { width: '70px' },
  '姓名': { width: '90px' },
  '國文': { width: '110px' },
  '英文': { width: '110px' },
  '數學': { width: '110px' },
  '社會': { width: '110px' },
  '自然': { width: '110px' },
};

// --- 3. 主應用程式元件 ---
export default function App() {
  const [view, setView] = useState('home'); 
  const [selectedGrade, setSelectedGrade] = useState(null); 
  const [adminPassword, setAdminPassword] = useState('');
  
  const [notification, setNotification] = useState(null); 
  const [clearConfirm, setClearConfirm] = useState(false);
  const [isLoading, setIsLoading] = useState(false);

  const [appSettings, setAppSettings] = useState({
    '7': { grade: defaultSettings, dist: defaultDistribution },
    '8': { grade: defaultSettings, dist: defaultDistribution },
    '9': { grade: defaultSettings, dist: defaultDistribution }
  });

  const [gridData, setGridData] = useState(() => {
    const grid = Array(INITIAL_GRID_ROWS).fill(0).map(() => Array(FIXED_HEADERS.length).fill(''));
    grid[0] = [...FIXED_HEADERS];
    return grid;
  });
  
  const settingFileInputRef = useRef(null);
  const distFileInputRef = useRef(null);
  const gridFileInputRef = useRef(null);

  useEffect(() => {
    if (!window.XLSX) {
      const script = document.createElement('script');
      script.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
      script.async = true;
      document.body.appendChild(script);
    }
  }, []);

  const showMsg = (type, text) => {
    setNotification({ type, text });
    setTimeout(() => setNotification(null), 5000);
  };

  // --- 修復：補回首頁點擊年級時的自動雲端抓取函數 ---
  const handleSelectGrade = async (grade) => {
    setSelectedGrade(grade);
    const gradeUrl = CLOUD_URLS[grade]?.grade;
    const distUrl = CLOUD_URLS[grade]?.dist;

    // 如果有設定雲端連結，則進行抓取
    if (gradeUrl || distUrl) {
      setIsLoading(true);
      showMsg('info', `正在自動同步 ${grade} 年級最新雲端標準...`);
      try {
        let newGradeText = appSettings[grade].grade;
        let newDistText = appSettings[grade].dist;

        if (gradeUrl) {
          const gradeRes = await fetch(gradeUrl);
          if (gradeRes.ok) newGradeText = await gradeRes.text();
        }
        if (distUrl) {
          const distRes = await fetch(distUrl);
          if (distRes.ok) newDistText = await distRes.text();
        }

        setAppSettings(prev => ({
          ...prev,
          [grade]: { grade: newGradeText, dist: newDistText }
        }));
        showMsg('success', `✅ 已自動套用 ${grade} 年級雲端最新標準！`);
      } catch (err) {
        showMsg('error', '⚠️ 雲端連線失敗，已載入系統預設標準。');
      } finally {
        setIsLoading(false);
        setView('input');
      }
    } else {
      // 若無連結則直接進入輸入頁面
      setView('input');
    }
  };

  const currentParsedSettings = useMemo(() => {
    if (!selectedGrade) return { settings: {}, subjects: [] };
    const parsed = parseCSV(appSettings[selectedGrade].grade);
    const settings = {};
    const subjects = Object.keys(parsed[0] || {}).filter(k => k !== '等級' && k !== '');
    subjects.forEach(subject => {
      settings[subject] = parsed
        .map(row => ({ level: row['等級'], minScore: parseFloat(row[subject]) }))
        .filter(item => !isNaN(item.minScore))
        .sort((a, b) => b.minScore - a.minScore);
    });
    return { settings, subjects };
  }, [appSettings, selectedGrade]);

  const currentDistMap = useMemo(() => {
    if (!selectedGrade) return [];
    return processDistribution(parseCSV(appSettings[selectedGrade].dist));
  }, [appSettings, selectedGrade]);

  const handleCellChange = (rIdx, cIdx, value) => {
    if (rIdx === 0) return; 
    const newGrid = [...gridData];
    newGrid[rIdx] = [...newGrid[rIdx]];
    newGrid[rIdx][cIdx] = value;
    setGridData(newGrid);
  };

  const handleGridPaste = (e, startRow, startCol) => {
    e.preventDefault();
    if (startRow === 0) return; 
    const pasteText = e.clipboardData.getData('text');
    if (!pasteText) return;
    const rows = pasteText.split(/\r?\n/);
    const newGrid = [...gridData];

    let pasteCount = 0;
    rows.forEach((rowStr, i) => {
      const targetRow = startRow + i;
      if (targetRow >= INITIAL_GRID_ROWS) return;
      const cells = rowStr.split(/\t|,/);
      if (cells.some(c => c.trim() !== '')) pasteCount++;
      newGrid[targetRow] = [...newGrid[targetRow]];
      cells.forEach((cellVal, j) => {
        const targetCol = startCol + j;
        if (targetCol < FIXED_HEADERS.length) { 
          newGrid[targetRow][targetCol] = cellVal.trim().replace(/^"|"$/g, '');
        }
      });
    });
    setGridData(newGrid);
    showMsg('success', `✅ 成功貼上 ${pasteCount} 筆資料`);
  };

  const handleClearGrid = () => {
    if (clearConfirm) {
      const grid = Array(INITIAL_GRID_ROWS).fill(0).map(() => Array(FIXED_HEADERS.length).fill(''));
      grid[0] = [...FIXED_HEADERS];
      setGridData(grid);
      setClearConfirm(false);
      showMsg('success', '✅ 表格資料已全部清空。');
    } else {
      setClearConfirm(true);
      setTimeout(() => setClearConfirm(false), 3000);
    }
  };

  const handleGridFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    if (!window.XLSX) { showMsg('error', "模組載入中，請稍後再試。"); return; }
    
    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const workbook = window.XLSX.read(new Uint8Array(event.target.result), { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const arr = window.XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });
        
        let headerRowIndex = -1;
        for (let i = 0; i < Math.min(arr.length, 10); i++) {
           const rowStr = arr[i].join('').toLowerCase();
           if (rowStr.includes('座號') && rowStr.includes('姓名')) { headerRowIndex = i; break; }
        }

        const dataStartRow = headerRowIndex !== -1 ? headerRowIndex + 1 : 1;
        const newGrid = Array(INITIAL_GRID_ROWS).fill(0).map(() => Array(FIXED_HEADERS.length).fill(''));
        newGrid[0] = [...FIXED_HEADERS];

        for (let i = dataStartRow; i < Math.min(arr.length, dataStartRow + INITIAL_GRID_ROWS - 1); i++) {
           const sourceRow = arr[i];
           const targetRowIdx = i - dataStartRow + 1;
           if (!sourceRow) continue;
           for(let j = 0; j < Math.min(sourceRow.length, FIXED_HEADERS.length); j++) {
               newGrid[targetRowIdx][j] = sourceRow[j] !== undefined ? String(sourceRow[j]) : '';
           }
        }
        setGridData(newGrid);
        showMsg('success', '✅ Excel 成績匯入成功！');
      } catch (err) {
        showMsg('error', '❌ 讀取 Excel 失敗，請確認檔案格式。');
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  const fetchCloudData = async (type) => {
    const url = CLOUD_URLS[selectedGrade]?.[type];
    
    if (!url) {
      showMsg('error', `⚠️ 系統尚未設定 ${selectedGrade} 年級的雲端連結，請在程式碼上方填寫 URL。`);
      return;
    }
    
    showMsg('info', `正在同步 ${selectedGrade} 年級雲端資料，請稍候...`);
    
    try {
      const response = await fetch(url);
      if (!response.ok) throw new Error('讀取失敗');
      const text = await response.text();
      
      if (!text.includes(',') && !text.includes('\t')) {
         throw new Error('格式不符');
      }
      
      setAppSettings(prev => ({
        ...prev,
        [selectedGrade]: { ...prev[selectedGrade], [type]: text }
      }));
      showMsg('success', '✅ 雲端資料同步成功！');
    } catch (err) {
      showMsg('error', '❌ 同步失敗！請確認該固定連結是否有效且已發佈為 CSV。');
    }
  };

  const handleAdminFileUpload = (e, type) => {
    const file = e.target.files[0];
    if (!file) return;
    const fileExt = file.name.split('.').pop().toLowerCase();
    const updateSetting = (text) => {
      setAppSettings(prev => ({...prev, [selectedGrade]: { ...prev[selectedGrade], [type]: text }}));
      showMsg('success', '✅ 本機檔案匯入成功！');
    };

    if (fileExt === 'csv') {
      const reader = new FileReader();
      reader.onload = (e) => updateSetting(e.target.result);
      reader.readAsText(file); 
    } else if (fileExt === 'xlsx' || fileExt === 'xls') {
      if (!window.XLSX) return;
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const workbook = window.XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
          updateSetting(window.XLSX.utils.sheet_to_csv(workbook.Sheets[workbook.SheetNames[0]]));
        } catch (err) { showMsg('error', '❌ 解析 Excel 失敗'); }
      };
      reader.readAsArrayBuffer(file);
    }
    e.target.value = '';
  };

  const goHome = () => {
    setView('home');
    setSelectedGrade(null);
  };

  // --- 報表資料計算 (包含加權、會考等級統計) ---
  const generateReportData = useMemo(() => {
    if (view !== 'result') return null;
    
    const headers = gridData[0];
    const excludeCols = ['座號', '姓名'];
    const subjects = headers.filter(h => h && !excludeCols.includes(h));

    const rawScoresData = gridData.slice(1).map(row => {
      const obj = {};
      let hasData = false;
      headers.forEach((h, i) => {
        obj[h] = row[i];
        if (row[i] && row[i].trim() !== '') hasData = true;
      });
      return hasData ? obj : null;
    }).filter(Boolean);

    if (rawScoresData.length === 0) return { error: "尚未輸入任何成績資料。" };

    const unmappedSubjects = subjects.filter(sub => !currentParsedSettings.settings[sub]);
    
    // 初始化各科統計物件
    const subjectStats = {};
    subjects.forEach(sub => {
      subjectStats[sub] = { 'A++':0, 'A+':0, 'A':0, 'B++':0, 'B+':0, 'B':0, 'C':0 };
    });

    let calculatedData = rawScoresData.map(student => {
      let totalWeightedScore = 0; 
      let totalWeight = 0;
      
      const gradeCounts = {};
      let mainACount = 0; let mainBCount = 0; let mainCCount = 0;
      
      const resultRow = { '座號': student['座號'] || '', '姓名': student['姓名'] || '' };
      
      subjects.forEach(subject => {
        const val = student[subject];
        const numScore = parseFloat(val);
        if (!isNaN(numScore)) {
          const weight = SUBJECT_WEIGHTS[subject] || 1;
          totalWeightedScore += numScore * weight;
          totalWeight += weight;

          if (currentParsedSettings.settings[subject]) {
             const grade = getGradeLevel(numScore, currentParsedSettings.settings[subject]);
             resultRow[subject] = `${numScore} (${grade})`;
             
             gradeCounts[grade] = (gradeCounts[grade] || 0) + 1;
             if (grade.startsWith('A')) mainACount++;
             if (grade.startsWith('B')) mainBCount++;
             if (grade.startsWith('C')) mainCCount++;
             
             if(subjectStats[subject][grade] !== undefined) {
                 subjectStats[subject][grade]++;
             }
          } else {
             resultRow[subject] = numScore;
          }
        } else { resultRow[subject] = val || ''; }
      });
      
      resultRow['加權平均'] = totalWeight > 0 ? parseFloat((totalWeightedScore / totalWeight).toFixed(2)) : 0;
      
      const gradeOrder = ['A++', 'A+', 'A', 'B++', 'B+', 'B', 'C'];
      let detailedSummary = gradeOrder.filter(g => gradeCounts[g]).map(g => `${gradeCounts[g]}${g}`).join('');
      let mainSummary = `${mainACount > 0 ? mainACount + 'A' : ''}${mainBCount > 0 ? mainBCount + 'B' : ''}${mainCCount > 0 ? mainCCount + 'C' : ''}`;
      
      resultRow['會考等級'] = detailedSummary ? `${mainSummary} (${detailedSummary})` : '';

      return resultRow;
    });

    calculatedData.sort((a, b) => b['加權平均'] - a['加權平均']);
    calculatedData.forEach((student, index) => {
      student['班排'] = (index > 0 && student['加權平均'] === calculatedData[index - 1]['加權平均']) ? calculatedData[index - 1]['班排'] : index + 1;
      student['預估校排'] = getSchoolRank(student['加權平均'], currentDistMap);
    });

    calculatedData.sort((a, b) => (parseInt(a['座號']) || 999) - (parseInt(b['座號']) || 999));
    
    return { data: calculatedData, unmappedSubjects, subjectStats, subjects };
  }, [view, gridData, currentParsedSettings, currentDistMap]);

  const handleCopyReport = () => {
    if (!generateReportData?.data) return;
    const tsvContent = [
      Object.keys(generateReportData.data[0]).join('\t'), 
      ...generateReportData.data.map(row => Object.values(row).join('\t'))
    ].join('\n');
    const textArea = document.createElement("textarea");
    textArea.value = tsvContent; document.body.appendChild(textArea); textArea.select();
    try { 
      document.execCommand('copy'); 
      showMsg('success', '✅ 已成功複製報表資料！請至 Excel 貼上。'); 
    } catch (err) { 
      showMsg('error', '❌ 複製失敗，請手動框選資料複製。'); 
    }
    document.body.removeChild(textArea);
  };

  return (
    <div className="min-h-screen bg-[#F4F7F9] text-slate-800 p-4 md:p-6 font-sans flex flex-col items-center">
      <div className="w-full max-w-[1250px] flex flex-col flex-grow space-y-4">
        
        <header className="w-full flex justify-between items-center bg-white p-4 md:p-5 rounded-2xl shadow-sm border border-slate-100">
          <div className="flex items-center space-x-4 cursor-pointer" onClick={goHome}>
            <div className="bg-blue-600 p-2.5 rounded-xl text-white shadow-md">
              <Calculator className="w-6 h-6" />
            </div>
            <h1 className="text-xl md:text-2xl font-black text-slate-800 tracking-tight">成績等級產生器</h1>
          </div>
          <div className="flex items-center gap-3">
            {view !== 'home' && view !== 'admin_login' && view !== 'admin_settings' && (
              <button onClick={() => setView('admin_login')} className="text-sm font-bold text-slate-500 hover:text-slate-800 px-3 py-2 flex items-center bg-slate-100 rounded-lg transition-colors">
                <Settings className="w-4 h-4 mr-2" /> 管理設定
              </button>
            )}
            {view !== 'home' && (
               <button onClick={goHome} className="text-sm font-bold text-blue-700 hover:text-white hover:bg-blue-600 px-4 py-2 flex items-center bg-blue-50 border border-blue-200 rounded-lg transition-all shadow-sm">
                 <Home className="w-4 h-4 mr-2" /> 返回首頁
               </button>
            )}
          </div>
        </header>

        {notification && (
          <div className={`w-full px-5 py-4 rounded-2xl flex items-center border text-base font-bold shadow-md transition-all animate-in slide-in-from-top-2 ${
            notification.type === 'error' ? 'bg-red-50 text-red-600 border-red-200' :
            notification.type === 'success' ? 'bg-emerald-50 text-emerald-700 border-emerald-200' :
            'bg-blue-50 text-blue-600 border-blue-200'
          }`}>
            {notification.type === 'error' && <AlertTriangle className="w-5 h-5 mr-3 flex-shrink-0" />}
            {notification.type === 'success' && <CheckCircle2 className="w-5 h-5 mr-3 flex-shrink-0" />}
            {notification.type === 'info' && <RefreshCw className="w-5 h-5 mr-3 flex-shrink-0 animate-spin" />}
            {notification.text}
          </div>
        )}

        <main className="w-full bg-white rounded-2xl shadow-sm border border-slate-100 flex flex-col flex-grow relative overflow-hidden items-center">
          
          {view === 'home' && (
            <div className="p-8 md:p-16 flex flex-col items-center justify-center flex-grow animate-in zoom-in-95 w-full">
              <div className="w-20 h-20 bg-blue-50 text-blue-600 rounded-full flex items-center justify-center mb-6 shadow-inner">
                <BookOpen className="w-10 h-10" />
              </div>
              <h2 className="text-3xl font-black text-slate-800 mb-3">請選擇欲輸入的年級</h2>
              <p className="text-slate-500 text-base mb-10 text-center font-medium">系統將自動載入該年級對應的全校組距標準與等級門檻</p>
              
              <div className="flex flex-wrap justify-center gap-6 w-full">
                {['7', '8', '9'].map(grade => (
                  <button key={grade} onClick={() => handleSelectGrade(grade)} disabled={isLoading}
                    className={`w-32 h-40 group flex flex-col items-center justify-center p-4 bg-white border-2 border-slate-200 hover:border-blue-500 hover:shadow-lg hover:bg-blue-50/50 rounded-2xl transition-all ${isLoading ? 'opacity-50 cursor-not-allowed' : ''}`}
                  >
                    <span className="text-5xl font-black text-slate-300 group-hover:text-blue-600 mb-2 transition-colors">{grade}</span>
                    <span className="text-base font-bold text-slate-500 group-hover:text-blue-800">年級專區</span>
                  </button>
                ))}
              </div>
            </div>
          )}

          {view === 'input' && selectedGrade && (
            <div className="flex flex-col items-center w-full flex-grow p-6 animate-in fade-in">
              <div className="w-[750px] flex flex-col">
                <div className="w-full bg-blue-50/60 border border-blue-200 rounded-xl p-4 mb-5 flex items-start shadow-sm">
                  <Info className="w-5 h-5 text-blue-600 mr-3 mt-0.5 flex-shrink-0" />
                  <div>
                    <h3 className="text-base font-bold text-blue-800 mb-1">【{selectedGrade}年級】成績輸入說明</h3>
                    <p className="text-sm text-blue-700 leading-relaxed font-medium">
                      1. 支援直接在下方表格內輸入或修改成績。<br />
                      2. 支援從 Excel 複製多筆資料，點選表格首格後按下 <kbd className="px-1.5 py-0.5 bg-white border border-blue-200 rounded text-xs mx-1">Ctrl+V</kbd> 貼上。<br />
                      3. 也可點選「匯入 Excel」按鈕直接上傳檔案（系統會自動對齊座號與姓名）。
                    </p>
                  </div>
                </div>

                <div className="w-full flex justify-between items-center mb-4">
                  <div className="flex gap-3">
                    <input type="file" accept=".csv, .xlsx, .xls" style={{ display: 'none' }} ref={gridFileInputRef} onChange={handleGridFileUpload} />
                    <button onClick={() => gridFileInputRef.current.click()} className="flex items-center justify-center px-4 py-2 bg-emerald-50 hover:bg-emerald-100 text-emerald-800 border border-emerald-200 font-bold rounded-xl transition-colors text-sm shadow-sm">
                      <FileSpreadsheet className="w-4 h-4 mr-2" /> 匯入 Excel
                    </button>
                    <button onClick={handleClearGrid} className={`px-4 py-2 font-bold rounded-xl transition-colors text-sm flex items-center justify-center shadow-sm ${clearConfirm ? 'bg-red-600 text-white hover:bg-red-700 border-red-700' : 'bg-slate-50 text-slate-600 hover:bg-slate-100 border border-slate-200'}`}>
                      <Trash2 className="w-4 h-4 mr-2" /> {clearConfirm ? '確定清空？' : '清空表格'}
                    </button>
                  </div>
                  <button onClick={() => gridData.slice(1).some(r => r.some(c => c)) ? setView('result') : showMsg('error', '請先輸入成績資料')} className="px-6 py-2.5 bg-slate-800 hover:bg-slate-900 text-white font-bold rounded-xl transition-all shadow-md text-base flex items-center justify-center">
                    產生報表 <ChevronRight className="w-5 h-5 ml-1" />
                  </button>
                </div>

                <div className="w-full border border-slate-300 rounded-xl overflow-hidden shadow-inner bg-slate-50 flex flex-col h-[400px]">
                  <div className="overflow-y-auto w-full custom-scrollbar">
                    <table className="w-[750px] text-base border-collapse table-fixed bg-white">
                      <thead className="sticky top-0 z-20 bg-slate-200 shadow-sm">
                        <tr>
                          <th style={COL_STYLES['#']} className="border-r border-slate-300 text-slate-600 py-3 text-center font-bold">#</th>
                          {FIXED_HEADERS.map((header, cIdx) => (
                            <th key={`h-${cIdx}`} style={COL_STYLES[header]} className="p-0 border-r border-slate-300 last:border-r-0 bg-slate-200">
                              <div className="w-full py-3 font-bold text-slate-800 text-center tracking-wider">
                                  {header}
                              </div>
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody className="bg-white">
                        {gridData.slice(1).map((row, rIdx) => {
                          const actualRowIdx = rIdx + 1;
                          const hasData = row.some(cell => cell.trim() !== '');
                          return (
                            <tr key={`r-${actualRowIdx}`} className={`${hasData ? 'bg-white' : 'bg-[#FDFDFD]'} hover:bg-blue-50/60 border-b border-slate-200 transition-colors`}>
                              <td className="bg-slate-100 border-r border-slate-200 text-slate-500 text-center font-mono text-sm py-2 font-medium">
                                {actualRowIdx}
                              </td>
                              {row.map((cell, cIdx) => {
                                const headerName = FIXED_HEADERS[cIdx];
                                const isScoreCol = headerName !== '座號' && headerName !== '姓名';
                                let isError = false;
                                if (isScoreCol && cell.trim() !== '') {
                                  const num = parseFloat(cell);
                                  if (isNaN(num) || num < 0 || num > 100) isError = true;
                                }

                                return (
                                  <td key={`c-${cIdx}`} className="p-0 border-r border-slate-200 last:border-r-0">
                                    <input
                                      type="text"
                                      className={`w-full h-full py-2.5 px-3 focus:outline-none focus:bg-blue-100 transition-all 
                                        ${isScoreCol ? 'text-right font-mono text-base' : 'text-center text-base'} 
                                        ${cell.trim() !== '' ? 'text-slate-800 font-bold' : 'text-slate-400'}
                                        ${isError ? 'bg-red-50 text-red-600 focus:bg-red-50' : ''}`}
                                      value={cell}
                                      onChange={(e) => handleCellChange(actualRowIdx, cIdx, e.target.value)}
                                      onPaste={(e) => handleGridPaste(e, actualRowIdx, cIdx)}
                                    />
                                  </td>
                                );
                              })}
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            </div>
          )}

          {view === 'result' && generateReportData && (
            <div className="w-full flex flex-col h-full flex-grow p-6 animate-in slide-in-from-bottom-4">
              
              <div className="w-full flex flex-col md:flex-row justify-between items-start md:items-center bg-teal-50 p-4 rounded-xl border border-teal-200 mb-5 gap-4 shadow-sm">
                <h2 className="text-xl font-black text-teal-900 flex items-center">
                  <CheckCircle2 className="w-6 h-6 text-teal-600 mr-2" />
                  {selectedGrade}年級 分析報表 <span className="text-teal-700 text-base font-bold ml-3 bg-teal-100 px-3 py-1 rounded-lg">共 {generateReportData.data?.length || 0} 筆紀錄</span>
                </h2>
                <div className="flex gap-3">
                  <button onClick={() => setView('input')} className="px-5 py-2.5 bg-white text-slate-700 border border-slate-300 rounded-xl hover:bg-slate-50 transition-all text-sm font-bold flex justify-center items-center shadow-sm">
                    <ArrowLeft className="w-4 h-4 mr-2" /> 修改資料
                  </button>
                  <button onClick={handleCopyReport} className="px-5 py-2.5 bg-white text-blue-700 border border-blue-200 rounded-xl hover:bg-blue-50 transition-all text-sm font-bold flex justify-center items-center shadow-sm">
                    <ClipboardCopy className="w-4 h-4 mr-2" /> 複製表格
                  </button>
                  <button onClick={() => exportToCSV(generateReportData.data, `${selectedGrade}年級_成績分析.csv`)} className="px-5 py-2.5 bg-blue-600 text-white rounded-xl hover:bg-blue-700 transition-all text-sm font-bold flex justify-center items-center shadow-md">
                    <FileDown className="w-4 h-4 mr-2" /> 匯出 CSV
                  </button>
                </div>
              </div>

              {generateReportData.error ? (
                 <div className="w-full p-12 text-center text-red-600 font-bold text-lg bg-red-50 rounded-xl border border-red-200">{generateReportData.error}</div>
              ) : (
                <div className="w-full flex-grow flex flex-col min-h-[300px]">
                  {generateReportData.unmappedSubjects && generateReportData.unmappedSubjects.length > 0 && (
                    <div className="w-full bg-amber-50 border border-amber-200 p-4 rounded-xl flex items-start text-sm mb-4 shadow-sm">
                      <AlertTriangle className="w-5 h-5 text-amber-500 mr-3 flex-shrink-0 mt-0.5" />
                      <p className="text-amber-800 font-medium">未設定等級門檻的科目：<strong className="text-amber-900 mx-1">{generateReportData.unmappedSubjects.join('、')}</strong>。已自動計入平均，但無法於報表中標示 ABC 等級。</p>
                    </div>
                  )}

                  <div className="w-full border border-slate-200 rounded-xl overflow-x-auto shadow-sm bg-white custom-scrollbar flex-grow">
                    <table className="w-max min-w-full text-[15px] text-center whitespace-nowrap table-auto border-collapse">
                      <thead className="text-slate-600 bg-slate-100 font-black sticky top-0 shadow-sm border-b-2 border-slate-300">
                        <tr>
                          {Object.keys(generateReportData.data[0] || {}).map((header, idx) => (
                            <th key={idx} className="px-5 py-3.5 border-r border-slate-200 last:border-r-0 tracking-wide">
                              {header}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100 font-medium">
                        {generateReportData.data.map((row, rowIndex) => (
                          <tr key={rowIndex} className="hover:bg-blue-50/70 transition-colors">
                            {Object.entries(row).map(([key, val], colIndex) => {
                              const isGrade = typeof val === 'string' && val.includes('(') && val.includes(')');
                              const gradeMatch = isGrade ? val.match(/\((.*?)\)/) : null;
                              const gradeLabel = gradeMatch ? gradeMatch[1] : '';
                              
                              let textClass = "text-slate-800";
                              if (gradeLabel.includes('A')) textClass = "text-emerald-600 font-black";
                              if (gradeLabel.includes('C')) textClass = "text-rose-600 font-black";
                              
                              if (key === '會考等級') textClass = "text-fuchsia-700 font-black tracking-wide";
                              if (key === '加權平均') textClass = "text-indigo-700 font-black";
                              if (key === '預估校排' || key === '班排') textClass = "text-blue-700 font-black";

                              return (
                                <td key={colIndex} className={`px-4 py-3 border-r border-slate-50 last:border-r-0 ${textClass}`}>
                                  {val}
                                </td>
                              );
                            })}
                          </tr>
                        ))}
                      </tbody>
                      {/* --- 各科統計 Footer --- */}
                      <tfoot className="bg-amber-50/60 border-t-2 border-slate-300">
                        <tr>
                          {Object.keys(generateReportData.data[0] || {}).map((header, idx) => {
                            if (idx === 0) {
                              return (
                                <td key={idx} colSpan={2} className="px-5 py-5 font-black text-slate-700 text-right tracking-widest border-r border-slate-200 align-top">
                                  各科等級人數統計
                                </td>
                              );
                            }
                            if (idx === 1) return null; // 被 colSpan 合併
                            
                            if (generateReportData.subjects.includes(header) && generateReportData.subjectStats[header]) {
                               const stats = generateReportData.subjectStats[header];
                               const hasAny = Object.values(stats).some(v => v > 0);
                               return (
                                 <td key={idx} className="px-4 py-4 text-left align-top border-r border-slate-200">
                                   {hasAny ? (
                                     <div className="flex flex-col gap-1.5 text-[13px] w-full min-w-[70px] mx-auto">
                                       {['A++', 'A+', 'A', 'B++', 'B+', 'B', 'C'].map(g => (
                                          stats[g] > 0 ? (
                                            <div key={g} className="flex justify-between items-center w-full">
                                              <span className={`font-black tracking-tighter ${g.includes('A') ? 'text-emerald-600' : g.includes('C') ? 'text-rose-600' : 'text-slate-600'}`}>{g}</span>
                                              <span className="text-slate-600 font-black bg-white px-2 py-0.5 rounded shadow-sm border border-slate-200 text-xs">{stats[g]}</span>
                                            </div>
                                          ) : null
                                       ))}
                                     </div>
                                   ) : (
                                     <span className="text-slate-400 text-xs flex justify-center">-</span>
                                   )}
                                 </td>
                               );
                            }
                            return <td key={idx} className="px-5 py-3 border-r border-slate-200 last:border-r-0"></td>;
                          })}
                        </tr>
                      </tfoot>
                    </table>
                  </div>
                </div>
              )}
            </div>
          )}

          {view === 'admin_login' && (
             <div className="w-full p-8 md:p-16 flex flex-col items-center justify-center flex-grow animate-in zoom-in-95">
               <div className="bg-slate-100 p-6 rounded-full mb-6 shadow-inner border border-slate-200">
                 <Lock className="w-12 h-12 text-slate-500" />
               </div>
               <h2 className="text-2xl font-black text-slate-800 mb-3">管理員身分驗證</h2>
               <p className="text-slate-500 text-base mb-8 text-center font-medium">請輸入系統密碼以進入全校標準設定頁面</p>
               <form onSubmit={(e) => {
                  e.preventDefault();
                  if (adminPassword === '690530') {
                    setView('admin_settings'); setAdminPassword('');
                  } else { showMsg('error', '密碼錯誤'); }
               }} className="flex flex-col w-full max-w-sm space-y-4">
                 <input type="password" autoFocus placeholder="請輸入密碼..." className="w-full px-5 py-3.5 text-center text-lg tracking-[0.2em] bg-white border-2 border-slate-300 rounded-xl focus:border-blue-500 focus:ring-4 focus:ring-blue-100 transition-all focus:outline-none" value={adminPassword} onChange={(e) => setAdminPassword(e.target.value)} />
                 <button type="submit" className="w-full py-3.5 bg-blue-600 hover:bg-blue-700 text-white font-black rounded-xl shadow-md flex justify-center items-center text-base transition-colors">
                   解鎖進入設定 <Unlock className="w-5 h-5 ml-2" />
                 </button>
               </form>
             </div>
          )}

          {view === 'admin_settings' && (
            <div className="w-full p-6 md:p-10 flex flex-col flex-grow animate-in fade-in overflow-y-auto">
              <div className="text-center pb-6 mb-6 border-b border-slate-200 flex flex-col items-center">
                <h2 className="text-2xl font-black text-slate-800 mb-4 flex items-center">
                   <Settings className="w-6 h-6 mr-2 text-slate-600" /> 全校標準與組距設定
                </h2>
                <div className="flex bg-slate-100 p-1.5 rounded-xl space-x-2 shadow-inner border border-slate-200">
                  {['7', '8', '9'].map(grade => (
                    <button key={`admin-${grade}`} onClick={() => setSelectedGrade(grade)} className={`px-6 py-2 rounded-lg font-black text-base transition-all ${selectedGrade === grade ? 'bg-white text-blue-700 shadow-sm border border-slate-200' : 'text-slate-500 hover:text-slate-700'}`}>
                      {grade} 年級設定
                    </button>
                  ))}
                </div>
              </div>
              {!selectedGrade ? (
                <div className="w-full text-center py-16 text-slate-400 font-bold text-lg flex flex-col items-center">
                  <ArrowLeft className="w-8 h-8 mb-3 text-slate-300 animate-pulse" />
                  請先在上方選擇欲設定的年級
                </div>
              ) : (
                <div className="w-full grid grid-cols-1 lg:grid-cols-2 gap-6 flex-grow">
                  
                  <div className="flex flex-col bg-[#FFFCF5] p-5 rounded-2xl border border-amber-200 shadow-sm">
                    <div className="flex justify-between items-center mb-3">
                      <span className="text-base font-black text-amber-900 flex items-center"><BookOpen className="w-5 h-5 mr-2 text-amber-600"/> 各科等級門檻</span>
                      <div>
                        <input type="file" accept=".csv, .xlsx, .xls" style={{ display: 'none' }} ref={settingFileInputRef} onChange={(e) => handleAdminFileUpload(e, 'grade')} />
                        <button onClick={() => settingFileInputRef.current.click()} className="px-3 py-1.5 bg-white text-amber-800 border border-amber-300 rounded-lg text-xs font-bold shadow-sm hover:bg-amber-50 transition-colors flex items-center">
                          <Upload className="w-3.5 h-3.5 mr-1.5"/> 本機檔案
                        </button>
                      </div>
                    </div>
                    <div className="flex justify-between items-center mb-4 bg-amber-50 p-3 rounded-xl border border-amber-100 shadow-inner">
                       <span className="text-sm font-bold text-amber-800">從固定的雲端試算表同步最新資料</span>
                       <button onClick={() => fetchCloudData('grade')} className="px-4 py-2 bg-amber-600 text-white rounded-lg text-sm font-bold hover:bg-amber-700 whitespace-nowrap flex items-center shadow-sm transition-colors">
                         <Cloud className="w-4 h-4 mr-1.5" /> 雲端同步
                       </button>
                    </div>
                    <textarea className="w-full flex-grow min-h-[300px] p-4 border border-amber-300/80 rounded-xl text-sm font-mono text-slate-700 focus:outline-none focus:ring-2 focus:ring-amber-400 leading-relaxed shadow-inner" value={appSettings[selectedGrade].grade} onChange={(e) => setAppSettings(prev => ({...prev, [selectedGrade]: {...prev[selectedGrade], grade: e.target.value}}))} />
                  </div>

                  <div className="flex flex-col bg-[#F5F9FF] p-5 rounded-2xl border border-blue-200 shadow-sm">
                    <div className="flex justify-between items-center mb-3">
                      <span className="text-base font-black text-blue-900 flex items-center"><Users className="w-5 h-5 mr-2 text-blue-600"/> 全校分數組距</span>
                      <div>
                        <input type="file" accept=".csv, .xlsx, .xls" style={{ display: 'none' }} ref={distFileInputRef} onChange={(e) => handleAdminFileUpload(e, 'dist')} />
                        <button onClick={() => distFileInputRef.current.click()} className="px-3 py-1.5 bg-white text-blue-800 border border-blue-300 rounded-lg text-xs font-bold shadow-sm hover:bg-blue-50 transition-colors flex items-center">
                          <Upload className="w-3.5 h-3.5 mr-1.5"/> 本機檔案
                        </button>
                      </div>
                    </div>
                    <div className="flex justify-between items-center mb-4 bg-blue-50 p-3 rounded-xl border border-blue-100 shadow-inner">
                       <span className="text-sm font-bold text-blue-800">從固定的雲端試算表同步最新資料</span>
                       <button onClick={() => fetchCloudData('dist')} className="px-4 py-2 bg-blue-600 text-white rounded-lg text-sm font-bold hover:bg-blue-700 whitespace-nowrap flex items-center shadow-sm transition-colors">
                         <Cloud className="w-4 h-4 mr-1.5" /> 雲端同步
                       </button>
                    </div>
                    <textarea className="w-full flex-grow min-h-[300px] p-4 border border-blue-300/80 rounded-xl text-sm font-mono text-slate-700 focus:outline-none focus:ring-2 focus:ring-blue-400 leading-relaxed shadow-inner" value={appSettings[selectedGrade].dist} onChange={(e) => setAppSettings(prev => ({...prev, [selectedGrade]: {...prev[selectedGrade], dist: e.target.value}}))} />
                  </div>

                </div>
              )}
            </div>
          )}
        </main>

        <footer className="flex flex-col gap-4 py-8">
           <div className="bg-amber-50/50 border border-amber-200/50 rounded-3xl p-6 flex items-start shadow-sm mx-auto w-full max-w-[850px]">
              <AlertTriangle className="w-6 h-6 text-amber-500 mr-4 mt-0.5 shrink-0" />
              <div>
                <h4 className="text-amber-900 font-black text-sm mb-1">系統免責聲明</h4>
                <p className="text-xs text-amber-800/70 leading-relaxed font-medium">
                  本工具僅供教師進行快速成績換算與預估排名參考，並非學校官方正式成績系統。所有計算結果（包含校排預估）請務必以教務處公告之正式紙本成績單或校務系統數據為準。
                </p>
              </div>
           </div>
           <div className="text-center mt-2">
             <div className="text-slate-300 text-xs font-bold tracking-widest">程式設計：蘇老爹</div>
           </div>
        </footer>
      </div>
    </div>
  );
}