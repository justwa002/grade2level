import React, { useState, useMemo, useRef, useEffect } from 'react';
import { 
  Upload, FileDown, ClipboardCopy, Calculator, CheckCircle2, 
  Settings, ArrowLeft, Lock, Unlock, AlertTriangle, Users, 
  BookOpen, ChevronRight, FileSpreadsheet, Trash2, Home, Info,
  Cloud, RefreshCw, DownloadCloud
} from 'lucide-react';

// --- 雲端試算表固定連結設定 ---
const CLOUD_URLS = {
  '7': { grade: "https://docs.google.com/spreadsheets/d/e/2PACX-1vR9LhxgNWTLkGftNnMkHQTR449Y_7M0NDr_IR_Oi5lTYZvCF9s01onsLaBWrxuA69DPntEwv0hFNU72/pub?gid=2077033678&single=true&output=csv", dist: "https://docs.google.com/spreadsheets/d/e/2PACX-1vR9LhxgNWTLkGftNnMkHQTR449Y_7M0NDr_IR_Oi5lTYZvCF9s01onsLaBWrxuA69DPntEwv0hFNU72/pub?gid=859457249&single=true&output=csv" },
  '8': { grade: "https://docs.google.com/spreadsheets/d/e/2PACX-1vR9LhxgNWTLkGftNnMkHQTR449Y_7M0NDr_IR_Oi5lTYZvCF9s01onsLaBWrxuA69DPntEwv0hFNU72/pub?gid=0&single=true&output=csv", dist: "https://docs.google.com/spreadsheets/d/e/2PACX-1vR9LhxgNWTLkGftNnMkHQTR449Y_7M0NDr_IR_Oi5lTYZvCF9s01onsLaBWrxuA69DPntEwv0hFNU72/pub?gid=853170505&single=true&output=csv" },
  '9': { grade: "https://docs.google.com/spreadsheets/d/e/2PACX-1vR9LhxgNWTLkGftNnMkHQTR449Y_7M0NDr_IR_Oi5lTYZvCF9s01onsLaBWrxuA69DPntEwv0hFNU72/pub?gid=1634530372&single=true&output=csv", dist: "https://docs.google.com/spreadsheets/d/e/2PACX-1vR9LhxgNWTLkGftNnMkHQTR449Y_7M0NDr_IR_Oi5lTYZvCF9s01onsLaBWrxuA69DPntEwv0hFNU72/pub?gid=1683092563&single=true&output=csv" }
};

const SUBJECT_WEIGHTS = {
  '國文': 5,
  '英文': 3,
  '數學': 4,
  '社會': 3,
  '自然': 3
};

const defaultSettings = `,等級,國文,英文,數學,社會,自然\n,A++,92,100,94,94,96\n,A+,89,98,89,88,92\n,A,84,95,79,80,86\n,B++,80,92,70,72,78\n,B+,74,88,62,64,68\n,B,52,50,28,34,36`;
const defaultDistribution = `分數組距,全校人數,累計人數\n100,0,0\n98-99.99,0,0\n96-97.99,23,23\n94-95.99,52,75\n92-93.99,76,151\n90-91.99,73,224\n87-90.99,102,326\n84-86.99,75,401\n80-83.99,117,518\n70-79.99,174,692\n60-69.99,127,819\n0-59.99,25,844`;

const FIXED_HEADERS = ['座號', '姓名', '國文', '英文', '數學', '社會', '自然'];
const INITIAL_GRID_ROWS = 45;

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

const exportReportToExcel = (reportData, gradeLevel, showMsg) => {
  if (!reportData || !reportData.data || reportData.data.length === 0) {
    if (showMsg) showMsg('error', '沒有可匯出的資料');
    return;
  }
  const { data, subjectStats, subjects } = reportData;
  const headers = Object.keys(data[0]);
  const exportData = [...data];

  const emptyRow = {}; headers.forEach(h => emptyRow[h] = ''); exportData.push(emptyRow);
  const titleRow = {}; headers.forEach((h, i) => titleRow[h] = i === 0 ? '各科等級人數統計' : ''); exportData.push(titleRow);

  const grades = ['A++', 'A+', 'A', 'B++', 'B+', 'B', 'C'];
  grades.forEach(g => {
    const row = {};
    headers.forEach((h, i) => {
      if (i === 0) row[h] = g;
      else if (subjects && subjects.includes(h) && subjectStats && subjectStats[h] !== undefined) row[h] = subjectStats[h][g] || 0;
      else row[h] = '';
    });
    exportData.push(row);
  });

  if (window.XLSX) {
    try {
      const ws = window.XLSX.utils.json_to_sheet(exportData);
      const wb = window.XLSX.utils.book_new();
      window.XLSX.utils.book_append_sheet(wb, ws, `${gradeLevel}年級成績分析`);
      window.XLSX.writeFile(wb, `${gradeLevel}年級_成績分析報表.xlsx`);
      if (showMsg) showMsg('success', '✅ Excel 報表匯出成功！');
      return;
    } catch (error) { console.error("Excel Error", error); }
  }

  const csvContent = [
    headers.join(','),
    ...exportData.map(row => headers.map(header => `"${row[header] !== undefined ? row[header] : ''}"`).join(','))
  ].join('\n');
  const blob = new Blob(['\uFEFF' + csvContent], { type: 'text/csv;charset=utf-8;' });
  const link = document.createElement('a'); link.href = URL.createObjectURL(blob); 
  link.download = `${gradeLevel}年級_成績分析報表.csv`;
  document.body.appendChild(link); link.click(); document.body.removeChild(link);
  if (showMsg) showMsg('success', '✅ 報表匯出成功！');
};

export default function App() {
  const [view, setView] = useState('home'); 
  const [selectedGrade, setSelectedGrade] = useState(null); 
  const [adminPassword, setAdminPassword] = useState('');
  const [notification, setNotification] = useState(null); 
  const [clearConfirm, setClearConfirm] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const [syncLogs, setSyncLogs] = useState([{ time: new Date().toLocaleTimeString(), msg: '系統初始化完成，準備就緒。' }]);

  const addLog = (msg) => {
    setSyncLogs(prev => [{ time: new Date().toLocaleTimeString(), msg }, ...prev].slice(0, 8)); // 保留最近8筆日誌
  };

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

  // 單一年級自動同步 (首頁點擊時)
  const handleSelectGrade = async (grade) => {
    setSelectedGrade(grade);
    const gradeUrl = CLOUD_URLS[grade]?.grade;
    const distUrl = CLOUD_URLS[grade]?.dist;

    if (gradeUrl || distUrl) {
      setIsLoading(true);
      showMsg('info', `正在同步 ${grade} 年級最新雲端標準...`);
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
        addLog(`自動載入 ${grade} 年級雲端標準成功`);
      } catch (err) {
        showMsg('error', '⚠️ 雲端連線失敗，已載入系統預設標準。');
        addLog(`讀取 ${grade} 年級雲端標準失敗，已使用預設`);
      } finally {
        setIsLoading(false);
        setView('input');
      }
    } else {
      setView('input');
    }
  };

  // 一鍵同步所有年級資料 (新功能)
  const handleSyncAllGrades = async () => {
    setIsLoading(true);
    showMsg('info', '正在批次同步所有年級(7,8,9)雲端資料，請稍候...');
    addLog('開始批次同步所有年級雲端資料...');
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
      if(successCount > 0) {
        showMsg('success', '✅ 所有年級雲端資料同步完成！');
        addLog(`批次同步完成 (成功獲取 ${successCount} 份檔案)`);
      } else {
        showMsg('error', '⚠️ 找不到有效的雲端連結，同步失敗。');
        addLog(`批次同步失敗 (找不到雲端連結)`);
      }
    } catch (err) {
      showMsg('error', '❌ 批次同步過程中發生網路錯誤。');
      addLog(`批次同步發生網路錯誤`);
    } finally {
      setIsLoading(false);
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
    if (!url) { showMsg('error', `⚠️ 尚未設定 ${selectedGrade} 年級的雲端連結。`); return; }
    showMsg('info', `正在同步 ${selectedGrade} 年級雲端資料...`);
    try {
      const response = await fetch(url);
      if (!response.ok) throw new Error('讀取失敗');
      const text = await response.text();
      setAppSettings(prev => ({
        ...prev,
        [selectedGrade]: { ...prev[selectedGrade], [type]: text }
      }));
      showMsg('success', '✅ 雲端資料同步成功！');
      addLog(`手動同步 ${selectedGrade} 年級 ${type === 'grade' ? '等級' : '組距'} 資料成功`);
    } catch (err) {
      showMsg('error', '❌ 同步失敗！請確認固定連結有效。');
      addLog(`手動同步 ${selectedGrade} 年級 ${type === 'grade' ? '等級' : '組距'} 資料失敗`);
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
    const { data, subjectStats, subjects } = generateReportData;
    const headers = Object.keys(data[0]);
    
    let tsvLines = [
      headers.join('\t'), 
      ...data.map(row => headers.map(h => row[h]).join('\t'))
    ];

    // 附加各科等級統計資料到底部
    tsvLines.push(headers.map(() => '').join('\t')); // 產生一列空行做區隔
    tsvLines.push(headers.map((h, i) => i === 0 ? '各科等級人數統計' : '').join('\t')); // 標題行
    
    const grades = ['A++', 'A+', 'A', 'B++', 'B+', 'B', 'C'];
    grades.forEach(g => {
      const rowData = headers.map((h, i) => {
        if (i === 0) return g; // 座號欄放置 A++, A+ 等級標籤
        if (subjects && subjects.includes(h) && subjectStats && subjectStats[h] !== undefined) return subjectStats[h][g] || 0;
        return '';
      });
      tsvLines.push(rowData.join('\t'));
    });

    const tsvContent = tsvLines.join('\n');
    const textArea = document.createElement("textarea");
    textArea.value = tsvContent; document.body.appendChild(textArea); textArea.select();
    try { 
      document.execCommand('copy'); showMsg('success', '✅ 已成功複製報表與統計資料！'); 
    } catch (err) { showMsg('error', '❌ 複製失敗，請手動框選複製。'); }
    document.body.removeChild(textArea);
  };

  return (
    <div className="min-h-screen bg-[#F4F7F9] text-slate-800 p-2 md:p-6 font-sans flex flex-col items-center">
      <div className="w-full max-w-[1250px] flex flex-col flex-grow space-y-4">
        
        <header className="w-full flex justify-between items-center bg-white p-4 md:p-5 rounded-2xl shadow-sm border border-slate-100 flex-wrap gap-2">
          <div className="flex items-center space-x-3 md:space-x-4 cursor-pointer" onClick={() => setView('home')}>
            <div className="bg-blue-600 p-2 md:p-2.5 rounded-xl text-white shadow-md">
              <Calculator className="w-5 h-5 md:w-6 md:h-6" />
            </div>
            <h1 className="text-lg md:text-2xl font-black text-slate-800 tracking-tight">成績等級產生器</h1>
          </div>
          <div className="flex items-center gap-2 md:gap-3">
            {view !== 'admin_login' && view !== 'admin_settings' && (
              <button onClick={() => setView('admin_login')} className="text-sm font-bold text-slate-500 hover:text-slate-800 px-3 py-2 flex items-center bg-slate-100 rounded-lg transition-colors">
                <Settings className="w-4 h-4 mr-2" /> <span className="hidden sm:inline">管理設定</span>
              </button>
            )}
            {view !== 'home' && (
               <button onClick={() => { setView('home'); setSelectedGrade(null); }} className="text-sm font-bold text-blue-700 hover:text-white hover:bg-blue-600 px-3 md:px-4 py-2 flex items-center bg-blue-50 border border-blue-200 rounded-lg transition-all shadow-sm">
                 <Home className="w-4 h-4 mr-1 md:mr-2" /> 返回
               </button>
            )}
          </div>
        </header>

        {notification && (
          <div className={`w-full px-5 py-3 md:py-4 rounded-2xl flex items-center border text-sm md:text-base font-bold shadow-md transition-all animate-in slide-in-from-top-2 ${
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
            <div className="p-6 md:p-16 flex flex-col items-center justify-center flex-grow animate-in zoom-in-95 w-full">
              <div className="w-16 h-16 md:w-20 md:h-20 bg-blue-50 text-blue-600 rounded-full flex items-center justify-center mb-6 shadow-inner">
                <BookOpen className="w-8 h-8 md:w-10 md:h-10" />
              </div>
              <h2 className="text-2xl md:text-3xl font-black text-slate-800 mb-3 text-center">請選擇欲輸入的年級</h2>
              <p className="text-slate-500 text-sm md:text-base mb-8 md:mb-10 text-center font-medium px-4">
                系統將自動載入該年級對應的全校組距標準與等級門檻
              </p>
              
              <div className="flex flex-wrap justify-center gap-4 md:gap-6 w-full px-4 mb-10">
                {['7', '8', '9'].map(grade => (
                  <button key={grade} onClick={() => handleSelectGrade(grade)} disabled={isLoading}
                    className={`w-28 h-36 md:w-32 md:h-40 group flex flex-col items-center justify-center p-4 bg-white border-2 border-slate-200 hover:border-blue-500 hover:shadow-lg hover:bg-blue-50/50 rounded-2xl transition-all ${isLoading ? 'opacity-50 cursor-not-allowed' : ''}`}
                  >
                    <span className="text-4xl md:text-5xl font-black text-slate-300 group-hover:text-blue-600 mb-2 transition-colors">{grade}</span>
                    <span className="text-sm md:text-base font-bold text-slate-500 group-hover:text-blue-800">年級專區</span>
                  </button>
                ))}
              </div>

              {/* 新增：系統更新日誌區塊 */}
              <div className="w-full max-w-md bg-white border border-slate-200 shadow-sm rounded-2xl p-4 md:p-5 flex flex-col">
                 <div className="flex items-center text-slate-700 font-black text-sm md:text-base mb-3 border-b border-slate-100 pb-2">
                   <RefreshCw className="w-4 h-4 md:w-5 md:h-5 mr-2 text-blue-500" />
                   系統數據更新日誌
                 </div>
                 <div className="flex flex-col space-y-2 h-32 overflow-y-auto custom-scrollbar pr-2">
                   {syncLogs.map((log, idx) => (
                     <div key={idx} className="flex items-start text-xs md:text-sm">
                       <span className="text-slate-400 font-mono w-[75px] md:w-[85px] flex-shrink-0">{log.time}</span>
                       <span className={`${idx === 0 ? 'text-blue-700 font-bold' : 'text-slate-600'}`}>{log.msg}</span>
                     </div>
                   ))}
                 </div>
              </div>
            </div>
          )}

          {view === 'input' && selectedGrade && (
            <div className="flex flex-col items-center w-full flex-grow p-3 md:p-6 animate-in fade-in">
              <div className="w-full max-w-5xl flex flex-col">
                <div className="w-full bg-blue-50/60 border border-blue-200 rounded-xl p-3 md:p-4 mb-4 flex items-start shadow-sm text-sm">
                  <Info className="w-5 h-5 text-blue-600 mr-2 md:mr-3 mt-0.5 flex-shrink-0" />
                  <div>
                    <h3 className="text-sm md:text-base font-bold text-blue-800 mb-1">【{selectedGrade}年級】成績輸入說明</h3>
                    <p className="text-xs md:text-sm text-blue-700 leading-relaxed font-medium">
                      1. 在手機上請「左右滑動表格」以輸入各科成績（最右側為自然）。<br />
                      2. 支援從 Excel 複製多筆資料（點選首格按 Ctrl+V 貼上），或點選「匯入 Excel」。
                    </p>
                  </div>
                </div>

                <div className="w-full flex flex-wrap justify-between items-center mb-4 gap-3">
                  <div className="flex gap-2 w-full sm:w-auto">
                    <input type="file" accept=".csv, .xlsx, .xls" style={{ display: 'none' }} ref={gridFileInputRef} onChange={handleGridFileUpload} />
                    <button onClick={() => gridFileInputRef.current.click()} className="flex-1 sm:flex-none flex items-center justify-center px-3 md:px-4 py-2 bg-emerald-50 hover:bg-emerald-100 text-emerald-800 border border-emerald-200 font-bold rounded-xl transition-colors text-xs md:text-sm shadow-sm">
                      <FileSpreadsheet className="w-4 h-4 mr-1.5" /> 匯入 Excel
                    </button>
                    <button onClick={handleClearGrid} className={`flex-1 sm:flex-none px-3 md:px-4 py-2 font-bold rounded-xl transition-colors text-xs md:text-sm flex items-center justify-center shadow-sm ${clearConfirm ? 'bg-red-600 text-white hover:bg-red-700 border-red-700' : 'bg-slate-50 text-slate-600 hover:bg-slate-100 border border-slate-200'}`}>
                      <Trash2 className="w-4 h-4 mr-1.5" /> {clearConfirm ? '確定清空？' : '清空表格'}
                    </button>
                  </div>
                  <button onClick={() => gridData.slice(1).some(r => r.some(c => c)) ? setView('result') : showMsg('error', '請先輸入成績資料')} className="w-full sm:w-auto px-6 py-2.5 bg-slate-800 hover:bg-slate-900 text-white font-bold rounded-xl transition-all shadow-md text-sm md:text-base flex items-center justify-center">
                    換算成績 <ChevronRight className="w-5 h-5 ml-1" />
                  </button>
                </div>

                {/* 響應式表格容器：解決家長反映的無法輸入自然成績問題 */}
                <div className="w-full border border-slate-300 rounded-xl shadow-inner bg-slate-50 flex flex-col h-[65vh] md:h-[500px]">
                  <div className="overflow-auto w-full h-full relative custom-scrollbar">
                    <table className="w-max min-w-full text-sm md:text-base border-collapse bg-white">
                      <thead className="sticky top-0 z-30 bg-slate-200 shadow-sm">
                        <tr>
                          <th className="sticky left-0 z-40 bg-slate-200 border-r border-b border-slate-300 text-slate-600 py-3 text-center font-bold min-w-[45px]">#</th>
                          {FIXED_HEADERS.map((header, cIdx) => (
                            <th 
                              key={`h-${cIdx}`} 
                              className={`p-0 border-r border-b border-slate-300 bg-slate-200 min-w-[70px] md:min-w-[100px]
                                ${header === '座號' ? 'sticky left-[45px] z-40 shadow-[1px_0_0_0_#cbd5e1]' : ''} 
                                ${header === '姓名' ? 'sticky left-[115px] md:left-[145px] z-40 shadow-[1px_0_0_0_#cbd5e1]' : ''} 
                              `}
                            >
                              <div className="w-full py-2 md:py-3 font-bold text-slate-800 text-center tracking-wider">
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
                              <td className="sticky left-0 z-20 bg-slate-100 border-r border-slate-300 text-slate-500 text-center font-mono text-xs md:text-sm py-2 font-medium">
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
                                  <td 
                                    key={`c-${cIdx}`} 
                                    className={`p-0 border-r border-slate-200 
                                      ${headerName === '座號' ? 'sticky left-[45px] z-20 bg-white shadow-[1px_0_0_0_#e2e8f0]' : ''}
                                      ${headerName === '姓名' ? 'sticky left-[115px] md:left-[145px] z-20 bg-white shadow-[1px_0_0_0_#e2e8f0]' : ''}
                                    `}
                                  >
                                    <input
                                      type="text"
                                      className={`w-full h-full py-2 md:py-2.5 px-2 md:px-3 focus:outline-none focus:bg-blue-100 transition-all min-w-[70px] md:min-w-[100px]
                                        ${isScoreCol ? 'text-right font-mono text-sm md:text-base' : 'text-center text-sm md:text-base'} 
                                        ${cell.trim() !== '' ? 'text-slate-800 font-bold' : 'text-slate-400'}
                                        ${isError ? 'bg-red-50 text-red-600 focus:bg-red-50' : 'bg-transparent'}`}
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
            <div className="w-full flex flex-col h-full flex-grow p-3 md:p-6 animate-in slide-in-from-bottom-4">
              <div className="w-full flex flex-col md:flex-row justify-between items-start md:items-center bg-teal-50 p-3 md:p-4 rounded-xl border border-teal-200 mb-4 md:mb-5 gap-3 shadow-sm">
                <h2 className="text-lg md:text-xl font-black text-teal-900 flex items-center flex-wrap gap-2">
                  <CheckCircle2 className="w-5 h-5 md:w-6 md:h-6 text-teal-600" />
                  {selectedGrade}年級 分析報表 
                  <span className="text-teal-700 text-xs md:text-base font-bold bg-teal-100 px-2 md:px-3 py-1 rounded-lg">共 {generateReportData.data?.length || 0} 筆</span>
                </h2>
                <div className="flex flex-wrap gap-2 w-full md:w-auto">
                  <button onClick={() => setView('input')} className="flex-1 md:flex-none px-3 md:px-5 py-2 bg-white text-slate-700 border border-slate-300 rounded-xl hover:bg-slate-50 transition-all text-xs md:text-sm font-bold flex justify-center items-center shadow-sm">
                    <ArrowLeft className="w-4 h-4 mr-1 md:mr-2" /> 修改
                  </button>
                  <button onClick={handleCopyReport} className="flex-1 md:flex-none px-3 md:px-5 py-2 bg-white text-blue-700 border border-blue-200 rounded-xl hover:bg-blue-50 transition-all text-xs md:text-sm font-bold flex justify-center items-center shadow-sm">
                    <ClipboardCopy className="w-4 h-4 mr-1 md:mr-2" /> 複製
                  </button>
                  <button onClick={() => exportReportToExcel(generateReportData, selectedGrade, showMsg)} className="flex-1 md:flex-none px-3 md:px-5 py-2 bg-blue-600 text-white rounded-xl hover:bg-blue-700 transition-all text-xs md:text-sm font-bold flex justify-center items-center shadow-md">
                    <FileDown className="w-4 h-4 mr-1 md:mr-2" /> 匯出
                  </button>
                </div>
              </div>

              {generateReportData.error ? (
                 <div className="w-full p-8 md:p-12 text-center text-red-600 font-bold text-base md:text-lg bg-red-50 rounded-xl border border-red-200">{generateReportData.error}</div>
              ) : (
                <div className="w-full flex-grow flex flex-col min-h-[300px]">
                  {generateReportData.unmappedSubjects && generateReportData.unmappedSubjects.length > 0 && (
                    <div className="w-full bg-amber-50 border border-amber-200 p-3 md:p-4 rounded-xl flex items-start text-xs md:text-sm mb-4 shadow-sm">
                      <AlertTriangle className="w-4 h-4 md:w-5 md:h-5 text-amber-500 mr-2 md:mr-3 flex-shrink-0 mt-0.5" />
                      <p className="text-amber-800 font-medium">未設定等級門檻的科目：<strong className="text-amber-900 mx-1">{generateReportData.unmappedSubjects.join('、')}</strong>。已自動計入平均，但無法於報表中標示 ABC 等級。</p>
                    </div>
                  )}

                  <div className="w-full border border-slate-200 rounded-xl overflow-x-auto shadow-sm bg-white custom-scrollbar flex-grow">
                    <table className="w-max min-w-full text-xs md:text-[15px] text-center whitespace-nowrap table-auto border-collapse">
                      <thead className="text-slate-600 bg-slate-100 font-black sticky top-0 shadow-sm border-b-2 border-slate-300">
                        <tr>
                          {Object.keys(generateReportData.data[0] || {}).map((header, idx) => (
                            <th key={idx} className="px-3 md:px-5 py-2.5 md:py-3.5 border-r border-slate-200 last:border-r-0 tracking-wide">
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
                                <td key={colIndex} className={`px-2 md:px-4 py-2 md:py-3 border-r border-slate-50 last:border-r-0 ${textClass}`}>
                                  {val}
                                </td>
                              );
                            })}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}
            </div>
          )}

          {view === 'admin_login' && (
             <div className="w-full p-8 md:p-16 flex flex-col items-center justify-center flex-grow animate-in zoom-in-95">
               <div className="bg-slate-100 p-5 md:p-6 rounded-full mb-6 shadow-inner border border-slate-200">
                 <Lock className="w-10 h-10 md:w-12 md:h-12 text-slate-500" />
               </div>
               <h2 className="text-xl md:text-2xl font-black text-slate-800 mb-3">管理員身分驗證</h2>
               <p className="text-slate-500 text-sm md:text-base mb-8 text-center font-medium">請輸入系統密碼以進入全校標準設定頁面</p>
               <form onSubmit={(e) => {
                  e.preventDefault();
                  if (adminPassword === '690530') {
                    setView('admin_settings'); setAdminPassword('');
                  } else { showMsg('error', '密碼錯誤'); }
               }} className="flex flex-col w-full max-w-sm space-y-4 px-4">
                 <input type="password" autoFocus placeholder="請輸入密碼..." className="w-full px-4 md:px-5 py-3 md:py-3.5 text-center text-base md:text-lg tracking-[0.2em] bg-white border-2 border-slate-300 rounded-xl focus:border-blue-500 focus:ring-4 focus:ring-blue-100 transition-all focus:outline-none" value={adminPassword} onChange={(e) => setAdminPassword(e.target.value)} />
                 <button type="submit" className="w-full py-3 md:py-3.5 bg-blue-600 hover:bg-blue-700 text-white font-black rounded-xl shadow-md flex justify-center items-center text-sm md:text-base transition-colors">
                   解鎖進入設定 <Unlock className="w-4 h-4 md:w-5 md:h-5 ml-2" />
                 </button>
               </form>
             </div>
          )}

          {view === 'admin_settings' && (
            <div className="w-full p-4 md:p-10 flex flex-col flex-grow animate-in fade-in overflow-y-auto">
              <div className="text-center pb-4 md:pb-6 mb-4 md:mb-6 border-b border-slate-200 flex flex-col items-center">
                <h2 className="text-xl md:text-2xl font-black text-slate-800 mb-4 flex items-center">
                   <Settings className="w-5 h-5 md:w-6 md:h-6 mr-2 text-slate-600" /> 全校標準與組距設定
                </h2>
                
                {/* 新增：將一鍵同步移至管理設定的首位 */}
                <button 
                  onClick={handleSyncAllGrades} 
                  disabled={isLoading}
                  className={`mb-6 px-6 py-2.5 bg-gradient-to-r from-amber-500 to-amber-600 hover:from-amber-600 hover:to-amber-700 text-white rounded-xl shadow-md font-black text-sm md:text-base flex items-center transition-all ${isLoading ? 'opacity-50 cursor-not-allowed' : ''}`}
                >
                  <DownloadCloud className={`w-5 h-5 mr-2 ${isLoading ? 'animate-bounce' : ''}`} /> 
                  一鍵批次同步所有年級雲端資料
                </button>

                <div className="flex bg-slate-100 p-1 md:p-1.5 rounded-xl space-x-1 md:space-x-2 shadow-inner border border-slate-200">
                  {['7', '8', '9'].map(grade => (
                    <button key={`admin-${grade}`} onClick={() => setSelectedGrade(grade)} className={`px-4 md:px-6 py-1.5 md:py-2 rounded-lg font-black text-sm md:text-base transition-all ${selectedGrade === grade ? 'bg-white text-blue-700 shadow-sm border border-slate-200' : 'text-slate-500 hover:text-slate-700'}`}>
                      {grade} 年級
                    </button>
                  ))}
                </div>
              </div>
              
              {!selectedGrade ? (
                <div className="w-full text-center py-16 text-slate-400 font-bold text-base md:text-lg flex flex-col items-center">
                  <ArrowLeft className="w-8 h-8 mb-3 text-slate-300 animate-pulse" />
                  請先在上方選擇欲設定的年級
                </div>
              ) : (
                <div className="w-full grid grid-cols-1 lg:grid-cols-2 gap-4 md:gap-6 flex-grow">
                  
                  <div className="flex flex-col bg-[#FFFCF5] p-4 md:p-5 rounded-2xl border border-amber-200 shadow-sm">
                    <div className="flex justify-between items-center mb-3">
                      <span className="text-sm md:text-base font-black text-amber-900 flex items-center"><BookOpen className="w-4 h-4 md:w-5 md:h-5 mr-2 text-amber-600"/> 各科等級門檻</span>
                      <div>
                        <input type="file" accept=".csv, .xlsx, .xls" style={{ display: 'none' }} ref={settingFileInputRef} onChange={(e) => handleAdminFileUpload(e, 'grade')} />
                        <button onClick={() => settingFileInputRef.current.click()} className="px-3 py-1.5 bg-white text-amber-800 border border-amber-300 rounded-lg text-xs font-bold shadow-sm hover:bg-amber-50 transition-colors flex items-center">
                          <Upload className="w-3.5 h-3.5 mr-1.5"/> 本機檔案
                        </button>
                      </div>
                    </div>
                    <div className="flex justify-between items-center mb-4 bg-amber-50 p-2 md:p-3 rounded-xl border border-amber-100 shadow-inner">
                       <span className="text-xs md:text-sm font-bold text-amber-800">從固定的雲端試算表同步</span>
                       <button onClick={() => fetchCloudData('grade')} className="px-3 md:px-4 py-1.5 md:py-2 bg-amber-600 text-white rounded-lg text-xs md:text-sm font-bold hover:bg-amber-700 whitespace-nowrap flex items-center shadow-sm transition-colors">
                         <Cloud className="w-4 h-4 mr-1.5" /> 雲端同步
                       </button>
                    </div>
                    <textarea className="w-full flex-grow min-h-[250px] md:min-h-[300px] p-3 md:p-4 border border-amber-300/80 rounded-xl text-xs md:text-sm font-mono text-slate-700 focus:outline-none focus:ring-2 focus:ring-amber-400 leading-relaxed shadow-inner custom-scrollbar" value={appSettings[selectedGrade].grade} onChange={(e) => setAppSettings(prev => ({...prev, [selectedGrade]: {...prev[selectedGrade], grade: e.target.value}}))} />
                  </div>

                  <div className="flex flex-col bg-[#F5F9FF] p-4 md:p-5 rounded-2xl border border-blue-200 shadow-sm">
                    <div className="flex justify-between items-center mb-3">
                      <span className="text-sm md:text-base font-black text-blue-900 flex items-center"><Users className="w-4 h-4 md:w-5 md:h-5 mr-2 text-blue-600"/> 全校分數組距</span>
                      <div>
                        <input type="file" accept=".csv, .xlsx, .xls" style={{ display: 'none' }} ref={distFileInputRef} onChange={(e) => handleAdminFileUpload(e, 'dist')} />
                        <button onClick={() => distFileInputRef.current.click()} className="px-3 py-1.5 bg-white text-blue-800 border border-blue-300 rounded-lg text-xs font-bold shadow-sm hover:bg-blue-50 transition-colors flex items-center">
                          <Upload className="w-3.5 h-3.5 mr-1.5"/> 本機檔案
                        </button>
                      </div>
                    </div>
                    <div className="flex justify-between items-center mb-4 bg-blue-50 p-2 md:p-3 rounded-xl border border-blue-100 shadow-inner">
                       <span className="text-xs md:text-sm font-bold text-blue-800">從固定的雲端試算表同步</span>
                       <button onClick={() => fetchCloudData('dist')} className="px-3 md:px-4 py-1.5 md:py-2 bg-blue-600 text-white rounded-lg text-xs md:text-sm font-bold hover:bg-blue-700 whitespace-nowrap flex items-center shadow-sm transition-colors">
                         <Cloud className="w-4 h-4 mr-1.5" /> 雲端同步
                       </button>
                    </div>
                    <textarea className="w-full flex-grow min-h-[250px] md:min-h-[300px] p-3 md:p-4 border border-blue-300/80 rounded-xl text-xs md:text-sm font-mono text-slate-700 focus:outline-none focus:ring-2 focus:ring-blue-400 leading-relaxed shadow-inner custom-scrollbar" value={appSettings[selectedGrade].dist} onChange={(e) => setAppSettings(prev => ({...prev, [selectedGrade]: {...prev[selectedGrade], dist: e.target.value}}))} />
                  </div>

                </div>
              )}
            </div>
          )}
        </main>

        <footer className="flex flex-col gap-3 md:gap-4 py-6 md:py-8">
           <div className="bg-amber-50/50 border border-amber-200/50 rounded-2xl md:rounded-3xl p-4 md:p-6 flex items-start shadow-sm mx-auto w-full max-w-[850px]">
              <AlertTriangle className="w-5 h-5 md:w-6 md:h-6 text-amber-500 mr-3 md:mr-4 mt-0.5 shrink-0" />
              <div>
                <h4 className="text-amber-900 font-black text-xs md:text-sm mb-1">系統免責聲明</h4>
                <p className="text-[10px] md:text-xs text-amber-800/70 leading-relaxed font-medium">
                  本工具僅供教師進行快速成績換算與預估排名參考，並非學校官方正式成績系統。所有計算結果請務必以教務處公告之正式紙本成績單為準。
                </p>
              </div>
           </div>
           <div className="text-center mt-2">
             <div className="text-slate-300 text-[10px] md:text-xs font-bold tracking-widest">程式設計：蘇老爹</div>
           </div>
        </footer>
      </div>
    </div>
  );
}