import React, { useState, useMemo, useRef, useEffect } from 'react';
import { 
  Upload, FileDown, ClipboardCopy, Calculator, CheckCircle2, 
  Settings, ArrowLeft, Lock, Unlock, AlertTriangle, Users, 
  BookOpen, ChevronRight, FileSpreadsheet, Trash2, Home, Info 
} from 'lucide-react';

// --- 1. 核心商業邏輯與工具函式 ---
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
        if (distMap[i].count === 1) return `${distMap[i].startRank}`;
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

// --- 2. 預設資料設定 ---
const defaultSettings = `,等級,國文,英文,數學,社會,自然\n,A++,92,100,94,94,96\n,A+,89,98,89,88,92\n,A,84,95,79,80,86\n,B++,80,92,70,72,78\n,B+,74,88,62,64,68\n,B,52,50,28,34,36`;
const defaultDistribution = `分數組距,全校人數,累計人數\n100,0,0\n98-99.99,0,0\n96-97.99,23,23\n94-95.99,52,75\n92-93.99,76,151\n90-91.99,73,224\n87-90.99,102,326\n84-86.99,75,401\n80-83.99,117,518\n70-79.99,174,692\n60-69.99,127,819\n0-59.99,25,844`;

const INITIAL_GRID_ROWS = 40;
const FIXED_HEADERS = ['座號', '姓名', '國文', '英文', '數學', '社會', '自然'];
const INITIAL_GRID_COLS = FIXED_HEADERS.length;

// 精確定義各欄位像素寬度 (加總為 750px)
const COL_WIDTHS = {
  '#': 'w-[50px]',
  '座號': 'w-[80px]',
  '姓名': 'w-[120px]',
  '國文': 'w-[100px]',
  '英文': 'w-[100px]',
  '數學': 'w-[100px]',
  '社會': 'w-[100px]',
  '自然': 'w-[100px]',
};

const createInitialGrid = () => {
  const grid = Array(INITIAL_GRID_ROWS).fill(0).map(() => Array(INITIAL_GRID_COLS).fill(''));
  grid[0] = [...FIXED_HEADERS];
  return grid;
};

// --- 3. 主應用程式元件 ---
export default function App() {
  const [view, setView] = useState('home'); 
  const [selectedGrade, setSelectedGrade] = useState(null); 
  const [error, setError] = useState(null);
  const [isAdminAuth, setIsAdminAuth] = useState(false);
  const [adminPassword, setAdminPassword] = useState('');
  
  const [appSettings, setAppSettings] = useState({
    '7': { grade: defaultSettings, dist: defaultDistribution },
    '8': { grade: defaultSettings, dist: defaultDistribution },
    '9': { grade: defaultSettings, dist: defaultDistribution }
  });

  const [gridData, setGridData] = useState(createInitialGrid());
  
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

  const currentParsedSettings = useMemo(() => {
    if (!selectedGrade) return { settings: {}, subjects: [] };
    return processSettings(parseCSV(appSettings[selectedGrade].grade));
  }, [appSettings, selectedGrade]);

  const currentDistMap = useMemo(() => {
    if (!selectedGrade) return [];
    return processDistribution(parseCSV(appSettings[selectedGrade].dist));
  }, [appSettings, selectedGrade]);

  // 表格變更與貼上處理
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

    rows.forEach((rowStr, i) => {
      if (!rowStr.trim() && i === rows.length - 1) return; 
      const cells = rowStr.split(/\t|,/);
      const targetRow = startRow + i;
      
      if (targetRow >= newGrid.length) {
        newGrid.push(Array(INITIAL_GRID_COLS).fill(''));
      }
      
      newGrid[targetRow] = [...newGrid[targetRow]];
      cells.forEach((cellVal, j) => {
        const targetCol = startCol + j;
        if (targetCol < INITIAL_GRID_COLS) { 
          newGrid[targetRow][targetCol] = cellVal.trim().replace(/^"|"$/g, '');
        }
      });
    });
    setGridData(newGrid);
  };

  // 匯入 Excel
  const handleGridFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    if (!window.XLSX) { setError("Excel 模組載入中，請稍後再試。"); return; }
    
    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const workbook = window.XLSX.read(new Uint8Array(event.target.result), { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const arr = window.XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });
        
        let headerRowIndex = -1;
        for (let i = 0; i < Math.min(arr.length, 10); i++) {
           const rowStr = arr[i].join('').toLowerCase();
           if (rowStr.includes('座號') && rowStr.includes('姓名')) {
              headerRowIndex = i; break;
           }
        }

        const dataStartRow = headerRowIndex !== -1 ? headerRowIndex + 1 : 1;
        const rowsCount = Math.max(INITIAL_GRID_ROWS, arr.length - dataStartRow + 1);
        
        const newGrid = Array(rowsCount).fill(0).map(() => Array(INITIAL_GRID_COLS).fill(''));
        newGrid[0] = [...FIXED_HEADERS];

        for (let i = dataStartRow; i < arr.length; i++) {
           const sourceRow = arr[i];
           const targetRowIdx = i - dataStartRow + 1;
           if (!sourceRow) continue;
           for(let j = 0; j < Math.min(sourceRow.length, INITIAL_GRID_COLS); j++) {
               newGrid[targetRowIdx][j] = sourceRow[j] !== undefined ? String(sourceRow[j]) : '';
           }
        }
        setGridData(newGrid);
        setError(null);
      } catch (err) {
        setError("讀取 Excel 失敗，請確認檔案格式。");
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  const handleClearGrid = () => {
    if(window.confirm("確定要清空目前表格內的所有資料嗎？")) {
      setGridData(createInitialGrid());
      setError(null);
    }
  };

  const handleAdminFileUpload = (e, type) => {
    const file = e.target.files[0];
    if (!file) return;
    const fileExt = file.name.split('.').pop().toLowerCase();
    const updateSetting = (text) => {
      setAppSettings(prev => ({...prev, [selectedGrade]: { ...prev[selectedGrade], [type]: text }}));
      setError(null);
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
        } catch (err) {}
      };
      reader.readAsArrayBuffer(file);
    }
    e.target.value = '';
  };

  const handleAdminLogin = (e) => {
    e.preventDefault();
    if (adminPassword === '690530') {
      setIsAdminAuth(true); setView('admin_settings'); setAdminPassword(''); setError(null);
    } else { setError("密碼錯誤。"); }
  };

  const goHome = () => {
    setView('home');
    setSelectedGrade(null);
    setError(null);
  };

  // 產生報表
  const generateReportData = useMemo(() => {
    if (view !== 'result') return null;
    const headers = gridData[0];
    const excludeCols = ['座號', '姓名'];
    const subjects = headers.filter(h => h && !excludeCols.includes(h));
    
    const rawScoresData = gridData.slice(1).map(row => {
      const obj = {};
      let hasData = false;
      headers.forEach((h, i) => {
        if (h) {
          obj[h] = row[i];
          if (row[i] && row[i].trim() !== '') hasData = true;
        }
      });
      return hasData ? obj : null;
    }).filter(Boolean);

    if (rawScoresData.length === 0) return { error: "沒有可計算的資料。" };

    const unmappedSubjects = subjects.filter(sub => !currentParsedSettings.settings[sub]);
    let calculatedData = rawScoresData.map(student => {
      let totalScore = 0; let validCount = 0;
      const resultRow = { '座號': student['座號'] || '', '姓名': student['姓名'] || '' };

      subjects.forEach(subject => {
        const val = student[subject];
        if (val !== undefined && val !== null && val.toString().trim() !== '') {
          const numScore = parseFloat(val);
          if (!isNaN(numScore)) {
            totalScore += numScore; validCount++;
            if (currentParsedSettings.settings[subject]) {
              resultRow[subject] = `${numScore} (${getGradeLevel(numScore, currentParsedSettings.settings[subject])})`;
            } else { resultRow[subject] = numScore; }
          } else { resultRow[subject] = val; }
        } else { resultRow[subject] = ''; }
      });

      resultRow['平均'] = validCount > 0 ? parseFloat((totalScore / validCount).toFixed(1)) : 0;
      return resultRow;
    });

    calculatedData.sort((a, b) => b['平均'] - a['平均']);
    calculatedData.forEach((student, index) => {
      student['班排'] = (index > 0 && student['平均'] === calculatedData[index - 1]['平均']) ? calculatedData[index - 1]['班排'] : index + 1;
      const schoolRank = getSchoolRank(student['平均'], currentDistMap);
      if (schoolRank) student['預估校排'] = schoolRank;
    });

    calculatedData.sort((a, b) => {
        const seatA = parseInt(a['座號'], 10) || 999;
        const seatB = parseInt(b['座號'], 10) || 999;
        return seatA - seatB;
    });

    return { data: calculatedData, subjects, unmappedSubjects };
  }, [view, gridData, currentParsedSettings, currentDistMap]);

  const handleGenerate = () => {
    const hasData = gridData.slice(1).some(row => row.some(cell => cell && cell.trim() !== ''));
    if (hasData) { setView('result'); setError(null); } 
    else { setError("請先在表格內輸入成績資料。"); }
  };

  const handleCopyReport = () => {
    if (!generateReportData?.data) return;
    const tsvContent = [
      Object.keys(generateReportData.data[0]).join('\t'), 
      ...generateReportData.data.map(row => Object.values(row).join('\t'))
    ].join('\n');
    const textArea = document.createElement("textarea");
    textArea.value = tsvContent; document.body.appendChild(textArea); textArea.select();
    try { document.execCommand('copy'); alert("✅ 已成功複製報表資料！"); } catch (err) { alert("複製失敗，請手動選取。"); }
    document.body.removeChild(textArea);
  };

  // --- 畫面渲染 ---
  return (
    <div className="min-h-screen bg-[#F4F7F9] text-slate-800 p-4 md:p-6 font-sans flex flex-col items-center">
      {/* 限制最大寬度為 900px，居中顯示 */}
      <div className="w-full max-w-[900px] flex flex-col flex-grow space-y-4">
        
        {/* Header */}
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

        {error && (
          <div className="w-full bg-red-50 text-red-600 px-5 py-4 rounded-xl flex items-center border border-red-200 text-base font-bold shadow-sm">
            <AlertTriangle className="w-5 h-5 mr-3 flex-shrink-0" /> {error}
          </div>
        )}

        {/* 主內容區塊 */}
        <main className="w-full bg-white rounded-2xl shadow-sm border border-slate-100 flex flex-col flex-grow relative overflow-hidden items-center">
          
          {/* 首頁 */}
          {view === 'home' && (
            <div className="p-8 md:p-16 flex flex-col items-center justify-center flex-grow animate-in zoom-in-95 w-full">
              <div className="w-20 h-20 bg-blue-50 text-blue-600 rounded-full flex items-center justify-center mb-6 shadow-inner">
                <BookOpen className="w-10 h-10" />
              </div>
              <h2 className="text-3xl font-black text-slate-800 mb-3">請選擇欲輸入的年級</h2>
              <p className="text-slate-500 text-base mb-10 text-center font-medium">系統將自動載入該年級對應的全校組距標準與等級門檻</p>
              
              <div className="flex flex-wrap justify-center gap-6 w-full">
                {['7', '8', '9'].map(grade => (
                  <button key={grade} onClick={() => { setSelectedGrade(grade); setView('input'); setError(null); }}
                    className="w-32 h-40 group flex flex-col items-center justify-center p-4 bg-white border-2 border-slate-200 hover:border-blue-500 hover:shadow-lg hover:bg-blue-50/50 rounded-2xl transition-all"
                  >
                    <span className="text-5xl font-black text-slate-300 group-hover:text-blue-600 mb-2 transition-colors">{grade}</span>
                    <span className="text-base font-bold text-slate-500 group-hover:text-blue-800">年級專區</span>
                  </button>
                ))}
              </div>
            </div>
          )}

          {/* 成績輸入區 */}
          {view === 'input' && selectedGrade && (
            <div className="flex flex-col items-center w-full flex-grow p-6 animate-in fade-in">
              
              {/* 精準寬度包裹層 (750px)，強制所有元素在此範圍內對齊 */}
              <div className="w-[750px] flex flex-col">
                
                {/* 操作說明區塊 */}
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

                {/* 操作工具列 */}
                <div className="w-full flex justify-between items-center mb-4">
                  <div className="flex gap-3">
                    <input type="file" accept=".csv, .xlsx, .xls" style={{ display: 'none' }} ref={gridFileInputRef} onChange={handleGridFileUpload} />
                    <button onClick={() => gridFileInputRef.current.click()} className="flex items-center justify-center px-4 py-2 bg-emerald-50 hover:bg-emerald-100 text-emerald-800 border border-emerald-200 font-bold rounded-xl transition-colors text-sm shadow-sm">
                      <FileSpreadsheet className="w-4 h-4 mr-2" /> 匯入 Excel
                    </button>
                    <button onClick={handleClearGrid} className="px-4 py-2 bg-slate-50 hover:bg-slate-100 text-slate-600 border border-slate-200 font-bold rounded-xl transition-colors text-sm flex items-center justify-center shadow-sm">
                      <Trash2 className="w-4 h-4 mr-2" /> 清空表格
                    </button>
                  </div>
                  <button onClick={handleGenerate} className="px-6 py-2.5 bg-slate-800 hover:bg-slate-900 text-white font-bold rounded-xl transition-all shadow-md text-base flex items-center justify-center">
                    產生報表 <ChevronRight className="w-5 h-5 ml-1" />
                  </button>
                </div>

                {/* 固定像素表格容器 */}
                <div className="w-full border border-slate-300 rounded-xl overflow-hidden shadow-inner bg-slate-50 flex flex-col h-[400px]">
                  <div className="overflow-y-auto w-full custom-scrollbar">
                    {/* table-fixed 搭配寫死的 w-[750px] 確保不會被撐開 */}
                    <table className="w-[750px] text-base border-collapse table-fixed">
                      <thead className="sticky top-0 z-20 bg-slate-200 shadow-sm">
                        <tr>
                          <th className={`${COL_WIDTHS['#']} border-r border-slate-300 text-slate-600 py-3 text-center font-bold`}>#</th>
                          {FIXED_HEADERS.map((header, cIdx) => (
                            <th key={`h-${cIdx}`} className={`${COL_WIDTHS[header]} p-0 border-r border-slate-300 last:border-r-0 bg-slate-200`}>
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
                              <td className={`${COL_WIDTHS['#']} bg-slate-100 border-r border-slate-200 text-slate-500 text-center font-mono text-sm py-2 font-medium`}>
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
                                  <td key={`c-${cIdx}`} className={`${COL_WIDTHS[headerName]} p-0 border-r border-slate-200 last:border-r-0`}>
                                    <input
                                      type="text"
                                      className={`w-full h-full py-2.5 px-3 focus:outline-none focus:bg-blue-100 transition-all 
                                        ${isScoreCol ? 'text-right font-mono text-base' : 'text-center text-base'} 
                                        ${cell.trim() !== '' ? 'text-slate-800 font-bold' : 'text-slate-400'}
                                        ${isError ? 'bg-red-50 text-red-600 focus:bg-red-50' : ''}`}
                                      value={cell}
                                      placeholder={isScoreCol ? "" : ""}
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

          {/* 結果報表區 */}
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
                  {generateReportData.unmappedSubjects.length > 0 && (
                    <div className="w-full bg-amber-50 border border-amber-200 p-4 rounded-xl flex items-start text-sm mb-4 shadow-sm">
                      <AlertTriangle className="w-5 h-5 text-amber-500 mr-3 flex-shrink-0 mt-0.5" />
                      <p className="text-amber-800 font-medium">未設定等級門檻的科目：<strong className="text-amber-900 mx-1">{generateReportData.unmappedSubjects.join('、')}</strong>。已自動計入平均，但無法於報表中標示 ABC 等級。</p>
                    </div>
                  )}

                  {/* 報表欄位多，可能超出 900px，因此加上 overflow-x-auto */}
                  <div className="w-full border border-slate-200 rounded-xl overflow-x-auto shadow-sm bg-white custom-scrollbar flex-grow">
                    <table className="w-max min-w-full text-[15px] text-left whitespace-nowrap table-auto">
                      <thead className="text-slate-600 bg-slate-100 font-black sticky top-0 shadow-sm border-b border-slate-200">
                        <tr>
                          {Object.keys(generateReportData.data[0]).map((header, idx) => (
                            <th key={idx} className="px-5 py-3.5 border-r border-slate-200 last:border-r-0 truncate tracking-wide">
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
                              if (key === '預估校排' || key === '班排') textClass = "text-blue-700 font-black";

                              return (
                                <td key={colIndex} className={`px-5 py-3 border-r border-slate-50 last:border-r-0 truncate ${textClass}`}>
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

          {/* 管理員登入區 */}
          {view === 'admin_login' && (
             <div className="w-full p-8 md:p-16 flex flex-col items-center justify-center flex-grow animate-in zoom-in-95">
               <div className="bg-slate-100 p-6 rounded-full mb-6 shadow-inner border border-slate-200">
                 <Lock className="w-12 h-12 text-slate-500" />
               </div>
               <h2 className="text-2xl font-black text-slate-800 mb-3">管理員身分驗證</h2>
               <p className="text-slate-500 text-base mb-8 text-center font-medium">請輸入系統密碼以進入全校標準設定頁面</p>
               <form onSubmit={handleAdminLogin} className="flex flex-col w-full max-w-sm space-y-4">
                 <input type="password" autoFocus placeholder="請輸入密碼..." className="w-full px-5 py-3.5 text-center text-lg tracking-[0.2em] bg-white border-2 border-slate-300 rounded-xl focus:border-blue-500 focus:ring-4 focus:ring-blue-100 transition-all focus:outline-none" value={adminPassword} onChange={(e) => setAdminPassword(e.target.value)} />
                 <button type="submit" className="w-full py-3.5 bg-blue-600 hover:bg-blue-700 text-white font-black rounded-xl shadow-md flex justify-center items-center text-base transition-colors">
                   解鎖進入設定 <Unlock className="w-5 h-5 ml-2" />
                 </button>
               </form>
             </div>
          )}

          {/* 管理員設定區 */}
          {view === 'admin_settings' && (
            <div className="w-full p-5 md:p-8 flex flex-col flex-grow animate-in fade-in overflow-y-auto">
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
                    <div className="flex justify-between items-center mb-4">
                      <span className="text-base font-black text-amber-900 flex items-center"><BookOpen className="w-5 h-5 mr-2 text-amber-600"/> 各科等級門檻 (CSV 格式)</span>
                      <div>
                        <input type="file" accept=".csv, .xlsx, .xls" style={{ display: 'none' }} ref={settingFileInputRef} onChange={(e) => handleAdminFileUpload(e, 'grade')} />
                        <button onClick={() => settingFileInputRef.current.click()} className="px-4 py-2 bg-white text-amber-800 border border-amber-300 rounded-lg text-sm font-bold shadow-sm hover:bg-amber-50 transition-colors flex items-center">
                          <Upload className="w-4 h-4 mr-1.5"/> 匯入設定
                        </button>
                      </div>
                    </div>
                    <textarea className="w-full flex-grow min-h-[300px] p-4 border border-amber-300/80 rounded-xl text-sm font-mono text-slate-700 focus:outline-none focus:ring-2 focus:ring-amber-400 leading-relaxed shadow-inner" value={appSettings[selectedGrade].grade} onChange={(e) => setAppSettings(prev => ({...prev, [selectedGrade]: {...prev[selectedGrade], grade: e.target.value}}))} />
                  </div>
                  <div className="flex flex-col bg-[#F5F9FF] p-5 rounded-2xl border border-blue-200 shadow-sm">
                    <div className="flex justify-between items-center mb-4">
                      <span className="text-base font-black text-blue-900 flex items-center"><Users className="w-5 h-5 mr-2 text-blue-600"/> 全校分數組距 (CSV 格式)</span>
                      <div>
                        <input type="file" accept=".csv, .xlsx, .xls" style={{ display: 'none' }} ref={distFileInputRef} onChange={(e) => handleAdminFileUpload(e, 'dist')} />
                        <button onClick={() => distFileInputRef.current.click()} className="px-4 py-2 bg-white text-blue-800 border border-blue-300 rounded-lg text-sm font-bold shadow-sm hover:bg-blue-50 transition-colors flex items-center">
                          <Upload className="w-4 h-4 mr-1.5"/> 匯入設定
                        </button>
                      </div>
                    </div>
                    <textarea className="w-full flex-grow min-h-[300px] p-4 border border-blue-300/80 rounded-xl text-sm font-mono text-slate-700 focus:outline-none focus:ring-2 focus:ring-blue-400 leading-relaxed shadow-inner" value={appSettings[selectedGrade].dist} onChange={(e) => setAppSettings(prev => ({...prev, [selectedGrade]: {...prev[selectedGrade], dist: e.target.value}}))} />
                  </div>
                </div>
              )}
            </div>
          )}
        </main>
        
        {/* Footer 免責聲明區塊 */}
        <footer className="w-full mt-2 space-y-3 pb-4">
          <div className="w-full bg-amber-50/80 border border-amber-200 rounded-xl p-3 flex items-start md:items-center text-amber-800 text-sm font-medium shadow-sm">
            <AlertTriangle className="w-5 h-5 mr-3 flex-shrink-0 text-amber-500" />
            <p className="leading-relaxed"><strong>免責說明：</strong>本程式為教育人員自行設計之快速換算輔助工具，計算結果包含線性插值之「預估」排名，並非學校官方正式成績系統。若對成績或排名有疑問，請依據學校教務處公告為準。</p>
          </div>
          <div className="w-full text-center text-slate-400 text-sm font-bold py-2 tracking-wide">
             由 蘇老爹 開發設計
          </div>
        </footer>

      </div>
    </div>
  );
}