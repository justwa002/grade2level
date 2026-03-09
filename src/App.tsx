import React, { useState, useMemo, useRef, useEffect } from 'react';
import { Upload, FileDown, ClipboardCopy, Calculator, CheckCircle2, RefreshCw, Settings, ArrowLeft, Lock, Unlock, AlertTriangle, Users, BookOpen, ChevronRight, FileSpreadsheet, Trash2 } from 'lucide-react';

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
  const link = document.createElement('url');
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url);
};

// --- 2. 預設資料設定 ---
const defaultSettings = `,等級,國文,英文,數學,社會,自然\n,A++,92,100,94,94,96\n,A+,89,98,89,88,92\n,A,84,95,79,80,86\n,B++,80,92,70,72,78\n,B+,74,88,62,64,68\n,B,52,50,28,34,36`;
const defaultDistribution = `分數組距,全校人數,累計人數\n100,0,0\n98-99.99,0,0\n96-97.99,23,23\n94-95.99,52,75\n92-93.99,76,151\n90-91.99,73,224\n87-90.99,102,326\n84-86.99,75,401\n80-83.99,117,518\n70-79.99,174,692\n60-69.99,127,819\n0-59.99,25,844`;

const INITIAL_GRID_ROWS = 40;
const INITIAL_GRID_COLS = 8; // 縮減為 8 欄，符合多數需求且能在單頁顯示
const defaultHeaders = ['座號', '姓名', '國文', '英文', '數學', '社會', '自然', ''];

const createInitialGrid = () => {
  const grid = Array(INITIAL_GRID_ROWS).fill(0).map(() => Array(INITIAL_GRID_COLS).fill(''));
  grid[0] = [...defaultHeaders];
  return grid;
};

// --- 3. 主應用程式元件 ---
export default function App() {
  const [view, setView] = useState('home'); // 'home' | 'input' | 'result' | 'admin_login' | 'admin_settings'
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
    const newGrid = [...gridData];
    newGrid[rIdx] = [...newGrid[rIdx]];
    newGrid[rIdx][cIdx] = value;
    setGridData(newGrid);
  };

  const handleGridPaste = (e, startRow, startCol) => {
    e.preventDefault();
    const pasteText = e.clipboardData.getData('text');
    if (!pasteText) return;
    const rows = pasteText.split(/\r?\n/);
    const newGrid = [...gridData];

    rows.forEach((rowStr, i) => {
      if (!rowStr.trim() && i === rows.length - 1) return; 
      const cells = rowStr.split(/\t|,/);
      const targetRow = startRow + i;
      
      // 動態擴展列數
      if (targetRow >= newGrid.length) {
        newGrid.push(Array(newGrid[0].length).fill(''));
      }
      
      newGrid[targetRow] = [...newGrid[targetRow]];
      cells.forEach((cellVal, j) => {
        const targetCol = startCol + j;
        // 動態擴展欄數
        if (targetCol >= newGrid[targetRow].length) {
           newGrid.forEach(r => r.push(''));
        }
        newGrid[targetRow][targetCol] = cellVal.trim().replace(/^"|"$/g, '');
      });
    });
    setGridData(newGrid);
  };

  // 匯入 Excel 到即時表格
  const handleGridFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    if (!window.XLSX) { setError("Excel 模組載入中，請稍後再試。"); return; }
    
    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const workbook = window.XLSX.read(new Uint8Array(event.target.result), { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        // 轉換為二維陣列
        const arr = window.XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });
        
        // 確保符合最小網格大小
        const rowsCount = Math.max(INITIAL_GRID_ROWS, arr.length);
        const colsCount = Math.max(INITIAL_GRID_COLS, arr.length > 0 ? arr[0].length : INITIAL_GRID_COLS);
        
        const newGrid = Array(rowsCount).fill(0).map((_, rIdx) => {
          const row = arr[rIdx] || [];
          return Array(colsCount).fill('').map((_, cIdx) => {
             return row[cIdx] !== undefined ? String(row[cIdx]) : '';
          });
        });

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

  // 管理員設定檔案上傳與登入
  const handleAdminFileUpload = (e, type) => {
    // ... [原有邏輯保留] ...
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
          if (row[i].trim() !== '') hasData = true;
        }
      });
      return hasData ? obj : null;
    }).filter(Boolean);

    if (rawScoresData.length === 0) return { error: "沒有可計算的資料。" };

    const unmappedSubjects = subjects.filter(sub => !currentParsedSettings.settings[sub]);
    const calculatedData = rawScoresData.map(student => {
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

    return { data: calculatedData, subjects, unmappedSubjects };
  }, [view, gridData, currentParsedSettings, currentDistMap]);

  const handleGenerate = () => {
    const hasData = gridData.slice(1).some(row => row.some(cell => cell.trim() !== ''));
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
    try { document.execCommand('copy'); alert("✅ 已成功複製！"); } catch (err) { alert("複製失敗。"); }
    document.body.removeChild(textArea);
  };

  // --- 畫面渲染 ---
  return (
    <div className="min-h-screen bg-[#F4F7F9] text-slate-800 p-2 md:p-6 font-sans selection:bg-blue-100 flex flex-col">
      <div className="max-w-[1200px] mx-auto w-full space-y-4 flex-grow flex flex-col">
        
        {/* 全域免責聲明 */}
        <div className="bg-amber-50 border border-amber-200/60 rounded-xl p-3 flex items-start md:items-center text-amber-800 text-xs md:text-sm font-medium shadow-sm">
          <AlertTriangle className="w-5 h-5 mr-3 flex-shrink-0 text-amber-500" />
          <p><strong>免責說明：</strong>本程式為老師自行設計之快速換算輔助工具，並非學校官方正式成績系統。若對成績或排名有疑問，請洽詢學校教務處。</p>
        </div>

        {/* 極簡 Header */}
        <header className="flex justify-between items-center bg-white p-3 md:p-4 rounded-2xl shadow-sm border border-slate-100">
          <div className="flex items-center space-x-3 cursor-pointer" onClick={() => { setView('home'); setSelectedGrade(null); }}>
            <div className="bg-blue-500 p-2 md:p-2.5 rounded-xl text-white shadow-sm">
              <Calculator className="w-5 h-5" />
            </div>
            <div>
              <h1 className="text-lg md:text-xl font-extrabold text-slate-800 tracking-tight">成績等級產生器</h1>
            </div>
          </div>
          <div className="flex items-center gap-2">
            {view !== 'home' && view !== 'admin_login' && view !== 'admin_settings' && (
              <button onClick={() => setView('admin_login')} className="text-xs md:text-sm font-semibold text-slate-500 hover:text-slate-800 px-2 md:px-3 py-1.5 flex items-center bg-slate-100 rounded-lg transition-colors">
                <Settings className="w-3.5 h-3.5 mr-1.5" /> 管理員設定
              </button>
            )}
            {view.includes('admin') && (
               <button onClick={() => setView('home')} className="text-xs md:text-sm font-semibold text-slate-600 hover:text-slate-800 px-3 py-1.5 flex items-center bg-slate-100 rounded-lg">
                 返回
               </button>
            )}
          </div>
        </header>

        {error && (
          <div className="bg-red-50 text-red-600 px-4 py-3 rounded-xl flex items-center border border-red-100 text-sm font-medium">
            <AlertTriangle className="w-4 h-4 mr-2 flex-shrink-0" /> {error}
          </div>
        )}

        {/* 主內容區塊 (動態填滿剩餘空間) */}
        <main className="bg-white rounded-2xl shadow-sm border border-slate-100 flex-grow flex flex-col relative overflow-hidden">
          
          {/* 介面一：首頁 (年級選擇) */}
          {view === 'home' && (
            <div className="p-6 md:p-12 flex flex-col items-center justify-center flex-grow animate-in zoom-in-95">
              <div className="w-14 h-14 bg-blue-50 text-blue-500 rounded-full flex items-center justify-center mb-4">
                <BookOpen className="w-7 h-7" />
              </div>
              <h2 className="text-2xl font-extrabold text-slate-800 mb-2">請選擇年級</h2>
              <p className="text-slate-500 text-sm mb-8 text-center">選擇年級以載入專屬全校標準與等級門檻</p>
              
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4 w-full max-w-2xl">
                {['7', '8', '9'].map(grade => (
                  <button key={grade} onClick={() => { setSelectedGrade(grade); setView('input'); setError(null); }}
                    className="group flex flex-col items-center justify-center p-6 bg-white border-2 border-slate-100 hover:border-blue-400 hover:shadow-md hover:bg-blue-50/30 rounded-2xl transition-all"
                  >
                    <span className="text-3xl font-black text-slate-300 group-hover:text-blue-500 mb-1">{grade}</span>
                    <span className="text-sm font-bold text-slate-600 group-hover:text-blue-700">年級專區</span>
                  </button>
                ))}
              </div>
            </div>
          )}

          {/* 介面二：極簡輸入區 (教師操作) */}
          {view === 'input' && selectedGrade && (
            <div className="flex flex-col h-full flex-grow p-4 md:p-5 animate-in fade-in">
              
              {/* 操作工具列 */}
              <div className="flex flex-col md:flex-row justify-between items-start md:items-center mb-3 gap-3">
                <div className="flex items-center">
                  <span className="px-3 py-1 bg-blue-100 text-blue-800 font-bold rounded-lg text-sm mr-3">
                    {selectedGrade} 年級
                  </span>
                  <span className="text-sm font-bold text-slate-600">填寫表格，或點擊右側匯入</span>
                </div>
                
                <div className="flex flex-wrap items-center gap-2 w-full md:w-auto">
                  <input type="file" accept=".csv, .xlsx, .xls" style={{ display: 'none' }} ref={gridFileInputRef} onChange={handleGridFileUpload} />
                  <button onClick={() => gridFileInputRef.current.click()} className="flex-1 md:flex-none flex items-center justify-center px-4 py-2 bg-emerald-50 hover:bg-emerald-100 text-emerald-700 border border-emerald-200 font-bold rounded-xl transition-colors text-sm">
                    <FileSpreadsheet className="w-4 h-4 mr-1.5" /> 匯入 Excel
                  </button>
                  <button onClick={handleClearGrid} className="px-3 py-2 bg-slate-50 hover:bg-slate-100 text-slate-500 border border-slate-200 rounded-xl transition-colors text-sm flex items-center justify-center" title="清空表格">
                    <Trash2 className="w-4 h-4" />
                  </button>
                  <button onClick={handleGenerate} className="flex-1 md:flex-none px-6 py-2 bg-slate-800 hover:bg-slate-900 text-white font-bold rounded-xl transition-all shadow-md text-sm flex items-center justify-center">
                    產生等級報表 <ChevronRight className="w-4 h-4 ml-1" />
                  </button>
                </div>
              </div>

              {/* 緊湊型即時表格 (Compact Grid) */}
              <div className="border border-slate-200 rounded-xl overflow-hidden shadow-inner flex-grow bg-slate-50 flex flex-col min-h-[400px]">
                <div className="overflow-auto flex-grow custom-scrollbar">
                  <table className="w-full text-[13px] border-collapse min-w-[600px] table-fixed">
                    <thead className="sticky top-0 z-20 bg-slate-200 shadow-[0_1px_0_0_#cbd5e1]">
                      <tr>
                        <th className="w-8 md:w-10 border-r border-slate-300 text-slate-500 py-1.5 text-center">#</th>
                        {gridData[0].map((header, cIdx) => (
                          <th key={`h-${cIdx}`} className="p-0 border-r border-slate-300 last:border-r-0 relative">
                            <input
                              type="text"
                              className="w-full bg-transparent p-1.5 font-bold text-slate-700 text-center focus:bg-white focus:outline-blue-500 placeholder-slate-400"
                              value={header} placeholder={`欄位 ${cIdx+1}`}
                              onChange={(e) => handleCellChange(0, cIdx, e.target.value)}
                              onPaste={(e) => handleGridPaste(e, 0, cIdx)}
                            />
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody className="bg-white">
                      {gridData.slice(1).map((row, rIdx) => {
                        const actualRowIdx = rIdx + 1;
                        const hasData = row.some(cell => cell.trim() !== '');
                        
                        return (
                          <tr key={`r-${actualRowIdx}`} className={`${hasData ? 'bg-white' : 'bg-[#FAFAFA]'} hover:bg-blue-50/40 border-b border-slate-100`}>
                            <td className="w-8 md:w-10 bg-slate-50 border-r border-slate-200 text-slate-400 text-center font-mono text-xs py-1">
                              {actualRowIdx}
                            </td>
                            {row.map((cell, cIdx) => {
                              const headerName = gridData[0][cIdx];
                              const isScoreCol = headerName && headerName !== '座號' && headerName !== '姓名';
                              
                              let liveGrade = '';
                              let isError = false;
                              if (isScoreCol && cell.trim() !== '') {
                                const num = parseFloat(cell);
                                if (!isNaN(num)) {
                                  const subjectSettings = currentParsedSettings.settings[headerName];
                                  if (subjectSettings) liveGrade = getGradeLevel(num, subjectSettings);
                                } else { isError = true; } // 非數字
                              }

                              return (
                                <td key={`c-${cIdx}`} className="p-0 border-r border-slate-200 last:border-r-0 relative">
                                  <div className="flex items-center px-1">
                                    <input
                                      type="text"
                                      className={`w-full py-1.5 px-1 focus:outline-none bg-transparent transition-all min-w-0 ${
                                        isScoreCol ? 'text-right font-mono' : 'text-center'
                                      } ${cell.trim() !== '' ? 'text-slate-800 font-medium' : 'text-slate-400'}
                                      ${isError ? 'text-red-500' : ''}`}
                                      value={cell}
                                      onChange={(e) => handleCellChange(actualRowIdx, cIdx, e.target.value)}
                                      onPaste={(e) => handleGridPaste(e, actualRowIdx, cIdx)}
                                    />
                                    {/* 緊湊型等級標籤，直接跟在輸入框旁邊 */}
                                    {isScoreCol && (
                                      <div className="w-6 shrink-0 flex justify-center items-center ml-0.5 pointer-events-none">
                                        {liveGrade && (
                                          <span className={`text-[10px] font-black leading-none ${
                                            liveGrade.includes('A') ? 'text-teal-600' :
                                            liveGrade.includes('C') ? 'text-rose-500' : 'text-slate-400'
                                          }`}>{liveGrade}</span>
                                        )}
                                      </div>
                                    )}
                                  </div>
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
          )}

          {/* 介面三：結果區 */}
          {view === 'result' && generateReportData && (
            <div className="p-4 md:p-6 flex flex-col h-full flex-grow animate-in slide-in-from-bottom-4">
              
              <div className="flex flex-col md:flex-row justify-between items-start md:items-center bg-slate-50 p-3 md:p-4 rounded-xl border border-slate-200 mb-4 gap-3">
                <h2 className="text-lg font-bold text-slate-800 flex items-center">
                  <CheckCircle2 className="w-5 h-5 text-teal-600 mr-2" />
                  {selectedGrade}年級 結果 <span className="text-slate-500 text-sm font-medium ml-2">({generateReportData.data?.length || 0}筆)</span>
                </h2>
                <div className="flex flex-wrap gap-2 w-full md:w-auto">
                  <button onClick={() => setView('input')} className="flex-1 md:flex-none px-4 py-2 bg-white text-slate-600 border border-slate-200 rounded-lg hover:bg-slate-100 transition-all text-sm font-bold flex justify-center items-center">
                    <ArrowLeft className="w-3.5 h-3.5 mr-1.5" /> 返回
                  </button>
                  <button onClick={handleCopyReport} className="flex-1 md:flex-none px-4 py-2 bg-blue-50 text-blue-700 border border-blue-200 rounded-lg hover:bg-blue-100 transition-all text-sm font-bold flex justify-center items-center">
                    <ClipboardCopy className="w-3.5 h-3.5 mr-1.5" /> 複製
                  </button>
                  <button onClick={() => exportToCSV(generateReportData.data, `${selectedGrade}年級_成績分析.csv`)} className="flex-1 md:flex-none px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-all text-sm font-bold flex justify-center items-center">
                    <FileDown className="w-3.5 h-3.5 mr-1.5" /> 匯出
                  </button>
                </div>
              </div>

              {generateReportData.error ? (
                 <div className="p-10 text-center text-red-500 font-bold">{generateReportData.error}</div>
              ) : (
                <div className="flex-grow flex flex-col min-h-[300px]">
                  {generateReportData.unmappedSubjects.length > 0 && (
                    <div className="bg-amber-50 border border-amber-200 p-3 rounded-lg flex items-start text-xs md:text-sm mb-3">
                      <AlertTriangle className="w-4 h-4 text-amber-500 mr-2 flex-shrink-0 mt-0.5" />
                      <p className="text-amber-800">未設定的科目：<strong>{generateReportData.unmappedSubjects.join('、')}</strong>。已計入平均，但無法標示等級。</p>
                    </div>
                  )}

                  <div className="border border-slate-200 rounded-xl overflow-hidden shadow-sm relative custom-scrollbar flex-grow bg-white">
                    <div className="overflow-auto absolute inset-0">
                      <table className="w-full text-[13px] text-left whitespace-nowrap table-fixed">
                        <thead className="text-slate-500 bg-slate-50 uppercase font-bold sticky top-0 z-10 shadow-[0_1px_0_0_#e2e8f0]">
                          <tr>
                            {Object.keys(generateReportData.data[0]).map((header, idx) => (
                              <th key={idx} className="px-4 py-3 border-r border-slate-100 last:border-r-0 truncate">
                                {header}
                              </th>
                            ))}
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100">
                          {generateReportData.data.map((row, rowIndex) => (
                            <tr key={rowIndex} className="hover:bg-blue-50/50">
                              {Object.entries(row).map(([key, val], colIndex) => {
                                const isGrade = typeof val === 'string' && val.includes('(') && val.includes(')');
                                const gradeMatch = isGrade ? val.match(/\((.*?)\)/) : null;
                                const gradeLabel = gradeMatch ? gradeMatch[1] : '';
                                
                                let textClass = "text-slate-700";
                                if (gradeLabel.includes('A')) textClass = "text-teal-600 font-bold";
                                if (gradeLabel.includes('C')) textClass = "text-rose-500 font-bold";
                                if (key === '預估校排') textClass = "text-blue-700 font-bold";

                                return (
                                  <td key={colIndex} className={`px-4 py-2 border-r border-slate-50 last:border-r-0 truncate ${textClass}`}>
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
                </div>
              )}
            </div>
          )}

          {/* VIEW: 管理員登入密碼區 */}
          {view === 'admin_login' && (
             <div className="p-8 md:p-16 flex flex-col items-center justify-center flex-grow animate-in zoom-in-95">
               <div className="bg-slate-50 p-5 rounded-full mb-4 shadow-inner">
                 <Lock className="w-10 h-10 text-slate-400" />
               </div>
               <h2 className="text-xl font-bold text-slate-800 mb-2">管理員登入</h2>
               <p className="text-slate-500 text-sm mb-6 text-center">輸入密碼以設定全校標準</p>
               <form onSubmit={handleAdminLogin} className="flex flex-col w-full max-w-xs space-y-3">
                 <input type="password" autoFocus placeholder="密碼..." className="w-full px-4 py-2 text-center tracking-widest bg-white border-2 border-slate-200 rounded-xl focus:border-teal-500 focus:outline-none" value={adminPassword} onChange={(e) => setAdminPassword(e.target.value)} />
                 <button type="submit" className="w-full py-2.5 bg-teal-600 hover:bg-teal-700 text-white font-bold rounded-xl shadow-sm flex justify-center items-center text-sm">
                   解鎖 <Unlock className="w-3.5 h-3.5 ml-1.5" />
                 </button>
               </form>
             </div>
          )}

          {/* VIEW: 管理員設定區 */}
          {view === 'admin_settings' && (
            <div className="p-4 md:p-6 flex flex-col flex-grow animate-in fade-in overflow-y-auto">
              <div className="text-center pb-4 mb-4 border-b border-slate-100 flex flex-col items-center">
                <h2 className="text-xl font-extrabold text-slate-800 mb-3">全校標準金庫</h2>
                <div className="flex bg-slate-100 p-1 rounded-lg space-x-1">
                  {['7', '8', '9'].map(grade => (
                    <button key={`admin-${grade}`} onClick={() => setSelectedGrade(grade)} className={`px-5 py-1.5 rounded-md font-bold text-sm transition-all ${selectedGrade === grade ? 'bg-white text-blue-700 shadow-sm' : 'text-slate-500'}`}>
                      {grade} 年級
                    </button>
                  ))}
                </div>
              </div>
              {!selectedGrade ? (
                <div className="text-center py-10 text-slate-400 font-medium text-sm">請先在上方選擇年級。</div>
              ) : (
                <div className="grid grid-cols-1 lg:grid-cols-2 gap-4 flex-grow">
                  <div className="flex flex-col bg-[#FFFCF5] p-4 rounded-2xl border border-amber-200/60 shadow-sm">
                    <div className="flex justify-between items-center mb-3">
                      <span className="text-sm font-bold text-amber-800">各科等級門檻</span>
                      <div>
                        <input type="file" accept=".csv, .xlsx, .xls" style={{ display: 'none' }} ref={settingFileInputRef} onChange={(e) => handleAdminFileUpload(e, 'grade')} />
                        <button onClick={() => settingFileInputRef.current.click()} className="px-3 py-1.5 bg-white text-amber-700 border border-amber-200 rounded-lg text-xs font-bold shadow-sm">匯入 Excel</button>
                      </div>
                    </div>
                    <textarea className="w-full flex-grow min-h-[250px] p-3 border border-amber-200/80 rounded-xl text-xs font-mono text-slate-700 focus:outline-none" value={appSettings[selectedGrade].grade} onChange={(e) => setAppSettings(prev => ({...prev, [selectedGrade]: {...prev[selectedGrade], grade: e.target.value}}))} />
                  </div>
                  <div className="flex flex-col bg-[#F5F9FF] p-4 rounded-2xl border border-blue-200/60 shadow-sm">
                    <div className="flex justify-between items-center mb-3">
                      <span className="text-sm font-bold text-blue-800">全校分數組距</span>
                      <div>
                        <input type="file" accept=".csv, .xlsx, .xls" style={{ display: 'none' }} ref={distFileInputRef} onChange={(e) => handleAdminFileUpload(e, 'dist')} />
                        <button onClick={() => distFileInputRef.current.click()} className="px-3 py-1.5 bg-white text-blue-700 border border-blue-200 rounded-lg text-xs font-bold shadow-sm">匯入 Excel</button>
                      </div>
                    </div>
                    <textarea className="w-full flex-grow min-h-[250px] p-3 border border-blue-200/80 rounded-xl text-xs font-mono text-slate-700 focus:outline-none" value={appSettings[selectedGrade].dist} onChange={(e) => setAppSettings(prev => ({...prev, [selectedGrade]: {...prev[selectedGrade], dist: e.target.value}}))} />
                  </div>
                </div>
              )}
            </div>
          )}
        </main>
        
        <footer className="text-center text-slate-400 text-xs font-medium py-1">
           由 蘇老爹 開發設計
        </footer>

      </div>
    </div>
  );
}