import React, { useState, useMemo, useRef, useEffect } from 'react';
import { Upload, FileDown, ClipboardCopy, Calculator, CheckCircle2, RefreshCw, Settings, ArrowLeft, Lock, Unlock, Info, AlertTriangle, Users, BookOpen, ChevronRight } from 'lucide-react';

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
  return 'C'; // 預設最低等級
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

// 匯出 CSV 工具
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

const INITIAL_GRID_ROWS = 45;
const INITIAL_GRID_COLS = 10;
const defaultHeaders = ['座號', '姓名', '國文', '英文', '數學', '社會', '自然', '', '', ''];

const createInitialGrid = () => {
  const grid = Array(INITIAL_GRID_ROWS).fill(0).map(() => Array(INITIAL_GRID_COLS).fill(''));
  grid[0] = [...defaultHeaders];
  return grid;
};

// --- 3. 主應用程式元件 ---
export default function App() {
  // 介面狀態: 'home' | 'input' | 'result' | 'admin_login' | 'admin_settings'
  const [view, setView] = useState('home'); 
  const [selectedGrade, setSelectedGrade] = useState(null); // '7', '8', '9'
  const [error, setError] = useState(null);
  
  // 管理員驗證狀態
  const [isAdminAuth, setIsAdminAuth] = useState(false);
  const [adminPassword, setAdminPassword] = useState('');
  
  // 各年級設定檔
  const [appSettings, setAppSettings] = useState({
    '7': { grade: defaultSettings, dist: defaultDistribution },
    '8': { grade: defaultSettings, dist: defaultDistribution },
    '9': { grade: defaultSettings, dist: defaultDistribution }
  });

  // 教師輸入網格 (動態表格)
  const [gridData, setGridData] = useState(createInitialGrid());
  
  const settingFileInputRef = useRef(null);
  const distFileInputRef = useRef(null);

  // 載入 SheetJS
  useEffect(() => {
    if (!window.XLSX) {
      const script = document.createElement('script');
      script.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
      script.async = true;
      document.body.appendChild(script);
    }
  }, []);

  // 取得當前年級的解析後設定 (供即時網格與報表使用)
  const currentParsedSettings = useMemo(() => {
    if (!selectedGrade) return { settings: {}, subjects: [] };
    return processSettings(parseCSV(appSettings[selectedGrade].grade));
  }, [appSettings, selectedGrade]);

  const currentDistMap = useMemo(() => {
    if (!selectedGrade) return [];
    return processDistribution(parseCSV(appSettings[selectedGrade].dist));
  }, [appSettings, selectedGrade]);

  // 表格輸入處理
  const handleCellChange = (rIdx, cIdx, value) => {
    const newGrid = [...gridData];
    newGrid[rIdx] = [...newGrid[rIdx]];
    newGrid[rIdx][cIdx] = value;
    setGridData(newGrid);
  };

  // 處理貼上事件 (Excel 直接貼上支援)
  const handleGridPaste = (e, startRow, startCol) => {
    e.preventDefault();
    const pasteText = e.clipboardData.getData('text');
    if (!pasteText) return;

    const rows = pasteText.split(/\r?\n/);
    const newGrid = [...gridData];

    rows.forEach((rowStr, i) => {
      if (!rowStr.trim() && i === rows.length - 1) return; // 忽略最後的空行
      const cells = rowStr.split(/\t|,/);
      const targetRow = startRow + i;
      
      if (targetRow < INITIAL_GRID_ROWS) {
        newGrid[targetRow] = [...newGrid[targetRow]];
        cells.forEach((cellVal, j) => {
          const targetCol = startCol + j;
          if (targetCol < INITIAL_GRID_COLS) {
            newGrid[targetRow][targetCol] = cellVal.trim().replace(/^"|"$/g, ''); // 移除可能的引號
          }
        });
      }
    });
    setGridData(newGrid);
  };

  // 清空表格
  const handleClearGrid = () => {
    if(window.confirm("確定要清空目前表格內的所有資料嗎？")) {
      setGridData(createInitialGrid());
      setError(null);
    }
  };

  // 管理員設定檔案上傳
  const handleAdminFileUpload = (e, type) => {
    const file = e.target.files[0];
    if (!file) return;
    const fileExt = file.name.split('.').pop().toLowerCase();

    const updateSetting = (text) => {
      setAppSettings(prev => ({
        ...prev,
        [selectedGrade]: { ...prev[selectedGrade], [type]: text }
      }));
      setError(null);
    };

    if (fileExt === 'csv') {
      const reader = new FileReader();
      reader.onload = (event) => updateSetting(event.target.result);
      reader.readAsText(file); 
    } else if (fileExt === 'xlsx' || fileExt === 'xls') {
      if (!window.XLSX) { setError("Excel 模組載入中，請稍後再試。"); return; }
      const reader = new FileReader();
      reader.onload = (event) => {
        try {
          const workbook = window.XLSX.read(new Uint8Array(event.target.result), { type: 'array' });
          updateSetting(window.XLSX.utils.sheet_to_csv(workbook.Sheets[workbook.SheetNames[0]]));
        } catch (err) { setError("讀取 Excel 失敗。"); }
      };
      reader.readAsArrayBuffer(file);
    } else {
      setError("不支援的檔案格式，請上傳 .xlsx 或 .csv。");
    }
    e.target.value = '';
  };

  // 驗證管理員密碼
  const handleAdminLogin = (e) => {
    e.preventDefault();
    if (adminPassword === '690530') {
      setIsAdminAuth(true);
      setView('admin_settings');
      setAdminPassword('');
      setError(null);
    } else {
      setError("密碼錯誤，請重新輸入。");
    }
  };

  // 產生報表邏輯
  const generateReportData = useMemo(() => {
    if (view !== 'result') return null;

    const headers = gridData[0];
    const excludeCols = ['座號', '姓名'];
    const subjects = headers.filter(h => h && !excludeCols.includes(h));
    
    // 將 Grid 轉換為物件陣列 (過濾空行)
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
            } else {
              resultRow[subject] = numScore; 
            }
          } else {
            resultRow[subject] = val; // 保留文字
          }
        } else {
          resultRow[subject] = '';
        }
      });

      resultRow['平均'] = validCount > 0 ? parseFloat((totalScore / validCount).toFixed(1)) : 0;
      return resultRow;
    });

    // 計算排名
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
    if (hasData) {
      setView('result'); setError(null);
    } else {
      setError("請先在表格內輸入成績資料。");
    }
  };

  const handleCopyReport = () => {
    if (!generateReportData || !generateReportData.data) return;
    const tsvContent = [
      Object.keys(generateReportData.data[0]).join('\t'), 
      ...generateReportData.data.map(row => Object.values(row).join('\t'))
    ].join('\n');
    const textArea = document.createElement("textarea");
    textArea.value = tsvContent;
    document.body.appendChild(textArea);
    textArea.select();
    try { document.execCommand('copy'); alert("✅ 已成功複製，可直接貼上至 Excel！"); } 
    catch (err) { alert("複製失敗。"); }
    document.body.removeChild(textArea);
  };

  // --- 畫面渲染 ---

  return (
    <div className="min-h-screen bg-[#F4F7F9] text-slate-800 p-4 md:p-6 font-sans selection:bg-blue-100 flex flex-col">
      <div className="max-w-[1280px] mx-auto w-full space-y-6 flex-grow">
        
        {/* 全域免責聲明 */}
        <div className="bg-amber-50 border border-amber-200/60 rounded-xl p-3 flex items-start md:items-center text-amber-800 text-sm font-medium shadow-sm">
          <AlertTriangle className="w-5 h-5 mr-3 flex-shrink-0 text-amber-500" />
          <p><strong>免責說明：</strong>本程式為老師自行設計之快速換算輔助工具，並非學校官方正式成績系統。若對成績或排名有任何疑問，請洽詢學校教務處確認。</p>
        </div>

        {/* Header */}
        <header className="flex flex-col md:flex-row justify-between items-center bg-white p-5 rounded-3xl shadow-sm border border-slate-100/50">
          <div className="flex items-center space-x-4 cursor-pointer hover:opacity-80 transition-opacity" onClick={() => { setView('home'); setSelectedGrade(null); }}>
            <div className="bg-gradient-to-br from-blue-500 to-teal-400 p-3.5 rounded-2xl text-white shadow-md shadow-blue-500/20">
              <Calculator className="w-6 h-6" />
            </div>
            <div>
              <h1 className="text-2xl font-extrabold text-slate-800 tracking-tight">成績分析與等級產生器</h1>
              <p className="text-sm text-slate-500 font-medium mt-0.5">自動套用全校標準 ‧ 智慧精算校排</p>
            </div>
          </div>
          <div className="mt-4 md:mt-0 flex flex-wrap gap-2">
            {view !== 'home' && view !== 'admin_login' && (
               <span className="px-4 py-2 bg-blue-50 text-blue-700 font-bold rounded-xl border border-blue-100 mr-2 flex items-center">
                 <Users className="w-4 h-4 mr-2"/> {selectedGrade} 年級模式
               </span>
            )}
            <button 
              onClick={() => {
                if (view === 'admin_settings' || view === 'admin_login') {
                  setView('home'); setSelectedGrade(null); setIsAdminAuth(false);
                } else {
                  setView('admin_login'); setSelectedGrade(null);
                }
              }}
              className={`px-5 py-2.5 rounded-xl font-semibold transition-all flex items-center ${
                view.includes('admin') 
                  ? 'bg-slate-800 text-white shadow-md' 
                  : 'bg-white border-2 border-slate-100 text-slate-500 hover:border-slate-200 hover:text-slate-700'
              }`}
            >
              {view.includes('admin') ? <><CheckCircle2 className="w-4 h-4 mr-2" /> 返回教師介面</> : <><Settings className="w-4 h-4 mr-2" /> 管理員設定</>}
            </button>
          </div>
        </header>

        {error && (
          <div className="bg-red-50 text-red-600 px-6 py-4 rounded-2xl flex items-center border border-red-100 shadow-sm animate-in fade-in slide-in-from-top-2">
            <AlertTriangle className="w-5 h-5 mr-3 flex-shrink-0" />
            <p className="font-medium">{error}</p>
          </div>
        )}

        {/* 主內容區塊 */}
        <main className="bg-white rounded-3xl shadow-sm border border-slate-100 overflow-hidden min-h-[600px] relative">
          
          {/* VIEW: 首頁 (年級選擇) */}
          {view === 'home' && (
            <div className="p-10 md:p-16 flex flex-col items-center justify-center h-full animate-in zoom-in-95">
              <div className="w-16 h-16 bg-blue-50 text-blue-500 rounded-full flex items-center justify-center mb-6">
                <BookOpen className="w-8 h-8" />
              </div>
              <h2 className="text-3xl font-extrabold text-slate-800 mb-3">歡迎使用成績分析系統</h2>
              <p className="text-slate-500 text-lg mb-10 text-center max-w-lg">請先選擇您要計算的年級，系統將載入對應的專屬全校成績標準與等級門檻。</p>
              
              <div className="grid grid-cols-1 md:grid-cols-3 gap-6 w-full max-w-3xl">
                {['7', '8', '9'].map(grade => (
                  <button
                    key={grade}
                    onClick={() => { setSelectedGrade(grade); setView('input'); setError(null); }}
                    className="group flex flex-col items-center justify-center p-8 bg-white border-2 border-slate-100 hover:border-blue-400 hover:shadow-lg hover:shadow-blue-500/10 rounded-3xl transition-all"
                  >
                    <span className="text-4xl font-black text-slate-300 group-hover:text-blue-500 transition-colors mb-2">
                      {grade}
                    </span>
                    <span className="text-xl font-bold text-slate-700 group-hover:text-blue-700">年級專區</span>
                    <ChevronRight className="w-5 h-5 text-slate-300 group-hover:text-blue-500 mt-4 opacity-0 group-hover:opacity-100 transform translate-x-[-10px] group-hover:translate-x-0 transition-all" />
                  </button>
                ))}
              </div>
            </div>
          )}

          {/* VIEW: 教師輸入區 (即時網格) */}
          {view === 'input' && selectedGrade && (
            <div className="p-6 md:p-8 flex flex-col h-full animate-in fade-in">
              
              <div className="flex flex-col lg:flex-row justify-between items-start gap-4 mb-6">
                <div className="bg-blue-50/50 rounded-2xl p-5 border border-blue-100/50 flex-1 w-full">
                  <h3 className="text-lg font-bold text-slate-800 flex items-center mb-2">
                    <Info className="w-5 h-5 text-blue-500 mr-2" /> 如何使用即時表格？
                  </h3>
                  <ul className="text-sm text-slate-600 font-medium space-y-1 ml-7 list-disc">
                    <li>您可以直接點擊下方表格格子<strong>「手動輸入分數」</strong>，系統會瞬間顯示對應等級。</li>
                    <li>支援從 Excel 複製整批成績，點擊表格左上角第一格，按下 <kbd className="bg-white border border-slate-200 rounded px-1.5 py-0.5 text-xs mx-1 shadow-sm">Ctrl+V</kbd> 直接貼上！</li>
                    <li>若科目名稱與管理員設定不同，將只計算平均無法標示等級。可以點擊最上排標題進行修改。</li>
                  </ul>
                </div>
                
                <div className="flex flex-col gap-3 min-w-[200px] w-full lg:w-auto">
                  <button onClick={handleGenerate} className="px-6 py-4 bg-slate-800 text-white font-bold rounded-xl hover:bg-slate-900 transition-all shadow-md flex justify-center items-center group">
                    產生排名與報表 <RefreshCw className="w-4 h-4 ml-2 text-teal-400 group-hover:rotate-180 transition-transform duration-500" />
                  </button>
                  <button onClick={handleClearGrid} className="px-6 py-2.5 bg-white border border-slate-200 text-slate-500 font-bold rounded-xl hover:bg-slate-50 hover:text-red-500 transition-all shadow-sm">
                    清空表格資料
                  </button>
                </div>
              </div>

              {/* 網格編輯區 */}
              <div className="border-2 border-slate-200 rounded-xl overflow-hidden shadow-inner flex-1 bg-slate-50 relative">
                <div className="overflow-auto max-h-[55vh] custom-scrollbar relative">
                  <table className="w-full text-sm border-collapse min-w-max">
                    <thead className="sticky top-0 z-20 bg-slate-100 shadow-[0_1px_0_0_#cbd5e1]">
                      <tr>
                        <th className="w-10 bg-slate-200 border-b border-r border-slate-300 text-slate-500 text-center py-2">#</th>
                        {gridData[0].map((header, cIdx) => (
                          <th key={`h-${cIdx}`} className="p-0 border-r border-slate-300 last:border-r-0 relative min-w-[100px]">
                            <input
                              type="text"
                              className="w-full bg-transparent p-2.5 font-bold text-slate-700 text-center focus:bg-white focus:outline-blue-500"
                              value={header}
                              placeholder={`欄位 ${cIdx + 1}`}
                              onChange={(e) => handleCellChange(0, cIdx, e.target.value)}
                              onPaste={(e) => handleGridPaste(e, 0, cIdx)}
                            />
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {gridData.slice(1).map((row, rIdx) => {
                        const actualRowIdx = rIdx + 1;
                        // 判斷該行是否有資料
                        const hasData = row.some(cell => cell.trim() !== '');
                        
                        return (
                          <tr key={`r-${actualRowIdx}`} className={`${hasData ? 'bg-white' : 'bg-[#FDFDFE]'} hover:bg-blue-50/30 transition-colors`}>
                            <td className="w-10 bg-slate-100 border-b border-r border-slate-200 text-slate-400 text-center font-mono text-xs select-none">
                              {actualRowIdx}
                            </td>
                            {row.map((cell, cIdx) => {
                              const headerName = gridData[0][cIdx];
                              const isScoreCol = headerName && headerName !== '座號' && headerName !== '姓名';
                              
                              // 即時計算等級
                              let liveGrade = '';
                              if (isScoreCol && cell.trim() !== '' && !isNaN(parseFloat(cell))) {
                                const subjectSettings = currentParsedSettings.settings[headerName];
                                if (subjectSettings) {
                                  liveGrade = getGradeLevel(parseFloat(cell), subjectSettings);
                                }
                              }

                              return (
                                <td key={`c-${cIdx}`} className="p-0 border-b border-r border-slate-200 last:border-r-0 relative group">
                                  <input
                                    type="text"
                                    className={`w-full p-2.5 focus:outline-none focus:ring-2 focus:ring-inset focus:ring-blue-500 bg-transparent transition-all ${
                                      isScoreCol ? 'text-right pr-10 font-mono' : 'text-center'
                                    } ${cell.trim() !== '' ? 'text-slate-800 font-medium' : 'text-slate-400'}`}
                                    value={cell}
                                    onChange={(e) => handleCellChange(actualRowIdx, cIdx, e.target.value)}
                                    onPaste={(e) => handleGridPaste(e, actualRowIdx, cIdx)}
                                  />
                                  {/* 即時等級徽章 */}
                                  {liveGrade && (
                                    <div className="absolute left-2 top-1/2 -translate-y-1/2 pointer-events-none select-none">
                                      <span className={`text-[11px] font-black px-1.5 py-0.5 rounded ${
                                        liveGrade.includes('A') ? 'bg-teal-100 text-teal-700' :
                                        liveGrade.includes('C') ? 'bg-rose-100 text-rose-700' :
                                        'bg-slate-100 text-slate-600'
                                      }`}>
                                        {liveGrade}
                                      </span>
                                    </div>
                                  )}
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

          {/* VIEW: 管理員登入密碼區 */}
          {view === 'admin_login' && (
             <div className="p-10 md:p-16 flex flex-col items-center justify-center h-full animate-in zoom-in-95">
               <div className="bg-slate-50 p-6 rounded-full mb-6 shadow-inner">
                 <Lock className="w-12 h-12 text-slate-400" />
               </div>
               <h2 className="text-2xl font-bold text-slate-800 mb-2">教務處 / 管理員登入</h2>
               <p className="text-slate-500 mb-8 text-center">請輸入管理員密碼以設定各年級的全校標準。</p>
               
               <form onSubmit={handleAdminLogin} className="flex flex-col w-full max-w-sm space-y-4">
                 <input 
                   type="password" autoFocus placeholder="請輸入密碼..." 
                   className="w-full px-5 py-3 text-center text-lg tracking-widest bg-white border-2 border-slate-200 rounded-xl focus:border-teal-500 focus:ring-4 focus:ring-teal-500/10 transition-all outline-none"
                   value={adminPassword} onChange={(e) => setAdminPassword(e.target.value)}
                 />
                 <button type="submit" className="w-full py-3 bg-teal-600 hover:bg-teal-700 text-white font-bold rounded-xl transition-all shadow-md flex justify-center items-center">
                   解鎖並進入設定 <Unlock className="w-4 h-4 ml-2" />
                 </button>
               </form>
             </div>
          )}

          {/* VIEW: 管理員設定區 */}
          {view === 'admin_settings' && (
            <div className="p-8 md:p-10 space-y-8 animate-in fade-in">
              <div className="text-center pb-6 border-b border-slate-100 flex flex-col items-center">
                <div className="inline-flex items-center justify-center bg-teal-100 text-teal-800 px-4 py-1.5 rounded-full text-sm font-bold mb-4">
                  <Unlock className="w-4 h-4 mr-2" /> 管理員金庫模式已解鎖
                </div>
                <h2 className="text-3xl font-extrabold text-slate-800 mb-4">全校標準設定</h2>
                
                {/* 管理員年級選擇 */}
                <div className="flex bg-slate-100 p-1.5 rounded-xl space-x-1">
                  {['7', '8', '9'].map(grade => (
                    <button
                      key={`admin-${grade}`}
                      onClick={() => setSelectedGrade(grade)}
                      className={`px-8 py-2 rounded-lg font-bold transition-all ${
                        selectedGrade === grade 
                          ? 'bg-white text-blue-700 shadow-sm' 
                          : 'text-slate-500 hover:text-slate-700'
                      }`}
                    >
                      {grade} 年級
                    </button>
                  ))}
                </div>
              </div>

              {!selectedGrade ? (
                <div className="text-center py-20 text-slate-400 font-medium animate-in fade-in">
                  請先在上方選擇要設定的年級。
                </div>
              ) : (
                <div className="grid grid-cols-1 lg:grid-cols-2 gap-8 animate-in slide-in-from-bottom-4">
                  {/* 等級設定 */}
                  <div className="space-y-4 bg-[#FFFCF5] p-6 rounded-3xl border border-amber-200/60 shadow-sm relative overflow-hidden">
                    <div className="absolute top-0 right-0 w-32 h-32 bg-amber-100/30 rounded-full blur-2xl -mr-10 -mt-10 pointer-events-none"></div>
                    <div className="flex justify-between items-center relative z-10">
                      <span className="text-lg font-bold text-amber-800 flex items-center">
                        【{selectedGrade}年級】各科等級門檻
                      </span>
                      <div>
                        <input type="file" accept=".csv, .xlsx, .xls" style={{ display: 'none' }} ref={settingFileInputRef} onChange={(e) => handleAdminFileUpload(e, 'grade')} />
                        <button onClick={() => settingFileInputRef.current.click()} className="flex items-center px-4 py-2 bg-white hover:bg-amber-50 text-amber-700 border border-amber-200 rounded-xl transition-colors font-semibold shadow-sm text-sm">
                          <Upload className="w-4 h-4 mr-2" /> 匯入 Excel
                        </button>
                      </div>
                    </div>
                    <textarea 
                      className="w-full h-80 p-5 bg-white border border-amber-200/80 focus:border-amber-400 focus:ring-4 focus:ring-amber-400/10 rounded-2xl text-sm font-mono transition-all resize-y shadow-inner text-slate-700 relative z-10"
                      value={appSettings[selectedGrade].grade} 
                      onChange={(e) => setAppSettings(prev => ({...prev, [selectedGrade]: {...prev[selectedGrade], grade: e.target.value}}))}
                    />
                  </div>

                  {/* 全校組距設定 */}
                  <div className="space-y-4 bg-[#F5F9FF] p-6 rounded-3xl border border-blue-200/60 shadow-sm relative overflow-hidden">
                    <div className="absolute top-0 right-0 w-32 h-32 bg-blue-100/30 rounded-full blur-2xl -mr-10 -mt-10 pointer-events-none"></div>
                    <div className="flex justify-between items-center relative z-10">
                      <span className="text-lg font-bold text-blue-800 flex items-center">
                        【{selectedGrade}年級】全校分數組距
                      </span>
                      <div>
                        <input type="file" accept=".csv, .xlsx, .xls" style={{ display: 'none' }} ref={distFileInputRef} onChange={(e) => handleAdminFileUpload(e, 'dist')} />
                        <button onClick={() => distFileInputRef.current.click()} className="flex items-center px-4 py-2 bg-white hover:bg-blue-50 text-blue-700 border border-blue-200 rounded-xl transition-colors font-semibold shadow-sm text-sm">
                          <Upload className="w-4 h-4 mr-2" /> 匯入 Excel
                        </button>
                      </div>
                    </div>
                    <textarea 
                      className="w-full h-80 p-5 bg-white border border-blue-200/80 focus:border-blue-400 focus:ring-4 focus:ring-blue-400/10 rounded-2xl text-sm font-mono transition-all resize-y shadow-inner text-slate-700 relative z-10"
                      value={appSettings[selectedGrade].dist} 
                      onChange={(e) => setAppSettings(prev => ({...prev, [selectedGrade]: {...prev[selectedGrade], dist: e.target.value}}))}
                    />
                  </div>
                </div>
              )}
            </div>
          )}

          {/* VIEW: 結果區 */}
          {view === 'result' && generateReportData && (
            <div className="p-8 md:p-10 space-y-6 animate-in slide-in-from-bottom-4">
              
              <div className="flex flex-col md:flex-row justify-between items-center bg-slate-50 p-4 rounded-2xl border border-slate-100">
                <h2 className="text-xl font-bold text-slate-800 flex items-center">
                  <span className="w-10 h-10 bg-teal-100 text-teal-600 rounded-full flex items-center justify-center mr-3 shadow-sm">
                    <CheckCircle2 className="w-6 h-6" />
                  </span>
                  {selectedGrade}年級 報表生成完畢 <span className="text-slate-500 text-base font-medium ml-2"> (共 {generateReportData.data?.length || 0} 筆)</span>
                </h2>
                <div className="flex flex-wrap gap-3 mt-4 md:mt-0">
                  <button onClick={() => setView('input')} className="flex items-center px-5 py-2.5 bg-white text-slate-600 border border-slate-200 rounded-xl hover:bg-slate-50 transition-all shadow-sm font-bold">
                    <ArrowLeft className="w-4 h-4 mr-2" /> 返回修改資料
                  </button>
                  <button onClick={handleCopyReport} className="flex items-center px-5 py-2.5 bg-white text-blue-700 border border-blue-200 rounded-xl hover:bg-blue-50 transition-all shadow-sm font-bold">
                    <ClipboardCopy className="w-4 h-4 mr-2" /> 複製完整表格
                  </button>
                  <button onClick={() => exportToCSV(generateReportData.data, `${selectedGrade}年級_成績分析結果.csv`)} className="flex items-center px-5 py-2.5 bg-blue-600 text-white rounded-xl hover:bg-blue-700 transition-all shadow-md shadow-blue-500/20 font-bold">
                    <FileDown className="w-4 h-4 mr-2" /> 下載 CSV
                  </button>
                </div>
              </div>

              {generateReportData.error ? (
                 <div className="p-10 text-center text-red-500 font-bold">{generateReportData.error}</div>
              ) : (
                <>
                  {/* 欄位名稱不符的警告 */}
                  {generateReportData.unmappedSubjects.length > 0 && (
                    <div className="bg-amber-50 border border-amber-200 p-4 rounded-xl flex items-start shadow-sm">
                      <AlertTriangle className="w-6 h-6 text-amber-500 mr-3 flex-shrink-0 mt-0.5" />
                      <div>
                        <h4 className="font-bold text-amber-800">發現未設定的科目</h4>
                        <p className="text-sm text-amber-700 mt-1">
                          您的表格包含：<strong>{generateReportData.unmappedSubjects.join('、')}</strong>。管理員尚未設定這些科目的門檻，系統已將其計入平均，但無法標示等級。
                        </p>
                      </div>
                    </div>
                  )}

                  <div className="overflow-x-auto overflow-y-auto max-h-[60vh] border-2 border-slate-100 rounded-2xl shadow-sm relative custom-scrollbar">
                    <table className="w-full text-sm text-left whitespace-nowrap">
                      <thead className="text-[13px] text-slate-500 bg-slate-50 uppercase font-extrabold tracking-wider sticky top-0 z-10 shadow-[0_1px_0_0_#e2e8f0]">
                        <tr>
                          {Object.keys(generateReportData.data[0]).map((header, idx) => (
                            <th key={idx} className="px-6 py-4 bg-[#F8FAFC]">
                              {header}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {generateReportData.data.map((row, rowIndex) => (
                          <tr key={rowIndex} className="hover:bg-blue-50/60 transition-colors bg-white">
                            {Object.entries(row).map(([key, val], colIndex) => {
                              const isGrade = typeof val === 'string' && val.includes('(') && val.includes(')');
                              const gradeMatch = isGrade ? val.match(/\((.*?)\)/) : null;
                              const gradeLabel = gradeMatch ? gradeMatch[1] : '';
                              
                              let badgeClass = "text-slate-700 font-medium";
                              if (gradeLabel.includes('A')) badgeClass = "text-teal-600 font-bold";
                              if (gradeLabel.includes('C')) badgeClass = "text-rose-500 font-bold";
                              if (key === '預估校排') badgeClass = "text-blue-700 font-bold bg-blue-50 px-2 py-1 rounded-md inline-block mt-1";

                              return (
                                <td key={colIndex} className="px-6 py-3.5">
                                  <span className={badgeClass}>{val}</span>
                                </td>
                              );
                            })}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </>
              )}
            </div>
          )}
        </main>
        
        {/* 頁尾開發者註記 */}
        <footer className="text-center pb-4 text-slate-400 text-sm font-medium pt-2">
           <p>由 蘇老爹 開發設計</p>
        </footer>

      </div>
    </div>
  );
}