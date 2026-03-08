import React, { useState, useMemo, useRef, useEffect } from 'react';
import { Upload, FileDown, ClipboardCopy, Calculator, CheckCircle2, RefreshCw, Settings, ArrowLeft, Lock, Unlock, Info, FileSpreadsheet, AlertCircle, AlertTriangle } from 'lucide-react';

// --- 1. 核心商業邏輯與工具函式 (Business Logic & Utilities) ---

// 解析 CSV/TSV 字串為物件陣列 (支援 Excel 直接貼上)
const parseCSV = (csvText) => {
  if (!csvText) return [];
  const lines = csvText.trim().split(/\r?\n/).filter(line => line.trim() !== '');
  if (lines.length === 0) return [];
  const delimiter = lines[0].includes('\t') ? '\t' : ',';
  const headers = lines[0].split(delimiter).map(h => h.trim());
  
  return lines.slice(1).map(line => {
    const values = line.split(delimiter).map(v => v.trim());
    const obj = {};
    headers.forEach((header, i) => {
      obj[header] = values[i];
    });
    return obj;
  });
};

// 將等級設定轉換為便於查詢的資料結構
const processSettings = (settingsData) => {
  const settings = {};
  if (!settingsData || settingsData.length === 0) return settings;

  const subjects = Object.keys(settingsData[0]).filter(k => k !== '等級' && k !== '');
  subjects.forEach(subject => {
    settings[subject] = settingsData
      .map(row => ({
        level: row['等級'],
        minScore: parseFloat(row[subject])
      }))
      .filter(item => !isNaN(item.minScore))
      .sort((a, b) => b.minScore - a.minScore);
  });
  return { settings, subjects };
};

// 根據分數取得等級
const getGradeLevel = (score, subjectSettings) => {
  if (isNaN(score)) return '';
  for (let i = 0; i < subjectSettings.length; i++) {
    if (score >= subjectSettings[i].minScore) return subjectSettings[i].level;
  }
  return 'C';
};

// 處理全校組距
const processDistribution = (distData) => {
  if (!distData || distData.length === 0) return [];
  let previousCumulative = 0;

  return distData.map(row => {
    const rangeStr = row['分數組距'];
    if (!rangeStr) return null;

    let min, max;
    if (rangeStr.includes('-')) {
      const parts = rangeStr.split('-');
      min = parseFloat(parts[0]);
      max = parseFloat(parts[1]);
    } else {
      min = parseFloat(rangeStr);
      max = parseFloat(rangeStr);
    }

    const count = parseInt(row['全校人數'] || '0', 10);
    const cumulative = parseInt(row['累計人數'] || '0', 10);
    const result = { min, max, count, cumulative, startRank: previousCumulative + 1 };
    previousCumulative = cumulative;
    return result;
  }).filter(Boolean);
};

// 計算內插法校排
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
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
};

// --- 2. 預設資料設定 ---
const defaultSettings = `,等級,國文,英文,數學,社會,自然
,A++,92,100,94,94,96
,A+,89,98,89,88,92
,A,84,95,79,80,86
,B++,80,92,70,72,78
,B+,74,88,62,64,68
,B,52,50,28,34,36`;

const defaultDistribution = `分數組距,全校人數,累計人數
100,0,0
98-99.99,0,0
96-97.99,23,23
94-95.99,52,75
92-93.99,76,151
90-91.99,73,224
87-90.99,102,326
84-86.99,75,401
80-83.99,117,518
70-79.99,174,692
60-69.99,127,819
0-59.99,25,844`;


// --- 3. 主應用程式元件 ---
export default function App() {
  const [rawScores, setRawScores] = useState('');
  const [gradeSettings, setGradeSettings] = useState(defaultSettings);
  const [distributionSettings, setDistributionSettings] = useState(defaultDistribution);
  
  // 介面狀態: 'input' | 'result' | 'settings_login' | 'settings'
  const [activeTab, setActiveTab] = useState('input'); 
  const [error, setError] = useState(null);
  
  // 管理員驗證狀態
  const [isAdminAuth, setIsAdminAuth] = useState(false);
  const [adminPassword, setAdminPassword] = useState('');
  
  const scoreFileInputRef = useRef(null);
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

  // 檔案上傳處理
  const handleFileUpload = (e, setter) => {
    const file = e.target.files[0];
    if (!file) return;
    const fileExt = file.name.split('.').pop().toLowerCase();

    if (fileExt === 'csv') {
      const reader = new FileReader();
      reader.onload = (event) => { setter(event.target.result); setError(null); };
      reader.onerror = () => setError("檔案讀取失敗。");
      reader.readAsText(file); 
    } else if (fileExt === 'xlsx' || fileExt === 'xls') {
      if (!window.XLSX) { setError("Excel 模組載入中，請稍後再試。"); return; }
      const reader = new FileReader();
      reader.onload = (event) => {
        try {
          const workbook = window.XLSX.read(new Uint8Array(event.target.result), { type: 'array' });
          setter(window.XLSX.utils.sheet_to_csv(workbook.Sheets[workbook.SheetNames[0]]));
          setError(null);
        } catch (err) { setError("讀取 Excel 失敗。"); }
      };
      reader.readAsArrayBuffer(file);
    } else {
      setError("不支援的檔案格式，請上傳 .xlsx 或 .csv。");
    }
    e.target.value = '';
  };

  const validStudentCount = useMemo(() => {
    if (!rawScores) return 0;
    return parseCSV(rawScores).filter(s => (s['姓名']?.trim() !== '') || (s['座號']?.toString().trim() !== '')).length;
  }, [rawScores]);

  // 核心處理邏輯
  const processedData = useMemo(() => {
    try {
      if (!rawScores || !gradeSettings) return null;
      let scoresData = parseCSV(rawScores).filter(s => (s['姓名']?.trim() !== '') || (s['座號']?.toString().trim() !== ''));
      if (scoresData.length === 0) return null;

      const { settings, subjects: adminSubjects } = processSettings(parseCSV(gradeSettings));
      const distMap = processDistribution(parseCSV(distributionSettings));
      
      // 動態抓取學生檔案的所有欄位 (排除基本欄位，剩下的就是科目)
      const excludeCols = ['座號', '姓名', '平均', '總分', '班排', '校排', '預估校排'];
      const studentHeaders = Object.keys(scoresData[0]);
      const studentSubjects = studentHeaders.filter(h => !excludeCols.includes(h));
      
      // 找出名字對不上的未知科目 (用於拋出警告)
      const unmappedSubjects = studentSubjects.filter(sub => !settings[sub]);
      
      const calculatedData = scoresData.map(student => {
        let totalScore = 0; let validCount = 0;
        const resultRow = { '座號': student['座號'] || '', '姓名': student['姓名'] || '' };

        studentSubjects.forEach(subject => {
          const val = student[subject];
          if (val !== undefined && val !== null && val.toString().trim() !== '') {
            const numScore = parseFloat(val);
            if (!isNaN(numScore)) {
              totalScore += numScore; validCount++;
              
              // 若管理員有設定該科目標籤，則進行標示；否則只顯示分數
              if (settings[subject]) {
                resultRow[subject] = `${numScore} (${getGradeLevel(numScore, settings[subject])})`;
              } else {
                resultRow[subject] = numScore; 
              }
            } else {
              resultRow[subject] = val; // 保留字串 (如: 缺考)
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
        const schoolRank = getSchoolRank(student['平均'], distMap);
        if (schoolRank) student['預估校排'] = schoolRank;
      });

      return { data: calculatedData, subjects: studentSubjects, unmappedSubjects };
    } catch (err) {
      setError("資料解析失敗，請確認標題與格式。"); return null;
    }
  }, [rawScores, gradeSettings, distributionSettings]);

  const handleGenerate = () => {
    if (processedData && processedData.data.length > 0) {
      setActiveTab('result'); setError(null);
    } else {
      setError("請先輸入有效的原始成績資料。");
    }
  };

  const handleCopy = () => {
    if (!processedData) return;
    const tsvContent = [Object.keys(processedData.data[0]).join('\t'), ...processedData.data.map(row => Object.values(row).join('\t'))].join('\n');
    const textArea = document.createElement("textarea");
    textArea.value = tsvContent;
    document.body.appendChild(textArea);
    textArea.select();
    try { document.execCommand('copy'); alert("✅ 已成功複製，可直接貼上至 Excel！"); } 
    catch (err) { alert("複製失敗。"); }
    document.body.removeChild(textArea);
  };

  const handleAdminLogin = (e) => {
    e.preventDefault();
    if (adminPassword === '690530') {
      setIsAdminAuth(true);
      setActiveTab('settings');
      setAdminPassword('');
      setError(null);
    } else {
      setError("密碼錯誤，請重新輸入。");
    }
  };

  const toggleAdminTab = () => {
    if (activeTab === 'input' || activeTab === 'result') {
      setActiveTab(isAdminAuth ? 'settings' : 'settings_login');
      setError(null);
    } else {
      setActiveTab('input');
      setError(null);
    }
  };

  return (
    <div className="min-h-screen bg-[#F4F7F9] text-slate-800 p-4 md:p-8 font-sans selection:bg-blue-100">
      <div className="max-w-[1200px] mx-auto space-y-6">
        
        {/* Header */}
        <header className="flex flex-col md:flex-row justify-between items-center bg-white p-5 rounded-3xl shadow-sm border border-slate-100/50">
          <div className="flex items-center space-x-4">
            <div className="bg-gradient-to-br from-blue-500 to-teal-400 p-3.5 rounded-2xl text-white shadow-md shadow-blue-500/20">
              <Calculator className="w-6 h-6" />
            </div>
            <div>
              <h1 className="text-2xl font-extrabold text-slate-800 tracking-tight">成績分析產生器</h1>
              <p className="text-sm text-slate-500 font-medium mt-0.5">自動套用全校標準 ‧ 智慧精算校排</p>
            </div>
          </div>
          <div className="mt-4 md:mt-0 flex flex-wrap gap-2">
            <button 
              onClick={toggleAdminTab}
              className={`px-5 py-2.5 rounded-xl font-semibold transition-all flex items-center ${
                activeTab.includes('settings') 
                  ? 'bg-slate-800 text-white shadow-md' 
                  : 'bg-white border-2 border-slate-100 text-slate-500 hover:border-slate-200 hover:text-slate-700'
              }`}
            >
              {activeTab.includes('settings') ? <><CheckCircle2 className="w-4 h-4 mr-2" /> 儲存並返回首頁</> : <><Settings className="w-4 h-4 mr-2" /> 管理員設定</>}
            </button>
          </div>
        </header>

        {error && (
          <div className="bg-red-50 text-red-600 px-6 py-4 rounded-2xl flex items-center border border-red-100 shadow-sm animate-in fade-in slide-in-from-top-2">
            <AlertCircle className="w-5 h-5 mr-3 flex-shrink-0" />
            <p className="font-medium">{error}</p>
          </div>
        )}

        {/* 主內容區塊 */}
        <main className="bg-white rounded-3xl shadow-sm border border-slate-100 overflow-hidden">
          
          {/* TAB 1: 老師輸入區 */}
          {activeTab === 'input' && (
            <div className="p-8 md:p-10 space-y-8 animate-in fade-in">
              
              {/* 資訊說明卡片 */}
              <div className="bg-gradient-to-r from-blue-50 to-teal-50 rounded-2xl p-6 border border-blue-100/50">
                <div className="flex items-start">
                  <Info className="w-6 h-6 text-teal-600 mr-3 mt-0.5 flex-shrink-0" />
                  <div>
                    <h3 className="text-lg font-bold text-slate-800 mb-2">如何使用？</h3>
                    <ul className="space-y-1.5 text-slate-700 text-sm font-medium">
                      <li>1. 請直接點擊右方按鈕 <strong>選擇 Excel 成績單</strong>，或將表格內容 <strong>複製貼上</strong> 到下方大空白處。</li>
                      <li>2. 系統會自動為各科成績標上 <strong>等級 (A++, B+ 等)</strong>，並透過線性內插法運算 <strong>「預估校排」</strong>。</li>
                      <li className="text-amber-700 mt-2 bg-amber-100/50 p-2 rounded-lg border border-amber-200/50">
                        ⚠️ <strong>重要提醒：</strong>成績單的「欄位名稱」與順序（座號、姓名、國文、英文、數學、社會、自然）必須與設定<strong>完全一致</strong>。若名稱不同（如：理化），會出現錯誤。
                      </li>
                    </ul>
                  </div>
                </div>
              </div>

              {/* 大輸入區 */}
              <div className="space-y-4">
                <div className="flex flex-col sm:flex-row justify-between items-start sm:items-end gap-4 border-b border-slate-100 pb-4">
                  <div className="flex items-center space-x-3">
                    <span className="text-xl font-bold text-slate-800">輸入班級成績</span>
                    {validStudentCount > 0 && (
                      <span className="px-3 py-1 bg-teal-100 text-teal-700 text-sm rounded-lg font-bold flex items-center shadow-sm">
                        <CheckCircle2 className="w-4 h-4 mr-1.5" /> 已成功讀取 {validStudentCount} 位學生
                      </span>
                    )}
                  </div>
                  <div>
                    <input type="file" accept=".csv, .xlsx, .xls" style={{ display: 'none' }} ref={scoreFileInputRef} onChange={(e) => handleFileUpload(e, setRawScores)} />
                    <button 
                      onClick={() => scoreFileInputRef.current.click()}
                      className="flex items-center px-6 py-2.5 bg-blue-600 hover:bg-blue-700 text-white font-semibold rounded-xl transition-all shadow-md shadow-blue-500/20 active:scale-95"
                    >
                      <FileSpreadsheet className="w-5 h-5 mr-2" /> 選擇 Excel / CSV 檔案
                    </button>
                  </div>
                </div>
                
                <textarea 
                  className="w-full min-h-[550px] p-6 bg-[#FAFAFA] border-2 border-slate-200 focus:border-blue-500 focus:ring-4 focus:ring-blue-500/10 rounded-2xl text-[15px] font-mono leading-relaxed transition-all resize-y shadow-inner text-slate-700"
                  value={rawScores}
                  onChange={(e) => setRawScores(e.target.value)}
                  placeholder="【直接貼上區】&#10;您也可以打開 Excel 框選成績單，直接在此處「貼上」！&#10;&#10;資料格式範例：&#10;座號, 姓名, 國文, 英文, 數學, 社會, 自然&#10;1, 王小明, 80, 90, 70, 85, 88&#10;2, 李小華, 92, 88, 95, 80, 90"
                />
              </div>

              {/* 產生按鈕 */}
              <div className="pt-2 flex justify-center">
                 <button 
                  onClick={handleGenerate}
                  className="px-10 py-4 bg-slate-800 text-white text-lg font-bold rounded-2xl hover:bg-slate-900 transition-all shadow-xl shadow-slate-900/20 flex items-center group disabled:opacity-50 disabled:cursor-not-allowed"
                  disabled={validStudentCount === 0}
                 >
                   套用全校標準並生成報表 
                   <RefreshCw className="w-5 h-5 ml-3 text-teal-400 group-hover:rotate-180 transition-transform duration-500" />
                 </button>
              </div>
            </div>
          )}

          {/* TAB: 管理員登入密碼區 */}
          {activeTab === 'settings_login' && (
            <div className="p-12 flex flex-col items-center justify-center min-h-[500px] animate-in zoom-in-95">
              <div className="bg-slate-50 p-6 rounded-full mb-6 shadow-inner">
                <Lock className="w-12 h-12 text-slate-400" />
              </div>
              <h2 className="text-2xl font-bold text-slate-800 mb-2">管理員身分驗證</h2>
              <p className="text-slate-500 mb-8 text-center">請輸入管理員密碼以設定全校等級與組距標準。</p>
              
              <form onSubmit={handleAdminLogin} className="flex flex-col w-full max-w-sm space-y-4">
                <input 
                  type="password" 
                  autoFocus
                  placeholder="請輸入密碼..." 
                  className="w-full px-5 py-3 text-center text-lg tracking-widest bg-white border-2 border-slate-200 rounded-xl focus:border-teal-500 focus:ring-4 focus:ring-teal-500/10 transition-all outline-none"
                  value={adminPassword}
                  onChange={(e) => setAdminPassword(e.target.value)}
                />
                <button 
                  type="submit"
                  className="w-full py-3 bg-teal-600 hover:bg-teal-700 text-white font-bold rounded-xl transition-all shadow-md flex justify-center items-center"
                >
                  解鎖並進入設定 <Unlock className="w-4 h-4 ml-2" />
                </button>
              </form>
            </div>
          )}

          {/* TAB: 管理員設定區 */}
          {activeTab === 'settings' && (
            <div className="p-8 md:p-10 space-y-8 animate-in fade-in">
              <div className="text-center space-y-2 mb-10 pb-6 border-b border-slate-100">
                <div className="inline-flex items-center justify-center bg-teal-100 text-teal-800 px-4 py-1.5 rounded-full text-sm font-bold mb-3">
                  <Unlock className="w-4 h-4 mr-2" /> 管理員模式已解鎖
                </div>
                <h2 className="text-3xl font-extrabold text-slate-800">全校標準設定金庫</h2>
                <p className="text-slate-500 font-medium">在此更新的數據將直接改變系統的計算邏輯。</p>
              </div>

              <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
                {/* 等級設定 */}
                <div className="space-y-4 bg-[#FFFCF5] p-6 rounded-3xl border border-amber-200/60 shadow-sm">
                  <div className="flex justify-between items-center">
                    <span className="text-lg font-bold text-amber-800 flex items-center">各科等級門檻</span>
                    <div>
                      <input type="file" accept=".csv, .xlsx, .xls" style={{ display: 'none' }} ref={settingFileInputRef} onChange={(e) => handleFileUpload(e, setGradeSettings)} />
                      <button onClick={() => settingFileInputRef.current.click()} className="flex items-center px-4 py-2 bg-white hover:bg-amber-50 text-amber-700 border border-amber-200 rounded-xl transition-colors font-semibold shadow-sm">
                        <Upload className="w-4 h-4 mr-2" /> 匯入 Excel
                      </button>
                    </div>
                  </div>
                  <textarea 
                    className="w-full h-96 p-5 bg-white border border-amber-200/80 focus:border-amber-400 focus:ring-4 focus:ring-amber-400/10 rounded-2xl text-sm font-mono transition-all resize-y shadow-inner text-slate-700"
                    value={gradeSettings} onChange={(e) => setGradeSettings(e.target.value)}
                  />
                </div>

                {/* 全校組距設定 */}
                <div className="space-y-4 bg-[#F5F9FF] p-6 rounded-3xl border border-blue-200/60 shadow-sm">
                  <div className="flex justify-between items-center">
                    <span className="text-lg font-bold text-blue-800 flex items-center">全校分數組距</span>
                    <div>
                      <input type="file" accept=".csv, .xlsx, .xls" style={{ display: 'none' }} ref={distFileInputRef} onChange={(e) => handleFileUpload(e, setDistributionSettings)} />
                      <button onClick={() => distFileInputRef.current.click()} className="flex items-center px-4 py-2 bg-white hover:bg-blue-50 text-blue-700 border border-blue-200 rounded-xl transition-colors font-semibold shadow-sm">
                        <Upload className="w-4 h-4 mr-2" /> 匯入 Excel
                      </button>
                    </div>
                  </div>
                  <textarea 
                    className="w-full h-96 p-5 bg-white border border-blue-200/80 focus:border-blue-400 focus:ring-4 focus:ring-blue-400/10 rounded-2xl text-sm font-mono transition-all resize-y shadow-inner text-slate-700"
                    value={distributionSettings} onChange={(e) => setDistributionSettings(e.target.value)}
                  />
                </div>
              </div>
            </div>
          )}

          {/* TAB 2: 結果區 */}
          {activeTab === 'result' && processedData && (
            <div className="p-8 md:p-10 space-y-6 animate-in slide-in-from-bottom-4">
              
              <div className="flex flex-col md:flex-row justify-between items-center bg-slate-50 p-4 rounded-2xl border border-slate-100">
                <h2 className="text-xl font-bold text-slate-800 flex items-center">
                  <span className="w-10 h-10 bg-teal-100 text-teal-600 rounded-full flex items-center justify-center mr-3 shadow-sm">
                    <CheckCircle2 className="w-6 h-6" />
                  </span>
                  報表生成完畢 <span className="text-slate-500 text-base font-medium ml-2"> (共 {processedData.data.length} 筆)</span>
                </h2>
                <div className="flex flex-wrap gap-3 mt-4 md:mt-0">
                  <button onClick={() => setActiveTab('input')} className="flex items-center px-5 py-2.5 bg-white text-slate-600 border border-slate-200 rounded-xl hover:bg-slate-50 transition-all shadow-sm font-bold">
                    <ArrowLeft className="w-4 h-4 mr-2" /> 修改資料
                  </button>
                  <button onClick={handleCopy} className="flex items-center px-5 py-2.5 bg-white text-blue-700 border border-blue-200 rounded-xl hover:bg-blue-50 transition-all shadow-sm font-bold">
                    <ClipboardCopy className="w-4 h-4 mr-2" /> 複製表格
                  </button>
                  <button onClick={() => exportToCSV(processedData.data, '成績分析結果.csv')} className="flex items-center px-5 py-2.5 bg-blue-600 text-white rounded-xl hover:bg-blue-700 transition-all shadow-md shadow-blue-500/20 font-bold">
                    <FileDown className="w-4 h-4 mr-2" /> 下載 CSV
                  </button>
                </div>
              </div>

              {/* 欄位名稱不符的警告提示 */}
              {processedData.unmappedSubjects.length > 0 && (
                <div className="bg-amber-50 border border-amber-200 p-4 rounded-xl flex items-start shadow-sm">
                  <AlertTriangle className="w-6 h-6 text-amber-500 mr-3 flex-shrink-0 mt-0.5" />
                  <div>
                    <h4 className="font-bold text-amber-800">發現未知科目名稱</h4>
                    <p className="text-sm text-amber-700 mt-1">
                      您的檔案中包含：<strong>{processedData.unmappedSubjects.join('、')}</strong>。由於管理員尚未設定這些科目的門檻，系統已將其計入平均，但<strong className="underline">無法標示等級</strong>。若需標示，請將檔案中的欄位名稱修改為與管理員設定完全一致（例如：將「理化」改為「自然」）。
                    </p>
                  </div>
                </div>
              )}

              <div className="overflow-x-auto overflow-y-auto max-h-[70vh] border-2 border-slate-100 rounded-2xl shadow-sm relative custom-scrollbar">
                <table className="w-full text-sm text-left whitespace-nowrap">
                  <thead className="text-[13px] text-slate-500 bg-slate-50 uppercase font-extrabold tracking-wider sticky top-0 z-10 shadow-[0_1px_0_0_#e2e8f0]">
                    <tr>
                      {Object.keys(processedData.data[0]).map((header, idx) => (
                        <th key={idx} className="px-6 py-4 bg-[#F8FAFC]">
                          {header}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {processedData.data.map((row, rowIndex) => (
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
            </div>
          )}
        </main>
      </div>
    </div>
  );
}