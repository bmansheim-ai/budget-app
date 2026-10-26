import React, { useState, useMemo, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import { 
  Upload, FileSpreadsheet, Copy, Check, Plus, Trash2, Tag, ChevronDown, 
  ChevronUp, Search, Download, FolderUp, CheckCircle2, 
  Scale, CreditCard, Building2, Table, UserCheck, Calendar, Folder, Settings, X, AlertTriangle, ArrowUpDown
} from "lucide-react";

// Exact 56 categories from the cat3 sheet
const DEFAULT_CATEGORIES = [
  "Arnona", "Bank charges", "Chashmal", "Car Service & Repairs", "Car Reg & Test", 
  "Car insurance Harel", "Chugim", "Natural gas", "Wedding", "Clothing & Shoes", 
  "Tzedaka", "Rav Kav/transport", "General Gifts", "Kid's/Grand Gifts", "Home Care", 
  "Home Repairs", "Water filters/Sherut plus", "School related", "Pharmacy/Medical", 
  "Vaad Bayit", "BY Brama", "Tmicha/tipulim", "Va'ad Habayit fund", "Beefland", 
  "Child support", "Misc ", "Chagim", "Health supplements", "Beauty Treatments", 
  "Braces", "Camps", "Cosmetics", "Couple's Gifts", "Couple's Vacation", "Dental", 
  "Entertainment", "Eye Care", "Family Vacation", "Pesach", "School Trips", 
  "Sefarim", "Sheital Care", "Sheital (New)", "Shtreimel Care", "Simcha Fund", 
  "Technology", "Uman other trips", "Uman Rosh Hashana", "US Trips", "House", 
  "Wealth Basket", "Safrus", "US taxes", "Kids monies", "Pension payout"
];

const DEFAULT_RULES = {
  "Pharmacy/Medical": ["בית מרקחת", "סופר פארם", "super-pharm", "סופר-פארם", "מכבי", "מאוחדת", "כללית"],
  "Car Service & Repairs": ["ספרינט מוטורס", "דלק", "פז", "סונול", "דור אלון", "מוסך", "שטיפה"],
  "Clothing & Shoes": ["מילא אופנה", "זארה", "zara", "קסטרו", "פוקס", "fox"],
  "Bank charges": ["דמי כרטיס", "עמלה", "דמי ניהול", "עמלת"],
  "Arnona": ["עיריית", "ארנונה"],
  "Chashmal": ["חברת החשמל", "חשמל"],
  "Natural gas": ["פזגז", "אמישראגז", "סופרגז"],
  "Rav Kav/transport": ["רב קו", "רב-קו", "רכבת", "אגד", "קווים"],
  "Misc ": ["סופרמרקט", "שופרסל", "רמי לוי", "אושר עד", "ויקטורי", "מחסני השוק", "יש חסד", "יש בשכונה"]
};

// Helper function to format Excel serial dates or Date objects into DD/MM/YYYY and extract Month-Year key & sortable timestamp
const formatDateAndMonth = (val) => {
  if (!val) return { dateStr: "", monthKey: "", timestamp: 0 };
  let dateObj = null;

  if (val instanceof Date) {
    dateObj = val;
  } else if (typeof val === 'number' && val > 30000 && val < 60000) {
    const excelEpoch = new Date(1899, 11, 30);
    dateObj = new Date(excelEpoch.getTime() + val * 86400000);
  } else if (typeof val === 'string' && val.includes('/')) {
    const parts = val.split('/');
    if (parts.length === 3) {
      dateObj = new Date(parts[2], parts[1] - 1, parts[0]);
    }
  }

  if (dateObj && !isNaN(dateObj.getTime())) {
    const day = String(dateObj.getDate()).padStart(2, '0');
    const month = String(dateObj.getMonth() + 1).padStart(2, '0');
    const year = dateObj.getFullYear();
    return { 
      dateStr: `${day}/${month}/${year}`, 
      monthKey: `${year}-${month}`,
      timestamp: dateObj.getTime()
    };
  }

  return { dateStr: String(val), monthKey: "", timestamp: 0 };
};

export default function App() {
  const [bankTransactions, setBankTransactions] = useState([]);
  const [ccTransactions, setCcTransactions] = useState([]);
  const [bankFileName, setBankFileName] = useState("");
  const [ccFileName, setCcFileName] = useState("");
  const [bankBalances, setBankBalances] = useState({ beginning: 0, ending: 0 });
  const [statementMonth, setStatementMonth] = useState("");
  const [savedMonthsList, setSavedMonthsList] = useState([]);
  
  // Settings Modal State
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [monthToPurge, setMonthToPurge] = useState("");

  // Table Sorting State
  const [bankSortConfig, setBankSortConfig] = useState({ key: 'timestamp', direction: 'asc' });
  const [ccSortConfig, setCcSortConfig] = useState({ key: 'timestamp', direction: 'asc' });

  // STATIC GLOBAL RULES (Applies across ALL months)
  const [rules, setRules] = useState(() => {
    const saved = localStorage.getItem("budget_rules");
    return saved ? JSON.parse(saved) : DEFAULT_RULES;
  });

  const [manualOverrides, setManualOverrides] = useState({});

  const [newKeyword, setNewKeyword] = useState("");
  const [selectedCategory, setSelectedCategory] = useState(DEFAULT_CATEGORIES[0]);
  const [copied, setCopied] = useState(false);
  const [activeTab, setActiveTab] = useState("summary");
  const [expandedCategory, setExpandedCategory] = useState(null);
  const [searchQuery, setSearchQuery] = useState("");
  const [tableSearchQuery, setTableSearchQuery] = useState("");
  
  const rulesSectionRef = useRef(null);
  const fileInputRulesRef = useRef(null);

  const refreshSavedMonths = () => {
    const keys = Object.keys(localStorage).filter(k => k.startsWith("budget_month_"));
    const months = keys.map(k => k.replace("budget_month_", "")).sort().reverse();
    setSavedMonthsList(months);
  };

  useEffect(() => {
    refreshSavedMonths();
  }, [statementMonth]);

  useEffect(() => {
    if (!statementMonth) return;
    const monthData = {
      bankTransactions,
      ccTransactions,
      bankFileName,
      ccFileName,
      bankBalances,
      manualOverrides
    };
    localStorage.setItem(`budget_month_${statementMonth}`, JSON.stringify(monthData));
    refreshSavedMonths();
  }, [statementMonth, bankTransactions, ccTransactions, bankFileName, ccFileName, bankBalances, manualOverrides]);

  const loadSavedMonth = (monthKey) => {
    if (!monthKey) return;
    const raw = localStorage.getItem(`budget_month_${monthKey}`);
    if (raw) {
      try {
        const parsed = JSON.parse(raw);
        setBankTransactions(parsed.bankTransactions || []);
        setCcTransactions(parsed.ccTransactions || []);
        setBankFileName(parsed.bankFileName || "");
        setCcFileName(parsed.ccFileName || "");
        setBankBalances(parsed.bankBalances || { beginning: 0, ending: 0 });
        setManualOverrides(parsed.manualOverrides || {});
        setStatementMonth(monthKey);
      } catch (err) {
        console.error("Failed to load month progress", err);
      }
    }
  };

  const handlePurgeMonthData = (monthKey) => {
    if (!monthKey) return;
    if (window.confirm(`Are you sure you want to delete all saved data and progress for month ${monthKey}?`)) {
      localStorage.removeItem(`budget_month_${monthKey}`);
      if (statementMonth === monthKey) {
        setBankTransactions([]);
        setCcTransactions([]);
        setBankFileName("");
        setCcFileName("");
        setBankBalances({ beginning: 0, ending: 0 });
        setManualOverrides({});
        setStatementMonth("");
      }
      refreshSavedMonths();
      setMonthToPurge("");
      alert(`Data for month ${monthKey} has been purged successfully.`);
    }
  };

  const saveRules = (updatedRules) => {
    setRules(updatedRules);
    localStorage.setItem("budget_rules", JSON.stringify(updatedRules));
  };

  const handleManualOverride = (txnKey, category) => {
    const updated = { ...manualOverrides };
    if (category === "Uncategorized") {
      delete updated[txnKey];
    } else {
      updated[txnKey] = category;
    }
    setManualOverrides(updated);
  };

  const handleAddRule = (e) => {
    e.preventDefault();
    if (!newKeyword.trim()) return;
    const updated = { ...rules };
    if (!updated[selectedCategory]) updated[selectedCategory] = [];
    if (!updated[selectedCategory].includes(newKeyword.trim())) {
      updated[selectedCategory].push(newKeyword.trim());
      saveRules(updated);
    }
    setNewKeyword("");
    setActiveTab("uncategorized");
  };

  const handleRemoveRule = (category, keyword) => {
    const updated = { ...rules };
    updated[category] = updated[category].filter(k => k !== keyword);
    saveRules(updated);
  };

  const exportRulesFile = () => {
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(rules, null, 2));
    const downloadAnchor = document.createElement('a');
    downloadAnchor.setAttribute("href", dataStr);
    downloadAnchor.setAttribute("download", "budget_rules.json");
    document.body.appendChild(downloadAnchor);
    downloadAnchor.click();
    downloadAnchor.remove();
  };

  const importRulesFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const parsed = JSON.parse(evt.target.result);
        saveRules(parsed);
        alert("Rules imported successfully!");
      } catch (err) {
        alert("Invalid JSON rules file.");
      }
    };
    reader.readAsText(file);
  };

  const handleMultiFileUpload = (e) => {
    const files = Array.from(e.target.files);
    if (files.length === 0) return;

    files.forEach(file => {
      const reader = new FileReader();
      reader.onload = (evt) => {
        const bstr = evt.target.result;
        const wb = XLSX.read(bstr, { type: "binary" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

        const isBank = data.some(row => row && row.some(cell => String(cell).includes("₪ יתרה")));
        const isCc = data.some(row => row && row.some(cell => String(cell).includes("בית עסק")));

        if (isBank) {
          setBankFileName(file.name);
          let headerIdx = data.findIndex(row => row && row.some(cell => String(cell).includes("תאריך")));
          if (headerIdx === -1) headerIdx = 7;
          const rows = data.slice(headerIdx + 1);
          
          let detectedMonth = "";
          const parsed = rows.map((r, idx) => {
            if (!r || r.length < 5) return null;
            const { dateStr, monthKey, timestamp } = formatDateAndMonth(r[0]);
            if (monthKey && !detectedMonth) detectedMonth = monthKey;
            const desc = String(r[2] || "");
            const amount = parseFloat(r[3]) || 0;
            const balance = parseFloat(r[4]) || 0;
            const txnKey = `${r[0]}_${desc}_${amount}_${idx}`;
            return { 
              id: txnKey,
              date: dateStr, 
              timestamp,
              desc: desc, 
              amount: amount, 
              balance: balance,
              source: "Bank" 
            };
          }).filter(item => item && item.desc && !isNaN(item.amount));

          setBankTransactions(parsed);
          if (detectedMonth) {
            setStatementMonth(detectedMonth);
            const existingMonthData = localStorage.getItem(`budget_month_${detectedMonth}`);
            if (existingMonthData) {
              try {
                const parsedMonth = JSON.parse(existingMonthData);
                setManualOverrides(parsedMonth.manualOverrides || {});
              } catch(e) {}
            } else {
              setManualOverrides({});
            }
          }

          if (parsed.length > 0) {
            const firstTxn = parsed[0];
            const lastTxn = parsed[parsed.length - 1];
            const begBal = firstTxn.balance - firstTxn.amount;
            const endBal = lastTxn.balance;
            setBankBalances({ beginning: begBal, ending: endBal });
          }
        } else if (isCc || !isBank) {
          setCcFileName(file.name);
          
          let headerIdx = data.findIndex(row => row && row.some(cell => String(cell).includes("בית עסק")));
          if (headerIdx === -1) headerIdx = 7;
          const headerRow = data[headerIdx] || [];
          let cardColIdx = headerRow.findIndex(cell => String(cell).includes("כרטיס"));

          const rows = data.slice(headerIdx + 1);
          
          const parsed = rows.map((r, idx) => {
            if (!r || r.length < 9) return null;
            let amt = parseFloat(r[8]) || 0;
            if (amt > 0) amt = -amt;
            const { dateStr, timestamp } = formatDateAndMonth(r[2]);
            const desc = String(r[1] || "");

            let ccLast4 = "CC";
            if (cardColIdx !== -1 && r[cardColIdx] !== undefined) {
              const cardVal = String(r[cardColIdx]);
              const matchDigits = cardVal.match(/\d{4}/);
              if (matchDigits) {
                ccLast4 = matchDigits[0];
              } else if (cardVal.trim().length > 0) {
                ccLast4 = cardVal.trim().slice(-4);
              }
            } else {
              const fileMatch = file.name.match(/\d{4}/);
              if (fileMatch) ccLast4 = fileMatch[0].slice(-4);
            }

            const txnKey = `${r[2]}_${desc}_${amt}_cc_${idx}`;
            return { 
              id: txnKey,
              date: dateStr, 
              timestamp,
              desc: desc, 
              amount: amt, 
              ccLast4: ccLast4,
              source: "Credit Card" 
            };
          }).filter(item => item && item.desc && !isNaN(item.amount));
          setCcTransactions(parsed);
        }
      };
      reader.readAsBinaryString(file);
    });
  };

  const getCategorizationDetails = (txn) => {
    if (manualOverrides[txn.id]) {
      return { category: manualOverrides[txn.id], type: 'manual' };
    }
    const lowerDesc = String(txn.desc).toLowerCase();
    for (const [category, keywords] of Object.entries(rules)) {
      if (keywords.some(k => lowerDesc.includes(k.toLowerCase()))) {
        return { category, type: 'rule' };
      }
    }
    return { category: 'Uncategorized', type: 'none' };
  };

  const nonCcBankTransactions = useMemo(() => {
    return bankTransactions.filter(t => {
      const desc = t.desc;
      return !desc.includes("חיוב לכרטיס") && !desc.includes("מזרחי כאל") && !desc.includes("חיוב מ כ.א.ל");
    });
  }, [bankTransactions]);

  const allTransactions = useMemo(() => {
    return [...nonCcBankTransactions, ...ccTransactions].map(t => {
      const { category, type } = getCategorizationDetails(t);
      return { 
        ...t, 
        category,
        categorizationType: type 
      };
    });
  }, [nonCcBankTransactions, ccTransactions, rules, manualOverrides]);

  const getOrderedCategoriesForTxn = () => {
    const previouslyUsedSet = new Set(Object.values(manualOverrides));
    const savedKeys = Object.keys(localStorage).filter(k => k.startsWith("budget_month_"));
    savedKeys.forEach(k => {
      try {
        const data = JSON.parse(localStorage.getItem(k));
        if (data && data.manualOverrides) {
          Object.values(data.manualOverrides).forEach(cat => previouslyUsedSet.add(cat));
        }
      } catch(e) {}
    });

    const priorList = Array.from(previouslyUsedSet);
    const remainingList = DEFAULT_CATEGORIES.filter(c => !priorList.includes(c));
    
    return {
      prior: priorList,
      remaining: remainingList
    };
  };

  const categoryTotals = useMemo(() => {
    const totals = {};
    DEFAULT_CATEGORIES.forEach(cat => totals[cat] = 0);
    totals["Uncategorized"] = 0;
    allTransactions.forEach(t => {
      if (totals[t.category] !== undefined) totals[t.category] += t.amount;
      else totals["Uncategorized"] += t.amount;
    });
    return totals;
  }, [allTransactions]);

  const transactionsByCategory = useMemo(() => {
    const map = {};
    DEFAULT_CATEGORIES.forEach(cat => map[cat] = []);
    map["Uncategorized"] = [];
    allTransactions.forEach(t => {
      if (map[t.category]) map[t.category].push(t);
      else map["Uncategorized"].push(t);
    });
    return map;
  }, [allTransactions]);

  const uncategorizedList = useMemo(() => {
    return allTransactions.filter(t => t.category === "Uncategorized");
  }, [allTransactions]);

  const totalExpenseSum = useMemo(() => {
    return Object.values(categoryTotals).reduce((a, b) => a + Math.abs(b), 0);
  }, [categoryTotals]);

  // Reconciliation Audit Values
  const bankBalanceDifference = bankBalances.ending - bankBalances.beginning;
  const totalBankStatementNet = useMemo(() => {
    return bankTransactions.reduce((acc, t) => acc + t.amount, 0);
  }, [bankTransactions]);
  const totalCategorizedNet = useMemo(() => {
    return allTransactions.reduce((acc, t) => acc + t.amount, 0);
  }, [allTransactions]);
  const reconciliationDifference = bankBalanceDifference - totalCategorizedNet;

  // Credit Card Audit Values
  const ccPaymentBankItems = useMemo(() => {
    return bankTransactions.filter(t => {
      const desc = t.desc;
      return desc.includes("חיוב לכרטיס") || desc.includes("מזרחי כאל") || desc.includes("חיוב מ כ.א.ל");
    });
  }, [bankTransactions]);

  const totalCcPaymentsFromBank = useMemo(() => {
    return ccPaymentBankItems.reduce((acc, t) => acc + Math.abs(t.amount), 0);
  }, [ccPaymentBankItems]);

  const totalCcStatementSum = useMemo(() => {
    return ccTransactions.reduce((acc, t) => acc + Math.abs(t.amount), 0);
  }, [ccTransactions]);

  const ccMatchingDifference = totalCcPaymentsFromBank - totalCcStatementSum;

  const copyForCat3SumFormulaWithSpacing = () => {
    const formattedRows = [];
    formattedRows.push("");

    DEFAULT_CATEGORIES.forEach((cat) => {
      const txns = transactionsByCategory[cat] || [];
      if (txns.length === 0) {
        formattedRows.push("");
      } else if (txns.length === 1) {
        formattedRows.push(`=${txns[0].amount.toFixed(2)}`);
      } else {
        const items = txns.map(t => t.amount.toFixed(2)).join(", ");
        formattedRows.push(`=-SUM(${items})`);
      }
      formattedRows.push("");
    });

    const outputText = formattedRows.join("\n");
    navigator.clipboard.writeText(outputText);
    setCopied(true);
    setTimeout(() => setCopied(false), 2500);
  };

  const sortedAndFilteredCategories = useMemo(() => {
    let baseList = DEFAULT_CATEGORIES;
    if (searchQuery.trim()) {
      baseList = DEFAULT_CATEGORIES.filter(cat => 
        cat.toLowerCase().includes(searchQuery.toLowerCase()) ||
        (transactionsByCategory[cat] && transactionsByCategory[cat].some(t => t.desc.toLowerCase().includes(searchQuery.toLowerCase())))
      );
    }

    const withExpenses = baseList.filter(cat => (categoryTotals[cat] || 0) !== 0);
    const withoutExpenses = baseList.filter(cat => (categoryTotals[cat] || 0) === 0);

    return [...withExpenses, ...withoutExpenses];
  }, [searchQuery, categoryTotals, transactionsByCategory]);

  const handleBankSort = (key) => {
    let direction = 'asc';
    if (bankSortConfig.key === key && bankSortConfig.direction === 'asc') {
      direction = 'desc';
    }
    setBankSortConfig({ key, direction });
  };

  const handleCcSort = (key) => {
    let direction = 'asc';
    if (ccSortConfig.key === key && ccSortConfig.direction === 'asc') {
      direction = 'desc';
    }
    setCcSortConfig({ key, direction });
  };

  const filteredBankTable = useMemo(() => {
    let list = bankTransactions;
    if (tableSearchQuery.trim()) {
      list = list.filter(t => t.desc.toLowerCase().includes(tableSearchQuery.toLowerCase()));
    }
    return [...list].sort((a, b) => {
      let aVal = a[bankSortConfig.key];
      let bVal = b[bankSortConfig.key];
      if (aVal === undefined) aVal = "";
      if (bVal === undefined) bVal = "";
      if (typeof aVal === 'string') {
        return bankSortConfig.direction === 'asc' ? aVal.localeCompare(bVal) : bVal.localeCompare(aVal);
      }
      return bankSortConfig.direction === 'asc' ? aVal - bVal : bVal - aVal;
    });
  }, [bankTransactions, tableSearchQuery, bankSortConfig]);

  const filteredCcTable = useMemo(() => {
    let list = ccTransactions;
    if (tableSearchQuery.trim()) {
      list = list.filter(t => t.desc.toLowerCase().includes(tableSearchQuery.toLowerCase()) || t.ccLast4.includes(tableSearchQuery.trim()));
    }
    return [...list].sort((a, b) => {
      let aVal = a[ccSortConfig.key];
      let bVal = b[ccSortConfig.key];
      if (aVal === undefined) aVal = "";
      if (bVal === undefined) bVal = "";
      if (typeof aVal === 'string') {
        return ccSortConfig.direction === 'asc' ? aVal.localeCompare(bVal) : bVal.localeCompare(aVal);
      }
      return ccSortConfig.direction === 'asc' ? aVal - bVal : bVal - aVal;
    });
  }, [ccTransactions, tableSearchQuery, ccSortConfig]);

  const goToRulesForKeyword = (keyword) => {
    setNewKeyword(keyword);
    setActiveTab("rules");
    setTimeout(() => {
      if (rulesSectionRef.current) {
        rulesSectionRef.current.scrollIntoView({ behavior: 'smooth' });
      }
    }, 100);
  };

  return (
    <div className="min-h-screen bg-slate-100 text-slate-800 font-sans pb-12">
      {/* Top Banner Header */}
      <div className="bg-gradient-to-r from-slate-900 via-blue-900 to-indigo-900 text-white shadow-md py-4 px-6 mb-6">
        <div className="max-w-7xl mx-auto flex flex-col sm:flex-row justify-between items-center gap-3">
          <div className="flex items-center gap-3">
            <div className="p-2 bg-blue-600 rounded-xl shadow-inner">
              <FileSpreadsheet className="w-6 h-6 text-white" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h1 className="text-xl font-extrabold tracking-tight">Mansheim Budget Processor</h1>
                {statementMonth && (
                  <span className="bg-blue-800 text-blue-200 text-xs font-semibold px-2.5 py-0.5 rounded-full flex items-center gap-1 border border-blue-700">
                    <Calendar className="w-3 h-3" /> Month: {statementMonth}
                  </span>
                )}
              </div>
              <p className="text-slate-300 text-xs">Automated Statement Categorization & Reconciliation</p>
            </div>
          </div>

          <div className="flex items-center gap-3">
            <div className="flex items-center gap-1.5 bg-blue-950/80 px-3 py-2 rounded-xl border border-blue-700/60 shadow-inner">
              <Folder className="w-4 h-4 text-blue-400" />
              <select
                value={statementMonth}
                onChange={(e) => loadSavedMonth(e.target.value)}
                className="bg-transparent text-xs font-semibold text-white outline-none cursor-pointer"
              >
                <option value="" className="text-slate-900">-- Select Saved Month --</option>
                {savedMonthsList.map(m => (
                  <option key={m} value={m} className="text-slate-900">Month: {m}</option>
                ))}
              </select>
            </div>

            <button
              onClick={() => setIsSettingsOpen(true)}
              className="p-2 bg-blue-800 hover:bg-blue-700 text-white rounded-xl shadow transition"
              title="Settings"
            >
              <Settings className="w-5 h-5" />
            </button>

            <button
              onClick={copyForCat3SumFormulaWithSpacing}
              className={`flex items-center gap-2 px-4 py-2 rounded-xl text-sm font-semibold shadow transition-all transform active:scale-95 ${
                copied ? 'bg-emerald-600 text-white' : 'bg-blue-600 hover:bg-blue-500 text-white'
              }`}
            >
              {copied ? <Check className="w-4 h-4" /> : <Copy className="w-4 h-4" />}
              {copied ? 'Copied!' : 'Copy cat3 Column'}
            </button>
          </div>
        </div>
      </div>

      {/* Settings Modal */}
      {isSettingsOpen && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl shadow-xl max-w-md w-full p-6 space-y-6 border border-slate-200">
            <div className="flex justify-between items-center border-b border-slate-100 pb-4">
              <div className="flex items-center gap-2">
                <div className="p-2 bg-blue-50 text-blue-600 rounded-lg">
                  <Settings className="w-5 h-5" />
                </div>
                <h3 className="text-base font-bold text-slate-900">Application Settings</h3>
              </div>
              <button 
                onClick={() => setIsSettingsOpen(false)}
                className="p-1.5 text-slate-400 hover:text-slate-600 rounded-lg hover:bg-slate-100 transition"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="space-y-3">
              <div className="flex items-center gap-2 text-rose-600">
                <AlertTriangle className="w-4 h-4" />
                <h4 className="text-xs font-bold uppercase tracking-wider">Purge Data for Specific Month</h4>
              </div>
              <p className="text-xs text-slate-500">
                Permanently delete all statements, transactions, and manual overrides associated with a specific saved month.
              </p>

              <div className="space-y-2 pt-1">
                <label className="block text-xs font-semibold text-slate-700">Select Month to Purge</label>
                <select
                  value={monthToPurge}
                  onChange={(e) => setMonthToPurge(e.target.value)}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl text-xs outline-none focus:border-blue-500 transition"
                >
                  <option value="">-- Choose Month --</option>
                  {savedMonthsList.map(m => (
                    <option key={m} value={m}>Month: {m}</option>
                  ))}
                </select>

                <button
                  onClick={() => handlePurgeMonthData(monthToPurge)}
                  disabled={!monthToPurge}
                  className={`w-full mt-2 py-2.5 rounded-xl text-xs font-semibold shadow-xs transition flex items-center justify-center gap-2 ${
                    monthToPurge 
                      ? 'bg-rose-600 hover:bg-rose-500 text-white cursor-pointer' 
                      : 'bg-slate-200 text-slate-400 cursor-not-allowed'
                  }`}
                >
                  <Trash2 className="w-4 h-4" /> Purge Selected Month Data
                </button>
              </div>
            </div>

            <div className="border-t border-slate-100 pt-4 flex justify-end">
              <button
                onClick={() => setIsSettingsOpen(false)}
                className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-semibold transition"
              >
                Close Settings
              </button>
            </div>
          </div>
        </div>
      )}

      <div className="max-w-7xl mx-auto px-4 space-y-4">
        
        {/* Combined Single-Action Upload Toolbar */}
        <div className="bg-white p-4 rounded-xl shadow-xs border border-slate-200/80 flex flex-col sm:flex-row items-center justify-between gap-4">
          <div className="flex items-center gap-3">
            <div className="p-2 bg-indigo-50 text-indigo-600 rounded-lg">
              <Upload className="w-5 h-5" />
            </div>
            <div>
              <h2 className="text-xs font-bold text-slate-800">Upload Statements (Bank & Credit Card)</h2>
              <p className="text-[11px] text-slate-400">Select both Excel files simultaneously in one action</p>
            </div>
          </div>

          <div className="flex items-center gap-3">
            {bankFileName && (
              <span className="text-[11px] bg-emerald-50 text-emerald-700 font-medium px-2.5 py-1 rounded-lg flex items-center gap-1 border border-emerald-200">
                <CheckCircle2 className="w-3.5 h-3.5" /> Bank: {bankFileName}
              </span>
            )}
            {ccFileName && (
              <span className="text-[11px] bg-emerald-50 text-emerald-700 font-medium px-2.5 py-1 rounded-lg flex items-center gap-1 border border-emerald-200">
                <CheckCircle2 className="w-3.5 h-3.5" /> CC: {ccFileName}
              </span>
            )}
            <label className="cursor-pointer bg-blue-600 hover:bg-blue-500 text-white px-4 py-2 rounded-xl text-xs font-semibold shadow-xs transition flex items-center gap-2">
              <Upload className="w-4 h-4" /> Upload Statements
              <input type="file" accept=".xlsx, .xls" multiple onChange={handleMultiFileUpload} className="hidden" />
            </label>
          </div>
        </div>

        {/* Tab Bar Navigation */}
        <div className="bg-white rounded-xl shadow-xs border border-slate-200/80 p-3 flex flex-col md:flex-row justify-between items-center gap-3">
          <div className="flex gap-2 w-full md:w-auto overflow-x-auto">
            <button
              onClick={() => setActiveTab('summary')}
              className={`px-4 py-2 rounded-lg text-xs font-semibold transition ${
                activeTab === 'summary' ? 'bg-blue-600 text-white shadow-xs' : 'text-slate-600 hover:bg-slate-100'
              }`}
            >
              Expenses ({DEFAULT_CATEGORIES.length})
            </button>
            <button
              onClick={() => setActiveTab('reconciliation')}
              className={`px-4 py-2 rounded-lg text-xs font-semibold transition flex items-center gap-1.5 ${
                activeTab === 'reconciliation' ? 'bg-blue-600 text-white shadow-xs' : 'text-slate-600 hover:bg-slate-100'
              }`}
            >
              <Scale className="w-3.5 h-3.5" /> Bank & CC Reconciliation
            </button>
            <button
              onClick={() => setActiveTab('uncategorized')}
              className={`px-4 py-2 rounded-lg text-xs font-semibold transition flex items-center gap-2 ${
                activeTab === 'uncategorized' ? 'bg-blue-600 text-white shadow-xs' : 'text-slate-600 hover:bg-slate-100'
              }`}
            >
              Uncategorized
              {uncategorizedList.length > 0 && (
                <span className="bg-amber-500 text-white text-[10px] px-1.5 py-0.5 rounded-full font-bold">
                  {uncategorizedList.length}
                </span>
              )}
            </button>
            <button
              onClick={() => setActiveTab('rules')}
              className={`px-4 py-2 rounded-lg text-xs font-semibold transition ${
                activeTab === 'rules' ? 'bg-blue-600 text-white shadow-xs' : 'text-slate-600 hover:bg-slate-100'
              }`}
            >
              Rules Manager
            </button>
          </div>

          {activeTab === 'summary' && (
            <div className="relative w-full md:w-64">
              <Search className="absolute left-3 top-2.5 w-3.5 h-3.5 text-slate-400" />
              <input
                type="text"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                placeholder="Search expenses or merchants..."
                className="w-full pl-9 pr-3 py-1.5 bg-slate-50 border border-slate-200 rounded-lg text-xs outline-none focus:border-blue-500 transition"
              />
            </div>
          )}
        </div>

        {/* TAB 1: EXPENSES TAB */}
        {activeTab === 'summary' && (
          <div className="space-y-4">
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
              {sortedAndFilteredCategories.map((cat) => {
                const amount = categoryTotals[cat] || 0;
                const txns = transactionsByCategory[cat] || [];
                const isExpanded = expandedCategory === cat;
                const hasExpenses = amount !== 0;

                return (
                  <div
                    key={cat}
                    className={`bg-white rounded-xl border transition-all shadow-2xs overflow-hidden ${
                      hasExpenses ? 'border-blue-300 ring-1 ring-blue-100 hover:border-blue-400' : 'border-slate-200/80 opacity-75 hover:opacity-100'
                    }`}
                  >
                    <div 
                      onClick={() => setExpandedCategory(isExpanded ? null : cat)}
                      className={`p-3.5 cursor-pointer flex justify-between items-center transition ${
                        hasExpenses ? 'bg-gradient-to-r from-blue-50/40 to-white hover:bg-blue-50/70' : 'bg-slate-50/40 hover:bg-slate-50'
                      }`}
                    >
                      <div className="flex items-center gap-2.5">
                        <span className={`text-[11px] font-mono font-bold px-1.5 py-0.5 rounded ${
                          hasExpenses ? 'text-blue-700 bg-blue-100/70' : 'text-slate-400 bg-slate-100'
                        }`}>
                          #{DEFAULT_CATEGORIES.indexOf(cat) + 1}
                        </span>
                        <div>
                          <h3 className="text-xs font-bold text-slate-900">{cat}</h3>
                          <p className="text-[10px] text-slate-400">{txns.length} transaction{txns.length !== 1 ? 's' : ''}</p>
                        </div>
                      </div>
                      <div className="flex items-center gap-2.5">
                        <span className={`text-xs font-mono font-bold ${amount < 0 ? 'text-rose-700' : amount > 0 ? 'text-blue-700' : 'text-slate-400'}`}>
                          ₪{amount.toFixed(2)}
                        </span>
                        {txns.length > 0 && (
                          <button className="p-0.5 rounded hover:bg-slate-200/60 text-slate-500 transition">
                            {isExpanded ? <ChevronUp className="w-3.5 h-3.5" /> : <ChevronDown className="w-3.5 h-3.5" />}
                          </button>
                        )}
                      </div>
                    </div>

                    {isExpanded && txns.length > 0 && (
                      <div className="border-t border-slate-100 bg-slate-50/90 p-3 space-y-2 max-h-56 overflow-y-auto">
                        <p className="text-[10px] font-semibold text-slate-500 uppercase tracking-wider mb-1">Individual Transactions:</p>
                        {txns.map((t) => {
                          const { prior, remaining } = getOrderedCategoriesForTxn();
                          return (
                            <div key={t.id} className="bg-white p-2.5 rounded-lg border border-slate-200/60 flex flex-col gap-2 text-[11px] shadow-2xs">
                              <div className="flex justify-between items-center">
                                <div className="space-y-0.5 pr-2 truncate">
                                  <p className="font-semibold text-slate-800 truncate">
                                    {t.desc} {t.ccLast4 ? `(CC: ${t.ccLast4})` : ""}
                                  </p>
                                  <div className="flex items-center gap-1.5 text-slate-400 text-[10px]">
                                    <span>{t.date}</span>
                                    <span>•</span>
                                    <span className="bg-slate-100 px-1 py-0.5 rounded text-[9px] font-medium">{t.source}</span>
                                    {t.categorizationType === 'rule' && (
                                      <span className="bg-blue-50 text-blue-700 px-1.5 py-0.5 rounded text-[9px] font-semibold">Matched by Rule</span>
                                    )}
                                    {t.categorizationType === 'manual' && (
                                      <span className="bg-indigo-50 text-indigo-700 px-1.5 py-0.5 rounded text-[9px] font-semibold">Manual Override</span>
                                    )}
                                  </div>
                                </div>
                                <span className={`font-mono font-bold whitespace-nowrap ${t.amount < 0 ? 'text-rose-600' : 'text-slate-900'}`}>
                                  ₪{t.amount.toFixed(2)}
                                </span>
                              </div>

                              {/* Manual Override Selector */}
                              <div className="flex items-center justify-between border-t border-slate-100 pt-1.5">
                                <span className="text-[10px] text-slate-400 flex items-center gap-1">
                                  <UserCheck className="w-3 h-3 text-blue-500" /> Override:
                                </span>
                                <select
                                  value={t.category}
                                  onChange={(e) => handleManualOverride(t.id, e.target.value)}
                                  className="bg-slate-50 border border-slate-200 rounded text-[10px] py-0.5 px-1 font-medium text-slate-700 outline-none focus:border-blue-500"
                                >
                                  <option value="Uncategorized">-- Uncategorized --</option>
                                  {prior.length > 0 && (
                                    <optgroup label="⭐ Previously Used">
                                      {prior.map(c => <option key={c} value={c}>{c}</option>)}
                                    </optgroup>
                                  )}
                                  <optgroup label="All Categories">
                                    {remaining.map(c => <option key={c} value={c}>{c}</option>)}
                                  </optgroup>
                                </select>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {/* TAB 2: BANK & CC RECONCILIATION TAB */}
        {activeTab === 'reconciliation' && (
          <div className="space-y-6">
            
            {/* 1. Bank Reconciliation Card */}
            <div className="bg-white rounded-xl shadow-xs border border-blue-200 p-5 space-y-4">
              <div className="flex justify-between items-center border-b border-slate-100 pb-3">
                <h3 className="text-sm font-bold text-slate-900 flex items-center gap-2">
                  <Building2 className="w-4 h-4 text-blue-600" /> Bank Balance Reconciliation Audit
                </h3>
                <span className="text-xs bg-blue-50 text-blue-700 px-2.5 py-1 rounded-full font-mono font-bold">
                  Statement Net: ₪{totalBankStatementNet.toFixed(2)}
                </span>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-4 gap-4 text-xs">
                <div className="bg-slate-50 p-3 rounded-xl border border-slate-200">
                  <p className="text-slate-500 font-medium">Beginning Balance</p>
                  <p className="text-base font-mono font-bold text-slate-800 mt-1">₪{bankBalances.beginning.toFixed(2)}</p>
                </div>
                <div className="bg-slate-50 p-3 rounded-xl border border-slate-200">
                  <p className="text-slate-500 font-medium">Ending Balance</p>
                  <p className="text-base font-mono font-bold text-slate-800 mt-1">₪{bankBalances.ending.toFixed(2)}</p>
                </div>
                <div className="bg-blue-50/50 p-3 rounded-xl border border-blue-100">
                  <p className="text-blue-700 font-medium">Target Difference (End - Beg)</p>
                  <p className={`text-base font-mono font-bold mt-1 ${bankBalanceDifference < 0 ? 'text-rose-600' : 'text-blue-900'}`}>
                    ₪{bankBalanceDifference.toFixed(2)}
                  </p>
                </div>
                <div className={`p-3 rounded-xl border ${Math.abs(reconciliationDifference) < 0.01 ? 'bg-emerald-50/50 border-emerald-200 text-emerald-800' : 'bg-amber-50/50 border-amber-200 text-amber-800'}`}>
                  <p className="font-medium">Unaccounted Difference</p>
                  <p className="text-base font-mono font-bold mt-1">
                    ₪{reconciliationDifference.toFixed(2)}
                  </p>
                </div>
              </div>
            </div>

            {/* 2. CC Verification Card */}
            <div className="bg-white rounded-xl shadow-xs border border-emerald-200 p-5 space-y-4">
              <div className="flex justify-between items-center border-b border-slate-100 pb-3">
                <h3 className="text-sm font-bold text-slate-900 flex items-center gap-2">
                  <CreditCard className="w-4 h-4 text-emerald-600" /> Credit Card Deductions vs Statement Verification
                </h3>
                <span className="text-xs bg-emerald-50 text-emerald-700 px-2.5 py-1 rounded-full font-mono font-bold">
                  {ccPaymentBankItems.length} Bank CC Charges Found
                </span>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 text-xs">
                <div className="bg-slate-50 p-3 rounded-xl border border-slate-200">
                  <p className="text-slate-500 font-medium">Total Deducted from Bank</p>
                  <p className="text-base font-mono font-bold text-slate-800 mt-1">₪{totalCcPaymentsFromBank.toFixed(2)}</p>
                </div>
                <div className="bg-slate-50 p-3 rounded-xl border border-slate-200">
                  <p className="text-slate-500 font-medium">Total CC Statement Sum</p>
                  <p className="text-base font-mono font-bold text-slate-800 mt-1">₪{totalCcStatementSum.toFixed(2)}</p>
                </div>
                <div className={`p-3 rounded-xl border ${Math.abs(ccMatchingDifference) < 0.01 ? 'bg-emerald-50/50 border-emerald-200 text-emerald-800' : 'bg-amber-50/50 border-amber-200 text-amber-800'}`}>
                  <p className="font-medium">Matching Difference</p>
                  <p className="text-base font-mono font-bold mt-1">
                    ₪{ccMatchingDifference.toFixed(2)} {Math.abs(ccMatchingDifference) < 0.01 ? "(Perfect Match!)" : ""}
                  </p>
                </div>
              </div>
            </div>

            {/* Table Search */}
            <div className="flex justify-between items-center bg-white p-3 rounded-xl border border-slate-200">
              <span className="text-xs font-bold text-slate-700 flex items-center gap-1.5">
                <Table className="w-4 h-4 text-slate-500" /> Statement Transactions Tables (Sortable by clicking headers)
              </span>
              <div className="relative w-64">
                <Search className="absolute left-3 top-2.5 w-3.5 h-3.5 text-slate-400" />
                <input
                  type="text"
                  value={tableSearchQuery}
                  onChange={(e) => setTableSearchQuery(e.target.value)}
                  placeholder="Filter table rows..."
                  className="w-full pl-9 pr-3 py-1.5 bg-slate-50 border border-slate-200 rounded-lg text-xs outline-none focus:border-blue-500 transition"
                />
              </div>
            </div>

            {/* 3. Bank Transactions Table (Sortable) */}
            <div className="bg-white rounded-xl shadow-xs border border-slate-200 overflow-hidden space-y-2 p-4">
              <h4 className="text-xs font-bold text-slate-800 uppercase tracking-wider flex items-center justify-between">
                <span>Bank Statement Transactions ({filteredBankTable.length})</span>
                <span className="text-slate-400 text-[11px] font-normal">Sorted by ascending date by default</span>
              </h4>
              <div className="overflow-x-auto max-h-80 border border-slate-100 rounded-lg">
                <table className="w-full text-left text-xs border-collapse">
                  <thead className="bg-slate-50 sticky top-0 border-b border-slate-200 text-slate-600 font-bold">
                    <tr>
                      <th onClick={() => handleBankSort('timestamp')} className="p-2.5 cursor-pointer hover:bg-slate-100 transition">
                        <div className="flex items-center gap-1">Date <ArrowUpDown className="w-3 h-3 text-slate-400" /></div>
                      </th>
                      <th onClick={() => handleBankSort('desc')} className="p-2.5 cursor-pointer hover:bg-slate-100 transition">
                        <div className="flex items-center gap-1">Description <ArrowUpDown className="w-3 h-3 text-slate-400" /></div>
                      </th>
                      <th onClick={() => handleBankSort('amount')} className="p-2.5 text-right cursor-pointer hover:bg-slate-100 transition">
                        <div className="flex items-center justify-end gap-1">Amount (₪) <ArrowUpDown className="w-3 h-3 text-slate-400" /></div>
                      </th>
                      <th onClick={() => handleBankSort('balance')} className="p-2.5 text-right cursor-pointer hover:bg-slate-100 transition">
                        <div className="flex items-center justify-end gap-1">Running Balance (₪) <ArrowUpDown className="w-3 h-3 text-slate-400" /></div>
                      </th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 font-mono">
                    {filteredBankTable.map((t) => (
                      <tr key={t.id} className="hover:bg-slate-50/80 transition">
                        <td className="p-2.5 text-slate-500 whitespace-nowrap">{t.date}</td>
                        <td className="p-2.5 font-sans font-medium text-slate-800">{t.desc}</td>
                        <td className={`p-2.5 text-right font-bold ${t.amount < 0 ? 'text-rose-600' : 'text-blue-600'}`}>
                          {t.amount.toFixed(2)}
                        </td>
                        <td className="p-2.5 text-right font-bold text-slate-700">{t.balance.toFixed(2)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            {/* 4. Credit Card Transactions Table (Sortable) */}
            <div className="bg-white rounded-xl shadow-xs border border-slate-200 overflow-hidden space-y-2 p-4">
              <h4 className="text-xs font-bold text-slate-800 uppercase tracking-wider flex items-center justify-between">
                <span>Credit Card Statement Transactions ({filteredCcTable.length})</span>
                <span className="text-slate-400 text-[11px] font-normal">Includes Card Last 4 Digits</span>
              </h4>
              <div className="overflow-x-auto max-h-80 border border-slate-100 rounded-lg">
                <table className="w-full text-left text-xs border-collapse">
                  <thead className="bg-slate-50 sticky top-0 border-b border-slate-200 text-slate-600 font-bold">
                    <tr>
                      <th onClick={() => handleCcSort('timestamp')} className="p-2.5 cursor-pointer hover:bg-slate-100 transition">
                        <div className="flex items-center gap-1">Date <ArrowUpDown className="w-3 h-3 text-slate-400" /></div>
                      </th>
                      <th onClick={() => handleCcSort('ccLast4')} className="p-2.5 cursor-pointer hover:bg-slate-100 transition">
                        <div className="flex items-center gap-1">Card (Last 4) <ArrowUpDown className="w-3 h-3 text-slate-400" /></div>
                      </th>
                      <th onClick={() => handleCcSort('desc')} className="p-2.5 cursor-pointer hover:bg-slate-100 transition">
                        <div className="flex items-center gap-1">Merchant / Business <ArrowUpDown className="w-3 h-3 text-slate-400" /></div>
                      </th>
                      <th onClick={() => handleCcSort('amount')} className="p-2.5 text-right cursor-pointer hover:bg-slate-100 transition">
                        <div className="flex items-center justify-end gap-1">Amount (₪) <ArrowUpDown className="w-3 h-3 text-slate-400" /></div>
                      </th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 font-mono">
                    {filteredCcTable.map((t) => (
                      <tr key={t.id} className="hover:bg-slate-50/80 transition">
                        <td className="p-2.5 text-slate-500 whitespace-nowrap">{t.date}</td>
                        <td className="p-2.5 font-bold text-indigo-700 whitespace-nowrap">{t.ccLast4}</td>
                        <td className="p-2.5 font-sans font-medium text-slate-800">{t.desc}</td>
                        <td className={`p-2.5 text-right font-bold ${t.amount < 0 ? 'text-rose-600' : 'text-blue-600'}`}>
                          {t.amount.toFixed(2)}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

          </div>
        )}

        {/* TAB 3: UNCATEGORIZED TAB */}
        {activeTab === 'uncategorized' && (
          <div className="bg-white rounded-xl shadow-xs border border-slate-200/80 p-5 space-y-4">
            <div className="flex justify-between items-center">
              <h3 className="text-sm font-bold text-slate-800">
                Uncategorized Transactions ({uncategorizedList.length})
              </h3>
              <p className="text-xs text-slate-400">Assign a one-time override or create a static rule</p>
            </div>
            {uncategorizedList.length === 0 ? (
              <div className="text-center py-10">
                <Check className="w-10 h-10 text-emerald-500 mx-auto mb-2 bg-emerald-50 p-2 rounded-full" />
                <p className="text-slate-800 text-xs font-semibold">All transactions are successfully categorized!</p>
              </div>
            ) : (
              <div className="divide-y divide-slate-100">
                {uncategorizedList.map((item) => {
                  const { prior, remaining } = getOrderedCategoriesForTxn();
                  return (
                    <div key={item.id} className="py-3.5 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 hover:bg-slate-50/60 px-2 rounded-lg transition">
                      <div className="space-y-0.5">
                        <p className="text-xs font-semibold text-slate-900">
                          {item.desc} {item.ccLast4 ? `(CC: ${item.ccLast4})` : ""}
                        </p>
                        <p className="text-[10px] text-slate-400">{item.source} • {item.date}</p>
                      </div>
                      <div className="flex items-center gap-3 w-full sm:w-auto justify-between">
                        <span className={`text-xs font-mono font-bold ${item.amount < 0 ? 'text-rose-600' : 'text-slate-900'}`}>
                          ₪{item.amount.toFixed(2)}
                        </span>
                        <div className="flex items-center gap-2">
                          <select
                            onChange={(e) => handleManualOverride(item.id, e.target.value)}
                            defaultValue=""
                            className="bg-slate-100 border border-slate-200 rounded text-xs py-1 px-2 font-medium text-slate-700 outline-none focus:border-blue-500"
                          >
                            <option value="" disabled>One-time Override...</option>
                            {prior.length > 0 && (
                              <optgroup label="⭐ Previously Used">
                                {prior.map(c => <option key={c} value={c}>{c}</option>)}
                              </optgroup>
                            )}
                            <optgroup label="All Categories">
                              {remaining.map(c => <option key={c} value={c}>{c}</option>)}
                            </optgroup>
                          </select>

                          <button
                            onClick={() => goToRulesForKeyword(item.desc)}
                            className="text-xs bg-blue-50 text-blue-700 hover:bg-blue-100 px-3 py-1 rounded-md font-semibold transition flex items-center gap-1 whitespace-nowrap"
                          >
                            <Tag className="w-3 h-3" /> Create Rule
                          </button>
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        )}

        {/* TAB 4: RULES MANAGER TAB */}
        {activeTab === 'rules' && (
          <div ref={rulesSectionRef} className="bg-white rounded-xl shadow-xs border border-slate-200/80 p-5 space-y-5">
            <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-3">
              <div>
                <h3 className="text-sm font-bold text-slate-800">
                  Global Categorization Rules
                </h3>
                <p className="text-xs text-slate-400 mt-0.5">Rules are static and apply automatically across all statement months.</p>
              </div>
              <div className="flex items-center gap-2">
                <button
                  onClick={exportRulesFile}
                  className="flex items-center gap-1.5 px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg text-xs font-semibold transition"
                >
                  <Download className="w-3.5 h-3.5" /> Export Rules File
                </button>
                <label className="flex items-center gap-1.5 px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg text-xs font-semibold cursor-pointer transition">
                  <FolderUp className="w-3.5 h-3.5" /> Import Rules File
                  <input type="file" accept=".json" ref={fileInputRulesRef} onChange={importRulesFile} className="hidden" />
                </label>
              </div>
            </div>

            <form onSubmit={handleAddRule} className="grid grid-cols-1 md:grid-cols-3 gap-3 bg-slate-50 p-3.5 rounded-xl border border-slate-200/80 items-end">
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Merchant Keyword / Name</label>
                <input
                  type="text"
                  value={newKeyword}
                  onChange={(e) => setNewKeyword(e.target.value)}
                  placeholder="e.g. שופרסל"
                  className="w-full px-3 py-2 bg-white border border-slate-300 rounded-lg text-xs focus:ring-2 focus:ring-blue-500 outline-none transition"
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Target cat3 Category</label>
                <select
                  value={selectedCategory}
                  onChange={(e) => setSelectedCategory(e.target.value)}
                  className="w-full px-3 py-2 bg-white border border-slate-300 rounded-lg text-xs focus:ring-2 focus:ring-blue-500 outline-none transition"
                >
                  {DEFAULT_CATEGORIES.map(cat => (
                    <option key={cat} value={cat}>{cat}</option>
                  ))}
                </select>
              </div>

              <button
                type="submit"
                className="w-full py-2 bg-blue-600 hover:bg-blue-500 text-white text-xs font-semibold rounded-lg shadow-xs transition flex items-center justify-center gap-1.5"
              >
                <Plus className="w-3.5 h-3.5" /> Save Rule & Return
              </button>
            </form>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
              {Object.entries(rules).map(([category, keywords]) => (
                <div key={category} className="p-3 border border-slate-200 rounded-lg bg-white shadow-2xs space-y-1.5">
                  <div className="flex justify-between items-center">
                    <h5 className="text-[11px] font-extrabold text-blue-900 tracking-wide uppercase">{category}</h5>
                    <span className="text-[9px] bg-slate-100 text-slate-600 px-1.5 py-0.5 rounded-full font-mono font-bold">{keywords.length} rules</span>
                  </div>
                  <div className="flex flex-wrap gap-1 pt-0.5">
                    {keywords.map(kw => (
                      <span
                        key={kw}
                        className="bg-slate-100 text-slate-800 border border-slate-200 text-[11px] px-2 py-0.5 rounded-md flex items-center gap-1 font-medium"
                      >
                        {kw}
                        <button
                          onClick={() => handleRemoveRule(category, kw)}
                          className="text-slate-400 hover:text-red-500 transition"
                        >
                          <Trash2 className="w-3 h-3" />
                        </button>
                      </span>
                    ))}
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

      </div>
    </div>
  );
}