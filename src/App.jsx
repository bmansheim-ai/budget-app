import React, { useState, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import { 
  Upload, FileSpreadsheet, Copy, Check, Plus, Trash2, Tag, ChevronDown, 
  ChevronUp, Search, Download, FolderUp, CheckCircle2, ShieldCheck, 
  Scale, ArrowDownRight, ArrowUpRight, CreditCard, Building2, Table
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

export default function App() {
  const [bankTransactions, setBankTransactions] = useState([]);
  const [ccTransactions, setCcTransactions] = useState([]);
  const [bankFileName, setBankFileName] = useState("");
  const [ccFileName, setCcFileName] = useState("");
  const [bankBalances, setBankBalances] = useState({ beginning: 0, ending: 0 });

  const [rules, setRules] = useState(() => {
    const saved = localStorage.getItem("budget_rules");
    return saved ? JSON.parse(saved) : DEFAULT_RULES;
  });
  const [newKeyword, setNewKeyword] = useState("");
  const [selectedCategory, setSelectedCategory] = useState(DEFAULT_CATEGORIES[0]);
  const [copied, setCopied] = useState(false);
  const [activeTab, setActiveTab] = useState("summary");
  const [expandedCategory, setExpandedCategory] = useState(null);
  const [searchQuery, setSearchQuery] = useState("");
  const [tableSearchQuery, setTableSearchQuery] = useState("");
  
  const rulesSectionRef = useRef(null);
  const fileInputRulesRef = useRef(null);

  const saveRules = (updatedRules) => {
    setRules(updatedRules);
    localStorage.setItem("budget_rules", JSON.stringify(updatedRules));
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

  const handleFileUpload = (e, type) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const wb = XLSX.read(bstr, { type: "binary" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

      if (type === "bank") {
        setBankFileName(file.name);
        
        let headerIdx = data.findIndex(row => row && row.some(cell => String(cell).includes("תאריך")));
        if (headerIdx === -1) headerIdx = 7;
        const rows = data.slice(headerIdx + 1);
        
        const parsed = rows.map(r => {
          if (!r || r.length < 5) return null;
          return { 
            date: r[0], 
            desc: String(r[2] || ""), 
            amount: parseFloat(r[3]) || 0, 
            balance: parseFloat(r[4]) || 0,
            source: "Bank" 
          };
        }).filter(item => item && item.desc && !isNaN(item.amount));

        setBankTransactions(parsed);

        if (parsed.length > 0) {
          const firstTxn = parsed[0];
          const lastTxn = parsed[parsed.length - 1];
          const begBal = firstTxn.balance - firstTxn.amount;
          const endBal = lastTxn.balance;
          setBankBalances({ beginning: begBal, ending: endBal });
        }
      } else if (type === "cc") {
        setCcFileName(file.name);
        let headerIdx = data.findIndex(row => row && row.some(cell => String(cell).includes("בית עסק")));
        if (headerIdx === -1) headerIdx = 7;
        const rows = data.slice(headerIdx + 1);
        const parsed = rows.map(r => {
          if (!r || r.length < 9) return null;
          let amt = parseFloat(r[8]) || 0;
          if (amt > 0) amt = -amt;
          return { date: r[2], desc: String(r[1] || ""), amount: amt, source: "Credit Card" };
        }).filter(item => item && item.desc && !isNaN(item.amount));
        setCcTransactions(parsed);
      }
    };
    reader.readAsBinaryString(file);
  };

  const categorize = (desc) => {
    const lowerDesc = String(desc).toLowerCase();
    for (const [category, keywords] of Object.entries(rules)) {
      if (keywords.some(k => lowerDesc.includes(k.toLowerCase()))) return category;
    }
    return "Uncategorized";
  };

  const nonCcBankTransactions = useMemo(() => {
    return bankTransactions.filter(t => {
      const desc = t.desc;
      return !desc.includes("חיוב לכרטיס") && !desc.includes("מזרחי כאל") && !desc.includes("חיוב מ כ.א.ל");
    });
  }, [bankTransactions]);

  const allTransactions = useMemo(() => {
    return [...nonCcBankTransactions, ...ccTransactions].map(t => ({ ...t, category: categorize(t.desc) }));
  }, [nonCcBankTransactions, ccTransactions, rules]);

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
        formattedRows.push(`=SUM(${items})`);
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

  const filteredBankTable = useMemo(() => {
    if (!tableSearchQuery.trim()) return bankTransactions;
    return bankTransactions.filter(t => t.desc.toLowerCase().includes(tableSearchQuery.toLowerCase()));
  }, [bankTransactions, tableSearchQuery]);

  const filteredCcTable = useMemo(() => {
    if (!tableSearchQuery.trim()) return ccTransactions;
    return ccTransactions.filter(t => t.desc.toLowerCase().includes(tableSearchQuery.toLowerCase()));
  }, [ccTransactions, tableSearchQuery]);

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
              <h1 className="text-xl font-extrabold tracking-tight">Mansheim Budget Processor</h1>
              <p className="text-slate-300 text-xs">Automated Statement Categorization & Reconciliation</p>
            </div>
          </div>

          <div className="flex items-center gap-4">
            <div className="text-right hidden sm:block">
              <p className="text-[11px] text-slate-400">Total Tracked Expenses</p>
              <p className="text-base font-mono font-bold text-emerald-400">₪{totalExpenseSum.toFixed(2)}</p>
            </div>
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

      <div className="max-w-7xl mx-auto px-4 space-y-4">
        
        {/* Compact File Upload Toolbar */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          <div className="bg-white px-4 py-3 rounded-xl shadow-xs border border-slate-200/80 flex items-center justify-between">
            <div className="flex items-center gap-2.5 truncate">
              {bankFileName ? (
                <CheckCircle2 className="w-5 h-5 text-emerald-500 flex-shrink-0" />
              ) : (
                <span className="w-2.5 h-2.5 rounded-full bg-indigo-500 flex-shrink-0"></span>
              )}
              <div className="truncate">
                <span className="text-xs font-bold text-slate-800">Bank (עובר ושב)</span>
                {bankFileName && (
                  <p className="text-[11px] text-slate-500 truncate">{bankFileName} ({bankTransactions.length} items)</p>
                )}
              </div>
            </div>
            <label className="cursor-pointer bg-slate-100 hover:bg-indigo-50 text-slate-700 hover:text-indigo-700 px-3 py-1.5 rounded-lg text-xs font-semibold flex items-center gap-1.5 transition flex-shrink-0">
              <Upload className="w-3.5 h-3.5" />
              {bankFileName ? "Change File" : "Upload Bank"}
              <input type="file" accept=".xlsx, .xls" onChange={(e) => handleFileUpload(e, 'bank')} className="hidden" />
            </label>
          </div>

          <div className="bg-white px-4 py-3 rounded-xl shadow-xs border border-slate-200/80 flex items-center justify-between">
            <div className="flex items-center gap-2.5 truncate">
              {ccFileName ? (
                <CheckCircle2 className="w-5 h-5 text-emerald-500 flex-shrink-0" />
              ) : (
                <span className="w-2.5 h-2.5 rounded-full bg-emerald-500 flex-shrink-0"></span>
              )}
              <div className="truncate">
                <span className="text-xs font-bold text-slate-800">Credit Card (כרטיסי אשראי)</span>
                {ccFileName && (
                  <p className="text-[11px] text-slate-500 truncate">{ccFileName} ({ccTransactions.length} items)</p>
                )}
              </div>
            </div>
            <label className="cursor-pointer bg-slate-100 hover:bg-emerald-50 text-slate-700 hover:text-emerald-700 px-3 py-1.5 rounded-lg text-xs font-semibold flex items-center gap-1.5 transition flex-shrink-0">
              <Upload className="w-3.5 h-3.5" />
              {ccFileName ? "Change File" : "Upload CC"}
              <input type="file" accept=".xlsx, .xls" onChange={(e) => handleFileUpload(e, 'cc')} className="hidden" />
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
                        {txns.map((t, i) => (
                          <div key={i} className="bg-white p-2.5 rounded-lg border border-slate-200/60 flex justify-between items-center text-[11px] shadow-2xs">
                            <div className="space-y-0.5 pr-2 truncate">
                              <p className="font-semibold text-slate-800 truncate">{t.desc}</p>
                              <div className="flex items-center gap-1.5 text-slate-400 text-[10px]">
                                <span>{t.date}</span>
                                <span>•</span>
                                <span className="bg-slate-100 px-1 py-0.5 rounded text-[9px] font-medium">{t.source}</span>
                              </div>
                            </div>
                            <span className={`font-mono font-bold whitespace-nowrap ${t.amount < 0 ? 'text-rose-600' : 'text-slate-900'}`}>
                              ₪{t.amount.toFixed(2)}
                            </span>
                          </div>
                        ))}
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
                <Table className="w-4 h-4 text-slate-500" /> Statement Transactions Tables
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

            {/* 3. Bank Transactions Table */}
            <div className="bg-white rounded-xl shadow-xs border border-slate-200 overflow-hidden space-y-2 p-4">
              <h4 className="text-xs font-bold text-slate-800 uppercase tracking-wider flex items-center justify-between">
                <span>Bank Statement Transactions ({filteredBankTable.length})</span>
                <span className="text-slate-400 text-[11px] font-normal">Col 5 Running Balances Included</span>
              </h4>
              <div className="overflow-x-auto max-h-80 border border-slate-100 rounded-lg">
                <table className="w-full text-left text-xs border-collapse">
                  <thead className="bg-slate-50 sticky top-0 border-b border-slate-200 text-slate-600 font-bold">
                    <tr>
                      <th className="p-2.5">Date</th>
                      <th className="p-2.5">Description</th>
                      <th className="p-2.5 text-right">Amount (₪)</th>
                      <th className="p-2.5 text-right">Running Balance (₪)</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 font-mono">
                    {filteredBankTable.map((t, idx) => (
                      <tr key={idx} className="hover:bg-slate-50/80 transition">
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

            {/* 4. Credit Card Transactions Table */}
            <div className="bg-white rounded-xl shadow-xs border border-slate-200 overflow-hidden space-y-2 p-4">
              <h4 className="text-xs font-bold text-slate-800 uppercase tracking-wider">
                Credit Card Statement Transactions ({filteredCcTable.length})
              </h4>
              <div className="overflow-x-auto max-h-80 border border-slate-100 rounded-lg">
                <table className="w-full text-left text-xs border-collapse">
                  <thead className="bg-slate-50 sticky top-0 border-b border-slate-200 text-slate-600 font-bold">
                    <tr>
                      <th className="p-2.5">Date</th>
                      <th className="p-2.5">Merchant / Business</th>
                      <th className="p-2.5 text-right">Amount (₪)</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 font-mono">
                    {filteredCcTable.map((t, idx) => (
                      <tr key={idx} className="hover:bg-slate-50/80 transition">
                        <td className="p-2.5 text-slate-500 whitespace-nowrap">{t.date}</td>
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
              <p className="text-xs text-slate-400">Clicking "Create Rule" pre-fills merchant name in Rules Manager</p>
            </div>
            {uncategorizedList.length === 0 ? (
              <div className="text-center py-10">
                <Check className="w-10 h-10 text-emerald-500 mx-auto mb-2 bg-emerald-50 p-2 rounded-full" />
                <p className="text-slate-800 text-xs font-semibold">All transactions are successfully categorized!</p>
              </div>
            ) : (
              <div className="divide-y divide-slate-100">
                {uncategorizedList.map((item, idx) => (
                  <div key={idx} className="py-3 flex items-center justify-between gap-4 hover:bg-slate-50/60 px-2 rounded-lg transition">
                    <div>
                      <p className="text-xs font-semibold text-slate-900">{item.desc}</p>
                      <p className="text-[10px] text-slate-400 mt-0.5">{item.source} • {item.date}</p>
                    </div>
                    <div className="flex items-center gap-3">
                      <span className={`text-xs font-mono font-bold ${item.amount < 0 ? 'text-rose-600' : 'text-slate-900'}`}>
                        ₪{item.amount.toFixed(2)}
                      </span>
                      <button
                        onClick={() => goToRulesForKeyword(item.desc)}
                        className="text-xs bg-blue-50 text-blue-700 hover:bg-blue-100 px-2.5 py-1 rounded-md font-semibold transition flex items-center gap-1"
                      >
                        <Tag className="w-3 h-3" /> Create Rule
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        )}

        {/* TAB 4: RULES MANAGER TAB */}
        {activeTab === 'rules' && (
          <div ref={rulesSectionRef} className="bg-white rounded-xl shadow-xs border border-slate-200/80 p-5 space-y-5">
            <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-3">
              <div>
                <h3 className="text-sm font-bold text-slate-800">Categorization Keyword Rules</h3>
                <p className="text-xs text-slate-400 mt-0.5">Saving a rule automatically takes you back to Uncategorized items.</p>
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
                          <Trash2 className="w-3.5 h-3.5" />
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