import React, { useState, useEffect } from 'react';

function App() {
  const [selectedMonth, setSelectedMonth] = useState('2026-10'); // Default active month format YYYY-MM
  const [transactions, setTransactions] = useState([]);
  const [reconciliationData, setReconciliationData] = useState(null);

  // Load saved data when selectedMonth changes
  useEffect(() => {
    if (!selectedMonth) return;
    const storageKey = `budget_month_${selectedMonth}`;
    const savedData = localStorage.getItem(storageKey);
    
    if (savedData) {
      try {
        const parsed = JSON.parse(savedData);
        setTransactions(parsed.transactions || []);
        setReconciliationData(parsed.reconciliationData || null);
      } catch (e) {
        console.error("Failed to parse saved month data", e);
      }
    } else {
      // Clear state if no data exists for this month
      setTransactions([]);
      setReconciliationData(null);
    }
  }, [selectedMonth]);

  // Handler to delete data for the currently selected month
  const handleDeleteMonthData = () => {
    if (!selectedMonth) return;
    
    const confirmDelete = window.confirm(
      `Are you sure you want to delete all saved data and overrides for ${selectedMonth}? Make sure you have already copied your expenses to Excel!`
    );
    
    if (confirmDelete) {
      const storageKey = `budget_month_${selectedMonth}`;
      localStorage.removeItem(storageKey);
      
      // Reset active states
      setTransactions([]);
      setReconciliationData(null);
      
      alert(`Data for ${selectedMonth} has been successfully cleared.`);
    }
  };

  // Example function demonstrating the Excel SUM formula generation with the minus sign
  const handleCopyExcelFormulas = (itemsRange) => {
    // Prepended minus sign ensuring outgoing totals are negative
    const excelFormula = `=-SUM(${itemsRange})`;
    
    navigator.clipboard.writeText(excelFormula).then(() => {
      alert("Copied Excel formula to clipboard: " + excelFormula);
    });
  };

  return (
    <div style={{ padding: '20px', fontFamily: 'Arial, sans-serif', maxWidth: '900px', margin: '0 auto' }}>
      <h1>Mansheim Budget Processor</h1>

      {/* Month Selector & Controls Panel */}
      <div style={{ 
        display: 'flex', 
        alignItems: 'center', 
        gap: '15px', 
        background: '#f8f9fa', 
        padding: '15px', 
        borderRadius: '8px',
        marginBottom: '20px',
        border: '1px solid #dee2e6'
      }}>
        <label style={{ fontWeight: 'bold' }}>
          Select Month:
          <input 
            type="month" 
            value={selectedMonth} 
            onChange={(e) => setSelectedMonth(e.target.value)}
            style={{ marginLeft: '10px', padding: '5px' }}
          />
        </label>

        {/* Delete Month Button */}
        <button
          onClick={handleDeleteMonthData}
          style={{
            backgroundColor: '#dc3545',
            color: '#ffffff',
            padding: '8px 14px',
            borderRadius: '4px',
            border: 'none',
            cursor: 'pointer',
            fontWeight: '600'
          }}
        >
          Delete Data for {selectedMonth}
        </button>

        <button
          onClick={() => handleCopyExcelFormulas("C2:C50")}
          style={{
            backgroundColor: '#28a745',
            color: '#ffffff',
            padding: '8px 14px',
            borderRadius: '4px',
            border: 'none',
            cursor: 'pointer',
            fontWeight: '600',
            marginLeft: 'auto'
          }}
        >
          Copy Excel SUM Formula
        </button>
      </div>

      {/* Main Workspace Dashboard */}
      <div style={{ background: '#ffffff', padding: '20px', borderRadius: '8px', border: '1px solid #dee2e6' }}>
        <h2>Active Statement Overview ({selectedMonth})</h2>
        {transactions.length === 0 ? (
          <p style={{ color: '#6c757d' }}>No statements or transactions loaded for this month yet. Upload your files to begin reconciliation.</p>
        ) : (
          <ul>
            {transactions.map((tx, index) => (
              <li key={index}>{tx.description} - {tx.amount}</li>
            ))}
          </ul>
        )}
      </div>
    </div>
  );
}

export default App;
