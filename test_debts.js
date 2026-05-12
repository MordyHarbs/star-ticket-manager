function getDebtsMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const debts = {};
  
  // Helper to add debt
  function addDebt(name, amount) {
    if (!name) return;
    const n = normalizeHebrew(name);
    if (!debts[n]) debts[n] = { originalName: name, sum: 0 };
    debts[n].sum += amount;
  }
  
  // ... loop sheets and add ...
}
