// ========================================
// 🏦 BUDGET TRACKING APP - MAIN SCRIPT
// ========================================
// A comprehensive budget management system built with Google Apps Script
// Features: Expense tracking, income management, category analysis, beautiful UI

// ========================================
// 📊 SETUP & INITIALIZATION
// ========================================

function setupBudgetApp() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create all necessary sheets
  createSheetIfNotExists(ss, 'Dashboard');
  createSheetIfNotExists(ss, 'Transactions');
  createSheetIfNotExists(ss, 'Categories');
  createSheetIfNotExists(ss, 'Budgets');
  createSheetIfNotExists(ss, 'Reports');
  createSheetIfNotExists(ss, 'Settings');
  
  // Initialize all sheets with proper structure
  setupDashboard();
  setupTransactions();
  setupCategories();
  setupBudgets();
  setupReports();
  setupSettings();
  
  // Create beautiful dashboard
  createBeautifulDashboard();
  
  SpreadsheetApp.getUi().alert('🎉 Budget App initialized successfully!\n\nYour budget tracking system is ready to use.');
}

function createSheetIfNotExists(ss, sheetName) {
  if (!ss.getSheetByName(sheetName)) {
    ss.insertSheet(sheetName);
  }
}

// ========================================
// 📋 TRANSACTIONS MANAGEMENT
// ========================================

function setupTransactions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions');
  
  // Clear existing data
  sheet.clear();
  
  // Set up headers with beautiful formatting
  const headers = [
    'Date', 'Type', 'Category', 'Description', 'Amount', 
    'Payment Method', 'Tags', 'Notes', 'Created Date', 'ID'
  ];
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  
  // Apply beautiful formatting
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('white');
  headerRange.setHorizontalAlignment('center');
  
  // Set up data validation for categories
  const categoryRange = sheet.getRange(2, 3, 1000, 1);
  categoryRange.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Categories').getRange('A2:A50'), true)
    .build());
  
  // Set up data validation for payment methods
  const paymentMethods = ['Cash', 'Credit Card', 'Debit Card', 'Bank Transfer', 'Digital Wallet', 'Check'];
  const paymentRange = sheet.getRange(2, 6, 1000, 1);
  paymentRange.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(paymentMethods, true)
    .build());
  
  // Auto-format columns
  sheet.getRange('A:A').setNumberFormat('mm/dd/yyyy');
  sheet.getRange('E:E').setNumberFormat('$#,##0.00');
  sheet.getRange('I:I').setNumberFormat('mm/dd/yyyy hh:mm:ss');
  
  // Freeze header row
  sheet.setFrozenRows(1);
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, headers.length);
}

function addTransaction(type, category, description, amount, paymentMethod, tags, notes) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions');
  const today = new Date();
  const id = Utilities.getUuid();
  
  const newRow = [
    today,                    // Date
    type,                     // Type (Income/Expense)
    category,                 // Category
    description,              // Description
    amount,                   // Amount
    paymentMethod,            // Payment Method
    tags,                     // Tags
    notes,                    // Notes
    today,                    // Created Date
    id                        // ID
  ];
  
  sheet.appendRow(newRow);
  
  // Apply conditional formatting for income vs expense
  const lastRow = sheet.getLastRow();
  const amountCell = sheet.getRange(lastRow, 5);
  
  if (type === 'Income') {
    amountCell.setBackground('#d4edda');
    amountCell.setFontColor('#155724');
  } else {
    amountCell.setBackground('#f8d7da');
    amountCell.setFontColor('#721c24');
  }
  
  // Update dashboard
  updateDashboard();
  
  return id;
}

// ========================================
// 🏷️ CATEGORIES MANAGEMENT
// ========================================

function setupCategories() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Categories');
  
  // Clear existing data
  sheet.clear();
  
  // Set up headers
  const headers = ['Category Name', 'Type', 'Color', 'Icon', 'Budget Limit', 'Description'];
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#34a853');
  headerRange.setFontColor('white');
  
  // Add default categories
  const defaultCategories = [
    ['Food & Dining', 'Expense', '#FF6B6B', '🍽️', 500, 'Restaurants, groceries, coffee'],
    ['Transportation', 'Expense', '#4ECDC4', '🚗', 300, 'Gas, public transport, rideshare'],
    ['Entertainment', 'Expense', '#45B7D1', '🎬', 200, 'Movies, games, hobbies'],
    ['Shopping', 'Expense', '#96CEB4', '🛍️', 400, 'Clothing, electronics, home'],
    ['Healthcare', 'Expense', '#FFEAA7', '🏥', 150, 'Medical, dental, prescriptions'],
    ['Utilities', 'Expense', '#DDA0DD', '💡', 250, 'Electricity, water, internet'],
    ['Salary', 'Income', '#98D8C8', '💰', 0, 'Regular employment income'],
    ['Freelance', 'Income', '#F7DC6F', '💼', 0, 'Contract work and side gigs'],
    ['Investment', 'Income', '#BB8FCE', '📈', 0, 'Dividends, interest, capital gains'],
    ['Gifts', 'Income', '#85C1E9', '🎁', 0, 'Birthday, holiday, special occasion']
  ];
  
  const dataRange = sheet.getRange(2, 1, defaultCategories.length, defaultCategories[0].length);
  dataRange.setValues(defaultCategories);
  
  // Apply color formatting
  for (let i = 0; i < defaultCategories.length; i++) {
    const colorCell = sheet.getRange(i + 2, 3);
    colorCell.setBackground(defaultCategories[i][2]);
  }
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, headers.length);
}

// ========================================
// 💰 BUDGET MANAGEMENT
// ========================================

function setupBudgets() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Budgets');
  
  // Clear existing data
  sheet.clear();
  
  // Set up headers
  const headers = ['Category', 'Monthly Budget', 'Spent', 'Remaining', 'Percentage Used', 'Status'];
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#EA4335');
  headerRange.setFontColor('white');
  
  // Set up formulas for budget tracking
  updateBudgetCalculations();
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, headers.length);
}

function updateBudgetCalculations() {
  const budgetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Budgets');
  const categoriesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Categories');
  const transactionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions');
  
  // Get categories with budget limits
  const categories = categoriesSheet.getDataRange().getValues();
  const budgetData = [];
  
  for (let i = 1; i < categories.length; i++) {
    const category = categories[i][0];
    const budgetLimit = categories[i][4] || 0;
    
    if (budgetLimit > 0) {
      // Calculate spent amount for this category this month
      const currentMonth = new Date().getMonth() + 1;
      const currentYear = new Date().getFullYear();
      
      const spent = calculateCategorySpending(category, currentMonth, currentYear);
      const remaining = budgetLimit - spent;
      const percentageUsed = budgetLimit > 0 ? (spent / budgetLimit) * 100 : 0;
      
      let status = '🟢 On Track';
      if (percentageUsed >= 90) status = '🔴 Over Budget';
      else if (percentageUsed >= 75) status = '🟡 Warning';
      
      budgetData.push([category, budgetLimit, spent, remaining, percentageUsed, status]);
    }
  }
  
  // Clear existing data and add new calculations
  const lastRow = budgetSheet.getLastRow();
  if (lastRow > 1) {
    budgetSheet.getRange(2, 1, lastRow - 1, 6).clear();
  }
  
  if (budgetData.length > 0) {
    const dataRange = budgetSheet.getRange(2, 1, budgetData.length, budgetData[0].length);
    dataRange.setValues(budgetData);
    
    // Apply conditional formatting
    const percentageRange = budgetSheet.getRange(2, 5, budgetData.length, 1);
    percentageRange.setNumberFormat('0.0%');
    
    // Color coding based on percentage
    const rule1 = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0.75)
      .setBackground('#d4edda')
      .setRanges([percentageRange])
      .build();
    
    const rule2 = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(0.75, 0.9)
      .setBackground('#fff3cd')
      .setRanges([percentageRange])
      .build();
    
    const rule3 = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0.9)
      .setBackground('#f8d7da')
      .setRanges([percentageRange])
      .build();
    
    budgetSheet.setConditionalFormatRules([rule1, rule2, rule3]);
  }
}

function calculateCategorySpending(category, month, year) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions');
  const data = sheet.getDataRange().getValues();
  
  let total = 0;
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const transactionDate = new Date(row[0]);
    const transactionCategory = row[2];
    const transactionType = row[1];
    const transactionAmount = parseFloat(row[4]) || 0;
    
    if (transactionCategory === category && 
        transactionDate.getMonth() + 1 === month && 
        transactionDate.getFullYear() === year &&
        transactionType === 'Expense') {
      total += transactionAmount;
    }
  }
  
  return total;
}

// ========================================
// 📊 DASHBOARD CREATION
// ========================================

function setupDashboard() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
  sheet.clear();
  
  // Set up dashboard layout
  sheet.getRange('A1').setValue('📊 BUDGET DASHBOARD');
  sheet.getRange('A1').setFontSize(24);
  sheet.getRange('A1').setFontWeight('bold');
  
  // Create summary cards
  createSummaryCards();
  
  // Create charts area
  createChartsArea();
  
  // Create recent transactions table
  createRecentTransactionsTable();
}

function createBeautifulDashboard() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
  
  // Clear and set up fresh dashboard
  sheet.clear();
  
  // Main title
  const titleCell = sheet.getRange('A1');
  titleCell.setValue('💰 SMART BUDGET TRACKER');
  titleCell.setFontSize(28);
  titleCell.setFontWeight('bold');
  titleCell.setFontColor('#1a73e8');
  sheet.getRange('A1:H1').merge();
  titleCell.setHorizontalAlignment('center');
  
  // Subtitle
  const subtitleCell = sheet.getRange('A2');
  subtitleCell.setValue('Track your finances with style and precision');
  subtitleCell.setFontSize(14);
  subtitleCell.setFontColor('#5f6368');
  sheet.getRange('A2:H2').merge();
  subtitleCell.setHorizontalAlignment('center');
  
  // Create summary cards
  createSummaryCards();
  
  // Create charts and visualizations
  createChartsArea();
  
  // Create recent transactions
  createRecentTransactionsTable();
  
  // Add action buttons area
  createActionButtons();
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, 8);
}

function createSummaryCards() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
  
  // Card 1: Total Balance
  sheet.getRange('A4').setValue('💳 TOTAL BALANCE');
  sheet.getRange('A5').setValue('=SUMIFS(Transactions!E:E,Transactions!B:B,"Income",Transactions!A:A,"<>")-SUMIFS(Transactions!E:E,Transactions!B:B,"Expense",Transactions!A:A,"<>")');
  sheet.getRange('A5').setNumberFormat('$#,##0.00');
  sheet.getRange('A4:A5').setFontWeight('bold');
  sheet.getRange('A4:B5').setBackground('#e8f5e8');
  sheet.getRange('A4:B5').setBorder(true, true, true, true, true, true);
  
  // Card 2: Monthly Income
  sheet.getRange('C4').setValue('💰 MONTHLY INCOME');
  sheet.getRange('C5').setValue('=SUMIFS(Transactions!E:E,Transactions!B:B,"Income",Transactions!A:A,">="&EOMONTH(TODAY(),-1)+1,Transactions!A:A,"<>")');
  sheet.getRange('C5').setNumberFormat('$#,##0.00');
  sheet.getRange('C4:C5').setFontWeight('bold');
  sheet.getRange('C4:D5').setBackground('#e3f2fd');
  sheet.getRange('C4:D5').setBorder(true, true, true, true, true, true);
  
  // Card 3: Monthly Expenses
  sheet.getRange('E4').setValue('💸 MONTHLY EXPENSES');
  sheet.getRange('E5').setValue('=SUMIFS(Transactions!E:E,Transactions!B:B,"Expense",Transactions!A:A,">="&EOMONTH(TODAY(),-1)+1,Transactions!A:A,"<>")');
  sheet.getRange('E5').setNumberFormat('$#,##0.00');
  sheet.getRange('E4:E5').setFontWeight('bold');
  sheet.getRange('E4:F5').setBackground('#fff3e0');
  sheet.getRange('E4:F5').setBorder(true, true, true, true, true, true);
  
  // Card 4: Savings Rate
  sheet.getRange('G4').setValue('📈 SAVINGS RATE');
  sheet.getRange('G5').setValue('=IF(C5>0,(C5-E5)/C5,0)');
  sheet.getRange('G5').setNumberFormat('0.0%');
  sheet.getRange('G4:G5').setFontWeight('bold');
  sheet.getRange('G4:H5').setBackground('#f3e5f5');
  sheet.getRange('G4:H5').setBorder(true, true, true, true, true, true);
}

function createChartsArea() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
  
  // Charts section title
  sheet.getRange('A7').setValue('📊 SPENDING ANALYTICS');
  sheet.getRange('A7').setFontSize(18);
  sheet.getRange('A7').setFontWeight('bold');
  sheet.getRange('A7').setFontColor('#1a73e8');
  
  // Add chart placeholders with instructions
  sheet.getRange('A9').setValue('📈 To create charts:');
  sheet.getRange('A10').setValue('1. Select data ranges below');
  sheet.getRange('A11').setValue('2. Insert > Chart');
  sheet.getRange('A12').setValue('3. Choose chart type (Pie, Bar, Line)');
  sheet.getRange('A13').setValue('4. Customize colors and labels');
  
  // Create data ranges for charts
  sheet.getRange('A15').setValue('Category Spending (This Month):');
  sheet.getRange('A16').setValue('=QUERY(Transactions!A:E,"select C, sum(E) where A >= date \'"&TEXT(EOMONTH(TODAY(),-1)+1,"yyyy-mm-dd")&"\' and B=\'Expense\' and A is not null group by C order by sum(E) desc",1)');
  
  // Add more chart data ranges
  sheet.getRange('A18').setValue('Monthly Income vs Expenses:');
  sheet.getRange('A19').setValue('=QUERY(Transactions!A:E,"select B, sum(E) where A >= date \'"&TEXT(EOMONTH(TODAY(),-1)+1,"yyyy-mm-dd")&"\' and A is not null group by B",1)');
  
  sheet.getRange('A21').setValue('Payment Method Distribution:');
  sheet.getRange('A22').setValue('=QUERY(Transactions!A:F,"select F, sum(E) where A >= date \'"&TEXT(EOMONTH(TODAY(),-1)+1,"yyyy-mm-dd")&"\' and B=\'Expense\' and A is not null group by F",1)');
}

function createRecentTransactionsTable() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
  
  // Recent transactions title
  sheet.getRange('A20').setValue('🕒 RECENT TRANSACTIONS');
  sheet.getRange('A20').setFontSize(18);
  sheet.getRange('A20').setFontWeight('bold');
  sheet.getRange('A20').setFontColor('#1a73e8');
  
  // Create recent transactions query
  sheet.getRange('A22').setValue('=QUERY(Transactions!A:J,"select A, C, D, E, B where A is not null order by A desc limit 10",1)');
  
  // Format the table
  const tableRange = sheet.getRange('A22:F31');
  tableRange.setBorder(true, true, true, true, true, true);
  
  // Header formatting
  const headerRange = sheet.getRange('A22:F22');
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#f8f9fa');
}

function createActionButtons() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
  
  // Action buttons area
  sheet.getRange('A35').setValue('⚡ QUICK ACTIONS');
  sheet.getRange('A35').setFontSize(18);
  sheet.getRange('A35').setFontWeight('bold');
  sheet.getRange('A35').setFontColor('#1a73e8');
  
  // Button instructions
  sheet.getRange('A37').setValue('💡 To add transactions:');
  sheet.getRange('A38').setValue('• Use the "Add Transaction" menu item');
  sheet.getRange('A39').setValue('• Or directly edit the Transactions sheet');
  sheet.getRange('A40').setValue('• Dashboard updates automatically');
}

// ========================================
// 📈 REPORTS & ANALYTICS
// ========================================

function setupReports() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Reports');
  
  // Clear existing data
  sheet.clear();
  
  // Set up headers
  const headers = ['Report Type', 'Date Range', 'Total Income', 'Total Expenses', 'Net Savings', 'Top Category'];
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#FBBC04');
  headerRange.setFontColor('white');
  
  // Generate monthly reports
  generateMonthlyReports();
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, headers.length);
}

function generateMonthlyReports() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Reports');
  const transactionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions');
  
  // Get last 6 months of data
  const today = new Date();
  const reports = [];
  
  for (let i = 0; i < 6; i++) {
    const reportDate = new Date(today.getFullYear(), today.getMonth() - i, 1);
    const monthName = reportDate.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
    
    const income = calculateMonthlyIncome(reportDate.getMonth() + 1, reportDate.getFullYear());
    const expenses = calculateMonthlyExpenses(reportDate.getMonth() + 1, reportDate.getFullYear());
    const netSavings = income - expenses;
    const topCategory = getTopSpendingCategory(reportDate.getMonth() + 1, reportDate.getFullYear());
    
    reports.push([
      'Monthly Report',
      monthName,
      income,
      expenses,
      netSavings,
      topCategory
    ]);
  }
  
  // Add reports to sheet
  if (reports.length > 0) {
    const dataRange = sheet.getRange(2, 1, reports.length, reports[0].length);
    dataRange.setValues(reports);
    
    // Format currency columns
    sheet.getRange(2, 3, reports.length, 3).setNumberFormat('$#,##0.00');
  }
}

function calculateMonthlyIncome(month, year) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions');
  const data = sheet.getDataRange().getValues();
  
  let total = 0;
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const transactionDate = new Date(row[0]);
    const transactionType = row[1];
    const transactionAmount = parseFloat(row[4]) || 0;
    
    if (transactionType === 'Income' && 
        transactionDate.getMonth() + 1 === month && 
        transactionDate.getFullYear() === year) {
      total += transactionAmount;
    }
  }
  
  return total;
}

function calculateMonthlyExpenses(month, year) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions');
  const data = sheet.getDataRange().getValues();
  
  let total = 0;
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const transactionDate = new Date(row[0]);
    const transactionType = row[1];
    const transactionAmount = parseFloat(row[4]) || 0;
    
    if (transactionType === 'Expense' && 
        transactionDate.getMonth() + 1 === month && 
        transactionDate.getFullYear() === year) {
      total += transactionAmount;
    }
  }
  
  return total;
}

function getTopSpendingCategory(month, year) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions');
  const data = sheet.getDataRange().getValues();
  
  const categoryTotals = {};
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const transactionDate = new Date(row[0]);
    const transactionType = row[1];
    const transactionCategory = row[2];
    const transactionAmount = parseFloat(row[4]) || 0;
    
    if (transactionType === 'Expense' && 
        transactionDate.getMonth() + 1 === month && 
        transactionDate.getFullYear() === year) {
      
      if (!categoryTotals[transactionCategory]) {
        categoryTotals[transactionCategory] = 0;
      }
      categoryTotals[transactionCategory] += transactionAmount;
    }
  }
  
  // Find category with highest spending
  let topCategory = 'None';
  let maxAmount = 0;
  
  for (const category in categoryTotals) {
    if (categoryTotals[category] > maxAmount) {
      maxAmount = categoryTotals[category];
      topCategory = category;
    }
  }
  
  return topCategory;
}

// ========================================
// ⚙️ SETTINGS & CONFIGURATION
// ========================================

function setupSettings() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  
  // Clear existing data
  sheet.clear();
  
  // Set up settings structure
  const settings = [
    ['Setting', 'Value', 'Description'],
    ['Currency', 'USD', 'Default currency for all transactions'],
    ['Date Format', 'MM/DD/YYYY', 'Date display format'],
    ['Auto Backup', 'Enabled', 'Automatic backup to Google Drive'],
    ['Email Notifications', 'Disabled', 'Receive budget alerts via email'],
    ['Dark Mode', 'Disabled', 'Dark theme for dashboard'],
    ['Default Payment Method', 'Credit Card', 'Default payment method for new transactions'],
    ['Budget Alert Threshold', '80%', 'Percentage when budget warnings appear'],
    ['Monthly Review Reminder', 'Last day of month', 'When to review monthly spending']
  ];
  
  const dataRange = sheet.getRange(1, 1, settings.length, settings[0].length);
  dataRange.setValues(settings);
  
  // Format headers
  const headerRange = sheet.getRange(1, 1, 1, 3);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#9c27b0');
  headerRange.setFontColor('white');
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, 3);
}

// ========================================
// 🔄 DASHBOARD UPDATE
// ========================================

function updateDashboard() {
  // Update budget calculations
  updateBudgetCalculations();
  
  // Refresh dashboard data
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
  
  // Force recalculation of all formulas by refreshing the sheet
  SpreadsheetApp.flush();
  
  // Update charts if they exist
  updateCharts();
  
  // Show success message
  console.log('Dashboard updated successfully');
}

function updateCharts() {
  // This function would update any charts on the dashboard
  // Implementation depends on specific chart requirements
  console.log('Charts updated');
}

// ========================================
// 🎯 USER INTERFACE FUNCTIONS
// ========================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // Create custom menu
  ui.createMenu('💰 Budget Tracker')
    .addItem('➕ Add Transaction', 'showAddTransactionDialog')
    .addItem('📊 Generate Report', 'generateCustomReport')
    .addItem('🔄 Update Dashboard', 'updateDashboard')
    .addItem('⚙️ Settings', 'showSettingsDialog')
    .addSeparator()
    .addItem('📈 View Analytics', 'showAnalyticsDialog')
    .addItem('💾 Export Data', 'exportBudgetData')
    .addItem('🆘 Help', 'showHelpDialog')
    .addToUi();
}

function showAddTransactionDialog() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.prompt(
    'Add New Transaction',
    'Enter transaction details (format: Type|Category|Description|Amount|PaymentMethod|Tags|Notes):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const input = result.getResponseText();
    const parts = input.split('|');
    
    if (parts.length >= 4) {
      const type = parts[0].trim();
      const category = parts[1].trim();
      const description = parts[2].trim();
      const amount = parseFloat(parts[3].trim());
      const paymentMethod = parts[4] ? parts[4].trim() : 'Cash';
      const tags = parts[5] ? parts[5].trim() : '';
      const notes = parts[6] ? parts[6].trim() : '';
      
      if (type && category && description && !isNaN(amount)) {
        addTransaction(type, category, description, amount, paymentMethod, tags, notes);
        ui.alert('✅ Transaction added successfully!');
      } else {
        ui.alert('❌ Invalid input format. Please check your data.');
      }
    } else {
      ui.alert('❌ Invalid input format. Please use: Type|Category|Description|Amount|PaymentMethod|Tags|Notes');
    }
  }
}

function generateCustomReport() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.prompt(
    'Generate Custom Report',
    'Enter month and year (format: MM/YYYY):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const input = result.getResponseText();
    const [month, year] = input.split('/');
    
    if (month && year) {
      const income = calculateMonthlyIncome(parseInt(month), parseInt(year));
      const expenses = calculateMonthlyExpenses(parseInt(month), parseInt(year));
      const netSavings = income - expenses;
      const savingsRate = income > 0 ? (netSavings / income) * 100 : 0;
      
      const report = `📊 Monthly Report for ${month}/${year}\n\n` +
                    `💰 Total Income: $${income.toFixed(2)}\n` +
                    `💸 Total Expenses: $${expenses.toFixed(2)}\n` +
                    `💳 Net Savings: $${netSavings.toFixed(2)}\n` +
                    `📈 Savings Rate: ${savingsRate.toFixed(1)}%\n\n` +
                    `Top Spending Category: ${getTopSpendingCategory(parseInt(month), parseInt(year))}`;
      
      ui.alert('Monthly Report', report, ui.ButtonSet.OK);
    } else {
      ui.alert('❌ Invalid date format. Please use MM/YYYY');
    }
  }
}

function showSettingsDialog() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    'Settings',
    'What would you like to configure?\n\n1. Currency settings\n2. Notification preferences\n3. Budget alerts\n4. Export settings',
    ui.ButtonSet.YES_NO_CANCEL
  );
  
  if (result === ui.Button.YES) {
    ui.alert('Settings', 'Settings dialog will be implemented in future versions.', ui.ButtonSet.OK);
  }
}

function showAnalyticsDialog() {
  const ui = SpreadsheetApp.getUi();
  
  // Calculate some analytics
  const today = new Date();
  const currentMonth = today.getMonth() + 1;
  const currentYear = today.getFullYear();
  
  const income = calculateMonthlyIncome(currentMonth, currentYear);
  const expenses = calculateMonthlyExpenses(currentMonth, currentYear);
  const netSavings = income - expenses;
  const savingsRate = income > 0 ? (netSavings / income) * 100 : 0;
  
  const analytics = `📈 Current Month Analytics\n\n` +
                   `💰 Income: $${income.toFixed(2)}\n` +
                   `💸 Expenses: $${expenses.toFixed(2)}\n` +
                   `💳 Net: $${netSavings.toFixed(2)}\n` +
                   `📊 Savings Rate: ${savingsRate.toFixed(1)}%\n\n` +
                   `🎯 Top Category: ${getTopSpendingCategory(currentMonth, currentYear)}\n` +
                   `📅 Days in Month: ${new Date(currentYear, currentMonth, 0).getDate()}`;
  
  ui.alert('Analytics', analytics, ui.ButtonSet.OK);
}

function exportBudgetData() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    'Export Data',
    'Export your budget data to CSV format?',
    ui.ButtonSet.YES_NO
  );
  
  if (result === ui.Button.YES) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions');
    const data = sheet.getDataRange().getValues();
    
    let csvContent = '';
    for (let i = 0; i < data.length; i++) {
      csvContent += data[i].join(',') + '\n';
    }
    
    // Create a temporary file
    const fileName = 'budget_data_' + new Date().toISOString().split('T')[0] + '.csv';
    const file = DriveApp.createFile(fileName, csvContent, MimeType.CSV);
    
    ui.alert('Export Complete', `Data exported to: ${file.getName()}\nFile ID: ${file.getId()}`, ui.ButtonSet.OK);
  }
}

function showHelpDialog() {
  const ui = SpreadsheetApp.getUi();
  
  const help = `🆘 Budget Tracker Help\n\n` +
               `📝 How to use:\n` +
               `1. Add transactions using the menu\n` +
               `2. View dashboard for overview\n` +
               `3. Check budgets for category limits\n` +
               `4. Generate reports for analysis\n\n` +
               `💡 Tips:\n` +
               `• Use categories to organize spending\n` +
               `• Set budget limits to stay on track\n` +
               `• Review reports monthly\n` +
               `• Export data for backup\n\n` +
               `🔧 Features:\n` +
               `• Automatic calculations\n` +
               `• Beautiful dashboard\n` +
               `• Budget tracking\n` +
               `• Spending analytics\n` +
               `• Export capabilities`;
  
  ui.alert('Help', help, ui.ButtonSet.OK);
}

// ========================================
// 🎨 UTILITY FUNCTIONS
// ========================================

function formatCurrency(amount) {
  return new Intl.NumberFormat('en-US', {
    style: 'currency',
    currency: 'USD'
  }).format(amount);
}

function getCurrentMonth() {
  return new Date().getMonth() + 1;
}

function getCurrentYear() {
  return new Date().getFullYear();
}

function isValidDate(dateString) {
  const date = new Date(dateString);
  return date instanceof Date && !isNaN(date);
}

// ========================================
// 🚀 INITIALIZATION
// ========================================

// Auto-run setup when script is first loaded
function initializeBudgetApp() {
  setupBudgetApp();
}
