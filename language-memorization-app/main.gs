// Add to Apps Script
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create sheets if they don't exist
  createSheetIfNotExists(ss, 'Words Database');
  createSheetIfNotExists(ss, 'Review Log');
  createSheetIfNotExists(ss, 'Analytics');
  
  // Set up headers
  setupHeaders();
}

function createSheetIfNotExists(ss, sheetName) {
  if (!ss.getSheetByName(sheetName)) {
    ss.insertSheet(sheetName);
  }
}

function findWordRow(word) {
  const wordsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Words Database');
  const data = wordsSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === word) { // Column A contains words
      return i + 1; // Return 1-based row number
    }
  }
  return -1; // Word not found
}

function setupHeaders() {
  const wordsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Words Database');
  const headers = ['Word', 'Definition', 'Context', 'Difficulty', 'Next Review', 'Interval', 'Success Rate', 'Total Reviews', 'Last Review', 'Status'];
  
  wordsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  wordsSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
}

function calculateNextInterval(word, result, currentInterval, difficulty) {
  // SuperMemo 2 Algorithm adaptation
  const easeFactor = getEaseFactor(word);
  const quality = result === 'Correct' ? 5 : 1; // 0-5 scale
  
  let newEaseFactor = easeFactor + (0.1 - (5 - quality) * (0.08 + (5 - quality) * 0.02));
  newEaseFactor = Math.max(1.3, newEaseFactor);
  
  let newInterval;
  if (result === 'Correct') {
    if (currentInterval === 0) {
      newInterval = 1;
    } else if (currentInterval === 1) {
      newInterval = 6;
    } else {
      newInterval = Math.round(currentInterval * newEaseFactor);
    }
  } else {
    newInterval = 1; // Reset to 1 day
  }
  
  // Adjust for difficulty
  const difficultyMultiplier = 1 + (difficulty - 3) * 0.2;
  newInterval = Math.round(newInterval * difficultyMultiplier);
  
  return { interval: newInterval, easeFactor: newEaseFactor };
}

function getEaseFactor(word) {
  // Get ease factor from hidden column or default to 2.5
  const wordsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Words Database');
  const wordRow = findWordRow(word);
  if (wordRow > 0) {
    return wordsSheet.getRange(wordRow, 11).getValue() || 2.5; // Column K for ease factor
  }
  return 2.5;
}

function startReview() {
  const wordsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Words Database');
  const today = new Date();
  
  // Get words due for review
  const dueWords = getDueWords(today);
  
  if (dueWords.length === 0) {
    SpreadsheetApp.getUi().alert('No words due for review today!');
    return;
  }
  
  // Create review interface
  createReviewInterface(dueWords);
}

function getDueWords(today) {
  const wordsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Words Database');
  const data = wordsSheet.getDataRange().getValues();
  const dueWords = [];
  
  for (let i = 1; i < data.length; i++) {
    const nextReview = data[i][4]; // Column E
    if (nextReview && new Date(nextReview) <= today) {
      dueWords.push({
        row: i + 1,
        word: data[i][0],
        definition: data[i][1],
        context: data[i][2]
      });
    }
  }
  
  return dueWords;
}

function createReviewInterface(words) {
  const ui = SpreadsheetApp.getUi();
  
  words.forEach((wordData, index) => {
    const response = ui.alert(
      `Review ${index + 1}/${words.length}`,
      `Word: ${wordData.word}\n\nClick "Yes" if you knew it, "No" if you didn't.`,
      ui.ButtonSet.YES_NO
    );
    
    const result = response === ui.Button.YES ? 'Correct' : 'Incorrect';
    processReviewResult(wordData, result);
  });
}

function processReviewResult(wordData, result) {
  const wordsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Words Database');
  const reviewLogSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Review Log');
  
  // Update word data
  const currentInterval = wordsSheet.getRange(wordData.row, 6).getValue() || 0;
  const difficulty = wordsSheet.getRange(wordData.row, 4).getValue() || 3;
  
  const newData = calculateNextInterval(wordData.word, result, currentInterval, difficulty);
  const nextReview = new Date();
  nextReview.setDate(nextReview.getDate() + newData.interval);
  
  // Update word row
  wordsSheet.getRange(wordData.row, 5).setValue(nextReview); // Next Review
  wordsSheet.getRange(wordData.row, 6).setValue(newData.interval); // Interval
  wordsSheet.getRange(wordData.row, 7).setValue(calculateSuccessRate(wordData.word)); // Success Rate
  wordsSheet.getRange(wordData.row, 8).setValue(wordsSheet.getRange(wordData.row, 8).getValue() + 1); // Total Reviews
  wordsSheet.getRange(wordData.row, 9).setValue(new Date()); // Last Review
  wordsSheet.getRange(wordData.row, 11).setValue(newData.easeFactor); // Ease Factor (hidden)
  
  // Log review
  const logRow = reviewLogSheet.getLastRow() + 1;
  reviewLogSheet.getRange(logRow, 1).setValue(new Date());
  reviewLogSheet.getRange(logRow, 2).setValue(wordData.word);
  reviewLogSheet.getRange(logRow, 3).setValue(result);
}

function calculateSuccessRate(word) {
  const reviewLogSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Review Log');
  const data = reviewLogSheet.getDataRange().getValues();
  
  let totalReviews = 0;
  let correctReviews = 0;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === word) {
      totalReviews++;
      if (data[i][2] === 'Correct') {
        correctReviews++;
      }
    }
  }
  
  return totalReviews > 0 ? (correctReviews / totalReviews * 100).toFixed(1) : 0;
}

function adjustDifficulty(word) {
  const wordsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Words Database');
  const wordRow = findWordRow(word);
  
  if (wordRow > 0) {
    const successRate = wordsSheet.getRange(wordRow, 7).getValue();
    const currentDifficulty = wordsSheet.getRange(wordRow, 4).getValue();
    
    let newDifficulty = currentDifficulty;
    if (successRate > 80) {
      newDifficulty = Math.max(1, currentDifficulty - 0.5);
    } else if (successRate < 50) {
      newDifficulty = Math.min(5, currentDifficulty + 0.5);
    }
    
    wordsSheet.getRange(wordRow, 4).setValue(newDifficulty);
  }
}

// Batch Import
function importWords(wordsArray) {
  const wordsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Words Database');
  
  wordsArray.forEach(wordData => {
    const newRow = wordsSheet.getLastRow() + 1;
    wordsSheet.getRange(newRow, 1).setValue(wordData.word);
    wordsSheet.getRange(newRow, 2).setValue(wordData.definition);
    wordsSheet.getRange(newRow, 3).setValue(wordData.context);
    wordsSheet.getRange(newRow, 4).setValue(wordData.difficulty || 3);
    wordsSheet.getRange(newRow, 5).setValue(new Date()); // Start reviewing today
    wordsSheet.getRange(newRow, 6).setValue(0); // Initial interval
    wordsSheet.getRange(newRow, 10).setValue('Active'); // Status
    wordsSheet.getRange(newRow, 11).setValue(2.5); // Initial ease factor
  });
}

// Analytics Dashboard
function updateAnalytics() {
  const analyticsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Analytics');
  const wordsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Words Database');
  const reviewLogSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Review Log');
  
  // Clear existing data
  analyticsSheet.clear();
  
  // Calculate statistics
  const totalWords = wordsSheet.getLastRow() - 1;
  const activeWords = getActiveWordsCount();
  const dueToday = getDueWordsCount(new Date());
  const avgSuccessRate = calculateAverageSuccessRate();
  
  // Create dashboard
  analyticsSheet.getRange('A1').setValue('Learning Dashboard');
  analyticsSheet.getRange('A3').setValue('Total Words:');
  analyticsSheet.getRange('B3').setValue(totalWords);
  analyticsSheet.getRange('A4').setValue('Active Words:');
  analyticsSheet.getRange('B4').setValue(activeWords);
  analyticsSheet.getRange('A5').setValue('Due Today:');
  analyticsSheet.getRange('B5').setValue(dueToday);
  analyticsSheet.getRange('A6').setValue('Average Success Rate:');
  analyticsSheet.getRange('B6').setValue(avgSuccessRate + '%');
}

function getActiveWordsCount() {
  const wordsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Words Database');
  const data = wordsSheet.getDataRange().getValues();
  let count = 0;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][9] === 'Active') { // Column J
      count++;
    }
  }
  
  return count;
}

function getDueWordsCount(date) {
  const wordsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Words Database');
  const data = wordsSheet.getDataRange().getValues();
  let count = 0;
  
  for (let i = 1; i < data.length; i++) {
    const nextReview = data[i][4];
    if (nextReview && new Date(nextReview) <= date) {
      count++;
    }
  }
  
  return count;
}

function calculateAverageSuccessRate() {
  const wordsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Words Database');
  const data = wordsSheet.getDataRange().getValues();
  let totalRate = 0;
  let wordCount = 0;
  
  for (let i = 1; i < data.length; i++) {
    const successRate = data[i][6];
    if (successRate > 0) {
      totalRate += successRate;
      wordCount++;
    }
  }
  
  return wordCount > 0 ? (totalRate / wordCount).toFixed(1) : 0;
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Memorization App')
    .addItem('Start Review', 'startReview')
    .addItem('Add New Word', 'addNewWord')
    .addItem('Update Analytics', 'updateAnalytics')
    .addItem('Export Data', 'exportData')
    .addItem('Import Data', 'importData')
    .addToUi();
}

function addNewWord() {
  const htmlTemplate = HtmlService.createTemplateFromFile('add-word-dialog');
  const html = htmlTemplate.evaluate()
    .setWidth(600)
    .setHeight(700)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  
  SpreadsheetApp.getUi().showModalDialog(html, '⚡ ADD NEW WORD - MEMORY FORGE ⚡');
}

function processNewWord(formData) {
  try {
    const wordsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Words Database');
    const newRow = wordsSheet.getLastRow() + 1;
    
    // Validate input
    if (!formData.word || !formData.definition) {
      return { success: false, message: 'Word and Definition are required!' };
    }
    
    // Calculate difficulty based on word length and complexity
    const difficulty = calculateWordDifficulty(formData.word, formData.definition);
    
    wordsSheet.getRange(newRow, 1).setValue(formData.word);
    wordsSheet.getRange(newRow, 2).setValue(formData.definition);
    wordsSheet.getRange(newRow, 3).setValue(formData.context || '');
    wordsSheet.getRange(newRow, 4).setValue(difficulty);
    wordsSheet.getRange(newRow, 5).setValue(new Date());
    wordsSheet.getRange(newRow, 6).setValue(0);
    wordsSheet.getRange(newRow, 7).setValue(0);
    wordsSheet.getRange(newRow, 8).setValue(0);
    wordsSheet.getRange(newRow, 9).setValue(new Date());
    wordsSheet.getRange(newRow, 10).setValue('Active');
    wordsSheet.getRange(newRow, 11).setValue(2.5);
    
    return { 
      success: true, 
      message: `🔥 WORD FORGED: "${formData.word}" has been added to your memory arsenal! 🔥`,
      difficulty: difficulty
    };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

function calculateWordDifficulty(word, definition) {
  // Calculate difficulty based on word length, definition complexity, and special characters
  let difficulty = 3; // Base difficulty
  
  // Word length factor
  if (word.length > 12) difficulty += 1;
  if (word.length > 20) difficulty += 1;
  
  // Definition complexity
  if (definition.length > 100) difficulty += 0.5;
  if (definition.includes(';') || definition.includes(':')) difficulty += 0.5;
  
  // Special characters and complexity
  const specialChars = (word.match(/[^a-zA-Z0-9\s]/g) || []).length;
  difficulty += specialChars * 0.3;
  
  // Capitalization patterns
  if (word !== word.toLowerCase() && word !== word.toUpperCase()) difficulty += 0.5;
  
  return Math.min(5, Math.max(1, Math.round(difficulty * 10) / 10));
}
function customForgettingCurve(interval, difficulty, successRate) {
  // Implement your own algorithm
  const baseInterval = interval;
  const difficultyMultiplier = 1 + (difficulty - 3) * 0.3;
  const successMultiplier = successRate > 80 ? 1.2 : successRate < 50 ? 0.8 : 1.0;
  
  return Math.round(baseInterval * difficultyMultiplier * successMultiplier);
}
function predictOptimalInterval(word, userHistory) {
  // Use historical data to predict optimal intervals
  const patterns = analyzeUserPatterns(userHistory);
  return calculateOptimalInterval(patterns, word);
}
function exportData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wordsSheet = ss.getSheetByName('Words Database');
  const reviewLogSheet = ss.getSheetByName('Review Log');
  
  const wordsData = wordsSheet.getDataRange().getValues();
  const reviewData = reviewLogSheet.getDataRange().getValues();
  
  const exportData = {
    words: wordsData,
    reviews: reviewData,
    exportDate: new Date().toISOString()
  };
  
  // Create downloadable JSON
  const blob = Utilities.newBlob(JSON.stringify(exportData), 'application/json', 'memorization_data.json');
  const url = DriveApp.createFile(blob).getDownloadUrl();
  
  SpreadsheetApp.getUi().alert('Data exported! Download URL: ' + url);
}

function importData() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Import Data', 'Paste JSON data:', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() === ui.Button.OK) {
    try {
      const data = JSON.parse(response.getResponseText());
      
      // Import words
      const wordsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Words Database');
      wordsSheet.clear();
      wordsSheet.getRange(1, 1, data.words.length, data.words[0].length).setValues(data.words);
      
      // Import reviews
      const reviewLogSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Review Log');
      reviewLogSheet.clear();
      reviewLogSheet.getRange(1, 1, data.reviews.length, data.reviews[0].length).setValues(data.reviews);
      
      ui.alert('Data imported successfully!');
    } catch (error) {
      ui.alert('Error importing data: ' + error.toString());
    }
  }
}