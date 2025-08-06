// ========================================
// 🎓 IT FRONTEND BOOTCAMP MANAGEMENT SYSTEM
// ========================================
// A comprehensive management system for IT bootcamp students
// Features: Technical learning tracking, English learning, job seeking management

// ========================================
// 📊 SETUP & INITIALIZATION
// ========================================

function setupBootcampSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create all necessary sheets
  createSheetIfNotExists(ss, 'Dashboard');
  createSheetIfNotExists(ss, 'Students');
  createSheetIfNotExists(ss, 'Technical_Learning');
  createSheetIfNotExists(ss, 'English_Learning');
  createSheetIfNotExists(ss, 'English_Corner');
  createSheetIfNotExists(ss, 'Job_Seeking');
  createSheetIfNotExists(ss, 'Projects');
  createSheetIfNotExists(ss, 'Assessments');
  createSheetIfNotExists(ss, 'Mentors');
  createSheetIfNotExists(ss, 'Reports');
  createSheetIfNotExists(ss, 'Settings');
  
  // Initialize all sheets with proper structure
  setupDashboard();
  setupStudents();
  setupTechnicalLearning();
  setupEnglishLearning();
  setupEnglishCorner();
  setupJobSeeking();
  setupProjects();
  setupAssessments();
  setupMentors();
  setupReports();
  setupSettings();
  
  // Create beautiful dashboard
  createBeautifulDashboard();
  
  SpreadsheetApp.getUi().alert('🎉 Bootcamp Management System initialized successfully!\n\nYour comprehensive learning and job seeking platform is ready.');
}

function createSheetIfNotExists(ss, sheetName) {
  if (!ss.getSheetByName(sheetName)) {
    ss.insertSheet(sheetName);
  }
}

// ========================================
// 👥 STUDENTS MANAGEMENT
// ========================================

function setupStudents() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Students');
  
  // Clear existing data
  sheet.clear();
  
  // Set up headers with beautiful formatting
  const headers = [
    'Student ID', 'Name', 'Email', 'Phone', 'Start Date', 'Current Level',
    'Technical Score', 'English Score', 'Job Readiness', 'Status', 'Mentor',
    'GitHub', 'LinkedIn', 'Portfolio', 'Notes', 'Created Date'
  ];
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  
  // Apply beautiful formatting
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('white');
  headerRange.setHorizontalAlignment('center');
  
  // Set up data validation for status
  const statusOptions = ['Active', 'Graduated', 'On Hold', 'Dropped Out'];
  const statusRange = sheet.getRange(2, 10, 1000, 1);
  statusRange.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(statusOptions, true)
    .build());
  
  // Set up data validation for current level
  const levelOptions = ['Beginner', 'Intermediate', 'Advanced', 'Expert'];
  const levelRange = sheet.getRange(2, 6, 1000, 1);
  levelRange.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(levelOptions, true)
    .build());
  
  // Auto-format columns
  sheet.getRange('E:E').setNumberFormat('mm/dd/yyyy');
  sheet.getRange('G:I').setNumberFormat('0.0');
  sheet.getRange('P:P').setNumberFormat('mm/dd/yyyy hh:mm:ss');
  
  // Freeze header row
  sheet.setFrozenRows(1);
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, headers.length);
}

function addStudent(name, email, phone, startDate, currentLevel, mentor) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Students');
  const today = new Date();
  const studentId = 'STU' + Utilities.getUuid().substring(0, 8).toUpperCase();
  
  const newRow = [
    studentId,           // Student ID
    name,               // Name
    email,              // Email
    phone,              // Phone
    startDate,          // Start Date
    currentLevel,       // Current Level
    0,                  // Technical Score
    0,                  // English Score
    0,                  // Job Readiness
    'Active',           // Status
    mentor,             // Mentor
    '',                 // GitHub
    '',                 // LinkedIn
    '',                 // Portfolio
    '',                 // Notes
    today               // Created Date
  ];
  
  sheet.appendRow(newRow);
  
  // Apply conditional formatting for scores
  const lastRow = sheet.getLastRow();
  const scoreRange = sheet.getRange(lastRow, 7, 1, 3);
  
  // Color coding based on scores
  const rule1 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(60)
    .setBackground('#f8d7da')
    .setRanges([scoreRange])
    .build();
  
  const rule2 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(60, 80)
    .setBackground('#fff3cd')
    .setRanges([scoreRange])
    .build();
  
  const rule3 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(80)
    .setBackground('#d4edda')
    .setRanges([scoreRange])
    .build();
  
  sheet.setConditionalFormatRules([rule1, rule2, rule3]);
  
  // Update dashboard
  updateDashboard();
  
  return studentId;
}

// ========================================
// 💻 TECHNICAL LEARNING MANAGEMENT
// ========================================

function setupTechnicalLearning() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Technical_Learning');
  
  // Clear existing data
  sheet.clear();
  
  // Set up headers
  const headers = [
    'Student ID', 'Topic', 'Category', 'Status', 'Progress %', 'Score',
    'Start Date', 'Completion Date', 'Instructor', 'Resources', 'Notes'
  ];
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#34a853');
  headerRange.setFontColor('white');
  
  // Set up data validation
  const categories = ['HTML/CSS', 'JavaScript', 'React', 'Node.js', 'Database', 'DevOps', 'Testing'];
  const statusOptions = ['Not Started', 'In Progress', 'Completed', 'Review Needed'];
  
  const categoryRange = sheet.getRange(2, 3, 1000, 1);
  categoryRange.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(categories, true)
    .build());
  
  const statusRange = sheet.getRange(2, 4, 1000, 1);
  statusRange.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(statusOptions, true)
    .build());
  
  // Auto-format columns
  sheet.getRange('E:E').setNumberFormat('0%');
  sheet.getRange('F:F').setNumberFormat('0.0');
  sheet.getRange('G:H').setNumberFormat('mm/dd/yyyy');
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, headers.length);
}

function addTechnicalLearning(studentId, topic, category, instructor) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Technical_Learning');
  const today = new Date();
  
  const newRow = [
    studentId,           // Student ID
    topic,              // Topic
    category,           // Category
    'Not Started',      // Status
    0,                  // Progress %
    0,                  // Score
    today,              // Start Date
    '',                 // Completion Date
    instructor,         // Instructor
    '',                 // Resources
    ''                  // Notes
  ];
  
  sheet.appendRow(newRow);
  
  // Update student's technical score
  updateStudentTechnicalScore(studentId);
  
  return true;
}

function updateStudentTechnicalScore(studentId) {
  const techSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Technical_Learning');
  const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Students');
  
  // Calculate average score for the student
  const data = techSheet.getDataRange().getValues();
  let totalScore = 0;
  let completedTopics = 0;
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0] === studentId && row[3] === 'Completed') {
      totalScore += parseFloat(row[5]) || 0;
      completedTopics++;
    }
  }
  
  const averageScore = completedTopics > 0 ? totalScore / completedTopics : 0;
  
  // Update student's technical score
  const studentsData = studentsSheet.getDataRange().getValues();
  for (let i = 1; i < studentsData.length; i++) {
    if (studentsData[i][0] === studentId) {
      studentsSheet.getRange(i + 1, 7).setValue(averageScore);
      break;
    }
  }
}

// ========================================
// 🗣️ ENGLISH LEARNING MANAGEMENT
// ========================================

function setupEnglishLearning() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('English_Learning');
  
  // Clear existing data
  sheet.clear();
  
  // Set up headers
  const headers = [
    'Student ID', 'Skill', 'Level', 'Status', 'Progress %', 'Score',
    'Start Date', 'Completion Date', 'Instructor', 'Resources', 'Notes'
  ];
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#FBBC04');
  headerRange.setFontColor('white');
  
  // Set up data validation
  const skills = ['Speaking', 'Listening', 'Reading', 'Writing', 'Interview Prep', 'Business English', 'English Corner'];
  const levels = ['Beginner', 'Intermediate', 'Advanced'];
  const statusOptions = ['Not Started', 'In Progress', 'Completed', 'Review Needed'];
  
  const skillRange = sheet.getRange(2, 2, 1000, 1);
  skillRange.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(skills, true)
    .build());
  
  const levelRange = sheet.getRange(2, 3, 1000, 1);
  levelRange.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(levels, true)
    .build());
  
  const statusRange = sheet.getRange(2, 4, 1000, 1);
  statusRange.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(statusOptions, true)
    .build());
  
  // Auto-format columns
  sheet.getRange('E:E').setNumberFormat('0%');
  sheet.getRange('F:F').setNumberFormat('0.0');
  sheet.getRange('G:H').setNumberFormat('mm/dd/yyyy');
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, headers.length);
}

function addEnglishLearning(studentId, skill, level, instructor) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('English_Learning');
  const today = new Date();
  
  const newRow = [
    studentId,           // Student ID
    skill,              // Skill
    level,              // Level
    'Not Started',      // Status
    0,                  // Progress %
    0,                  // Score
    today,              // Start Date
    '',                 // Completion Date
    instructor,         // Instructor
    '',                 // Resources
    ''                  // Notes
  ];
  
  sheet.appendRow(newRow);
  
  // Update student's English score
  updateStudentEnglishScore(studentId);
  
  return true;
}

function updateStudentEnglishScore(studentId) {
  const englishSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('English_Learning');
  const englishCornerSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('English_Corner');
  const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Students');
  
  // Calculate average score for the student from regular English learning
  const data = englishSheet.getDataRange().getValues();
  let totalScore = 0;
  let completedSkills = 0;
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0] === studentId && row[3] === 'Completed') {
      totalScore += parseFloat(row[5]) || 0;
      completedSkills++;
    }
  }
  
  // Calculate English Corner participation score
  const cornerData = englishCornerSheet.getDataRange().getValues();
  let cornerScore = 0;
  let cornerCount = 0;
  
  for (let i = 1; i < cornerData.length; i++) {
    const row = cornerData[i];
    if (row[0] === studentId && row[4] === 'Completed') {
      cornerScore += parseFloat(row[5]) || 0;
      cornerCount++;
    }
  }
  
  // Combine regular English and English Corner scores
  const totalEnglishScore = totalScore + cornerScore;
  const totalCompleted = completedSkills + cornerCount;
  const averageScore = totalCompleted > 0 ? totalEnglishScore / totalCompleted : 0;
  
  // Update student's English score
  const studentsData = studentsSheet.getDataRange().getValues();
  for (let i = 1; i < studentsData.length; i++) {
    if (studentsData[i][0] === studentId) {
      studentsSheet.getRange(i + 1, 8).setValue(averageScore);
      break;
    }
  }
}

// ========================================
// 🎤 ENGLISH CORNER MANAGEMENT
// ========================================

function setupEnglishCorner() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('English_Corner');
  
  // Clear existing data
  sheet.clear();
  
  // Set up headers
  const headers = [
    'Session ID', 'Student ID', 'Topic', 'Video URL', 'Status', 'Score',
    'Upload Date', 'Review Date', 'Instructor', 'Duration (min)', 'Feedback Count',
    'Timestamps', 'Overall Feedback', 'Improvement Areas', 'Notes'
  ];
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#FF6B9D');
  headerRange.setFontColor('white');
  
  // Set up data validation
  const topics = [
    'Self Introduction', 'Daily Conversation', 'Job Interview Practice',
    'Presentation Skills', 'Debate Practice', 'Storytelling',
    'Pronunciation Practice', 'Grammar Review', 'Vocabulary Building',
    'Cultural Discussion', 'News Discussion', 'Book Review'
  ];
  const statusOptions = ['Uploaded', 'Under Review', 'Feedback Given', 'Completed', 'Needs Improvement'];
  
  const topicRange = sheet.getRange(2, 3, 1000, 1);
  topicRange.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(topics, true)
    .build());
  
  const statusRange = sheet.getRange(2, 5, 1000, 1);
  statusRange.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(statusOptions, true)
    .build());
  
  // Auto-format columns
  sheet.getRange('F:F').setNumberFormat('0.0');
  sheet.getRange('G:H').setNumberFormat('mm/dd/yyyy');
  sheet.getRange('J:J').setNumberFormat('0');
  sheet.getRange('K:K').setNumberFormat('0');
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, headers.length);
}

function addEnglishCornerSession(studentId, topic, videoUrl, duration) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('English_Corner');
  const today = new Date();
  const sessionId = 'EC' + Utilities.getUuid().substring(0, 8).toUpperCase();
  
  const newRow = [
    sessionId,           // Session ID
    studentId,           // Student ID
    topic,               // Topic
    videoUrl,            // Video URL
    'Uploaded',          // Status
    0,                   // Score
    today,               // Upload Date
    '',                  // Review Date
    '',                  // Instructor
    duration,            // Duration (min)
    0,                   // Feedback Count
    '',                  // Timestamps
    '',                  // Overall Feedback
    '',                  // Improvement Areas
    ''                   // Notes
  ];
  
  sheet.appendRow(newRow);
  
  // Update student's English score
  updateStudentEnglishScore(studentId);
  
  return sessionId;
}

function addTimestampFeedback(sessionId, timestamp, feedback, instructor) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('English_Corner');
  const data = sheet.getDataRange().getValues();
  
  // Find the session
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0] === sessionId) {
      const currentTimestamps = row[11] || '';
      const currentFeedback = row[12] || '';
      const feedbackCount = parseInt(row[10]) || 0;
      
      // Add new timestamp feedback
      const newTimestamp = `${timestamp} - ${feedback} (by ${instructor})`;
      const updatedTimestamps = currentTimestamps ? currentTimestamps + '; ' + newTimestamp : newTimestamp;
      
      // Update the row
      sheet.getRange(i + 1, 12).setValue(updatedTimestamps);
      sheet.getRange(i + 1, 11).setValue(feedbackCount + 1);
      
      // Update status to 'Feedback Given' if it was 'Under Review'
      if (row[4] === 'Under Review') {
        sheet.getRange(i + 1, 5).setValue('Feedback Given');
      }
      
      break;
    }
  }
}

function addOverallFeedback(sessionId, overallFeedback, improvementAreas, score, instructor) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('English_Corner');
  const data = sheet.getDataRange().getValues();
  
  // Find the session
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0] === sessionId) {
      const today = new Date();
      
      // Update the row with overall feedback
      sheet.getRange(i + 1, 13).setValue(overallFeedback);
      sheet.getRange(i + 1, 14).setValue(improvementAreas);
      sheet.getRange(i + 1, 6).setValue(score);
      sheet.getRange(i + 1, 8).setValue(today);
      sheet.getRange(i + 1, 9).setValue(instructor);
      sheet.getRange(i + 1, 5).setValue('Completed');
      
      // Update student's English score
      updateStudentEnglishScore(row[1]);
      
      break;
    }
  }
}

function getEnglishCornerStats(studentId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('English_Corner');
  const data = sheet.getDataRange().getValues();
  
  let totalSessions = 0;
  let completedSessions = 0;
  let totalScore = 0;
  let totalDuration = 0;
  let totalFeedback = 0;
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[1] === studentId) {
      totalSessions++;
      totalDuration += parseInt(row[9]) || 0;
      totalFeedback += parseInt(row[10]) || 0;
      
      if (row[4] === 'Completed') {
        completedSessions++;
        totalScore += parseFloat(row[5]) || 0;
      }
    }
  }
  
  const averageScore = completedSessions > 0 ? totalScore / completedSessions : 0;
  const completionRate = totalSessions > 0 ? (completedSessions / totalSessions) * 100 : 0;
  
  return {
    totalSessions: totalSessions,
    completedSessions: completedSessions,
    averageScore: averageScore,
    completionRate: completionRate,
    totalDuration: totalDuration,
    totalFeedback: totalFeedback
  };
}

function generateEnglishCornerReport(studentId) {
  const stats = getEnglishCornerStats(studentId);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('English_Corner');
  const data = sheet.getDataRange().getValues();
  
  let recentSessions = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[1] === studentId) {
      recentSessions.push({
        topic: row[2],
        status: row[4],
        score: row[5],
        uploadDate: row[6],
        feedbackCount: row[10]
      });
    }
  }
  
  // Sort by upload date (most recent first)
  recentSessions.sort((a, b) => new Date(b.uploadDate) - new Date(a.uploadDate));
  
  let report = `🎤 English Corner Report for Student ${studentId}\n\n`;
  report += `📊 Overall Statistics:\n`;
  report += `• Total Sessions: ${stats.totalSessions}\n`;
  report += `• Completed Sessions: ${stats.completedSessions}\n`;
  report += `• Average Score: ${stats.averageScore.toFixed(1)}/10\n`;
  report += `• Completion Rate: ${stats.completionRate.toFixed(1)}%\n`;
  report += `• Total Duration: ${stats.totalDuration} minutes\n`;
  report += `• Total Feedback Points: ${stats.totalFeedback}\n\n`;
  
  report += `📝 Recent Sessions:\n`;
  const recentCount = Math.min(5, recentSessions.length);
  for (let i = 0; i < recentCount; i++) {
    const session = recentSessions[i];
    report += `• ${session.topic} (${session.status}) - Score: ${session.score}/10 - Feedback: ${session.feedbackCount} points\n`;
  }
  
  return report;
}

// ========================================
// 💼 JOB SEEKING MANAGEMENT
// ========================================

function setupJobSeeking() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Job_Seeking');
  
  // Clear existing data
  sheet.clear();
  
  // Set up headers
  const headers = [
    'Student ID', 'Company', 'Position', 'Status', 'Application Date',
    'Interview Date', 'Response', 'Salary Range', 'Location', 'Notes'
  ];
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#EA4335');
  headerRange.setFontColor('white');
  
  // Set up data validation
  const statusOptions = ['Applied', 'Interview Scheduled', 'Interview Completed', 'Offer Received', 'Rejected', 'Accepted'];
  const responseOptions = ['Pending', 'Positive', 'Negative', 'No Response'];
  
  const statusRange = sheet.getRange(2, 4, 1000, 1);
  statusRange.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(statusOptions, true)
    .build());
  
  const responseRange = sheet.getRange(2, 7, 1000, 1);
  responseRange.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(responseOptions, true)
    .build());
  
  // Auto-format columns
  sheet.getRange('E:F').setNumberFormat('mm/dd/yyyy');
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, headers.length);
}

function addJobApplication(studentId, company, position, salaryRange, location) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Job_Seeking');
  const today = new Date();
  
  const newRow = [
    studentId,           // Student ID
    company,             // Company
    position,            // Position
    'Applied',           // Status
    today,               // Application Date
    '',                  // Interview Date
    'Pending',           // Response
    salaryRange,         // Salary Range
    location,            // Location
    ''                   // Notes
  ];
  
  sheet.appendRow(newRow);
  
  // Update student's job readiness score
  updateStudentJobReadiness(studentId);
  
  return true;
}

function updateStudentJobReadiness(studentId) {
  const jobSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Job_Seeking');
  const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Students');
  
  // Calculate job readiness based on applications and responses
  const data = jobSheet.getDataRange().getValues();
  let totalApplications = 0;
  let positiveResponses = 0;
  let interviewsScheduled = 0;
  let offersReceived = 0;
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0] === studentId) {
      totalApplications++;
      
      if (row[6] === 'Positive') positiveResponses++;
      if (row[3] === 'Interview Scheduled' || row[3] === 'Interview Completed') interviewsScheduled++;
      if (row[3] === 'Offer Received' || row[3] === 'Accepted') offersReceived++;
    }
  }
  
  // Calculate job readiness score (0-100)
  let jobReadiness = 0;
  if (totalApplications > 0) {
    const responseRate = (positiveResponses / totalApplications) * 30;
    const interviewRate = (interviewsScheduled / totalApplications) * 40;
    const offerRate = (offersReceived / totalApplications) * 30;
    jobReadiness = responseRate + interviewRate + offerRate;
  }
  
  // Update student's job readiness score
  const studentsData = studentsSheet.getDataRange().getValues();
  for (let i = 1; i < studentsData.length; i++) {
    if (studentsData[i][0] === studentId) {
      studentsSheet.getRange(i + 1, 9).setValue(jobReadiness);
      break;
    }
  }
}

// ========================================
// 📊 DASHBOARD CREATION
// ========================================

function setupDashboard() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
  sheet.clear();
  
  // Set up dashboard layout
  sheet.getRange('A1').setValue('📊 BOOTCAMP DASHBOARD');
  sheet.getRange('A1').setFontSize(24);
  sheet.getRange('A1').setFontWeight('bold');
  
  // Create summary cards
  createSummaryCards();
  
  // Create charts area
  createChartsArea();
  
  // Create recent activities table
  createRecentActivitiesTable();
}

function createBeautifulDashboard() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
  
  // Clear and set up fresh dashboard
  sheet.clear();
  
  // Main title
  const titleCell = sheet.getRange('A1');
  titleCell.setValue('🎓 IT FRONTEND BOOTCAMP MANAGEMENT');
  titleCell.setFontSize(28);
  titleCell.setFontWeight('bold');
  titleCell.setFontColor('#1a73e8');
  sheet.getRange('A1:H1').merge();
  titleCell.setHorizontalAlignment('center');
  
  // Subtitle
  const subtitleCell = sheet.getRange('A2');
  subtitleCell.setValue('Comprehensive learning and job seeking platform');
  subtitleCell.setFontSize(14);
  subtitleCell.setFontColor('#5f6368');
  sheet.getRange('A2:H2').merge();
  subtitleCell.setHorizontalAlignment('center');
  
  // Create summary cards
  createSummaryCards();
  
  // Create charts and visualizations
  createChartsArea();
  
  // Create recent activities
  createRecentActivitiesTable();
  
  // Add action buttons area
  createActionButtons();
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, 8);
}

function createSummaryCards() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
  
  // Card 1: Total Students
  sheet.getRange('A4').setValue('👥 TOTAL STUDENTS');
  sheet.getRange('A5').setValue('=COUNTA(Students!A:A)-1');
  sheet.getRange('A4:A5').setFontWeight('bold');
  sheet.getRange('A4:B5').setBackground('#e8f5e8');
  sheet.getRange('A4:B5').setBorder(true, true, true, true, true, true);
  
  // Card 2: Active Students
  sheet.getRange('C4').setValue('✅ ACTIVE STUDENTS');
  sheet.getRange('C5').setValue('=COUNTIFS(Students!J:J,"Active")');
  sheet.getRange('C4:C5').setFontWeight('bold');
  sheet.getRange('C4:D5').setBackground('#e3f2fd');
  sheet.getRange('C4:D5').setBorder(true, true, true, true, true, true);
  
  // Card 3: Average Technical Score
  sheet.getRange('E4').setValue('💻 AVG TECHNICAL SCORE');
  sheet.getRange('E5').setValue('=AVERAGEIFS(Students!G:G,Students!J:J,"Active")');
  sheet.getRange('E5').setNumberFormat('0.0');
  sheet.getRange('E4:E5').setFontWeight('bold');
  sheet.getRange('E4:F5').setBackground('#fff3e0');
  sheet.getRange('E4:F5').setBorder(true, true, true, true, true, true);
  
  // Card 4: Average English Score
  sheet.getRange('G4').setValue('🗣️ AVG ENGLISH SCORE');
  sheet.getRange('G5').setValue('=AVERAGEIFS(Students!H:H,Students!J:J,"Active")');
  sheet.getRange('G5').setNumberFormat('0.0');
  sheet.getRange('G4:G5').setFontWeight('bold');
  sheet.getRange('G4:H5').setBackground('#f3e5f5');
  sheet.getRange('G4:H5').setBorder(true, true, true, true, true, true);
}

function createChartsArea() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
  
  // Charts section title
  sheet.getRange('A7').setValue('📊 LEARNING ANALYTICS');
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
  sheet.getRange('A15').setValue('Student Progress by Level:');
  sheet.getRange('A16').setValue('=QUERY(Students!A:F,"select F, count(A) where A is not null group by F",1)');
  
  // Add more chart data ranges
  sheet.getRange('A18').setValue('Technical Learning Progress:');
  sheet.getRange('A19').setValue('=QUERY(Technical_Learning!A:D,"select D, count(A) where A is not null group by D",1)');
  
  sheet.getRange('A21').setValue('Job Application Status:');
  sheet.getRange('A22').setValue('=QUERY(Job_Seeking!A:D,"select D, count(A) where A is not null group by D",1)');
}

function createRecentActivitiesTable() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
  
  // Recent activities title
  sheet.getRange('A20').setValue('🕒 RECENT ACTIVITIES');
  sheet.getRange('A20').setFontSize(18);
  sheet.getRange('A20').setFontWeight('bold');
  sheet.getRange('A20').setFontColor('#1a73e8');
  
  // Create recent activities query
  sheet.getRange('A22').setValue('=QUERY({Students!A:B,Students!P:P},"select A, B, C where A is not null order by C desc limit 10",1)');
  
  // Format the table
  const tableRange = sheet.getRange('A22:C31');
  tableRange.setBorder(true, true, true, true, true, true);
  
  // Header formatting
  const headerRange = sheet.getRange('A22:C22');
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
  sheet.getRange('A37').setValue('💡 Quick actions:');
  sheet.getRange('A38').setValue('• Add new student');
  sheet.getRange('A39').setValue('• Track learning progress');
  sheet.getRange('A40').setValue('• Monitor job applications');
  sheet.getRange('A41').setValue('• Generate reports');
}

// ========================================
// 🔄 DASHBOARD UPDATE
// ========================================

function updateDashboard() {
  // Refresh dashboard data
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
  
  // Force recalculation of all formulas by refreshing the sheet
  SpreadsheetApp.flush();
  
  // Show success message
  console.log('Dashboard updated successfully');
}

// ========================================
// 📁 PROJECTS MANAGEMENT
// ========================================

function setupProjects() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
  
  // Clear existing data
  sheet.clear();
  
  // Set up headers
  const headers = [
    'Project ID', 'Student ID', 'Project Name', 'Category', 'Status', 'Progress %',
    'Start Date', 'Due Date', 'Completion Date', 'Technologies', 'GitHub Link',
    'Live Demo', 'Score', 'Feedback', 'Notes'
  ];
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#9c27b0');
  headerRange.setFontColor('white');
  
  // Set up data validation
  const categories = ['Portfolio', 'E-commerce', 'Social Media', 'Dashboard', 'API', 'Mobile App'];
  const statusOptions = ['Planning', 'In Progress', 'Under Review', 'Completed', 'On Hold'];
  
  const categoryRange = sheet.getRange(2, 4, 1000, 1);
  categoryRange.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(categories, true)
    .build());
  
  const statusRange = sheet.getRange(2, 5, 1000, 1);
  statusRange.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(statusOptions, true)
    .build());
  
  // Auto-format columns
  sheet.getRange('F:F').setNumberFormat('0%');
  sheet.getRange('G:I').setNumberFormat('mm/dd/yyyy');
  sheet.getRange('N:N').setNumberFormat('0.0');
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, headers.length);
}

function addProject(studentId, projectName, category, dueDate, technologies) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
  const today = new Date();
  const projectId = 'PRJ' + Utilities.getUuid().substring(0, 8).toUpperCase();
  
  const newRow = [
    projectId,           // Project ID
    studentId,           // Student ID
    projectName,         // Project Name
    category,            // Category
    'Planning',          // Status
    0,                   // Progress %
    today,               // Start Date
    dueDate,             // Due Date
    '',                  // Completion Date
    technologies,        // Technologies
    '',                  // GitHub Link
    '',                  // Live Demo
    0,                   // Score
    '',                  // Feedback
    ''                   // Notes
  ];
  
  sheet.appendRow(newRow);
  
  return projectId;
}

// ========================================
// 📝 ASSESSMENTS MANAGEMENT
// ========================================

function setupAssessments() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assessments');
  
  // Clear existing data
  sheet.clear();
  
  // Set up headers
  const headers = [
    'Assessment ID', 'Student ID', 'Type', 'Subject', 'Score', 'Max Score',
    'Date', 'Instructor', 'Status', 'Feedback', 'Retake Date', 'Notes'
  ];
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#FF6B6B');
  headerRange.setFontColor('white');
  
  // Set up data validation
  const types = ['Quiz', 'Exam', 'Project Review', 'Interview Prep', 'Technical Test'];
  const subjects = ['HTML/CSS', 'JavaScript', 'React', 'Node.js', 'Database', 'English', 'Interview Skills'];
  const statusOptions = ['Scheduled', 'Completed', 'Passed', 'Failed', 'Retake'];
  
  const typeRange = sheet.getRange(2, 3, 1000, 1);
  typeRange.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(types, true)
    .build());
  
  const subjectRange = sheet.getRange(2, 4, 1000, 1);
  subjectRange.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(subjects, true)
    .build());
  
  const statusRange = sheet.getRange(2, 9, 1000, 1);
  statusRange.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(statusOptions, true)
    .build());
  
  // Auto-format columns
  sheet.getRange('E:F').setNumberFormat('0.0');
  sheet.getRange('G:G').setNumberFormat('mm/dd/yyyy');
  sheet.getRange('K:K').setNumberFormat('mm/dd/yyyy');
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, headers.length);
}

function addAssessment(studentId, type, subject, maxScore, instructor) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assessments');
  const today = new Date();
  const assessmentId = 'ASS' + Utilities.getUuid().substring(0, 8).toUpperCase();
  
  const newRow = [
    assessmentId,        // Assessment ID
    studentId,           // Student ID
    type,                // Type
    subject,             // Subject
    0,                   // Score
    maxScore,            // Max Score
    today,               // Date
    instructor,          // Instructor
    'Scheduled',         // Status
    '',                  // Feedback
    '',                  // Retake Date
    ''                   // Notes
  ];
  
  sheet.appendRow(newRow);
  
  return assessmentId;
}

// ========================================
// 👨‍🏫 MENTORS MANAGEMENT
// ========================================

function setupMentors() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Mentors');
  
  // Clear existing data
  sheet.clear();
  
  // Set up headers
  const headers = [
    'Mentor ID', 'Name', 'Email', 'Phone', 'Specialization', 'Experience Years',
    'Students Assigned', 'Status', 'Hourly Rate', 'Availability', 'Notes'
  ];
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4ECDC4');
  headerRange.setFontColor('white');
  
  // Set up data validation
  const specializations = ['Frontend', 'Backend', 'Full Stack', 'DevOps', 'English', 'Interview Prep'];
  const statusOptions = ['Active', 'Inactive', 'On Leave'];
  
  const specRange = sheet.getRange(2, 5, 1000, 1);
  specRange.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(specializations, true)
    .build());
  
  const statusRange = sheet.getRange(2, 8, 1000, 1);
  statusRange.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(statusOptions, true)
    .build());
  
  // Auto-format columns
  sheet.getRange('F:F').setNumberFormat('0');
  sheet.getRange('G:G').setNumberFormat('0');
  sheet.getRange('I:I').setNumberFormat('$#,##0.00');
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, headers.length);
}

function addMentor(name, email, phone, specialization, experienceYears, hourlyRate) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Mentors');
  const mentorId = 'MEN' + Utilities.getUuid().substring(0, 8).toUpperCase();
  
  const newRow = [
    mentorId,            // Mentor ID
    name,                // Name
    email,               // Email
    phone,               // Phone
    specialization,      // Specialization
    experienceYears,     // Experience Years
    0,                   // Students Assigned
    'Active',            // Status
    hourlyRate,          // Hourly Rate
    'Available',         // Availability
    ''                   // Notes
  ];
  
  sheet.appendRow(newRow);
  
  return mentorId;
}

// ========================================
// 📈 REPORTS & ANALYTICS
// ========================================

function setupReports() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Reports');
  
  // Clear existing data
  sheet.clear();
  
  // Set up headers
  const headers = ['Report Type', 'Date Range', 'Total Students', 'Active Students', 'Graduation Rate', 'Job Placement Rate'];
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#FF9800');
  headerRange.setFontColor('white');
  
  // Generate monthly reports
  generateMonthlyReports();
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, headers.length);
}

function generateMonthlyReports() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Reports');
  const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Students');
  
  // Get last 6 months of data
  const today = new Date();
  const reports = [];
  
  for (let i = 0; i < 6; i++) {
    const reportDate = new Date(today.getFullYear(), today.getMonth() - i, 1);
    const monthName = reportDate.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
    
    const totalStudents = calculateTotalStudents(reportDate.getMonth() + 1, reportDate.getFullYear());
    const activeStudents = calculateActiveStudents(reportDate.getMonth() + 1, reportDate.getFullYear());
    const graduationRate = calculateGraduationRate(reportDate.getMonth() + 1, reportDate.getFullYear());
    const jobPlacementRate = calculateJobPlacementRate(reportDate.getMonth() + 1, reportDate.getFullYear());
    
    reports.push([
      'Monthly Report',
      monthName,
      totalStudents,
      activeStudents,
      graduationRate,
      jobPlacementRate
    ]);
  }
  
  // Add reports to sheet
  if (reports.length > 0) {
    const dataRange = sheet.getRange(2, 1, reports.length, reports[0].length);
    dataRange.setValues(reports);
    
    // Format percentage columns
    sheet.getRange(2, 5, reports.length, 2).setNumberFormat('0.0%');
  }
}

function calculateTotalStudents(month, year) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Students');
  const data = sheet.getDataRange().getValues();
  
  let total = 0;
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const startDate = new Date(row[4]);
    
    if (startDate.getMonth() + 1 <= month && startDate.getFullYear() <= year) {
      total++;
    }
  }
  
  return total;
}

function calculateActiveStudents(month, year) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Students');
  const data = sheet.getDataRange().getValues();
  
  let active = 0;
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const startDate = new Date(row[4]);
    const status = row[9];
    
    if (startDate.getMonth() + 1 <= month && startDate.getFullYear() <= year && status === 'Active') {
      active++;
    }
  }
  
  return active;
}

function calculateGraduationRate(month, year) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Students');
  const data = sheet.getDataRange().getValues();
  
  let total = 0;
  let graduated = 0;
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const startDate = new Date(row[4]);
    const status = row[9];
    
    if (startDate.getMonth() + 1 <= month && startDate.getFullYear() <= year) {
      total++;
      if (status === 'Graduated') {
        graduated++;
      }
    }
  }
  
  return total > 0 ? graduated / total : 0;
}

function calculateJobPlacementRate(month, year) {
  const jobSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Job_Seeking');
  const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Students');
  
  const jobData = jobSheet.getDataRange().getValues();
  const studentsData = studentsSheet.getDataRange().getValues();
  
  let totalStudents = 0;
  let placedStudents = 0;
  
  // Count students who started in this month/year
  for (let i = 1; i < studentsData.length; i++) {
    const row = studentsData[i];
    const startDate = new Date(row[4]);
    
    if (startDate.getMonth() + 1 <= month && startDate.getFullYear() <= year) {
      totalStudents++;
      
      // Check if this student has accepted a job
      const studentId = row[0];
      for (let j = 1; j < jobData.length; j++) {
        const jobRow = jobData[j];
        if (jobRow[0] === studentId && jobRow[3] === 'Accepted') {
          placedStudents++;
          break;
        }
      }
    }
  }
  
  return totalStudents > 0 ? placedStudents / totalStudents : 0;
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
    ['Bootcamp Name', 'IT Frontend Bootcamp', 'Name of the bootcamp program'],
    ['Location', 'Online', 'Bootcamp location or delivery method'],
    ['Duration', '6 months', 'Program duration'],
    ['Max Students', '50', 'Maximum number of students per cohort'],
    ['Passing Score', '70%', 'Minimum score to pass assessments'],
    ['Job Placement Target', '80%', 'Target job placement rate'],
    ['English Proficiency', 'B2', 'Required English proficiency level'],
    ['Technical Stack', 'HTML/CSS, JavaScript, React, Node.js', 'Core technical skills'],
    ['Assessment Frequency', 'Weekly', 'How often assessments are conducted'],
    ['Mentor Ratio', '1:10', 'Mentor to student ratio']
  ];
  
  const dataRange = sheet.getRange(1, 1, settings.length, settings[0].length);
  dataRange.setValues(settings);
  
  // Format headers
  const headerRange = sheet.getRange(1, 1, 1, 3);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#607d8b');
  headerRange.setFontColor('white');
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, 3);
}

// ========================================
// 🎨 UTILITY FUNCTIONS
// ========================================

function generateCustomReport() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.prompt(
    'Generate Custom Report',
    'Enter report type (1=Student Progress, 2=Learning Analytics, 3=Job Placement):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const input = result.getResponseText();
    const reportType = parseInt(input);
    
    let report = '';
    
    switch(reportType) {
      case 1:
        report = generateStudentProgressReport();
        break;
      case 2:
        report = generateLearningAnalyticsReport();
        break;
      case 3:
        report = generateJobPlacementReport();
        break;
      default:
        ui.alert('❌ Invalid report type. Please enter 1, 2, or 3.');
        return;
    }
    
    ui.alert('Custom Report', report, ui.ButtonSet.OK);
  }
}

function generateStudentProgressReport() {
  const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Students');
  const data = studentsSheet.getDataRange().getValues();
  
  let totalStudents = 0;
  let activeStudents = 0;
  let graduatedStudents = 0;
  let avgTechnicalScore = 0;
  let avgEnglishScore = 0;
  let totalTechnicalScore = 0;
  let totalEnglishScore = 0;
  let activeCount = 0;
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    totalStudents++;
    
    if (row[9] === 'Active') {
      activeStudents++;
      totalTechnicalScore += parseFloat(row[6]) || 0;
      totalEnglishScore += parseFloat(row[7]) || 0;
      activeCount++;
    } else if (row[9] === 'Graduated') {
      graduatedStudents++;
    }
  }
  
  avgTechnicalScore = activeCount > 0 ? totalTechnicalScore / activeCount : 0;
  avgEnglishScore = activeCount > 0 ? totalEnglishScore / activeCount : 0;
  
  return `📊 Student Progress Report\n\n` +
         `👥 Total Students: ${totalStudents}\n` +
         `✅ Active Students: ${activeStudents}\n` +
         `🎓 Graduated Students: ${graduatedStudents}\n` +
         `💻 Average Technical Score: ${avgTechnicalScore.toFixed(1)}\n` +
         `🗣️ Average English Score: ${avgEnglishScore.toFixed(1)}\n` +
         `📈 Graduation Rate: ${((graduatedStudents / totalStudents) * 100).toFixed(1)}%`;
}

function generateLearningAnalyticsReport() {
  const techSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Technical_Learning');
  const englishSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('English_Learning');
  
  const techData = techSheet.getDataRange().getValues();
  const englishData = englishSheet.getDataRange().getValues();
  
  let techCompleted = 0;
  let techTotal = 0;
  let englishCompleted = 0;
  let englishTotal = 0;
  
  for (let i = 1; i < techData.length; i++) {
    techTotal++;
    if (techData[i][3] === 'Completed') {
      techCompleted++;
    }
  }
  
  for (let i = 1; i < englishData.length; i++) {
    englishTotal++;
    if (englishData[i][3] === 'Completed') {
      englishCompleted++;
    }
  }
  
  const techCompletionRate = techTotal > 0 ? (techCompleted / techTotal) * 100 : 0;
  const englishCompletionRate = englishTotal > 0 ? (englishCompleted / englishTotal) * 100 : 0;
  
  return `📈 Learning Analytics Report\n\n` +
         `💻 Technical Learning:\n` +
         `   • Total Topics: ${techTotal}\n` +
         `   • Completed: ${techCompleted}\n` +
         `   • Completion Rate: ${techCompletionRate.toFixed(1)}%\n\n` +
         `🗣️ English Learning:\n` +
         `   • Total Skills: ${englishTotal}\n` +
         `   • Completed: ${englishCompleted}\n` +
         `   • Completion Rate: ${englishCompletionRate.toFixed(1)}%`;
}

function generateJobPlacementReport() {
  const jobSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Job_Seeking');
  const data = jobSheet.getDataRange().getValues();
  
  let totalApplications = 0;
  let interviewsScheduled = 0;
  let offersReceived = 0;
  let acceptedJobs = 0;
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    totalApplications++;
    
    if (row[3] === 'Interview Scheduled' || row[3] === 'Interview Completed') {
      interviewsScheduled++;
    }
    if (row[3] === 'Offer Received' || row[3] === 'Accepted') {
      offersReceived++;
    }
    if (row[3] === 'Accepted') {
      acceptedJobs++;
    }
  }
  
  const interviewRate = totalApplications > 0 ? (interviewsScheduled / totalApplications) * 100 : 0;
  const offerRate = totalApplications > 0 ? (offersReceived / totalApplications) * 100 : 0;
  const acceptanceRate = totalApplications > 0 ? (acceptedJobs / totalApplications) * 100 : 0;
  
  return `💼 Job Placement Report\n\n` +
         `📝 Total Applications: ${totalApplications}\n` +
         `📅 Interviews Scheduled: ${interviewsScheduled}\n` +
         `🎯 Offers Received: ${offersReceived}\n` +
         `✅ Jobs Accepted: ${acceptedJobs}\n\n` +
         `📊 Success Rates:\n` +
         `   • Interview Rate: ${interviewRate.toFixed(1)}%\n` +
         `   • Offer Rate: ${offerRate.toFixed(1)}%\n` +
         `   • Acceptance Rate: ${acceptanceRate.toFixed(1)}%`;
}

function showAnalyticsDialog() {
  const ui = SpreadsheetApp.getUi();
  
  const analytics = `📊 Bootcamp Analytics Overview\n\n` +
                   `🎓 Current Status:\n` +
                   `• Total Students: ${getTotalStudents()}\n` +
                   `• Active Students: ${getActiveStudents()}\n` +
                   `• Average Technical Score: ${getAverageTechnicalScore().toFixed(1)}\n` +
                   `• Average English Score: ${getAverageEnglishScore().toFixed(1)}\n\n` +
                   `📈 Learning Progress:\n` +
                   `• Technical Completion: ${getTechnicalCompletionRate().toFixed(1)}%\n` +
                   `• English Completion: ${getEnglishCompletionRate().toFixed(1)}%\n` +
                   `• Job Placement Rate: ${getJobPlacementRate().toFixed(1)}%`;
  
  ui.alert('Analytics', analytics, ui.ButtonSet.OK);
}

function exportBootcampData() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    'Export Data',
    'Export bootcamp data to CSV format?',
    ui.ButtonSet.YES_NO
  );
  
  if (result === ui.Button.YES) {
    const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Students');
    const data = studentsSheet.getDataRange().getValues();
    
    let csvContent = '';
    for (let i = 0; i < data.length; i++) {
      csvContent += data[i].join(',') + '\n';
    }
    
    // Create a temporary file
    const fileName = 'bootcamp_data_' + new Date().toISOString().split('T')[0] + '.csv';
    const file = DriveApp.createFile(fileName, csvContent, MimeType.CSV);
    
    ui.alert('Export Complete', `Data exported to: ${file.getName()}\nFile ID: ${file.getId()}`, ui.ButtonSet.OK);
  }
}

function showHelpDialog() {
  const ui = SpreadsheetApp.getUi();
  
  const help = `🆘 Bootcamp Management System Help\n\n` +
               `📝 How to use:\n` +
               `1. Add students using the menu\n` +
               `2. Track technical and English learning\n` +
               `3. Monitor job applications\n` +
               `4. Generate reports for analysis\n\n` +
               `💡 Features:\n` +
               `• Student management\n` +
               `• Learning progress tracking\n` +
               `• Job application monitoring\n` +
               `• Assessment management\n` +
               `• Mentor assignment\n` +
               `• Analytics and reporting\n\n` +
               `🔧 Quick Actions:\n` +
               `• Add Student: Name|Email|Phone|StartDate|Level|Mentor\n` +
               `• Add Technical: StudentID|Topic|Category|Instructor\n` +
               `• Add English: StudentID|Skill|Level|Instructor\n` +
               `• Add English Corner: StudentID|Topic|VideoURL|Duration\n` +
               `• Add Job: StudentID|Company|Position|Salary|Location`;
  
  ui.alert('Help', help, ui.ButtonSet.OK);
}

function showSettingsDialog() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    'Settings',
    'What would you like to configure?\n\n1. Bootcamp settings\n2. Assessment criteria\n3. Learning objectives\n4. Export settings',
    ui.ButtonSet.YES_NO_CANCEL
  );
  
  if (result === ui.Button.YES) {
    ui.alert('Settings', 'Settings dialog will be implemented in future versions.', ui.ButtonSet.OK);
  }
}

// ========================================
// 📊 UTILITY FUNCTIONS
// ========================================

function getTotalStudents() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Students');
  return sheet.getLastRow() - 1;
}

function getActiveStudents() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Students');
  const data = sheet.getDataRange().getValues();
  let count = 0;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][9] === 'Active') {
      count++;
    }
  }
  
  return count;
}

function getAverageTechnicalScore() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Students');
  const data = sheet.getDataRange().getValues();
  let total = 0;
  let count = 0;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][9] === 'Active') {
      total += parseFloat(data[i][6]) || 0;
      count++;
    }
  }
  
  return count > 0 ? total / count : 0;
}

function getAverageEnglishScore() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Students');
  const data = sheet.getDataRange().getValues();
  let total = 0;
  let count = 0;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][9] === 'Active') {
      total += parseFloat(data[i][7]) || 0;
      count++;
    }
  }
  
  return count > 0 ? total / count : 0;
}

function getTechnicalCompletionRate() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Technical_Learning');
  const data = sheet.getDataRange().getValues();
  let total = 0;
  let completed = 0;
  
  for (let i = 1; i < data.length; i++) {
    total++;
    if (data[i][3] === 'Completed') {
      completed++;
    }
  }
  
  return total > 0 ? (completed / total) * 100 : 0;
}

function getEnglishCompletionRate() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('English_Learning');
  const data = sheet.getDataRange().getValues();
  let total = 0;
  let completed = 0;
  
  for (let i = 1; i < data.length; i++) {
    total++;
    if (data[i][3] === 'Completed') {
      completed++;
    }
  }
  
  return total > 0 ? (completed / total) * 100 : 0;
}

function getJobPlacementRate() {
  const jobSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Job_Seeking');
  const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Students');
  
  const jobData = jobSheet.getDataRange().getValues();
  const studentsData = studentsSheet.getDataRange().getValues();
  
  let totalStudents = 0;
  let placedStudents = 0;
  
  for (let i = 1; i < studentsData.length; i++) {
    if (studentsData[i][9] === 'Active' || studentsData[i][9] === 'Graduated') {
      totalStudents++;
      
      const studentId = studentsData[i][0];
      for (let j = 1; j < jobData.length; j++) {
        if (jobData[j][0] === studentId && jobData[j][3] === 'Accepted') {
          placedStudents++;
          break;
        }
      }
    }
  }
  
  return totalStudents > 0 ? (placedStudents / totalStudents) * 100 : 0;
}

// ========================================
// 🎯 USER INTERFACE FUNCTIONS
// ========================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // Create custom menu
  ui.createMenu('🎓 Bootcamp Manager')
    .addItem('➕ Add Student', 'showAddStudentDialog')
    .addItem('💻 Add Technical Learning', 'showAddTechnicalDialog')
    .addItem('🗣️ Add English Learning', 'showAddEnglishDialog')
    .addItem('🎤 Add English Corner Session', 'showAddEnglishCornerDialog')
    .addItem('💼 Add Job Application', 'showAddJobDialog')
    .addItem('📊 Generate Report', 'generateCustomReport')
    .addItem('🔄 Update Dashboard', 'updateDashboard')
    .addItem('⚙️ Settings', 'showSettingsDialog')
    .addSeparator()
    .addItem('📈 View Analytics', 'showAnalyticsDialog')
    .addItem('💾 Export Data', 'exportBootcampData')
    .addItem('🆘 Help', 'showHelpDialog')
    .addToUi();
}

function showAddStudentDialog() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.prompt(
    'Add New Student',
    'Enter student details (format: Name|Email|Phone|StartDate|CurrentLevel|Mentor):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const input = result.getResponseText();
    const parts = input.split('|');
    
    if (parts.length >= 5) {
      const name = parts[0].trim();
      const email = parts[1].trim();
      const phone = parts[2].trim();
      const startDate = new Date(parts[3].trim());
      const currentLevel = parts[4].trim();
      const mentor = parts[5] ? parts[5].trim() : '';
      
      if (name && email && !isNaN(startDate.getTime())) {
        const studentId = addStudent(name, email, phone, startDate, currentLevel, mentor);
        ui.alert('✅ Student added successfully!\nStudent ID: ' + studentId);
      } else {
        ui.alert('❌ Invalid input format. Please check your data.');
      }
    } else {
      ui.alert('❌ Invalid input format. Please use: Name|Email|Phone|StartDate|CurrentLevel|Mentor');
    }
  }
}

function showAddTechnicalDialog() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.prompt(
    'Add Technical Learning',
    'Enter learning details (format: StudentID|Topic|Category|Instructor):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const input = result.getResponseText();
    const parts = input.split('|');
    
    if (parts.length >= 4) {
      const studentId = parts[0].trim();
      const topic = parts[1].trim();
      const category = parts[2].trim();
      const instructor = parts[3].trim();
      
      if (studentId && topic && category) {
        addTechnicalLearning(studentId, topic, category, instructor);
        ui.alert('✅ Technical learning added successfully!');
      } else {
        ui.alert('❌ Invalid input format. Please check your data.');
      }
    } else {
      ui.alert('❌ Invalid input format. Please use: StudentID|Topic|Category|Instructor');
    }
  }
}

function showAddEnglishDialog() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.prompt(
    'Add English Learning',
    'Enter learning details (format: StudentID|Skill|Level|Instructor):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const input = result.getResponseText();
    const parts = input.split('|');
    
    if (parts.length >= 4) {
      const studentId = parts[0].trim();
      const skill = parts[1].trim();
      const level = parts[2].trim();
      const instructor = parts[3].trim();
      
      if (studentId && skill && level) {
        addEnglishLearning(studentId, skill, level, instructor);
        ui.alert('✅ English learning added successfully!');
      } else {
        ui.alert('❌ Invalid input format. Please check your data.');
      }
    } else {
      ui.alert('❌ Invalid input format. Please use: StudentID|Skill|Level|Instructor');
    }
  }
}

function showAddEnglishCornerDialog() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.prompt(
    'Add English Corner Session',
    'Enter session details (format: StudentID|Topic|VideoURL|Duration):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const input = result.getResponseText();
    const parts = input.split('|');
    
    if (parts.length >= 4) {
      const studentId = parts[0].trim();
      const topic = parts[1].trim();
      const videoUrl = parts[2].trim();
      const duration = parseInt(parts[3].trim());
      
      if (studentId && topic && videoUrl && !isNaN(duration)) {
        const sessionId = addEnglishCornerSession(studentId, topic, videoUrl, duration);
        ui.alert('✅ English Corner session added successfully!\nSession ID: ' + sessionId);
      } else {
        ui.alert('❌ Invalid input format. Please check your data.');
      }
    } else {
      ui.alert('❌ Invalid input format. Please use: StudentID|Topic|VideoURL|Duration');
    }
  }
}

function showAddJobDialog() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.prompt(
    'Add Job Application',
    'Enter job details (format: StudentID|Company|Position|SalaryRange|Location):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const input = result.getResponseText();
    const parts = input.split('|');
    
    if (parts.length >= 5) {
      const studentId = parts[0].trim();
      const company = parts[1].trim();
      const position = parts[2].trim();
      const salaryRange = parts[3].trim();
      const location = parts[4].trim();
      
      if (studentId && company && position) {
        addJobApplication(studentId, company, position, salaryRange, location);
        ui.alert('✅ Job application added successfully!');
      } else {
        ui.alert('❌ Invalid input format. Please check your data.');
      }
    } else {
      ui.alert('❌ Invalid input format. Please use: StudentID|Company|Position|SalaryRange|Location');
    }
  }
}

// ========================================
// 🚀 INITIALIZATION
// ========================================

// Auto-run setup when script is first loaded
function initializeBootcampSystem() {
  setupBootcampSystem();
}
