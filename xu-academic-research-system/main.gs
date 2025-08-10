/**
 * XU Exponential University Academic Research Management System
 * A comprehensive system for tracking academic progress, abilities, and research projects
 * 
 * Author: AI Assistant
 * Version: 1.0
 * Date: 2024
 */

// System Configuration
const SYSTEM_CONFIG = {
  SHEET_NAMES: {
    MEMBERS: 'Members',
    LEARNING_ACTIVITIES: 'Learning_Activities', 
    ABILITY_ASSESSMENTS: 'Ability_Assessments',
    PROJECTS: 'Projects',
    ACADEMIC_IDEAS: 'Academic_Ideas',
    EVENTS: 'Events',
    SETTINGS: 'Settings'
  },
  ABILITY_CATEGORIES: [
    'Academic Rigor',
    'Mathematics & Statistics',
    'Computer Science',
    'Natural Sciences',
    'Social Sciences',
    'Engineering',
    'Software Development',
    'Research Methodology',
    'Critical Thinking',
    'Communication Skills'
  ],
  LEARNING_TYPES: [
    'YouTube Video',
    'Online Course',
    'Book',
    'Academic Paper',
    'Conference',
    'Workshop',
    'Seminar',
    'Peer Discussion'
  ],
  // Cyberpunk Theme Configuration
  CYBERPUNK_THEME: {
    COLORS: {
      DARK_BG: '#0a0a0a',           // Deep black background
      MATRIX_GREEN: '#00ff41',      // Classic matrix green
      NEON_CYAN: '#00ffff',         // Bright cyan
      NEON_PINK: '#ff0080',         // Hot pink
      NEON_PURPLE: '#8000ff',       // Electric purple
      NEON_BLUE: '#0080ff',         // Electric blue
      ORANGE_GLOW: '#ff8000',       // Orange highlight
      DARK_GRAY: '#1a1a1a',         // Dark gray for alternating rows
      TEXT_WHITE: '#ffffff',        // Pure white text
      TEXT_GRAY: '#cccccc',         // Light gray text
      BORDER_GLOW: '#00ff41'        // Glowing border color
    },
    FONTS: {
      HEADER: 'Courier New',        // Monospace for cyberpunk feel
      BODY: 'Arial'                 // Clean sans-serif for readability
    }
  }
};

/**
 * Add sample learning activities for testing
 */
function addSampleLearningActivities() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const membersSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.MEMBERS);
    
    // Get the first member for testing
    let memberId = 'TEST001';
    if (membersSheet && membersSheet.getLastRow() > 1) {
      const memberData = membersSheet.getRange(2, 1, 1, 1).getValue();
      memberId = memberData;
    }
    
    const sampleActivities = [
      {
        memberId: memberId,
        type: 'Online Course',
        title: 'Machine Learning Fundamentals',
        description: 'Comprehensive course on ML basics and algorithms',
        source: 'https://coursera.org/ml-course',
        duration: 40,
        progress: 85,
        rating: 4.5,
        notes: 'Excellent course, highly recommended for beginners',
        skillsGained: 'Python, Scikit-learn, Neural Networks'
      },
      {
        memberId: memberId,
        type: 'YouTube Video',
        title: 'Deep Learning with TensorFlow',
        description: 'Tutorial on building neural networks',
        source: 'https://youtube.com/watch?v=tensorflow-tutorial',
        duration: 2.5,
        progress: 100,
        rating: 4.0,
        notes: 'Great practical examples',
        skillsGained: 'TensorFlow, Neural Networks'
      },
      {
        memberId: memberId,
        type: 'Book',
        title: 'The Art of Computer Programming',
        description: 'Classic text on algorithms and data structures',
        source: 'Donald Knuth',
        duration: 60,
        progress: 30,
        rating: 5.0,
        notes: 'Challenging but rewarding read',
        skillsGained: 'Algorithms, Data Structures, Mathematical Thinking'
      },
      {
        memberId: memberId,
        type: 'Academic Paper',
        title: 'Attention Is All You Need',
        description: 'Transformer architecture paper',
        source: 'arXiv:1706.03762',
        duration: 8,
        progress: 100,
        rating: 4.8,
        notes: 'Revolutionary paper in NLP',
        skillsGained: 'Transformer Architecture, Attention Mechanisms'
      },
      {
        memberId: memberId,
        type: 'Workshop',
        title: 'Research Methodology Workshop',
        description: 'Hands-on workshop on research design',
        source: 'University Research Center',
        duration: 6,
        progress: 75,
        rating: 4.2,
        notes: 'Very practical and interactive',
        skillsGained: 'Research Design, Statistical Analysis'
      }
    ];
    
    let successCount = 0;
    sampleActivities.forEach(activity => {
      const result = addLearningActivity(activity);
      if (result.success) {
        successCount++;
        console.log('Added sample activity:', activity.title);
      } else {
        console.error('Failed to add activity:', result.message);
      }
    });
    
    console.log(`Successfully added ${successCount} sample learning activities`);
    return `Added ${successCount} sample learning activities for testing`;
    
  } catch (error) {
    console.error('Error adding sample activities:', error);
    return 'Error: ' + error.message;
  }
}

/**
 * Initialize the XU Academic Research Management System
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // Create custom navigation menu
  const menu = ui.createMenu('🎓 XU Academic System')
    .addItem('📊 Dashboard', 'showDashboard')
    .addSeparator()
    .addSubMenu(ui.createMenu('👥 Members')
      .addItem('Add New Member', 'showAddMemberDialog')
      .addItem('View All Members', 'showMembersView')
      .addItem('Member Profile', 'showMemberProfile'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📚 Learning Tracking')
      .addItem('Log Learning Activity', 'showLearningActivityDialog')
      .addItem('View Learning Progress', 'showLearningProgress')
      .addItem('Learning Analytics', 'showLearningAnalytics'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🧠 Abilities & Skills')
      .addItem('Assess Abilities', 'showAbilityAssessmentDialog')
      .addItem('View Skill Matrix', 'showSkillMatrix')
      .addItem('Competency Reports', 'showCompetencyReports'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🚀 Projects & Research')
      .addItem('Create Project', 'showCreateProjectDialog')
      .addItem('Project Hiring', 'showProjectHiring')
      .addItem('Submit Academic Idea', 'showAcademicIdeaDialog')
      .addItem('Idea Repository', 'showIdeaRepository'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📅 Events & Activities')
      .addItem('Schedule Event', 'showEventDialog')
      .addItem('Event Calendar', 'showEventCalendar')
      .addItem('RSVP Management', 'showRSVPManagement'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🧪 Testing')
      .addItem('Add Sample Learning Activities', 'addSampleLearningActivities'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🌟 Cyberpunk Theme')
      .addItem('🚀 Apply Cyberpunk Theme', 'applyCyberpunkThemeToAllSheets')
      .addItem('🔄 Reset to Default', 'resetSheetsToDefault')
      .addItem('⚡ Re-style Current Sheet', 'restyle_current_sheet'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🔧 Utilities')
      .addItem('Update Google Email Column', 'updateCurrentUserGoogleEmail')
      .addItem('Test Google Email Integration', 'testGoogleEmailIntegration')
      .addItem('Test Project Hiring Data', 'testProjectHiringData')
      .addItem('Add All ENUM Dropdowns', 'addAllEnumDropdowns')
      .addItem('Add Status Dropdowns', 'addStatusDropdowns')
      .addItem('Inspect Sheet Structures', 'inspectSheetStructures')
      .addItem('Create Sample Project Data', 'createSampleProjectData')
              .addItem('Test Ideas Repository Data', 'testIdeasRepositoryData')
        .addItem('Test Calendar Data', 'testCalendarData')
        .addItem('Test RSVP Data', 'testRSVPData'))
    .addSeparator()
    .addItem('⚙️ System Settings', 'showSystemSettings')
    .addItem('📈 Analytics Hub', 'showAnalyticsHub');
    
  menu.addToUi();
  
  // Initialize system on first run
  initializeSystem();
}

/**
 * Initialize system sheets and data structures
 */
function initializeSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create all necessary sheets
  Object.values(SYSTEM_CONFIG.SHEET_NAMES).forEach(sheetName => {
    if (!ss.getSheetByName(sheetName)) {
      const sheet = ss.insertSheet(sheetName);
      setupSheetHeaders(sheet, sheetName);
    }
  });
  
  console.log('XU Academic Research System initialized successfully');
}

/**
 * Setup headers for each sheet type
 */
function setupSheetHeaders(sheet, sheetName) {
  const headers = getSheetHeaders(sheetName);
  if (headers.length > 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    
    // Apply cyberpunk theme to the new sheet
    applyCyberpunkTheme(sheet, sheetName);
  }
}

/**
 * Get appropriate headers for each sheet type
 */
function getSheetHeaders(sheetName) {
  const headerMap = {
    [SYSTEM_CONFIG.SHEET_NAMES.MEMBERS]: [
      'Member ID', 'Name', 'Email', 'Role', 'Department', 'Join Date', 
      'Status', 'Research Interests', 'Contact Info', 'Last Active', 'Google Email'
    ],
    [SYSTEM_CONFIG.SHEET_NAMES.LEARNING_ACTIVITIES]: [
      'Activity ID', 'Member ID', 'Type', 'Title', 'Description', 
      'URL/Source', 'Duration (hrs)', 'Progress %', 'Start Date', 
      'Completion Date', 'Rating', 'Notes', 'Skills Gained'
    ],
    [SYSTEM_CONFIG.SHEET_NAMES.ABILITY_ASSESSMENTS]: [
      'Assessment ID', 'Member ID', 'Category', 'Skill', 'Level (1-10)', 
      'Assessment Date', 'Assessor', 'Evidence', 'Next Review Date', 'Goals'
    ],
    [SYSTEM_CONFIG.SHEET_NAMES.PROJECTS]: [
      'Project ID', 'Title', 'Description', 'Status', 'Leader ID', 
      'Team Members', 'Skills Required', 'Start Date', 'End Date', 
      'Budget', 'Progress %', 'Repository URL'
    ],
    [SYSTEM_CONFIG.SHEET_NAMES.ACADEMIC_IDEAS]: [
      'Idea ID', 'Submitter ID', 'Title', 'Abstract', 'Category', 
      'Keywords', 'Potential Impact', 'Resources Needed', 'Submit Date', 
      'Status', 'Reviewer Comments', 'Vote Score'
    ],
    [SYSTEM_CONFIG.SHEET_NAMES.EVENTS]: [
      'Event ID', 'Title', 'Type', 'Description', 'Date', 'Time', 
      'Location', 'Organizer ID', 'Attendees', 'Max Capacity', 'Status'
    ],
    [SYSTEM_CONFIG.SHEET_NAMES.SETTINGS]: [
      'Setting Key', 'Setting Value', 'Description', 'Modified Date', 'Modified By'
    ]
  };
  
  return headerMap[sheetName] || [];
}

/**
 * Show main dashboard with system overview
 */
function showDashboard() {
  const htmlOutput = HtmlService.createTemplateFromFile('dashboard');
  htmlOutput.data = getDashboardData();
  
  const html = htmlOutput.evaluate()
    .setWidth(1200)
    .setHeight(800)
    .setTitle('🎓 XU Academic System Dashboard');
    
  SpreadsheetApp.getUi().showModalDialog(html, '🎓 XU Academic System Dashboard');
}

/**
 * Get dashboard data for display
 */
function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get member count
  const membersSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.MEMBERS);
  const memberCount = membersSheet ? Math.max(0, membersSheet.getLastRow() - 1) : 0;
  
  // Get active learning activities
  const learningSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.LEARNING_ACTIVITIES);
  const activeLearning = learningSheet ? Math.max(0, learningSheet.getLastRow() - 1) : 0;
  
  // Get project count
  const projectsSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.PROJECTS);
  const projectCount = projectsSheet ? Math.max(0, projectsSheet.getLastRow() - 1) : 0;
  
  // Get recent ideas
  const ideasSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.ACADEMIC_IDEAS);
  const ideaCount = ideasSheet ? Math.max(0, ideasSheet.getLastRow() - 1) : 0;
  
  return {
    memberCount,
    activeLearning,
    projectCount,
    ideaCount,
    timestamp: new Date().toLocaleString()
  };
}

/**
 * Show add member dialog
 */
function showAddMemberDialog() {
  // Get current user's Google email for pre-population
  const currentUserEmail = Session.getActiveUser().getEmail();
  
  const html = HtmlService.createTemplateFromFile('add-member-dialog');
  html.currentUserEmail = currentUserEmail;
  
  const htmlOutput = html.evaluate()
    .setWidth(600)
    .setHeight(500)
    .setTitle('Add New Member');
    
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Add New Member');
}

/**
 * Add a new member to the system
 */
function addMember(memberData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.MEMBERS);
    
    if (!sheet) {
      throw new Error('Members sheet not found');
    }
    
    const memberId = 'MEM' + Date.now();
    const timestamp = new Date();
    
    // Get current user's Google email for automatic identification
    const currentUserEmail = Session.getActiveUser().getEmail();
    
    const rowData = [
      memberId,
      memberData.name,
      memberData.email, // Institutional email
      memberData.role,
      memberData.department,
      timestamp,
      'Active',
      memberData.researchInterests,
      memberData.contactInfo,
      timestamp,
      memberData.googleEmail || currentUserEmail // Google email for user identification
    ];
    
    sheet.appendRow(rowData);
    
    return {
      success: true,
      message: 'Member added successfully',
      memberId: memberId
    };
  } catch (error) {
    return {
      success: false,
      message: 'Error adding member: ' + error.message
    };
  }
}

/**
 * Show learning activity dialog
 */
function showLearningActivityDialog() {
  const html = HtmlService.createTemplateFromFile('learning-activity-dialog')
    .evaluate()
    .setWidth(700)
    .setHeight(600)
    .setTitle('Log Learning Activity');
    
  SpreadsheetApp.getUi().showModalDialog(html, 'Log Learning Activity');
}

/**
 * Add learning activity
 */
function addLearningActivity(activityData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.LEARNING_ACTIVITIES);
    
    if (!sheet) {
      throw new Error('Learning Activities sheet not found');
    }
    
    const activityId = 'ACT' + Date.now();
    const timestamp = new Date();
    
    const rowData = [
      activityId,
      activityData.memberId,
      activityData.type,
      activityData.title,
      activityData.description,
      activityData.source,
      activityData.duration,
      activityData.progress,
      timestamp,
      activityData.completionDate || '',
      activityData.rating || '',
      activityData.notes || '',
      activityData.skillsGained || ''
    ];
    
    sheet.appendRow(rowData);
    
    return {
      success: true,
      message: 'Learning activity logged successfully',
      activityId: activityId
    };
  } catch (error) {
    return {
      success: false,
      message: 'Error logging activity: ' + error.message
    };
  }
}

/**
 * Show ability assessment dialog
 */
function showAbilityAssessmentDialog() {
  const html = HtmlService.createTemplateFromFile('ability-assessment-dialog')
    .evaluate()
    .setWidth(800)
    .setHeight(700)
    .setTitle('Ability Assessment');
    
  SpreadsheetApp.getUi().showModalDialog(html, 'Ability Assessment');
}

/**
 * Submit ability assessment
 */
function submitAbilityAssessment(assessmentData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.ABILITY_ASSESSMENTS);
    
    if (!sheet) {
      throw new Error('Ability Assessments sheet not found');
    }
    
    const assessmentId = 'ASS' + Date.now();
    const timestamp = new Date();
    const nextReview = new Date(timestamp.getTime() + (90 * 24 * 60 * 60 * 1000)); // 90 days later
    
    const rowData = [
      assessmentId,
      assessmentData.memberId,
      assessmentData.category,
      assessmentData.skill,
      assessmentData.level,
      timestamp,
      assessmentData.assessor,
      assessmentData.evidence,
      nextReview,
      assessmentData.goals
    ];
    
    sheet.appendRow(rowData);
    
    return {
      success: true,
      message: 'Ability assessment submitted successfully',
      assessmentId: assessmentId
    };
  } catch (error) {
    return {
      success: false,
      message: 'Error submitting assessment: ' + error.message
    };
  }
}

/**
 * Show create project dialog
 */
function showCreateProjectDialog() {
  const html = HtmlService.createTemplateFromFile('create-project-dialog')
    .evaluate()
    .setWidth(800)
    .setHeight(700)
    .setTitle('Create New Project');
    
  SpreadsheetApp.getUi().showModalDialog(html, 'Create New Project');
}

/**
 * Create new project
 */
function createProject(projectData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.PROJECTS);
    
    if (!sheet) {
      throw new Error('Projects sheet not found');
    }
    
    const projectId = 'PRJ' + Date.now();
    const timestamp = new Date();
    
    const rowData = [
      projectId,
      projectData.title,
      projectData.description,
      'Planning',
      projectData.leaderId,
      projectData.teamMembers,
      projectData.skillsRequired,
      timestamp,
      projectData.endDate,
      projectData.budget,
      0,
      projectData.repositoryUrl || ''
    ];
    
    sheet.appendRow(rowData);
    
    return {
      success: true,
      message: 'Project created successfully',
      projectId: projectId
    };
  } catch (error) {
    return {
      success: false,
      message: 'Error creating project: ' + error.message
    };
  }
}

/**
 * Show academic idea submission dialog
 */
function showAcademicIdeaDialog() {
  const html = HtmlService.createTemplateFromFile('academic-idea-dialog')
    .evaluate()
    .setWidth(800)
    .setHeight(700)
    .setTitle('Submit Academic Idea');
    
  SpreadsheetApp.getUi().showModalDialog(html, 'Submit Academic Idea');
}

/**
 * Submit academic idea
 */
function submitAcademicIdea(ideaData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.ACADEMIC_IDEAS);
    
    if (!sheet) {
      throw new Error('Academic Ideas sheet not found');
    }
    
    const ideaId = 'IDEA' + Date.now();
    const timestamp = new Date();
    
    const rowData = [
      ideaId,
      ideaData.submitterId,
      ideaData.title,
      ideaData.abstract,
      ideaData.category,
      ideaData.keywords,
      ideaData.potentialImpact,
      ideaData.resourcesNeeded,
      timestamp,
      'Under Review',
      '',
      0
    ];
    
    sheet.appendRow(rowData);
    
    return {
      success: true,
      message: 'Academic idea submitted successfully',
      ideaId: ideaId
    };
  } catch (error) {
    return {
      success: false,
      message: 'Error submitting idea: ' + error.message
    };
  }
}

/**
 * Get all members for dropdown selections
 */
function getAllMembers() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.MEMBERS);
    
    console.log('getAllMembers: Looking for sheet:', SYSTEM_CONFIG.SHEET_NAMES.MEMBERS);
    
    if (!sheet) {
      console.log('getAllMembers: Sheet not found. Available sheets:', ss.getSheets().map(s => s.getName()));
      return [];
    }
    
    const lastRow = sheet.getLastRow();
    console.log('getAllMembers: Last row:', lastRow);
    
    if (lastRow <= 1) {
      console.log('getAllMembers: No data rows found');
      return [];
    }
    
    // Get more columns to ensure we have all needed data (including Google email)
    const data = sheet.getRange(2, 1, lastRow - 1, Math.min(11, sheet.getLastColumn())).getValues();
    console.log('getAllMembers: Retrieved', data.length, 'rows');
    
    const members = data
      .filter((row, index) => {
        // Skip completely empty rows
        if (!row[0] && !row[1] && !row[2]) {
          console.log('getAllMembers: Skipping empty row', index + 2);
          return false;
        }
        return true;
      })
      .map((row, index) => {
        const member = {
          id: String(row[0] || `member_${index + 1}`),
          name: String(row[1] || 'Unknown Member'),
          email: String(row[2] || 'no-email@example.com'), // Institutional email
          role: String(row[3] || 'Unknown'),
          department: String(row[4] || ''),
          status: String(row[6] || 'Active'),
          googleEmail: String(row[10] || '') // Google email for user identification
        };
        console.log('getAllMembers: Processed member:', member.name, '(' + member.id + ') - Google email:', member.googleEmail);
        return member;
      });
    
    console.log('getAllMembers: Returning', members.length, 'members');
    return members;
  } catch (error) {
    console.error('getAllMembers: Error occurred:', error.toString());
    console.error('getAllMembers: Full error details:', error);
    
    // Return empty array instead of null/undefined
    console.log('getAllMembers: Returning empty array due to error');
    return [];
  }
}

/**
 * Get system configuration for client-side use
 */
function getSystemConfig() {
  return SYSTEM_CONFIG;
}

/**
 * Test Google email integration
 */
function testGoogleEmailIntegration() {
  console.log('=== TESTING GOOGLE EMAIL INTEGRATION ===');
  
  // Test 1: Get current user email
  const currentUserEmail = Session.getActiveUser().getEmail();
  console.log('Current user email:', currentUserEmail);
  
  // Test 2: Check if Members sheet has Google email column
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const membersSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.MEMBERS);
  
  if (membersSheet) {
    console.log('Members sheet found');
    console.log('- Last row:', membersSheet.getLastRow());
    console.log('- Last column:', membersSheet.getLastColumn());
    
    if (membersSheet.getLastColumn() >= 11) {
      console.log('✅ Google email column exists (column 11)');
      
      // Test 3: Check if current user exists in members
      const allMembers = getAllMembers();
      const currentMember = allMembers.find(member => 
        (member.email && member.email.toLowerCase() === currentUserEmail.toLowerCase()) ||
        (member.googleEmail && member.googleEmail.toLowerCase() === currentUserEmail.toLowerCase())
      );
      
      if (currentMember) {
        console.log('✅ Current user found in members:', currentMember.name);
        console.log('- Institutional email:', currentMember.email);
        console.log('- Google email:', currentMember.googleEmail);
      } else {
        console.log('⚠️ Current user not found in members list');
        console.log('Available members:');
        allMembers.forEach(member => {
          console.log(`  - ${member.name}: ${member.email} / ${member.googleEmail}`);
        });
      }
    } else {
      console.log('❌ Google email column missing (only', membersSheet.getLastColumn(), 'columns)');
    }
  } else {
    console.log('❌ Members sheet not found');
  }
  
  console.log('=== END TEST ===');
}

/**
 * Debug function to check sheet status and data
 */
function debugSheetStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  
  console.log('=== SHEET STATUS DEBUG ===');
  console.log('Spreadsheet name:', ss.getName());
  console.log('Total sheets:', allSheets.length);
  console.log('Available sheets:', allSheets.map(s => s.getName()));
  console.log('Looking for Members sheet:', SYSTEM_CONFIG.SHEET_NAMES.MEMBERS);
  
  const membersSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.MEMBERS);
  if (membersSheet) {
    console.log('Members sheet found!');
    console.log('- Last row:', membersSheet.getLastRow());
    console.log('- Last column:', membersSheet.getLastColumn());
    
    if (membersSheet.getLastRow() > 1) {
      const sampleData = membersSheet.getRange(2, 1, Math.min(3, membersSheet.getLastRow() - 1), membersSheet.getLastColumn()).getValues();
      console.log('- Sample data (first 3 rows):');
      sampleData.forEach((row, index) => {
        console.log(`  Row ${index + 2}:`, row);
      });
      
      // Test the actual function call
      console.log('- Testing getMembersData() function:');
      const result = getMembersData();
      console.log('  - Result:', result);
    } else {
      console.log('- No data rows found');
    }
  } else {
    console.log('Members sheet NOT found!');
  }
  
  console.log('=== END DEBUG ===');
  
  return {
    sheetsFound: allSheets.map(s => s.getName()),
    membersSheetExists: !!membersSheet,
    membersDataRows: membersSheet ? membersSheet.getLastRow() - 1 : 0
  };
}

/**
 * Quick test function to get members data and return it
 */
function testGetMembersData() {
  console.log('=== TESTING getMembersData() ===');
  const result = getMembersData();
  console.log('Function returned:', JSON.stringify(result, null, 2));
  return result;
}

/**
 * Update existing members with Google email for current user
 */
function updateCurrentUserGoogleEmail() {
  try {
    const currentUserEmail = Session.getActiveUser().getEmail();
    console.log('updateCurrentUserGoogleEmail: Current user email:', currentUserEmail);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const membersSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.MEMBERS);
    
    if (!membersSheet) {
      console.log('updateCurrentUserGoogleEmail: Members sheet not found');
      return { success: false, message: 'Members sheet not found' };
    }
    
    const lastRow = membersSheet.getLastColumn();
    console.log('updateCurrentUserGoogleEmail: Last column:', lastRow);
    
    // If Google email column doesn't exist, add it
    if (lastRow < 11) {
      console.log('updateCurrentUserGoogleEmail: Adding Google email column');
      membersSheet.getRange(1, 11).setValue('Google Email');
    }
    
    // Find current user in the sheet
    const data = membersSheet.getRange(2, 1, membersSheet.getLastRow() - 1, 10).getValues();
    let userRowIndex = -1;
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const institutionalEmail = String(row[2] || '');
      if (institutionalEmail.toLowerCase() === currentUserEmail.toLowerCase()) {
        userRowIndex = i + 2; // +2 because we start from row 2 and i is 0-based
        break;
      }
    }
    
    if (userRowIndex === -1) {
      console.log('updateCurrentUserGoogleEmail: Current user not found in members list');
      return { success: false, message: 'Current user not found in members list' };
    }
    
    // Update the Google email column for the current user
    membersSheet.getRange(userRowIndex, 11).setValue(currentUserEmail);
    console.log('updateCurrentUserGoogleEmail: Updated Google email for user at row', userRowIndex);
    
    return { success: true, message: 'Google email updated successfully' };
    
  } catch (error) {
    console.error('updateCurrentUserGoogleEmail: Error:', error);
    return { success: false, message: 'Error updating Google email: ' + error.message };
  }
}

/**
 * Test getMemberProfileData function specifically
 */
function testGetMemberProfileData() {
  console.log('=== TESTING getMemberProfileData() ===');
  
  const testId = 'MEM1754685273988'; // Tian Shao's ID
  console.log('Testing with member ID:', testId);
  
  try {
    const result = getMemberProfileData(testId);
    console.log('getMemberProfileData result:', result);
    console.log('getMemberProfileData result type:', typeof result);
    return result;
  } catch (error) {
    console.error('getMemberProfileData failed:', error);
    return { error: error.toString() };
  }
}

/**
 * Test getAllMembers function specifically
 */
function testGetAllMembers() {
  console.log('=== TESTING getAllMembers() ===');
  
  try {
    const result = getAllMembers();
    console.log('getAllMembers result:', result);
    console.log('getAllMembers result type:', typeof result);
    console.log('getAllMembers result length:', result ? result.length : 'null/undefined');
    return result;
  } catch (error) {
    console.error('getAllMembers failed:', error);
    return { error: error.toString() };
  }
}

/**
 * Simple function to test member data access
 */
function testMemberAccess() {
  console.log('=== SIMPLE MEMBER ACCESS TEST ===');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    console.log('Spreadsheet name:', ss.getName());
    
    // List all sheets
    const sheets = ss.getSheets();
    console.log('Available sheets:', sheets.map(s => s.getName()));
    
    // Test getMembersData function
    console.log('Testing getMembersData()...');
    const result = getMembersData();
    console.log('Result:', result);
    
    return result;
  } catch (error) {
    console.error('Test failed:', error);
    return { error: error.toString() };
  }
}

/**
 * Find member data in any sheet
 */
function findMemberData() {
  console.log('=== SEARCHING FOR MEMBER DATA ===');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    
    console.log('Checking', sheets.length, 'sheets for member data...');
    
    for (let i = 0; i < sheets.length; i++) {
      const sheet = sheets[i];
      const sheetName = sheet.getName();
      console.log(`Checking sheet: ${sheetName}`);
      
      if (sheet.getLastRow() > 1) {
        // Get first row to check headers
        const headerRow = sheet.getRange(1, 1, 1, Math.min(10, sheet.getLastColumn())).getValues()[0];
        console.log(`${sheetName} headers:`, headerRow);
        
        // Check if this looks like member data
        const hasNameColumn = headerRow.some(header => 
          header && String(header).toLowerCase().includes('name')
        );
        const hasEmailColumn = headerRow.some(header => 
          header && String(header).toLowerCase().includes('email')
        );
        const hasMemberIdColumn = headerRow.some(header => 
          header && (String(header).toLowerCase().includes('member') || String(header).toLowerCase().includes('id'))
        );
        
        if (hasNameColumn && hasEmailColumn) {
          console.log(`✅ Found potential member data in sheet: ${sheetName}`);
          
          // Get sample data
          const sampleData = sheet.getRange(2, 1, Math.min(3, sheet.getLastRow() - 1), Math.min(10, sheet.getLastColumn())).getValues();
          console.log(`Sample data from ${sheetName}:`, sampleData);
          
          return {
            foundSheet: sheetName,
            headers: headerRow,
            sampleData: sampleData,
            totalRows: sheet.getLastRow() - 1
          };
        }
      }
    }
    
    console.log('❌ No member data found in any sheet');
    return { foundSheet: null, message: 'No member data found' };
    
  } catch (error) {
    console.error('Search failed:', error);
    return { error: error.toString() };
  }
}

/**
 * Simple debug function to check what's in the Members sheet
 */
function debugMembersSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    console.log('=== MEMBERS SHEET DEBUG ===');
    console.log('Spreadsheet name:', ss.getName());
    
    // List all sheets
    const allSheets = ss.getSheets();
    console.log('All sheets:', allSheets.map(s => s.getName()));
    
    // Try to get the Members sheet
    const membersSheet = ss.getSheetByName('Members');
    if (membersSheet) {
      console.log('✅ Members sheet found');
      console.log('Last row:', membersSheet.getLastRow());
      console.log('Last column:', membersSheet.getLastColumn());
      
      if (membersSheet.getLastRow() > 1) {
        // Get first few rows of data
        const data = membersSheet.getRange(1, 1, Math.min(5, membersSheet.getLastRow()), membersSheet.getLastColumn()).getValues();
        console.log('Raw sheet data:');
        data.forEach((row, index) => {
          console.log(`Row ${index + 1}:`, row);
        });
        
        // Test the actual function
        console.log('--- Testing getMembersData() ---');
        const result = getMembersData();
        console.log('getMembersData() result:', result);
        
        return {
          success: true,
          sheetExists: true,
          lastRow: membersSheet.getLastRow(),
          lastColumn: membersSheet.getLastColumn(),
          sampleData: data,
          functionResult: result
        };
      } else {
        console.log('❌ No data in Members sheet');
        return {
          success: false,
          sheetExists: true,
          lastRow: membersSheet.getLastRow(),
          message: 'Sheet exists but has no data'
        };
      }
    } else {
      console.log('❌ Members sheet not found');
      // Try other possible names
      const possibleNames = ['Members', 'Member', 'members', 'Team', 'People'];
      const foundSheets = [];
      possibleNames.forEach(name => {
        const sheet = ss.getSheetByName(name);
        if (sheet) {
          foundSheets.push({
            name: name,
            lastRow: sheet.getLastRow(),
            lastColumn: sheet.getLastColumn()
          });
        }
      });
      
      return {
        success: false,
        sheetExists: false,
        allSheets: allSheets.map(s => s.getName()),
        possibleMemberSheets: foundSheets,
        message: 'Members sheet not found'
      };
    }
  } catch (error) {
    console.error('debugMembersSheet error:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Create sample data for testing if sheets are empty
 */
function createSampleData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const membersSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.MEMBERS);
    
    if (!membersSheet) {
      console.log('createSampleData: Members sheet not found, creating...');
      setupSystemSheets();
      return;
    }
    
    if (membersSheet.getLastRow() > 1) {
      console.log('createSampleData: Members sheet already has data');
      return;
    }
    
    console.log('createSampleData: Adding sample member data...');
    
    const sampleMembers = [
      ['MEM001', 'Dr. Sarah Chen', 'sarah.chen@university.edu', 'Professor', 'Computer Science', new Date('2020-01-15'), 'Active', 'Machine Learning, AI Ethics', '+1-555-0101', new Date()],
      ['MEM002', 'Alex Rodriguez', 'alex.rodriguez@university.edu', 'PhD Student', 'Computer Science', new Date('2021-09-01'), 'Active', 'Data Science, Statistics', '+1-555-0102', new Date()],
      ['MEM003', 'Dr. Michael Johnson', 'michael.johnson@university.edu', 'Faculty', 'Mathematics', new Date('2018-08-20'), 'Active', 'Statistical Analysis, Research Methods', '+1-555-0103', new Date()],
      ['MEM004', 'Emma Wilson', 'emma.wilson@university.edu', 'Undergraduate', 'Computer Science', new Date('2022-01-10'), 'Active', 'Web Development, UI/UX', '+1-555-0104', new Date()],
      ['MEM005', 'Dr. James Park', 'james.park@university.edu', 'Professor', 'Engineering', new Date('2019-03-05'), 'Active', 'Systems Engineering, Project Management', '+1-555-0105', new Date()]
    ];
    
    sampleMembers.forEach((member, index) => {
      membersSheet.appendRow(member);
      console.log('createSampleData: Added member:', member[1]);
    });
    
    // Apply cyberpunk theme to the new data
    applyCyberpunkTheme(membersSheet, SYSTEM_CONFIG.SHEET_NAMES.MEMBERS);
    
    console.log('createSampleData: Sample data created successfully!');
    return { success: true, message: 'Sample data created successfully!' };
  } catch (error) {
    console.error('createSampleData: Error occurred:', error.toString());
    return { success: false, message: 'Error creating sample data: ' + error.message };
  }
}

/**
 * Include HTML files for dialogs
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * =====================================
 * CYBERPUNK THEME STYLING FUNCTIONS
 * =====================================
 */

/**
 * Apply cyberpunk theme to a specific sheet
 */
function applyCyberpunkTheme(sheet, sheetName) {
  const theme = SYSTEM_CONFIG.CYBERPUNK_THEME;
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  
  try {
    // Set sheet background to dark
    const fullRange = sheet.getRange(1, 1, Math.max(lastRow, 100), lastCol);
    fullRange.setBackground(theme.COLORS.DARK_BG);
    fullRange.setFontColor(theme.COLORS.TEXT_WHITE);
    fullRange.setFontFamily(theme.FONTS.BODY);
    
    // Style header row with matrix green glow effect
    if (lastRow >= 1) {
      const headerRange = sheet.getRange(1, 1, 1, lastCol);
      headerRange.setBackground(theme.COLORS.MATRIX_GREEN);
      headerRange.setFontColor(theme.COLORS.DARK_BG);
      headerRange.setFontFamily(theme.FONTS.HEADER);
      headerRange.setFontWeight('bold');
      headerRange.setFontSize(12);
      
      // Add cyberpunk border effects
      headerRange.setBorder(true, true, true, true, false, false, 
                           theme.COLORS.NEON_CYAN, SpreadsheetApp.BorderStyle.SOLID_THICK);
    }
    
    // Apply alternating row colors for cyberpunk effect
    for (let row = 2; row <= Math.max(lastRow, 50); row++) {
      const rowRange = sheet.getRange(row, 1, 1, lastCol);
      if (row % 2 === 0) {
        rowRange.setBackground(theme.COLORS.DARK_GRAY);
      } else {
        rowRange.setBackground(theme.COLORS.DARK_BG);
      }
      
      // Add subtle neon borders every 5 rows for grid effect
      if (row % 5 === 0) {
        rowRange.setBorder(false, false, true, false, false, false,
                          theme.COLORS.NEON_BLUE, SpreadsheetApp.BorderStyle.DOTTED);
      }
    }
    
    // Apply sheet-specific cyberpunk styling
    applyCyberpunkSheetSpecificStyling(sheet, sheetName);
    
    console.log(`Cyberpunk theme applied to ${sheetName} sheet`);
    
  } catch (error) {
    console.error(`Error applying cyberpunk theme to ${sheetName}:`, error);
  }
}

/**
 * Apply sheet-specific cyberpunk styling based on data type
 */
function applyCyberpunkSheetSpecificStyling(sheet, sheetName) {
  const theme = SYSTEM_CONFIG.CYBERPUNK_THEME;
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow <= 1) return; // No data to style
  
  try {
    switch (sheetName) {
      case SYSTEM_CONFIG.SHEET_NAMES.MEMBERS:
        // Highlight member roles with different neon colors
        highlightMemberRoles(sheet, theme);
        break;
        
      case SYSTEM_CONFIG.SHEET_NAMES.PROJECTS:
        // Color-code project status
        highlightProjectStatus(sheet, theme);
        break;
        
      case SYSTEM_CONFIG.SHEET_NAMES.ABILITY_ASSESSMENTS:
        // Create skill level heat map
        createSkillLevelHeatMap(sheet, theme);
        break;
        
      case SYSTEM_CONFIG.SHEET_NAMES.LEARNING_ACTIVITIES:
        // Highlight learning progress
        highlightLearningProgress(sheet, theme);
        break;
        
      case SYSTEM_CONFIG.SHEET_NAMES.ACADEMIC_IDEAS:
        // Color-code idea status and votes
        highlightAcademicIdeas(sheet, theme);
        break;
        
      case SYSTEM_CONFIG.SHEET_NAMES.EVENTS:
        // Highlight upcoming events
        highlightEvents(sheet, theme);
        break;
    }
  } catch (error) {
    console.error(`Error applying specific styling to ${sheetName}:`, error);
  }
}

/**
 * Highlight member roles with cyberpunk colors
 */
function highlightMemberRoles(sheet, theme) {
  const roleColumnIndex = 4; // Assuming role is in column D
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) return;
  
  const roleRange = sheet.getRange(2, roleColumnIndex, lastRow - 1, 1);
  const values = roleRange.getValues();
  
  values.forEach((row, index) => {
    const role = row[0];
    const cellRange = sheet.getRange(index + 2, roleColumnIndex);
    
    switch (role) {
      case 'Professor':
      case 'Faculty':
        cellRange.setBackground(theme.COLORS.NEON_PURPLE);
        cellRange.setFontColor(theme.COLORS.DARK_BG);
        break;
      case 'PhD Student':
        cellRange.setBackground(theme.COLORS.NEON_BLUE);
        cellRange.setFontColor(theme.COLORS.TEXT_WHITE);
        break;
      case 'Research Assistant':
        cellRange.setBackground(theme.COLORS.NEON_CYAN);
        cellRange.setFontColor(theme.COLORS.DARK_BG);
        break;
      case 'Undergraduate':
        cellRange.setBackground(theme.COLORS.MATRIX_GREEN);
        cellRange.setFontColor(theme.COLORS.DARK_BG);
        break;
      default:
        cellRange.setBackground(theme.COLORS.ORANGE_GLOW);
        cellRange.setFontColor(theme.COLORS.DARK_BG);
    }
    cellRange.setFontWeight('bold');
  });
}

/**
 * Highlight project status with cyberpunk colors
 */
function highlightProjectStatus(sheet, theme) {
  const statusColumnIndex = 4; // Assuming status is in column D
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) return;
  
  const statusRange = sheet.getRange(2, statusColumnIndex, lastRow - 1, 1);
  const values = statusRange.getValues();
  
  values.forEach((row, index) => {
    const status = row[0];
    const cellRange = sheet.getRange(index + 2, statusColumnIndex);
    
    switch (status) {
      case 'Active':
      case 'In Progress':
        cellRange.setBackground(theme.COLORS.MATRIX_GREEN);
        cellRange.setFontColor(theme.COLORS.DARK_BG);
        break;
      case 'Planning':
        cellRange.setBackground(theme.COLORS.NEON_BLUE);
        cellRange.setFontColor(theme.COLORS.TEXT_WHITE);
        break;
      case 'Completed':
        cellRange.setBackground(theme.COLORS.NEON_CYAN);
        cellRange.setFontColor(theme.COLORS.DARK_BG);
        break;
      case 'On Hold':
        cellRange.setBackground(theme.COLORS.ORANGE_GLOW);
        cellRange.setFontColor(theme.COLORS.DARK_BG);
        break;
      case 'Cancelled':
        cellRange.setBackground(theme.COLORS.NEON_PINK);
        cellRange.setFontColor(theme.COLORS.TEXT_WHITE);
        break;
    }
    cellRange.setFontWeight('bold');
  });
}

/**
 * Create cyberpunk skill level heat map
 */
function createSkillLevelHeatMap(sheet, theme) {
  const levelColumnIndex = 5; // Assuming level is in column E
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) return;
  
  const levelRange = sheet.getRange(2, levelColumnIndex, lastRow - 1, 1);
  const values = levelRange.getValues();
  
  values.forEach((row, index) => {
    const level = parseInt(row[0]);
    const cellRange = sheet.getRange(index + 2, levelColumnIndex);
    
    if (level >= 9) {
      cellRange.setBackground(theme.COLORS.NEON_PURPLE);
      cellRange.setFontColor(theme.COLORS.TEXT_WHITE);
    } else if (level >= 7) {
      cellRange.setBackground(theme.COLORS.NEON_CYAN);
      cellRange.setFontColor(theme.COLORS.DARK_BG);
    } else if (level >= 5) {
      cellRange.setBackground(theme.COLORS.MATRIX_GREEN);
      cellRange.setFontColor(theme.COLORS.DARK_BG);
    } else if (level >= 3) {
      cellRange.setBackground(theme.COLORS.ORANGE_GLOW);
      cellRange.setFontColor(theme.COLORS.DARK_BG);
    } else {
      cellRange.setBackground(theme.COLORS.NEON_PINK);
      cellRange.setFontColor(theme.COLORS.TEXT_WHITE);
    }
    cellRange.setFontWeight('bold');
  });
}

/**
 * Highlight learning progress with cyberpunk styling
 */
function highlightLearningProgress(sheet, theme) {
  const progressColumnIndex = 8; // Assuming progress is in column H
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) return;
  
  const progressRange = sheet.getRange(2, progressColumnIndex, lastRow - 1, 1);
  const values = progressRange.getValues();
  
  values.forEach((row, index) => {
    const progress = parseInt(row[0]);
    const cellRange = sheet.getRange(index + 2, progressColumnIndex);
    
    if (progress >= 100) {
      cellRange.setBackground(theme.COLORS.MATRIX_GREEN);
      cellRange.setFontColor(theme.COLORS.DARK_BG);
    } else if (progress >= 75) {
      cellRange.setBackground(theme.COLORS.NEON_CYAN);
      cellRange.setFontColor(theme.COLORS.DARK_BG);
    } else if (progress >= 50) {
      cellRange.setBackground(theme.COLORS.NEON_BLUE);
      cellRange.setFontColor(theme.COLORS.TEXT_WHITE);
    } else if (progress >= 25) {
      cellRange.setBackground(theme.COLORS.ORANGE_GLOW);
      cellRange.setFontColor(theme.COLORS.DARK_BG);
    } else {
      cellRange.setBackground(theme.COLORS.NEON_PINK);
      cellRange.setFontColor(theme.COLORS.TEXT_WHITE);
    }
    cellRange.setFontWeight('bold');
  });
}

/**
 * Highlight academic ideas with cyberpunk styling
 */
function highlightAcademicIdeas(sheet, theme) {
  const statusColumnIndex = 10; // Assuming status is in column J
  const voteColumnIndex = 12; // Assuming vote score is in column L
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) return;
  
  // Highlight status
  const statusRange = sheet.getRange(2, statusColumnIndex, lastRow - 1, 1);
  const statusValues = statusRange.getValues();
  
  statusValues.forEach((row, index) => {
    const status = row[0];
    const cellRange = sheet.getRange(index + 2, statusColumnIndex);
    
    switch (status) {
      case 'Approved':
        cellRange.setBackground(theme.COLORS.MATRIX_GREEN);
        cellRange.setFontColor(theme.COLORS.DARK_BG);
        break;
      case 'Under Review':
        cellRange.setBackground(theme.COLORS.NEON_BLUE);
        cellRange.setFontColor(theme.COLORS.TEXT_WHITE);
        break;
      case 'Rejected':
        cellRange.setBackground(theme.COLORS.NEON_PINK);
        cellRange.setFontColor(theme.COLORS.TEXT_WHITE);
        break;
      case 'Implemented':
        cellRange.setBackground(theme.COLORS.NEON_PURPLE);
        cellRange.setFontColor(theme.COLORS.TEXT_WHITE);
        break;
    }
    cellRange.setFontWeight('bold');
  });
  
  // Highlight vote scores
  if (sheet.getLastColumn() >= voteColumnIndex) {
    const voteRange = sheet.getRange(2, voteColumnIndex, lastRow - 1, 1);
    const voteValues = voteRange.getValues();
    
    voteValues.forEach((row, index) => {
      const score = parseInt(row[0]) || 0;
      const cellRange = sheet.getRange(index + 2, voteColumnIndex);
      
      if (score >= 10) {
        cellRange.setBackground(theme.COLORS.NEON_PURPLE);
        cellRange.setFontColor(theme.COLORS.TEXT_WHITE);
      } else if (score >= 5) {
        cellRange.setBackground(theme.COLORS.NEON_CYAN);
        cellRange.setFontColor(theme.COLORS.DARK_BG);
      } else if (score >= 1) {
        cellRange.setBackground(theme.COLORS.MATRIX_GREEN);
        cellRange.setFontColor(theme.COLORS.DARK_BG);
      } else {
        cellRange.setBackground(theme.COLORS.DARK_GRAY);
        cellRange.setFontColor(theme.COLORS.TEXT_GRAY);
      }
      cellRange.setFontWeight('bold');
    });
  }
}

/**
 * Highlight events with cyberpunk styling
 */
function highlightEvents(sheet, theme) {
  const dateColumnIndex = 5; // Assuming date is in column E
  const statusColumnIndex = 11; // Assuming status is in column K
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) return;
  
  const today = new Date();
  const dateRange = sheet.getRange(2, dateColumnIndex, lastRow - 1, 1);
  const dateValues = dateRange.getValues();
  
  // Highlight dates based on proximity
  dateValues.forEach((row, index) => {
    const eventDate = new Date(row[0]);
    const cellRange = sheet.getRange(index + 2, dateColumnIndex);
    const daysDiff = Math.ceil((eventDate - today) / (1000 * 60 * 60 * 24));
    
    if (daysDiff < 0) {
      // Past events
      cellRange.setBackground(theme.COLORS.DARK_GRAY);
      cellRange.setFontColor(theme.COLORS.TEXT_GRAY);
    } else if (daysDiff <= 7) {
      // This week
      cellRange.setBackground(theme.COLORS.NEON_PINK);
      cellRange.setFontColor(theme.COLORS.TEXT_WHITE);
    } else if (daysDiff <= 30) {
      // This month
      cellRange.setBackground(theme.COLORS.ORANGE_GLOW);
      cellRange.setFontColor(theme.COLORS.DARK_BG);
    } else {
      // Future events
      cellRange.setBackground(theme.COLORS.NEON_BLUE);
      cellRange.setFontColor(theme.COLORS.TEXT_WHITE);
    }
    cellRange.setFontWeight('bold');
  });
}

/**
 * Apply cyberpunk theme to all existing sheets
 */
function applyCyberpunkThemeToAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    // Only apply to system sheets, not user's other sheets
    if (Object.values(SYSTEM_CONFIG.SHEET_NAMES).includes(sheetName)) {
      applyCyberpunkTheme(sheet, sheetName);
    }
  });
  
  SpreadsheetApp.getUi().alert('🌟 Cyberpunk Theme Applied!', 
    'All system sheets have been transformed with the cyberpunk theme. Welcome to the matrix! 🚀', 
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Reset sheets to default Google Sheets styling
 */
function resetSheetsToDefault() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    if (Object.values(SYSTEM_CONFIG.SHEET_NAMES).includes(sheetName)) {
      const lastRow = Math.max(sheet.getLastRow(), 1);
      const lastCol = Math.max(sheet.getLastColumn(), 1);
      
      // Reset to default white background
      const fullRange = sheet.getRange(1, 1, Math.max(lastRow, 100), lastCol);
      fullRange.setBackground('#ffffff');
      fullRange.setFontColor('#000000');
      fullRange.setBorder(false, false, false, false, false, false);
      
      // Reset header to default
      if (lastRow >= 1) {
        const headerRange = sheet.getRange(1, 1, 1, lastCol);
        headerRange.setBackground('#f3f3f3');
        headerRange.setFontColor('#000000');
        headerRange.setFontWeight('bold');
      }
    }
  });
  
  SpreadsheetApp.getUi().alert('Reset Complete', 
    'All sheets have been reset to default Google Sheets styling.', 
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Re-style the currently active sheet with cyberpunk theme
 */
function restyle_current_sheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheet = ss.getActiveSheet();
  const sheetName = currentSheet.getName();
  
  // Check if it's a system sheet
  if (Object.values(SYSTEM_CONFIG.SHEET_NAMES).includes(sheetName)) {
    applyCyberpunkTheme(currentSheet, sheetName);
    SpreadsheetApp.getUi().alert('✨ Sheet Re-styled!', 
      `${sheetName} has been updated with the cyberpunk theme!`, 
      SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    SpreadsheetApp.getUi().alert('⚠️ Sheet Not Supported', 
      'Cyberpunk theme can only be applied to system sheets.', 
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// Additional utility functions for data analysis and reporting would go here...

/**
 * Show members view
 */
function showMembersView() {
  // Pre-load the data and pass it to the template
  const membersData = getMembersData();
  console.log('showMembersView: Pre-loading data:', membersData);
  
  const html = HtmlService.createTemplateFromFile('members-view');
  html.membersData = membersData; // Pass data to template
  
  const htmlOutput = html.evaluate()
    .setWidth(1000)
    .setHeight(600)
    .setTitle('Members Overview');
    
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Members Overview');
}

/**
 * Get members data for the HTML interface
 */
function getMembersData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const membersSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.MEMBERS);
    
    console.log('getMembersData: Looking for sheet:', SYSTEM_CONFIG.SHEET_NAMES.MEMBERS);
    
    if (!membersSheet) {
      console.log('getMembersData: Sheet not found. Available sheets:', ss.getSheets().map(s => s.getName()));
      return {
        members: [],
        stats: { total: 0, active: 0, faculty: 0, students: 0 }
      };
    }
    
    const lastRow = membersSheet.getLastRow();
    console.log('getMembersData: Last row:', lastRow);
    
    if (lastRow <= 1) {
      console.log('getMembersData: No data rows found');
      return {
        members: [],
        stats: { total: 0, active: 0, faculty: 0, students: 0 }
      };
    }
    
    // Get all member data
    const data = membersSheet.getRange(2, 1, lastRow - 1, membersSheet.getLastColumn()).getValues();
    console.log('getMembersData: Retrieved', data.length, 'rows with', membersSheet.getLastColumn(), 'columns');
    
    const members = [];
    let activeCount = 0;
    let facultyCount = 0;
    let studentCount = 0;
    
    data.forEach((row, index) => {
      // Skip completely empty rows
      if (!row[0] && !row[1] && !row[2]) {
        console.log('getMembersData: Skipping empty row', index + 2);
        return;
      }
      
      const member = {
        id: String(row[0] || `member_${index + 1}`),
        name: String(row[1] || 'Unknown'),
        email: String(row[2] || ''), // Institutional email
        role: String(row[3] || 'Unknown'),
        department: String(row[4] || ''),
        joinDate: row[5] || new Date(),
        status: String(row[6] || 'Active'),
        researchInterests: String(row[7] || ''),
        contactInfo: String(row[8] || ''),
        lastActive: row[9] || new Date(),
        googleEmail: String(row[10] || '') // Google email for user identification
      };
      
      console.log('getMembersData: Processing member:', member.name, '(' + member.id + ') - Status:', member.status, 'Role:', member.role);
      members.push(member);
      
      // Count stats (case insensitive)
      const status = member.status.toLowerCase();
      if (status === 'active' || status === 'Active') activeCount++;
      
      const role = member.role.toLowerCase();
      if (role.includes('professor') || role.includes('faculty') || role.includes('admin')) {
        facultyCount++;
      } else if (role.includes('student') || role.includes('undergraduate') || role.includes('phd')) {
        studentCount++;
      }
    });
    
    const result = {
      members: members,
      stats: {
        total: members.length,
        active: activeCount,
        faculty: facultyCount,
        students: studentCount
      }
    };
    
    console.log('getMembersData: Returning', result.members.length, 'members with stats:', result.stats);
    return result;
  } catch (error) {
    console.error('getMembersData: Error occurred:', error.toString());
    console.error('getMembersData: Full error details:', error);
    
    // Return a proper structure even on error
    const errorResult = {
      members: [],
      stats: { total: 0, active: 0, faculty: 0, students: 0 },
      error: error.toString()
    };
    
    console.log('getMembersData: Returning error result:', errorResult);
    return errorResult;
  }
}

/**
 * Show learning progress
 */
function showLearningProgress() {
  // Get current user email
  const currentUserEmail = Session.getActiveUser().getEmail();
  console.log('showLearningProgress: Current user email:', currentUserEmail);
  
  // Pre-load the members list
  const allMembers = getAllMembers();
  console.log('showLearningProgress: Pre-loading members:', allMembers);
  
  // Find current user in members list (check both institutional and Google email)
  const currentMember = allMembers.find(member => 
    (member.email && member.email.toLowerCase() === currentUserEmail.toLowerCase()) ||
    (member.googleEmail && member.googleEmail.toLowerCase() === currentUserEmail.toLowerCase())
  );
  console.log('showLearningProgress: Current member found:', currentMember);
  
  // Pre-load learning progress for all members
  const learningProgressData = {};
  allMembers.forEach(member => {
    console.log('showLearningProgress: Pre-loading progress for:', member.id);
    learningProgressData[member.id] = getLearningProgressDataForMember(member.id);
  });
  console.log('showLearningProgress: Pre-loaded progress for', Object.keys(learningProgressData).length, 'members');
  
  const html = HtmlService.createTemplateFromFile('learning-progress');
  html.allMembers = allMembers;
  html.learningProgressData = learningProgressData;
  html.currentUserEmail = currentUserEmail;
  html.currentMember = currentMember;
  
  const htmlOutput = html.evaluate()
    .setWidth(1000)
    .setHeight(700)
    .setTitle('Learning Progress');
    
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Learning Progress');
}

/**
 * Get learning progress data for a specific member
 */
function getLearningProgressDataForMember(memberId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const learningSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.LEARNING_ACTIVITIES);
    
    if (!learningSheet || learningSheet.getLastRow() <= 1) {
      return {
        activities: [],
        stats: { total: 0, completed: 0, avgProgress: 0, totalHours: 0 },
        filters: { members: [], types: [] },
        error: null
      };
    }
    
    // Get member's learning activities
    const learning = getMemberLearning(learningSheet, memberId);
    
    // Calculate stats
    const totalActivities = learning.length;
    const completedActivities = learning.filter(activity => activity.progress >= 100).length;
    const avgProgress = totalActivities > 0 ? 
      learning.reduce((sum, activity) => sum + activity.progress, 0) / totalActivities : 0;
    const totalHours = learning.reduce((sum, activity) => sum + activity.duration, 0);
    
    // Get filters data
    const filters = getLearningFilters();
    
    return {
      activities: learning,
      stats: {
        total: totalActivities,
        completed: completedActivities,
        avgProgress: Math.round(avgProgress * 10) / 10,
        totalHours: totalHours
      },
      filters: filters,
      error: null
    };
    
  } catch (error) {
    console.error('getLearningProgressDataForMember: Error for member', memberId, ':', error);
    return {
      activities: [],
      stats: { total: 0, completed: 0, avgProgress: 0, totalHours: 0 },
      filters: { members: [], types: [] },
      error: error.toString()
    };
  }
}

/**
 * Get learning filters for the interface
 */
function getLearningFilters() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const learningSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.LEARNING_ACTIVITIES);
  const membersSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.MEMBERS);
  
  const uniqueMembers = new Set();
  const uniqueTypes = new Set();
  
  // Get member names
  if (membersSheet && membersSheet.getLastRow() > 1) {
    const memberData = membersSheet.getRange(2, 1, membersSheet.getLastRow() - 1, 2).getValues();
    memberData.forEach(row => {
      uniqueMembers.add(row[1] || row[0]);
    });
  }
  
  // Get learning types
  if (learningSheet && learningSheet.getLastRow() > 1) {
    const data = learningSheet.getRange(2, 3, learningSheet.getLastRow() - 1, 1).getValues();
    data.forEach(row => {
      if (row[0]) uniqueTypes.add(row[0]);
    });
  }
  
  return {
    members: Array.from(uniqueMembers).sort(),
    types: Array.from(uniqueTypes).sort()
  };
}

/**
 * Get learning progress data for the HTML interface (legacy)
 */
function getLearningProgressData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const learningSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.LEARNING_ACTIVITIES);
  const membersSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.MEMBERS);
  
  if (!learningSheet || learningSheet.getLastRow() <= 1) {
    return {
      activities: [],
      stats: { total: 0, completed: 0, avgProgress: 0, totalHours: 0 },
      filters: { members: [], types: [] }
    };
  }
  
  // Get member names mapping
  const memberNames = {};
  if (membersSheet && membersSheet.getLastRow() > 1) {
    const memberData = membersSheet.getRange(2, 1, membersSheet.getLastRow() - 1, 2).getValues();
    memberData.forEach(row => {
      memberNames[row[0]] = row[1] || row[0];
    });
  }
  
  // Get learning activities data
  const data = learningSheet.getRange(2, 1, learningSheet.getLastRow() - 1, learningSheet.getLastColumn()).getValues();
  
  const activities = [];
  let totalProgress = 0;
  let completedCount = 0;
  let totalHours = 0;
  const uniqueMembers = new Set();
  const uniqueTypes = new Set();
  
  data.forEach(row => {
    const memberId = row[1] || 'Unknown';
    const memberName = memberNames[memberId] || memberId;
    const progress = parseInt(row[7]) || 0;
    const duration = parseFloat(row[6]) || 0;
    const type = row[2] || 'Unknown';
    
    const activity = {
      id: row[0] || '',
      memberId: memberId,
      memberName: memberName,
      type: type,
      title: row[3] || 'Untitled Activity',
      description: row[4] || '',
      source: row[5] || '',
      duration: duration,
      progress: progress,
      startDate: row[8] || new Date(),
      completionDate: row[9] || null,
      rating: parseFloat(row[10]) || null,
      notes: row[11] || '',
      skillsGained: row[12] || ''
    };
    
    activities.push(activity);
    
    // Calculate stats
    totalProgress += progress;
    totalHours += duration;
    if (progress === 100) completedCount++;
    
    uniqueMembers.add(memberName);
    uniqueTypes.add(type);
  });
  
  const stats = {
    total: activities.length,
    completed: completedCount,
    avgProgress: activities.length > 0 ? Math.round(totalProgress / activities.length) : 0,
    totalHours: Math.round(totalHours * 10) / 10
  };
  
  return {
    activities: activities,
    stats: stats,
    filters: {
      members: Array.from(uniqueMembers).sort(),
      types: Array.from(uniqueTypes).sort()
    }
  };
}

/**
 * Show skill matrix
 */
function showSkillMatrix() {
  const html = HtmlService.createTemplateFromFile('skill-matrix')
    .evaluate()
    .setWidth(1200)
    .setHeight(800)
    .setTitle('Skill Matrix');
    
  SpreadsheetApp.getUi().showModalDialog(html, 'Skill Matrix');
}

/**
 * Get skill matrix data for the HTML interface
 */
function getSkillMatrixData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const assessmentSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.ABILITY_ASSESSMENTS);
  const membersSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.MEMBERS);
  
  if (!assessmentSheet || assessmentSheet.getLastRow() <= 1) {
    return {
      skillSummary: {},
      stats: { totalMembers: 0, totalSkills: 0, avgLevel: 0, expertCount: 0 }
    };
  }
  
  // Get member names mapping
  const memberNames = {};
  if (membersSheet && membersSheet.getLastRow() > 1) {
    const memberData = membersSheet.getRange(2, 1, membersSheet.getLastRow() - 1, 2).getValues();
    memberData.forEach(row => {
      memberNames[row[0]] = row[1] || row[0]; // Map ID to Name, fallback to ID
    });
  }
  
  // Get assessment data and create a summary
  const data = assessmentSheet.getRange(2, 1, assessmentSheet.getLastRow() - 1, assessmentSheet.getLastColumn()).getValues();
  const skillSummary = {};
  let totalSkills = 0;
  let totalLevel = 0;
  let expertCount = 0;
  
  data.forEach(row => {
    const memberId = row[1] || 'Unknown';
    const memberName = memberNames[memberId] || memberId;
    const category = row[2] || 'General';
    const skill = row[3] || 'Unknown Skill';
    const level = parseInt(row[4]) || 0;
    
    if (!skillSummary[memberName]) {
      skillSummary[memberName] = {};
    }
    if (!skillSummary[memberName][category]) {
      skillSummary[memberName][category] = [];
    }
    
    skillSummary[memberName][category].push({
      name: skill,
      level: level
    });
    
    totalSkills++;
    totalLevel += level;
    if (level >= 9) expertCount++;
  });
  
  const stats = {
    totalMembers: Object.keys(skillSummary).length,
    totalSkills: totalSkills,
    avgLevel: totalSkills > 0 ? totalLevel / totalSkills : 0,
    expertCount: expertCount
  };
  
  return {
    skillSummary: skillSummary,
    stats: stats
  };
}

/**
 * Show project hiring interface
 */
function showProjectHiring() {
  // Pre-load project hiring data
  const projectHiringData = getProjectHiringData();
  
  const html = HtmlService.createTemplateFromFile('project-hiring');
  html.projectHiringData = projectHiringData;
  
  const htmlOutput = html.evaluate()
    .setWidth(1000)
    .setHeight(700)
    .setTitle('Project Hiring');
    
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Project Hiring');
}

/**
 * Get project hiring data for the HTML interface
 */
function getProjectHiringData() {
  console.log('=== getProjectHiringData() ===');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const projectsSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.PROJECTS);
  const membersSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.MEMBERS);
  
  console.log('getProjectHiringData: Looking for sheet:', SYSTEM_CONFIG.SHEET_NAMES.PROJECTS);
  console.log('getProjectHiringData: Projects sheet found:', !!projectsSheet);
  console.log('getProjectHiringData: Members sheet found:', !!membersSheet);
  
  if (!projectsSheet || projectsSheet.getLastRow() <= 1) {
    console.log('getProjectHiringData: No projects data available');
    return {
      projects: [],
      stats: { total: 0, urgent: 0, positions: 0, avgBudget: 0 },
      skills: []
    };
  }
  
  // Get member names mapping
  const memberNames = {};
  if (membersSheet && membersSheet.getLastRow() > 1) {
    const memberData = membersSheet.getRange(2, 1, membersSheet.getLastRow() - 1, 2).getValues();
    memberData.forEach(row => {
      memberNames[row[0]] = row[1] || row[0];
    });
    console.log('getProjectHiringData: Member names mapping created, count:', Object.keys(memberNames).length);
  }
  
  // Get project data
  console.log('getProjectHiringData: Last row:', projectsSheet.getLastRow());
  console.log('getProjectHiringData: Last column:', projectsSheet.getLastColumn());
  const data = projectsSheet.getRange(2, 1, projectsSheet.getLastRow() - 1, projectsSheet.getLastColumn()).getValues();
  console.log('getProjectHiringData: Retrieved', data.length, 'rows with', data[0]?.length || 0, 'columns');
  const hiringProjects = data.filter(row => row[3] === 'Active' || row[3] === 'Planning' || row[3] === 'Recruiting');
  console.log('getProjectHiringData: Filtered to', hiringProjects.length, 'hiring projects');
  
  const projects = [];
  let totalBudget = 0;
  let budgetCount = 0;
  let urgentCount = 0;
  const allSkills = new Set();
  
  hiringProjects.forEach((row, index) => {
    const leaderId = row[4] || '';
    const leaderName = memberNames[leaderId] || leaderId || 'TBD';
    const budget = parseFloat(row[9]) || 0;
    const skillsRequired = row[6] || '';
    const teamMembers = row[5] || '';
    const teamSize = teamMembers ? teamMembers.split(',').length : 1;
    
    const project = {
      id: row[0] || '',
      title: row[1] || 'Untitled Project',
      description: row[2] || '',
      status: row[3] || 'Unknown',
      leaderId: leaderId,
      leaderName: leaderName,
      teamMembers: teamMembers,
      teamSize: teamSize,
      skillsRequired: skillsRequired,
      startDate: row[7] || new Date(),
      endDate: row[8] || null,
      budget: budget,
      progress: parseInt(row[10]) || 0,
      repositoryUrl: row[11] || ''
    };
    
    console.log(`getProjectHiringData: Processing project ${index + 1}: ${project.title} (${project.id}) - Status: ${project.status} Leader: ${leaderName}`);
    
    projects.push(project);
    
    // Calculate stats
    if (budget > 0) {
      totalBudget += budget;
      budgetCount++;
    }
    
    if (project.status === 'Recruiting' || skillsRequired.toLowerCase().includes('urgent')) {
      urgentCount++;
    }
    
    // Extract skills
    if (skillsRequired) {
      skillsRequired.split(',').forEach(skill => {
        allSkills.add(skill.trim());
      });
    }
  });
  
  const stats = {
    total: projects.length,
    urgent: urgentCount,
    positions: projects.reduce((sum, p) => sum + p.teamSize, 0),
    avgBudget: budgetCount > 0 ? Math.round(totalBudget / budgetCount) : 0
  };
  
  console.log('getProjectHiringData: Returning', projects.length, 'projects with stats:', stats);
  console.log('getProjectHiringData: Skills found:', Array.from(allSkills).sort());
  
  const result = {
    projects: projects,
    stats: stats,
    skills: Array.from(allSkills).sort()
  };
  
  console.log('getProjectHiringData result:', result);
  return result;
}

/**
 * Show idea repository
 */
function showIdeaRepository() {
  console.log('showIdeaRepository: Starting pre-load of ideas repository data');
  
  // Pre-load ideas repository data
  const ideasRepositoryData = getIdeasRepositoryData();
  console.log('showIdeaRepository: Pre-loaded data:', ideasRepositoryData);
  
  const html = HtmlService.createTemplateFromFile('idea-repository');
  html.ideasRepositoryData = ideasRepositoryData; // Pass data to template
  
  const template = html.evaluate()
    .setWidth(1200)
    .setHeight(800)
    .setTitle('Academic Ideas Repository');
    
  SpreadsheetApp.getUi().showModalDialog(template, 'Academic Ideas Repository');
}

/**
 * Get ideas repository data for the HTML interface
 */
function getIdeasRepositoryData() {
  console.log('=== GETTING IDEAS REPOSITORY DATA ===');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ideasSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.ACADEMIC_IDEAS);
  const membersSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.MEMBERS);
  
  console.log('getIdeasRepositoryData: Looking for sheet:', SYSTEM_CONFIG.SHEET_NAMES.ACADEMIC_IDEAS);
  console.log('getIdeasRepositoryData: Ideas sheet found:', !!ideasSheet);
  console.log('getIdeasRepositoryData: Members sheet found:', !!membersSheet);
  
  if (!ideasSheet || ideasSheet.getLastRow() <= 1) {
    console.log('getIdeasRepositoryData: No ideas sheet or no data, returning empty result');
    return {
      ideas: [],
      stats: { total: 0, approved: 0, implemented: 0, avgVotes: 0 },
      categories: []
    };
  }
  
  // Get member names mapping
  const memberNames = {};
  if (membersSheet && membersSheet.getLastRow() > 1) {
    const memberData = membersSheet.getRange(2, 1, membersSheet.getLastRow() - 1, 2).getValues();
    memberData.forEach(row => {
      memberNames[row[0]] = row[1] || row[0];
    });
    console.log('getIdeasRepositoryData: Member names mapping:', memberNames);
  }
  
  // Get ideas data
  const lastRow = ideasSheet.getLastRow();
  const lastColumn = ideasSheet.getLastColumn();
  console.log('getIdeasRepositoryData: Ideas sheet dimensions - Last row:', lastRow, 'Last column:', lastColumn);
  
  const data = ideasSheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
  console.log('getIdeasRepositoryData: Retrieved', data.length, 'rows with', data[0]?.length || 0, 'columns');
  
  const ideas = [];
  let approvedCount = 0;
  let implementedCount = 0;
  let totalVotes = 0;
  const uniqueCategories = new Set();
  
  data.forEach((row, index) => {
    const submitterId = row[1] || '';
    const submitterName = memberNames[submitterId] || submitterId || 'Anonymous';
    const status = row[9] || 'Under Review';
    const voteScore = parseInt(row[11]) || 0;
    
    console.log(`getIdeasRepositoryData: Processing idea ${index + 1}: ${row[2] || 'Untitled'} (${row[0]}) - Status: ${status} - Submitter: ${submitterName}`);
    
    const idea = {
      id: row[0] || '',
      submitterId: submitterId,
      submitterName: submitterName,
      title: row[2] || 'Untitled Idea',
      abstract: row[3] || '',
      category: row[4] || 'General',
      keywords: row[5] || '',
      potentialImpact: row[6] || '',
      resourcesNeeded: row[7] || '',
      submitDate: row[8] || new Date(),
      status: status,
      reviewerComments: row[10] || '',
      voteScore: voteScore
    };
    
    ideas.push(idea);
    
    // Calculate stats
    if (status === 'Approved') approvedCount++;
    if (status === 'Implemented') implementedCount++;
    totalVotes += voteScore;
    uniqueCategories.add(idea.category);
  });
  
  const stats = {
    total: ideas.length,
    approved: approvedCount,
    implemented: implementedCount,
    avgVotes: ideas.length > 0 ? totalVotes / ideas.length : 0
  };
  
  console.log('getIdeasRepositoryData: Returning', ideas.length, 'ideas with stats:', stats);
  console.log('getIdeasRepositoryData: Categories found:', Array.from(uniqueCategories).sort());
  
  return {
    ideas: ideas,
    stats: stats,
    categories: Array.from(uniqueCategories).sort()
  };
}

/**
 * Vote on an academic idea
 */
function voteOnIdea(ideaId, vote) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ideasSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.ACADEMIC_IDEAS);
    
    if (!ideasSheet) return { success: false, message: 'Ideas sheet not found' };
    
    // Find the idea row
    const data = ideasSheet.getRange(1, 1, ideasSheet.getLastRow(), ideasSheet.getLastColumn()).getValues();
    const ideaRowIndex = data.findIndex(row => row[0] === ideaId);
    
    if (ideaRowIndex === -1) {
      return { success: false, message: 'Idea not found' };
    }
    
    // Update vote score (assuming vote score is in column L, index 11)
    const currentVotes = parseInt(data[ideaRowIndex][11]) || 0;
    const newVotes = Math.max(0, currentVotes + vote); // Prevent negative votes
    
    ideasSheet.getRange(ideaRowIndex + 1, 12).setValue(newVotes);
    
    return { success: true, message: 'Vote recorded successfully' };
  } catch (error) {
    return { success: false, message: 'Error recording vote: ' + error.message };
  }
}

/**
 * Show event dialog
 */
function showEventDialog() {
  const html = HtmlService.createTemplateFromFile('event-dialog')
    .evaluate()
    .setWidth(700)
    .setHeight(600)
    .setTitle('Schedule Event');
    
  SpreadsheetApp.getUi().showModalDialog(html, 'Schedule Event');
}

/**
 * Show analytics hub
 */
function showAnalyticsHub() {
  const html = HtmlService.createTemplateFromFile('analytics-hub')
    .evaluate()
    .setWidth(1400)
    .setHeight(900)
    .setTitle('Analytics Hub');
    
  SpreadsheetApp.getUi().showModalDialog(html, 'Analytics Hub');
}

/**
 * Placeholder functions for menu items that don't have HTML files yet
 */

function showMemberProfile() {
  // Pre-load the members list
  const allMembers = getAllMembers();
  console.log('showMemberProfile: Pre-loading members:', allMembers);
  
  // Pre-load all member profiles
  const memberProfiles = {};
  allMembers.forEach(member => {
    console.log('showMemberProfile: Pre-loading profile for:', member.id);
    memberProfiles[member.id] = getMemberProfileData(member.id);
  });
  console.log('showMemberProfile: Pre-loaded profiles:', Object.keys(memberProfiles));
  
  const html = HtmlService.createTemplateFromFile('member-profile');
  html.allMembers = allMembers; // Pass members list to template
  html.memberProfiles = memberProfiles; // Pass all profiles to template
  
  const htmlOutput = html.evaluate()
    .setWidth(1400)
    .setHeight(900)
    .setTitle('Member Profile');
    
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Member Profile');
}

/**
 * Get comprehensive member profile data
 */
function getMemberProfileData(memberId) {
  console.log('getMemberProfileData: Starting with ID:', memberId);
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const membersSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.MEMBERS);
    const assessmentsSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.ABILITY_ASSESSMENTS);
    const learningSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.LEARNING_ACTIVITIES);
    const projectsSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.PROJECTS);
    const ideasSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.ACADEMIC_IDEAS);
    
    console.log('getMemberProfileData: Sheets found:', {
      members: !!membersSheet,
      assessments: !!assessmentsSheet,
      learning: !!learningSheet,
      projects: !!projectsSheet,
      ideas: !!ideasSheet
    });
    
    // Get member basic info
    console.log('getMemberProfileData: Getting basic info for ID:', memberId);
    const memberData = getMemberBasicInfo(membersSheet, memberId);
    console.log('getMemberProfileData: Basic info result:', memberData);
    
    if (!memberData) {
      console.log('getMemberProfileData: No member data found, returning null');
      return null;
    }
  
    // Get member's skills
    console.log('getMemberProfileData: Getting skills for ID:', memberId);
    const skills = getMemberSkills(assessmentsSheet, memberId);
    console.log('getMemberProfileData: Skills result:', skills);
    
    // Get member's learning activities
    console.log('getMemberProfileData: Getting learning activities for ID:', memberId);
    const learning = getMemberLearning(learningSheet, memberId);
    console.log('getMemberProfileData: Learning result:', learning);
    
    // Get member's project involvement
    console.log('getMemberProfileData: Getting projects for ID:', memberId);
    const projects = getMemberProjects(projectsSheet, memberId);
    console.log('getMemberProfileData: Projects result:', projects);
    
    // Get member's submitted ideas
    console.log('getMemberProfileData: Getting ideas count for ID:', memberId);
    const ideasCount = getMemberIdeasCount(ideasSheet, memberId);
    console.log('getMemberProfileData: Ideas count result:', ideasCount);
    
    // Calculate statistics
    console.log('getMemberProfileData: Calculating stats');
    const stats = calculateMemberStats(skills, learning, projects, ideasCount);
    console.log('getMemberProfileData: Stats result:', stats);
    
    const result = {
      member: memberData,
      skills: skills,
      learning: learning,
      projects: projects,
      stats: stats
    };
    
    console.log('getMemberProfileData: Returning result:', result);
    return result;
    
  } catch (error) {
    console.error('getMemberProfileData: Error occurred:', error);
    return null;
  }
}

/**
 * Get member basic information
 */
function getMemberBasicInfo(sheet, memberId) {
  try {
    console.log('getMemberBasicInfo: Looking for member ID:', memberId);
    
    if (!sheet) {
      console.log('getMemberBasicInfo: Sheet is null');
      return null;
    }
    
    if (sheet.getLastRow() <= 1) {
      console.log('getMemberBasicInfo: No data rows in sheet');
      return null;
    }
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    console.log('getMemberBasicInfo: Retrieved', data.length, 'rows, looking for member:', memberId);
    
    // Log all member IDs for debugging
    data.forEach((row, index) => {
      console.log('getMemberBasicInfo: Row', index + 1, 'ID:', row[0], 'Name:', row[1]);
    });
    
    const memberRow = data.find(row => {
      const rowId = row[0];
      console.log('getMemberBasicInfo: Comparing', rowId, 'with', memberId, '(types:', typeof rowId, typeof memberId, ')');
      return String(rowId) === String(memberId);
    });
    
    if (!memberRow) {
      console.log('getMemberBasicInfo: Member not found with ID:', memberId);
      return null;
    }
    
    const memberInfo = {
      id: memberRow[0],
      name: memberRow[1] || 'Unknown',
      email: memberRow[2] || '', // Institutional email
      role: memberRow[3] || 'Unknown',
      department: memberRow[4] || '',
      joinDate: memberRow[5] || new Date(),
      status: memberRow[6] || 'Active',
      researchInterests: memberRow[7] || '',
      contactInfo: memberRow[8] || '',
      lastActive: memberRow[9] || new Date(),
      googleEmail: memberRow[10] || '' // Google email for user identification
    };
    
    console.log('getMemberBasicInfo: Found member:', memberInfo.name, '(' + memberInfo.id + ')');
    return memberInfo;
  } catch (error) {
    console.error('getMemberBasicInfo: Error occurred:', error.toString());
    return null;
  }
}

/**
 * Get member's skills organized by category
 */
function getMemberSkills(sheet, memberId) {
  if (!sheet || sheet.getLastRow() <= 1) return {};
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const memberAssessments = data.filter(row => row[1] === memberId);
  
  const skillsByCategory = {};
  
  memberAssessments.forEach(row => {
    const category = row[2] || 'General';
    const skill = row[3] || 'Unknown Skill';
    const level = parseInt(row[4]) || 0;
    
    if (!skillsByCategory[category]) {
      skillsByCategory[category] = [];
    }
    
    skillsByCategory[category].push({
      name: skill,
      level: level,
      assessmentDate: row[5] || new Date(),
      evidence: row[7] || ''
    });
  });
  
  return skillsByCategory;
}

/**
 * Get member's learning activities
 */
function getMemberLearning(sheet, memberId) {
  if (!sheet || sheet.getLastRow() <= 1) return [];
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const memberActivities = data.filter(row => row[1] === memberId);
  
  // Get member name for display
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const membersSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.MEMBERS);
  let memberName = memberId;
  if (membersSheet && membersSheet.getLastRow() > 1) {
    const memberData = membersSheet.getRange(2, 1, membersSheet.getLastRow() - 1, 2).getValues();
    const memberRow = memberData.find(row => row[0] === memberId);
    if (memberRow) {
      memberName = memberRow[1] || memberId;
    }
  }
  
  return memberActivities.map(row => ({
    id: row[0],
    memberId: memberId,
    memberName: memberName,
    type: row[2] || 'Unknown',
    title: row[3] || 'Untitled Activity',
    description: row[4] || '',
    source: row[5] || '',
    duration: parseFloat(row[6]) || 0,
    progress: parseInt(row[7]) || 0,
    startDate: row[8] || new Date(),
    completionDate: row[9] || null,
    rating: parseFloat(row[10]) || null,
    notes: row[11] || '',
    skillsGained: row[12] || ''
  })).sort((a, b) => new Date(b.startDate) - new Date(a.startDate));
}

/**
 * Get member's project involvement
 */
function getMemberProjects(sheet, memberId) {
  if (!sheet || sheet.getLastRow() <= 1) return [];
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const memberProjects = [];
  
  data.forEach(row => {
    const leaderId = row[4];
    const teamMembers = (row[5] || '').split(',').map(m => m.trim());
    
    let role = 'Member';
    if (leaderId === memberId) {
      role = 'Leader';
    } else if (teamMembers.includes(memberId)) {
      role = 'Team Member';
    } else {
      return; // Member not involved in this project
    }
    
    memberProjects.push({
      id: row[0],
      title: row[1] || 'Untitled Project',
      description: row[2] || '',
      status: row[3] || 'Unknown',
      role: role,
      progress: parseInt(row[10]) || 0,
      startDate: row[7] || new Date()
    });
  });
  
  return memberProjects.sort((a, b) => new Date(b.startDate) - new Date(a.startDate));
}

/**
 * Get count of member's submitted ideas
 */
function getMemberIdeasCount(sheet, memberId) {
  if (!sheet || sheet.getLastRow() <= 1) return 0;
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  return data.filter(row => row[1] === memberId).length;
}

/**
 * Calculate member statistics
 */
function calculateMemberStats(skills, learning, projects, ideasCount) {
  const allSkills = Object.values(skills).flat();
  const totalSkills = allSkills.length;
  const avgSkillLevel = totalSkills > 0 ? allSkills.reduce((sum, skill) => sum + skill.level, 0) / totalSkills : 0;
  const totalHours = learning.reduce((sum, activity) => sum + activity.duration, 0);
  
  return {
    totalSkills: totalSkills,
    avgSkillLevel: avgSkillLevel,
    learningActivities: learning.length,
    projectsInvolved: projects.length,
    ideasSubmitted: ideasCount,
    totalHours: Math.round(totalHours * 10) / 10
  };
}

function showLearningAnalytics() {
  const html = HtmlService.createTemplateFromFile('learning-analytics')
    .evaluate()
    .setWidth(1600)
    .setHeight(1000)
    .setTitle('Learning Analytics');
    
  SpreadsheetApp.getUi().showModalDialog(html, 'Learning Analytics');
}

/**
 * Get learning analytics data with advanced metrics
 */
function getLearningAnalyticsData(filters = {}) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const learningSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.LEARNING_ACTIVITIES);
  const membersSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.MEMBERS);
  
  if (!learningSheet || learningSheet.getLastRow() <= 1) {
    return {
      totalActivities: 0,
      metrics: { totalHours: 0, completionRate: 0, activeLearners: 0, avgRating: 0 },
      filters: { departments: [], types: [] },
      charts: {},
      insights: []
    };
  }
  
  // Get all data
  const learningData = learningSheet.getRange(2, 1, learningSheet.getLastRow() - 1, learningSheet.getLastColumn()).getValues();
  const membersData = getMembersDataForAnalytics(membersSheet);
  
  // Filter data based on time period
  const filteredData = filterDataByPeriod(learningData, filters.period);
  
  // Calculate metrics
  const metrics = calculateAnalyticsMetrics(filteredData, membersData);
  
  // Generate charts data
  const charts = generateChartsData(filteredData, membersData);
  
  // Generate AI insights
  const insights = generateAIInsights(metrics, charts);
  
  // Get filter options
  const filterOptions = getAnalyticsFilters(learningData, membersData);
  
  return {
    totalActivities: filteredData.length,
    metrics: metrics,
    filters: filterOptions,
    charts: charts,
    insights: insights
  };
}

/**
 * Get members data for analytics
 */
function getMembersDataForAnalytics(sheet) {
  if (!sheet || sheet.getLastRow() <= 1) return {};
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const membersMap = {};
  
  data.forEach(row => {
    membersMap[row[0]] = {
      name: row[1] || 'Unknown',
      department: row[4] || 'Unknown',
      role: row[3] || 'Unknown'
    };
  });
  
  return membersMap;
}

/**
 * Filter data by time period
 */
function filterDataByPeriod(data, period) {
  if (period === 'all') return data;
  
  const now = new Date();
  const daysBack = parseInt(period) || 30;
  const cutoffDate = new Date(now.getTime() - (daysBack * 24 * 60 * 60 * 1000));
  
  return data.filter(row => {
    const startDate = new Date(row[8] || row[8]); // Start date column
    return startDate >= cutoffDate;
  });
}

/**
 * Calculate comprehensive analytics metrics
 */
function calculateAnalyticsMetrics(data, membersData) {
  if (data.length === 0) {
    return {
      totalHours: 0,
      completionRate: 0,
      activeLearners: 0,
      avgRating: 0,
      hoursChange: 0,
      completionChange: 0,
      learnersChange: 0,
      ratingChange: 0
    };
  }
  
  const totalHours = data.reduce((sum, row) => sum + (parseFloat(row[6]) || 0), 0);
  const completedActivities = data.filter(row => (parseInt(row[7]) || 0) === 100).length;
  const completionRate = (completedActivities / data.length) * 100;
  const uniqueLearners = new Set(data.map(row => row[1])).size;
  const ratingsData = data.filter(row => row[10] && parseFloat(row[10]) > 0);
  const avgRating = ratingsData.length > 0 ? 
    ratingsData.reduce((sum, row) => sum + parseFloat(row[10]), 0) / ratingsData.length : 0;
  
  // Calculate change percentages (simplified - would need historical data for real implementation)
  const hoursChange = Math.random() * 20 - 10; // Placeholder
  const completionChange = Math.random() * 15 - 7.5; // Placeholder
  const learnersChange = Math.random() * 25 - 12.5; // Placeholder
  const ratingChange = Math.random() * 10 - 5; // Placeholder
  
  return {
    totalHours,
    completionRate,
    activeLearners: uniqueLearners,
    avgRating,
    hoursChange,
    completionChange,
    learnersChange,
    ratingChange
  };
}

/**
 * Generate charts data for visualization
 */
function generateChartsData(data, membersData) {
  // Learners by department
  const departmentCounts = {};
  data.forEach(row => {
    const memberId = row[1];
    const dept = membersData[memberId]?.department || 'Unknown';
    departmentCounts[dept] = (departmentCounts[dept] || 0) + 1;
  });
  
  const learners = Object.entries(departmentCounts).map(([dept, count]) => ({
    label: dept,
    value: count
  }));
  
  // Rating distribution
  const ratingCounts = {};
  for (let i = 1; i <= 5; i++) {
    ratingCounts[i] = 0;
  }
  
  data.forEach(row => {
    const rating = Math.floor(parseFloat(row[10]) || 0);
    if (rating >= 1 && rating <= 5) {
      ratingCounts[rating]++;
    }
  });
  
  const ratings = Object.entries(ratingCounts).map(([rating, count]) => ({
    label: `${rating} ⭐`,
    value: count
  }));
  
  // Learning types distribution
  const typeCounts = {};
  data.forEach(row => {
    const type = row[2] || 'Unknown';
    typeCounts[type] = (typeCounts[type] || 0) + 1;
  });
  
  const types = Object.entries(typeCounts).map(([type, count]) => ({
    label: type,
    value: count
  }));
  
  // Weekly progress (last 7 days)
  const weekly = [];
  const now = new Date();
  for (let i = 6; i >= 0; i--) {
    const date = new Date(now.getTime() - (i * 24 * 60 * 60 * 1000));
    const dayData = data.filter(row => {
      const startDate = new Date(row[8]);
      return startDate.toDateString() === date.toDateString();
    });
    
    weekly.push({
      label: date.toLocaleDateString('en-US', { weekday: 'short' }),
      value: dayData.length
    });
  }
  
  return {
    learners,
    ratings,
    types,
    weekly
  };
}

/**
 * Generate AI-powered insights
 */
function generateAIInsights(metrics, charts) {
  const insights = [];
  
  // Completion rate insights
  if (metrics.completionRate > 80) {
    insights.push("Excellent completion rate! Your team is highly engaged with learning activities.");
  } else if (metrics.completionRate < 50) {
    insights.push("Low completion rate detected. Consider reviewing activity difficulty or engagement strategies.");
  }
  
  // Learning hours insights
  if (metrics.totalHours > 100) {
    insights.push("High learning volume indicates strong commitment to professional development.");
  }
  
  // Rating insights
  if (metrics.avgRating > 4.0) {
    insights.push("High satisfaction ratings suggest well-curated learning content.");
  } else if (metrics.avgRating < 3.0) {
    insights.push("Consider reviewing learning materials quality based on low ratings.");
  }
  
  // Learning type insights
  const topType = charts.types.sort((a, b) => b.value - a.value)[0];
  if (topType) {
    insights.push(`${topType.label} is the most popular learning format with ${topType.value} activities.`);
  }
  
  // Engagement insights
  if (metrics.activeLearners > 10) {
    insights.push("Strong team participation across multiple departments.");
  }
  
  // Add default insights if none generated
  if (insights.length === 0) {
    insights.push("Continue tracking learning activities to generate more detailed insights.");
    insights.push("Encourage team members to rate their learning experiences for better analytics.");
  }
  
  return insights;
}

/**
 * Get filter options for analytics
 */
function getAnalyticsFilters(learningData, membersData) {
  const departments = new Set();
  const types = new Set();
  
  learningData.forEach(row => {
    const memberId = row[1];
    const dept = membersData[memberId]?.department;
    if (dept) departments.add(dept);
    
    const type = row[2];
    if (type) types.add(type);
  });
  
  return {
    departments: Array.from(departments).sort(),
    types: Array.from(types).sort()
  };
}

function showCompetencyReports() {
  const html = HtmlService.createTemplateFromFile('competency-reports')
    .evaluate()
    .setWidth(1600)
    .setHeight(1000)
    .setTitle('Competency Reports');
    
  SpreadsheetApp.getUi().showModalDialog(html, 'Competency Reports');
}

/**
 * Get skill categories for competency reports
 */
function getSkillCategories() {
  return SYSTEM_CONFIG.ABILITY_CATEGORIES;
}

/**
 * Generate comprehensive competency report
 */
function generateCompetencyReport(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const assessmentsSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.ABILITY_ASSESSMENTS);
  const membersSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.MEMBERS);
  
  if (!assessmentsSheet || assessmentsSheet.getLastRow() <= 1) {
    return null;
  }
  
  // Get data
  const assessmentsData = assessmentsSheet.getRange(2, 1, assessmentsSheet.getLastRow() - 1, assessmentsSheet.getLastColumn()).getValues();
  const membersData = getMembersDataForAnalytics(membersSheet);
  
  // Filter by focus area if specified
  const filteredAssessments = params.focusArea ? 
    assessmentsData.filter(row => row[2] === params.focusArea) : 
    assessmentsData;
  
  // Generate matrix data
  const matrix = generateCompetencyMatrix(filteredAssessments, membersData, params);
  
  // Generate radar data
  const radar = generateSkillsRadarData(filteredAssessments, membersData);
  
  // Analyze skill gaps
  const gaps = analyzeSkillGaps(filteredAssessments, membersData, params.minLevel);
  
  // Generate recommendations
  const recommendations = generateCompetencyRecommendations(gaps, matrix);
  
  return {
    matrix: matrix,
    radar: radar,
    gaps: gaps,
    recommendations: recommendations
  };
}

/**
 * Generate competency matrix for heat map visualization
 */
function generateCompetencyMatrix(assessments, membersData, params) {
  const memberSkills = {};
  const allSkills = new Set();
  
  // Organize data by member and skill
  assessments.forEach(row => {
    const memberId = row[1];
    const skill = row[3];
    const level = parseInt(row[4]) || 0;
    
    if (level >= params.minLevel && membersData[memberId]) {
      if (!memberSkills[memberId]) {
        memberSkills[memberId] = {};
      }
      memberSkills[memberId][skill] = Math.max(memberSkills[memberId][skill] || 0, level);
      allSkills.add(skill);
    }
  });
  
  // Convert to matrix format
  const members = Object.keys(memberSkills).map(memberId => ({
    id: memberId,
    name: membersData[memberId]?.name || 'Unknown',
    skills: memberSkills[memberId]
  }));
  
  const skills = Array.from(allSkills).sort();
  
  return {
    members: members,
    skills: skills
  };
}

/**
 * Generate skills radar data for categories
 */
function generateSkillsRadarData(assessments, membersData) {
  const categoryData = {};
  
  // Group by category
  assessments.forEach(row => {
    const category = row[2] || 'General';
    const memberId = row[1];
    const level = parseInt(row[4]) || 0;
    
    if (membersData[memberId]) {
      if (!categoryData[category]) {
        categoryData[category] = [];
      }
      categoryData[category].push(level);
    }
  });
  
  // Calculate averages
  return Object.entries(categoryData).map(([category, levels]) => ({
    category: category,
    avgLevel: levels.reduce((sum, level) => sum + level, 0) / levels.length,
    memberCount: new Set(levels).size
  }));
}

/**
 * Analyze skill gaps in the team
 */
function analyzeSkillGaps(assessments, membersData, minLevel) {
  const skillLevels = {};
  
  // Collect all skill levels
  assessments.forEach(row => {
    const skill = row[3];
    const level = parseInt(row[4]) || 0;
    const memberId = row[1];
    
    if (membersData[memberId]) {
      if (!skillLevels[skill]) {
        skillLevels[skill] = [];
      }
      skillLevels[skill].push(level);
    }
  });
  
  // Identify gaps
  const gaps = [];
  Object.entries(skillLevels).forEach(([skill, levels]) => {
    const avgLevel = levels.reduce((sum, level) => sum + level, 0) / levels.length;
    const maxLevel = Math.max(...levels);
    const coverage = levels.filter(level => level >= minLevel).length / levels.length;
    
    // Calculate severity (0-10 scale)
    let severity = 0;
    if (avgLevel < 3) severity += 3;
    if (maxLevel < 7) severity += 3;
    if (coverage < 0.5) severity += 4;
    
    if (severity > 0) {
      gaps.push({
        skill: skill,
        avgLevel: avgLevel,
        maxLevel: maxLevel,
        coverage: coverage,
        severity: severity
      });
    }
  });
  
  return gaps.sort((a, b) => b.severity - a.severity).slice(0, 10); // Top 10 gaps
}

/**
 * Generate AI-powered competency recommendations
 */
function generateCompetencyRecommendations(gaps, matrix) {
  const recommendations = [];
  
  // Training recommendations based on gaps
  gaps.forEach(gap => {
    if (gap.severity >= 8) {
      recommendations.push({
        priority: 'HIGH',
        text: `Urgent: Organize intensive training for ${gap.skill}. Only ${(gap.coverage * 100).toFixed(0)}% of team meets minimum standards.`
      });
    } else if (gap.severity >= 6) {
      recommendations.push({
        priority: 'MED',
        text: `Schedule skill development sessions for ${gap.skill}. Current team average is ${gap.avgLevel.toFixed(1)}/10.`
      });
    }
  });
  
  // Mentorship recommendations
  const expertMembers = matrix.members.filter(member => {
    const avgSkill = Object.values(member.skills).reduce((sum, level) => sum + level, 0) / Object.keys(member.skills).length;
    return avgSkill >= 7;
  });
  
  if (expertMembers.length > 0) {
    recommendations.push({
      priority: 'LOW',
      text: `Leverage expertise of ${expertMembers.map(m => m.name).join(', ')} for peer mentoring programs.`
    });
  }
  
  // Cross-training recommendations
  if (matrix.skills.length > 5) {
    recommendations.push({
      priority: 'MED',
      text: 'Implement cross-training program to reduce single points of failure in critical skills.'
    });
  }
  
  // General recommendations
  if (recommendations.length === 0) {
    recommendations.push({
      priority: 'LOW',
      text: 'Team competency levels are generally healthy. Continue regular skill assessments.'
    });
  }
  
  return recommendations.slice(0, 8); // Limit to 8 recommendations
}

/**
 * Export competency report in various formats
 */
function exportCompetencyReport(params) {
  try {
    // This would typically generate actual files
    // For now, we'll simulate the export process
    
    const reportData = generateCompetencyReport(params);
    
    if (!reportData) {
      return { success: false, message: 'No data available for export' };
    }
    
    // Simulate export processing
    Utilities.sleep(1000); // Simulate processing time
    
    return { 
      success: true, 
      message: `${params.format.toUpperCase()} report generated with ${reportData.matrix.members.length} members and ${reportData.matrix.skills.length} skills analyzed.`
    };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function showEventCalendar() {
  console.log('showEventCalendar: Starting pre-load of calendar data');
  const currentDate = new Date();
  const calendarData = getCalendarData(currentDate.getFullYear(), currentDate.getMonth() + 1);
  console.log('showEventCalendar: Pre-loaded data:', calendarData);
  
  const html = HtmlService.createTemplateFromFile('event-calendar');
  html.calendarData = calendarData; // Pass data to template
  const template = html.evaluate()
    .setWidth(1600)
    .setHeight(900)
    .setTitle('Event Calendar');
    
  SpreadsheetApp.getUi().showModalDialog(template, 'Event Calendar');
}

/**
 * Get calendar data for a specific month
 */
function getCalendarData(year, month) {
  console.log('=== GETTING CALENDAR DATA ===');
  console.log('getCalendarData: Year:', year, 'Month:', month);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eventsSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.EVENTS);
  const membersSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.MEMBERS);
  
  console.log('getCalendarData: Looking for sheet:', SYSTEM_CONFIG.SHEET_NAMES.EVENTS);
  console.log('getCalendarData: Events sheet found:', !!eventsSheet);
  console.log('getCalendarData: Members sheet found:', !!membersSheet);
  
  if (!eventsSheet || eventsSheet.getLastRow() <= 1) {
    console.log('getCalendarData: No events sheet or no data, returning empty result');
    return { events: [] };
  }
  
  // Get member names mapping
  const memberNames = {};
  if (membersSheet && membersSheet.getLastRow() > 1) {
    const memberData = membersSheet.getRange(2, 1, membersSheet.getLastRow() - 1, 2).getValues();
    memberData.forEach(row => {
      memberNames[row[0]] = row[1] || row[0];
    });
    console.log('getCalendarData: Member names mapping:', memberNames);
  }
  
  // Get events data
  const lastRow = eventsSheet.getLastRow();
  const lastColumn = eventsSheet.getLastColumn();
  console.log('getCalendarData: Events sheet dimensions - Last row:', lastRow, 'Last column:', lastColumn);
  
  const data = eventsSheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
  console.log('getCalendarData: Retrieved', data.length, 'rows with', data[0]?.length || 0, 'columns');
  
  // Filter events for the requested month (and adjacent months for display)
  const startDate = new Date(year, month - 1, 1); // Previous month
  const endDate = new Date(year, month + 2, 0); // Next month end
  
  console.log('getCalendarData: Date range - Start:', startDate, 'End:', endDate);
  
  const events = [];
  data.forEach((row, index) => {
    try {
      const eventDate = new Date(row[4]); // Date column
      console.log(`getCalendarData: Processing event ${index + 1}: ${row[1] || 'Untitled'} (${row[0]}) - Date: ${eventDate}`);
      
      if (eventDate >= startDate && eventDate <= endDate) {
        const organizerId = row[7];
        const organizerName = memberNames[organizerId] || organizerId || 'Unknown';
        
        const event = {
          id: row[0] || '',
          title: row[1] || 'Untitled Event',
          type: (row[2] || 'other').toLowerCase(),
          description: row[3] || '',
          date: row[4],
          time: row[5] || '',
          location: row[6] || '',
          organizer: organizerName,
          attendees: row[8] || '',
          maxCapacity: parseInt(row[9]) || 0,
          status: row[10] || 'Active'
        };
        
        events.push(event);
        console.log(`getCalendarData: Added event: ${event.title} (${event.type}) on ${eventDate}`);
      }
    } catch (error) {
      console.log(`getCalendarData: Error processing row ${index + 1}:`, error.message);
    }
  });
  
  console.log('getCalendarData: Returning', events.length, 'events');
  return { events: events };
}

/**
 * Add a quick event from the calendar interface
 */
function addQuickEvent(eventData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const eventsSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.EVENTS);
    
    if (!eventsSheet) {
      throw new Error('Events sheet not found');
    }
    
    const eventId = 'EVT' + Date.now();
    const currentUser = Session.getActiveUser().getEmail();
    
    const rowData = [
      eventId,
      eventData.title,
      eventData.type,
      eventData.description,
      new Date(eventData.date),
      eventData.time,
      eventData.location,
      currentUser, // Organizer (current user)
      '', // Attendees (empty for quick add)
      0, // Max capacity (0 for quick add)
      'Active'
    ];
    
    eventsSheet.appendRow(rowData);
    
    return {
      success: true,
      message: 'Event added successfully',
      eventId: eventId
    };
  } catch (error) {
    return {
      success: false,
      message: 'Error adding event: ' + error.message
    };
  }
}

function showRSVPManagement() {
  console.log('showRSVPManagement: Starting pre-load of RSVP data');
  const rsvpData = getRSVPData();
  console.log('showRSVPManagement: Pre-loaded data:', rsvpData);

  const html = HtmlService.createTemplateFromFile('rsvp-management');
  html.rsvpData = rsvpData; // Pass data to template
  const template = html.evaluate()
    .setWidth(1600)
    .setHeight(1000)
    .setTitle('RSVP Management');
    
  SpreadsheetApp.getUi().showModalDialog(template, 'RSVP Management');
}

/**
 * Get RSVP data for all events with attendee tracking
 */
function getRSVPData() {
  console.log('=== GETTING RSVP DATA ===');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eventsSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.EVENTS);
  const membersSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.MEMBERS);
  
  console.log('getRSVPData: Looking for sheet:', SYSTEM_CONFIG.SHEET_NAMES.EVENTS);
  console.log('getRSVPData: Events sheet found:', !!eventsSheet);
  console.log('getRSVPData: Members sheet found:', !!membersSheet);
  
  if (!eventsSheet || eventsSheet.getLastRow() <= 1) {
    console.log('getRSVPData: No events sheet or no data, returning empty result');
    return { events: [], summary: { totalEvents: 0, totalAttendees: 0, avgAttendance: 0, pendingRSVPs: 0 } };
  }
  
  // Get member names mapping
  const memberNames = {};
  if (membersSheet && membersSheet.getLastRow() > 1) {
    const memberData = membersSheet.getRange(2, 1, membersSheet.getLastRow() - 1, 2).getValues();
    memberData.forEach(row => {
      memberNames[row[0]] = row[1] || row[0];
    });
    console.log('getRSVPData: Member names mapping:', memberNames);
  }
  
  // Get events data
  const lastRow = eventsSheet.getLastRow();
  const lastColumn = eventsSheet.getLastColumn();
  console.log('getRSVPData: Events sheet dimensions - Last row:', lastRow, 'Last column:', lastColumn);
  
  const data = eventsSheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
  console.log('getRSVPData: Retrieved', data.length, 'rows with', data[0]?.length || 0, 'columns');
  
  const events = [];
  let totalAttendees = 0;
  let totalConfirmed = 0;
  let totalPending = 0;
  
  data.forEach((row, index) => {
    try {
      const eventId = row[0];
      const title = row[1] || 'Untitled Event';
      const type = row[2] || 'other';
      const description = row[3] || '';
      const date = row[4];
      const time = row[5] || '';
      const location = row[6] || '';
      const organizerId = row[7];
      const attendeesStr = row[8] || '';
      const maxCapacity = parseInt(row[9]) || 0;
      const status = row[10] || 'Active';
      
      console.log(`getRSVPData: Processing event ${index + 1}: ${title} (${eventId}) - Date: ${date}`);
    
      // Parse attendees string (format: "member1:confirmed,member2:pending,member3:declined")
      const attendees = [];
      let confirmedCount = 0;
      let pendingCount = 0;
      let declinedCount = 0;
      
      console.log(`getRSVPData: Attendees string for "${title}": "${attendeesStr}"`);
      
      if (attendeesStr) {
        const attendeeEntries = attendeesStr.split(',');
        console.log(`getRSVPData: Parsed ${attendeeEntries.length} attendee entries`);
        
        attendeeEntries.forEach((entry, entryIndex) => {
          const [memberId, rsvpStatus] = entry.split(':');
          if (memberId && rsvpStatus) {
            const memberName = memberNames[memberId.trim()] || memberId.trim();
            const status = rsvpStatus.trim().toLowerCase();
            
            attendees.push({
              id: memberId.trim(),
              name: memberName,
              status: status
            });
            
            if (status === 'confirmed') confirmedCount++;
            else if (status === 'pending') pendingCount++;
            else if (status === 'declined') declinedCount++;
            
            console.log(`getRSVPData: Attendee ${entryIndex + 1}: ${memberName} (${memberId.trim()}) - ${status}`);
          }
        });
      }
      
      console.log(`getRSVPData: Event "${title}" - Confirmed: ${confirmedCount}, Pending: ${pendingCount}, Declined: ${declinedCount}`);
    
      totalAttendees += confirmedCount + pendingCount + declinedCount;
      totalConfirmed += confirmedCount;
      totalPending += pendingCount;
      
      events.push({
        id: eventId,
        title: title,
        type: type,
        description: description,
        date: date,
        time: time,
        location: location,
        organizer: memberNames[organizerId] || organizerId || 'Unknown',
        attendees: attendees,
        confirmedCount: confirmedCount,
        pendingCount: pendingCount,
        declinedCount: declinedCount,
        maxCapacity: maxCapacity,
        status: status
      });
      
      console.log(`getRSVPData: Added event: ${title} with ${attendees.length} attendees`);
    } catch (error) {
      console.log(`getRSVPData: Error processing row ${index + 1}:`, error.message);
    }
  });
  
  // Calculate summary statistics
  const avgAttendance = events.length > 0 ? (totalConfirmed / Math.max(totalAttendees, 1)) * 100 : 0;
  
  const summary = {
    totalEvents: events.length,
    totalAttendees: totalAttendees,
    avgAttendance: avgAttendance,
    pendingRSVPs: totalPending
  };
  
  console.log('getRSVPData: Returning', events.length, 'events with summary:', summary);
  return {
    events: events,
    summary: summary
  };
}

/**
 * Export attendee list for a specific event
 */
function exportEventAttendees(eventId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const eventsSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.EVENTS);
    
    if (!eventsSheet) {
      throw new Error('Events sheet not found');
    }
    
    // Find the event
    const data = eventsSheet.getRange(2, 1, eventsSheet.getLastRow() - 1, eventsSheet.getLastColumn()).getValues();
    const eventRow = data.find(row => row[0] === eventId);
    
    if (!eventRow) {
      throw new Error('Event not found');
    }
    
    // This would typically create a new sheet or export to a file
    // For now, we'll simulate the export process
    const attendeesStr = eventRow[8] || '';
    const attendeeCount = attendeesStr ? attendeesStr.split(',').length : 0;
    
    return {
      success: true,
      message: `Exported ${attendeeCount} attendee records for "${eventRow[1]}"`
    };
  } catch (error) {
    return {
      success: false,
      message: 'Error exporting attendees: ' + error.message
    };
  }
}

function showSystemSettings() {
  const html = HtmlService.createTemplateFromFile('system-settings')
    .evaluate()
    .setWidth(1400)
    .setHeight(1000)
    .setTitle('System Settings');
    
  SpreadsheetApp.getUi().showModalDialog(html, 'System Settings');
}

/**
 * Get current system settings
 */
function getSystemSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // This would typically read from a settings sheet or properties service
  // For now, return default settings
  return {
    cyberpunkTheme: true,
    animationSpeed: 'normal',
    language: 'en',
    autoRefresh: 5,
    itemsPerPage: 25,
    showAdvancedMetrics: true,
    twoFactorAuth: false,
    sessionTimeout: 60,
    passwordPolicy: 'medium',
    activityLogging: true,
    logRetention: 90,
    newEventAlerts: true,
    dailyDigest: false,
    systemAlerts: true,
    eventReminders: 24,
    deadlineWarnings: 3,
    autoBackup: true,
    backupRetention: 7,
    autoArchive: 365,
    cleanTempFiles: true
  };
}

/**
 * Save system settings
 */
function saveSystemSettings(settings) {
  try {
    // This would typically save to PropertiesService or a settings sheet
    const properties = PropertiesService.getDocumentProperties();
    
    Object.keys(settings).forEach(key => {
      properties.setProperty('setting_' + key, JSON.stringify(settings[key]));
    });
    
    return {
      success: true,
      message: 'Settings saved successfully'
    };
  } catch (error) {
    return {
      success: false,
      message: 'Error saving settings: ' + error.message
    };
  }
}

/**
 * Reset system settings to defaults
 */
function resetSystemSettings() {
  try {
    const properties = PropertiesService.getDocumentProperties();
    
    // Clear all setting properties
    const allProperties = properties.getProperties();
    Object.keys(allProperties).forEach(key => {
      if (key.startsWith('setting_')) {
        properties.deleteProperty(key);
      }
    });
    
    return {
      success: true,
      message: 'Settings reset to defaults'
    };
  } catch (error) {
    return {
      success: false,
      message: 'Error resetting settings: ' + error.message
    };
  }
}

/**
 * Test function to debug project hiring data
 */
function testProjectHiringData() {
  console.log('=== TESTING PROJECT HIRING DATA ===');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const projectsSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.PROJECTS);
  
  if (!projectsSheet) {
    console.log('❌ Projects sheet not found');
    return;
  }
  
  console.log('✅ Projects sheet found');
  console.log('Last row:', projectsSheet.getLastRow());
  console.log('Last column:', projectsSheet.getLastColumn());
  
  if (projectsSheet.getLastRow() <= 1) {
    console.log('❌ No project data (only header row)');
    return;
  }
  
  // Get all project data
  const data = projectsSheet.getRange(2, 1, projectsSheet.getLastRow() - 1, projectsSheet.getLastColumn()).getValues();
  console.log('📊 Retrieved', data.length, 'project rows');
  
  // Check each project's status
  data.forEach((row, index) => {
    const projectId = row[0] || 'No ID';
    const projectTitle = row[1] || 'No Title';
    const projectStatus = row[3] || 'No Status';
    console.log(`Project ${index + 1}: ${projectTitle} (${projectId}) - Status: "${projectStatus}"`);
  });
  
  // Check what statuses are being filtered for
  const hiringStatuses = ['Active', 'Planning', 'Recruiting'];
  console.log('🔍 Filtering for statuses:', hiringStatuses);
  
  const hiringProjects = data.filter(row => {
    const status = row[3] || '';
    const isHiring = hiringStatuses.includes(status);
    console.log(`Status "${status}" is hiring: ${isHiring}`);
    return isHiring;
  });
  
  console.log('✅ Found', hiringProjects.length, 'hiring projects');
  console.log('=== END TEST ===');
}

/**
 * Add comprehensive dropdown menus for all ENUM type data in the system
 */
function addAllEnumDropdowns() {
  console.log('=== ADDING ALL ENUM DROPDOWNS ===');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Define all ENUM configurations
  const enumConfigs = {
    // Status dropdowns
    PROJECTS: {
      column: 4, // Status column
      options: ['Planning', 'Active', 'Recruiting', 'Completed', 'On Hold', 'Cancelled'],
      description: 'Project Status'
    },
    MEMBERS: {
      column: 7, // Status column (assuming 7th column is status)
      options: ['Active', 'Inactive', 'Graduated', 'On Leave'],
      description: 'Member Status'
    },
    ACADEMIC_IDEAS: {
      column: 10, // Status column
      options: ['Under Review', 'Approved', 'Rejected', 'Implemented'],
      description: 'Idea Status'
    },
    EVENTS: {
      column: 11, // Status column
      options: ['Scheduled', 'In Progress', 'Completed', 'Cancelled'],
      description: 'Event Status'
    },
    LEARNING_ACTIVITIES: {
      column: 8, // Status column
      options: ['Not Started', 'In Progress', 'Completed', 'On Hold'],
      description: 'Learning Status'
    },
    
    // Type/Category dropdowns
    ABILITY_ASSESSMENTS: {
      column: 3, // Category column
      options: SYSTEM_CONFIG.ABILITY_CATEGORIES,
      description: 'Ability Category'
    },
    LEARNING_ACTIVITIES_TYPE: {
      column: 2, // Type column
      options: SYSTEM_CONFIG.LEARNING_TYPES,
      description: 'Learning Type'
    },
    
    // Role dropdowns
    MEMBERS_ROLE: {
      column: 4, // Role column
      options: ['Professor', 'PhD Student', 'Faculty', 'Undergraduate', 'Graduate Student', 'Research Assistant', 'Admin'],
      description: 'Member Role'
    }
  };
  
  // Apply dropdowns to each sheet
  Object.keys(enumConfigs).forEach(sheetKey => {
    const config = enumConfigs[sheetKey];
    const sheetName = SYSTEM_CONFIG.SHEET_NAMES[sheetKey] || sheetKey;
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      console.log(`❌ Sheet ${sheetName} not found`);
      return;
    }
    
    console.log(`✅ Processing ${sheetName} sheet`);
    console.log(`📋 Applying ${config.description} dropdown to column ${config.column}`);
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      console.log(`⚠️ No data in ${sheetName} sheet`);
      return;
    }
    
    try {
      // Create data validation rule
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(config.options, true)
        .setAllowInvalid(false)
        .setHelpText(`Select a ${config.description.toLowerCase()} from the dropdown`)
        .build();
      
      // Apply to the specified column (skip header row)
      const targetRange = sheet.getRange(2, config.column, lastRow - 1, 1);
      targetRange.setDataValidation(rule);
      
      console.log(`✅ Applied ${config.description} dropdown to ${sheetName} (${lastRow - 1} rows)`);
    } catch (error) {
      console.log(`❌ Error applying dropdown to ${sheetName}: ${error.message}`);
    }
  });
  
  // Special handling for Learning Activities - apply both type and status dropdowns
  const learningSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.LEARNING_ACTIVITIES);
  if (learningSheet) {
    console.log('✅ Applying additional dropdowns to Learning Activities sheet');
    
    // Apply type dropdown (column 2)
    const typeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(SYSTEM_CONFIG.LEARNING_TYPES, true)
      .setAllowInvalid(false)
      .setHelpText('Select a learning type from the dropdown')
      .build();
    
    const lastRow = learningSheet.getLastRow();
    if (lastRow > 1) {
      learningSheet.getRange(2, 2, lastRow - 1, 1).setDataValidation(typeRule);
      console.log('✅ Applied Learning Type dropdown');
    }
  }
  
  console.log('=== ALL ENUM DROPDOWNS COMPLETE ===');
}

/**
 * Test function for debugging ideas repository data
 */
function testIdeasRepositoryData() {
  console.log('=== TESTING IDEAS REPOSITORY DATA ===');
  const result = getIdeasRepositoryData();
  console.log('testIdeasRepositoryData result:', result);
  console.log('testIdeasRepositoryData result type:', typeof result);
  console.log('testIdeasRepositoryData ideas count:', result.ideas.length);
  if (result.ideas.length > 0) {
    console.log('testIdeasRepositoryData first idea:', result.ideas[0]);
  }
  console.log('=== IDEAS REPOSITORY DATA TEST COMPLETE ===');
}

function testCalendarData() {
  console.log('=== TESTING CALENDAR DATA ===');
  const currentDate = new Date();
  const result = getCalendarData(currentDate.getFullYear(), currentDate.getMonth() + 1);
  console.log('testCalendarData result:', result);
  console.log('testCalendarData result type:', typeof result);
  console.log('testCalendarData events count:', result.events.length);
  if (result.events.length > 0) {
    console.log('testCalendarData first event:', result.events[0]);
  }
  console.log('=== CALENDAR DATA TEST COMPLETE ===');
}

function testRSVPData() {
  console.log('=== TESTING RSVP DATA ===');
  const result = getRSVPData();
  console.log('testRSVPData result:', result);
  console.log('testRSVPData result type:', typeof result);
  console.log('testRSVPData events count:', result.events.length);
  console.log('testRSVPData summary:', result.summary);
  if (result.events.length > 0) {
    console.log('testRSVPData first event:', result.events[0]);
  }
  console.log('=== RSVP DATA TEST COMPLETE ===');
}

/**
 * Add status dropdowns (legacy function for backward compatibility)
 */
function addStatusDropdowns() {
  console.log('=== ADDING STATUS DROPDOWNS ===');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const statusOptions = {
    PROJECTS: ['Planning', 'Active', 'Recruiting', 'Completed', 'On Hold', 'Cancelled'],
    MEMBERS: ['Active', 'Inactive', 'Graduated', 'On Leave'],
    ACADEMIC_IDEAS: ['Under Review', 'Approved', 'Rejected', 'Implemented'],
    EVENTS: ['Scheduled', 'In Progress', 'Completed', 'Cancelled'],
    LEARNING_ACTIVITIES: ['Not Started', 'In Progress', 'Completed', 'On Hold']
  };
  
  Object.keys(statusOptions).forEach(sheetName => {
    const sheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES[sheetName]);
    if (!sheet) {
      console.log(`❌ Sheet ${sheetName} not found`);
      return;
    }
    
    console.log(`✅ Processing ${sheetName} sheet`);
    let statusColumn = 4; // Default status column
    
    // Override column numbers for specific sheets
    if (sheetName === 'ACADEMIC_IDEAS') statusColumn = 10;
    if (sheetName === 'EVENTS') statusColumn = 11;
    if (sheetName === 'LEARNING_ACTIVITIES') statusColumn = 8;
    if (sheetName === 'MEMBERS') statusColumn = 7;
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      console.log(`⚠️ No data in ${sheetName} sheet`);
      return;
    }
    
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(statusOptions[sheetName], true)
      .setAllowInvalid(false)
      .setHelpText('Select a status from the dropdown')
      .build();
    
    const statusRange = sheet.getRange(2, statusColumn, lastRow - 1, 1);
    statusRange.setDataValidation(rule);
    console.log(`✅ Applied dropdown to ${sheetName} status column (${lastRow - 1} rows)`);
  });
  
  console.log('=== STATUS DROPDOWNS COMPLETE ===');
}

/**
 * Inspect sheet structure to verify column positions
 */
function inspectSheetStructures() {
  console.log('=== INSPECTING SHEET STRUCTURES ===');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  Object.keys(SYSTEM_CONFIG.SHEET_NAMES).forEach(sheetKey => {
    const sheetName = SYSTEM_CONFIG.SHEET_NAMES[sheetKey];
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      console.log(`❌ Sheet ${sheetName} not found`);
      return;
    }
    
    console.log(`\n📋 ${sheetName} Sheet Structure:`);
    console.log(`Last row: ${sheet.getLastRow()}`);
    console.log(`Last column: ${sheet.getLastColumn()}`);
    
    if (sheet.getLastRow() > 0) {
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      console.log('Headers:');
      headers.forEach((header, index) => {
        console.log(`  Column ${index + 1}: "${header}"`);
      });
      
      // Show sample data if available
      if (sheet.getLastRow() > 1) {
        console.log('Sample data (first 2 rows):');
        const sampleData = sheet.getRange(2, 1, Math.min(2, sheet.getLastRow() - 1), sheet.getLastColumn()).getValues();
        sampleData.forEach((row, rowIndex) => {
          console.log(`  Row ${rowIndex + 2}:`, row.slice(0, 5).map(cell => `"${cell}"`).join(' | '));
        });
      }
    }
  });
  
  console.log('=== SHEET STRUCTURE INSPECTION COMPLETE ===');
}

/**
 * Create sample project data for testing
 */
function createSampleProjectData() {
  console.log('=== CREATING SAMPLE PROJECT DATA ===');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const projectsSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEET_NAMES.PROJECTS);
  
  if (!projectsSheet) {
    console.log('❌ Projects sheet not found');
    return;
  }
  
  // Clear existing data (keep headers)
  const lastRow = projectsSheet.getLastRow();
  if (lastRow > 1) {
    projectsSheet.getRange(2, 1, lastRow - 1, projectsSheet.getLastColumn()).clear();
  }
  
  // Sample project data
  const sampleProjects = [
    ['PROJ001', 'AI Research Assistant Development', 'Developing an AI-powered research assistant for academic literature analysis', 'Active', 'MEM001', 'MEM001,MEM002', 'Python, Machine Learning, NLP', '2024-01-15', '2024-12-31', 50000, 75, 'https://github.com/xu-research/ai-assistant'],
    ['PROJ002', 'Data Visualization Platform', 'Creating interactive data visualization tools for research presentations', 'Recruiting', 'MEM002', 'MEM002', 'JavaScript, D3.js, React', '2024-03-01', '2024-08-31', 30000, 25, 'https://github.com/xu-research/viz-platform'],
    ['PROJ003', 'Academic Paper Management System', 'Building a system to organize and track academic papers and citations', 'Planning', 'MEM003', '', 'Python, Django, PostgreSQL', '2024-06-01', '2025-01-31', 40000, 0, 'https://github.com/xu-research/paper-mgmt'],
    ['PROJ004', 'Research Collaboration Network', 'Developing a platform for connecting researchers and facilitating collaborations', 'Active', 'MEM001', 'MEM001,MEM004', 'Node.js, MongoDB, React', '2024-02-01', '2024-11-30', 60000, 60, 'https://github.com/xu-research/collab-network'],
    ['PROJ005', 'Statistical Analysis Toolkit', 'Creating a comprehensive toolkit for statistical analysis in research', 'Recruiting', 'MEM003', 'MEM003', 'R, Python, Shiny', '2024-04-01', '2024-09-30', 35000, 30, 'https://github.com/xu-research/stats-toolkit']
  ];
  
  // Add sample data
  if (sampleProjects.length > 0) {
    const range = projectsSheet.getRange(2, 1, sampleProjects.length, sampleProjects[0].length);
    range.setValues(sampleProjects);
    console.log(`✅ Added ${sampleProjects.length} sample projects`);
    
    // Log the projects for verification
    sampleProjects.forEach((project, index) => {
      console.log(`Project ${index + 1}: ${project[1]} (${project[0]}) - Status: ${project[3]}`);
    });
  }
  
  console.log('=== SAMPLE PROJECT DATA CREATED ===');
}
