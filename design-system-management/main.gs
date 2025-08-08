/**
 * Design System Management Tool for HRIS SaaS Company
 * 
 * This Google Apps Script manages the communication protocol and mapping between:
 * 1. Client's global design system variables
 * 2. Company's Figma global variables  
 * 3. Developer-generated CSS global variables
 * 
 * Features:
 * - Multi-tier mapping system
 * - Management reporting dashboard
 * - Communication protocol automation
 * - Client style guide analysis
 * - Conflict detection and resolution
 */

// Global Constants
const SHEET_NAMES = {
  DASHBOARD: 'Management Dashboard',
  CLIENT_VARIABLES: 'Client Design Variables',
  COMPANY_MAPPING: 'Company-Client Mapping',
  CSS_VARIABLES: 'CSS Variable Generation',
  COMMUNICATION_LOG: 'Communication Log',
  STYLE_CONFLICTS: 'Style Conflicts',
  PROJECT_TRACKER: 'Project Tracker',
  TEMPLATES: 'Communication Templates'
};

const PORTAL_TYPES = {
  HIRING_MANAGER: 'Hiring Manager',
  CAREERS: 'Careers',
  EMPLOYEE: 'Employee Portal',
  ADMIN: 'Admin Portal',
  REPORTS: 'Reports Portal'
};

const DESIGN_CATEGORIES = {
  COLORS: 'Colors',
  TYPOGRAPHY: 'Typography', 
  SPACING: 'Spacing',
  COMPONENTS: 'Components',
  LAYOUT: 'Layout',
  ANIMATIONS: 'Animations'
};

/**
 * Initialize the spreadsheet with all necessary sheets and structure
 */
function initializeDesignSystemTool() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create all necessary sheets
  Object.values(SHEET_NAMES).forEach(sheetName => {
    if (!ss.getSheetByName(sheetName)) {
      ss.insertSheet(sheetName);
    }
  });
  
  setupDashboardSheet();
  setupClientVariablesSheet();
  setupMappingSheet();
  setupCSSVariablesSheet();
  setupCommunicationLogSheet();
  setupConflictsSheet();
  setupProjectTrackerSheet();
  setupTemplatesSheet();
  
  Logger.log('Design System Management Tool initialized successfully');
}

/**
 * Setup Management Dashboard Sheet
 */
function setupDashboardSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.DASHBOARD);
  sheet.clear();
  
  // Header
  sheet.getRange('A1').setValue('Design System Management Dashboard');
  sheet.getRange('A1:H1').merge().setBackground('#2E7D32').setFontColor('white').setFontWeight('bold');
  
  // Key Metrics Section
  const metricsHeaders = [
    ['Metric', 'Count', 'Status', 'Last Updated'],
    ['Active Projects', '=COUNTA(\'Project Tracker\'!B:B)-1', 'Tracking', '=NOW()'],
    ['Pending Mappings', '=COUNTIF(\'Company-Client Mapping\'!D:D,"Pending")', 'Review Needed', '=NOW()'],
    ['Style Conflicts', '=COUNTA(\'Style Conflicts\'!A:A)-1', 'Action Required', '=NOW()'],
    ['Completed Variables', '=COUNTIF(\'CSS Variable Generation\'!E:E,"Generated")', 'On Track', '=NOW()']
  ];
  
  sheet.getRange('A3:D7').setValues(metricsHeaders);
  sheet.getRange('A3:D3').setBackground('#E8F5E8').setFontWeight('bold');
  
  // Recent Activity Section  
  sheet.getRange('A9').setValue('Recent Communication Activity');
  sheet.getRange('A9').setFontWeight('bold').setFontSize(14);
  
  const activityHeaders = [
    ['Date', 'Client', 'Type', 'Status', 'Next Action', 'Assigned To']
  ];
  sheet.getRange('A10:F10').setValues(activityHeaders);
  sheet.getRange('A10:F10').setBackground('#E3F2FD').setFontWeight('bold');
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, 8);
}

/**
 * Setup Client Design Variables Sheet
 */
function setupClientVariablesSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.CLIENT_VARIABLES);
  sheet.clear();
  
  const headers = [
    ['Client Name', 'Portal Type', 'Category', 'Variable Name', 'Value', 'Type', 'Usage Context', 'Priority', 'Status', 'Notes', 'Date Added']
  ];
  
  sheet.getRange('A1:K1').setValues(headers);
  sheet.getRange('A1:K1').setBackground('#1976D2').setFontColor('white').setFontWeight('bold');
  
  // Add data validation for dropdowns
  const clientRange = sheet.getRange('A2:A1000');
  const portalTypeRange = sheet.getRange('B2:B1000');
  const categoryRange = sheet.getRange('C2:C1000');
  const priorityRange = sheet.getRange('H2:H1000');
  const statusRange = sheet.getRange('I2:I1000');
  
  const portalTypeValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(Object.values(PORTAL_TYPES))
    .build();
  portalTypeRange.setDataValidation(portalTypeValidation);
  
  const categoryValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(Object.values(DESIGN_CATEGORIES))
    .build();
  categoryRange.setDataValidation(categoryValidation);
  
  const priorityValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['High', 'Medium', 'Low'])
    .build();
  priorityRange.setDataValidation(priorityValidation);
  
  const statusValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Received', 'Analyzing', 'Mapped', 'Approved', 'Implemented'])
    .build();
  statusRange.setDataValidation(statusValidation);
  
  sheet.autoResizeColumns(1, 11);
}

/**
 * Setup Company-Client Mapping Sheet
 */
function setupMappingSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.COMPANY_MAPPING);
  sheet.clear();
  
  const headers = [
    ['Client Variable', 'Company Figma Variable', 'CSS Variable Name', 'Mapping Status', 'Conflict Level', 'Resolution', 'Approved By', 'Implementation Date', 'Notes']
  ];
  
  sheet.getRange('A1:I1').setValues(headers);
  sheet.getRange('A1:I1').setBackground('#4CAF50').setFontColor('white').setFontWeight('bold');
  
  // Add conditional formatting for conflict levels
  const conflictRange = sheet.getRange('E2:E1000');
  const lowConflictRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Low')
    .setBackground('#C8E6C9')
    .setRanges([conflictRange])
    .build();
  const mediumConflictRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Medium')
    .setBackground('#FFE0B2')
    .setRanges([conflictRange])
    .build();
  const highConflictRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('High')
    .setBackground('#FFCDD2')
    .setRanges([conflictRange])
    .build();
  
  sheet.setConditionalFormatRules([lowConflictRule, mediumConflictRule, highConflictRule]);
  
  sheet.autoResizeColumns(1, 9);
}

/**
 * Setup CSS Variables Generation Sheet
 */
function setupCSSVariablesSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.CSS_VARIABLES);
  sheet.clear();
  
  const headers = [
    ['Portal Type', 'Category', 'CSS Variable Name', 'Value', 'Generation Status', 'Command Used', 'File Path', 'Last Generated', 'Validation Status']
  ];
  
  sheet.getRange('A1:I1').setValues(headers);
  sheet.getRange('A1:I1').setBackground('#FF9800').setFontColor('white').setFontWeight('bold');
  
  // Add sample CSS variable generation patterns
  const sampleData = [
    ['Hiring Manager', 'Colors', '--t-gs-color-primary', '#1976D2', 'Generated', 'figma-tokens generate', '/styles/tokens.css', new Date(), 'Valid'],
    ['Careers', 'Typography', '--t-gs-font-family-primary', 'Inter, sans-serif', 'Generated', 'figma-tokens generate', '/styles/tokens.css', new Date(), 'Valid'],
    ['Employee Portal', 'Spacing', '--t-gs-spacing-medium', '16px', 'Pending', '', '', '', 'Pending']
  ];
  
  sheet.getRange('A2:I4').setValues(sampleData);
  sheet.autoResizeColumns(1, 9);
}

/**
 * Setup Communication Log Sheet
 */
function setupCommunicationLogSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.COMMUNICATION_LOG);
  sheet.clear();
  
  const headers = [
    ['Date', 'Client', 'Communication Type', 'Subject', 'Participants', 'Key Points', 'Action Items', 'Follow-up Date', 'Status', 'Attachments']
  ];
  
  sheet.getRange('A1:J1').setValues(headers);
  sheet.getRange('A1:J1').setBackground('#9C27B0').setFontColor('white').setFontWeight('bold');
  
  sheet.autoResizeColumns(1, 10);
}

/**
 * Setup Style Conflicts Sheet
 */
function setupConflictsSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.STYLE_CONFLICTS);
  sheet.clear();
  
  const headers = [
    ['Conflict ID', 'Client', 'Element Type', 'Client Specification', 'Company Standard', 'Conflict Description', 'Severity', 'Proposed Solution', 'Client Response', 'Resolution Status', 'Date Resolved']
  ];
  
  sheet.getRange('A1:K1').setValues(headers);
  sheet.getRange('A1:K1').setBackground('#F44336').setFontColor('white').setFontWeight('bold');
  
  sheet.autoResizeColumns(1, 11);
}

/**
 * Setup Project Tracker Sheet
 */
function setupProjectTrackerSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.PROJECT_TRACKER);
  sheet.clear();
  
  const headers = [
    ['Project ID', 'Client Name', 'Portal Types', 'Start Date', 'Expected Completion', 'Current Phase', 'Progress %', 'Team Members', 'Blockers', 'Next Milestone', 'Status']
  ];
  
  sheet.getRange('A1:K1').setValues(headers);
  sheet.getRange('A1:K1').setBackground('#607D8B').setFontColor('white').setFontWeight('bold');
  
  sheet.autoResizeColumns(1, 11);
}

/**
 * Setup Communication Templates Sheet
 */
function setupTemplatesSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.TEMPLATES);
  sheet.clear();
  
  const headers = [
    ['Template Name', 'Purpose', 'Template Content', 'Variables', 'Usage Instructions']
  ];
  
  sheet.getRange('A1:E1').setValues(headers);
  sheet.getRange('A1:E1').setBackground('#795548').setFontColor('white').setFontWeight('bold');
  
  // Add sample templates
  const templates = [
    [
      'Style Guide Request',
      'Initial client style guide collection',
      `Dear {{CLIENT_NAME}},

Thank you for choosing our HRIS portal development services. To ensure we deliver a solution that perfectly aligns with your brand identity, we need to collect your design system specifications.

Please provide the following:

1. **Global Variables**: Color palette, typography scales, spacing values
2. **Component Library**: Buttons, forms, navigation, modals, cards, tables
3. **Brand Guidelines**: Logo usage, imagery guidelines, tone of voice
4. **Portal-Specific Requirements**: Any unique styling for different user roles

**Delivery Format**: Please organize these in the following structure:
- Figma file with organized components and variables
- Style guide documentation (PDF or online)
- Asset library (logos, icons, images)

**Timeline**: We'll need these materials within {{TIMELINE}} to maintain project schedule.

Our design system mapping process will ensure consistency across all portal applications while maintaining your brand integrity.

Best regards,
{{TEAM_MEMBER_NAME}}`,
      '{{CLIENT_NAME}}, {{TIMELINE}}, {{TEAM_MEMBER_NAME}}',
      'Use when starting new client engagement'
    ],
    [
      'Style Conflict Resolution',
      'Address conflicts between client and company standards',
      `Dear {{CLIENT_NAME}},

We've identified some areas where your design requirements may conflict with our platform's technical constraints or UX best practices:

**Conflicts Identified:**
{{CONFLICT_LIST}}

**Our Recommendations:**
{{RECOMMENDATIONS}}

**Next Steps:**
1. Review the attached comparison document
2. Schedule a design review call within {{TIMEFRAME}}
3. Confirm your preferences for each identified conflict

We're committed to finding solutions that maintain your brand integrity while ensuring optimal user experience and technical feasibility.

Best regards,
{{TEAM_MEMBER_NAME}}`,
      '{{CLIENT_NAME}}, {{CONFLICT_LIST}}, {{RECOMMENDATIONS}}, {{TIMEFRAME}}, {{TEAM_MEMBER_NAME}}',
      'Use when conflicts are detected during mapping process'
    ]
  ];
  
  sheet.getRange('A2:E3').setValues(templates);
  sheet.autoResizeColumns(1, 5);
}

/**
 * Add new client design variable
 */
function addClientVariable(clientName, portalType, category, variableName, value, type, context, priority = 'Medium', notes = '') {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.CLIENT_VARIABLES);
  const newRow = [
    clientName,
    portalType,
    category,
    variableName,
    value,
    type,
    context,
    priority,
    'Received',
    notes,
    new Date()
  ];
  
  sheet.appendRow(newRow);
  
  // Trigger analysis for potential conflicts
  analyzeVariableConflicts(variableName, value, category, clientName);
  
  return `Added variable ${variableName} for ${clientName}`;
}

/**
 * Analyze potential conflicts between client variables and company standards
 */
function analyzeVariableConflicts(variableName, value, category, clientName) {
  // This would integrate with your company's design system API or database
  // For now, implementing basic conflict detection logic
  
  const conflictsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.STYLE_CONFLICTS);
  
  // Sample conflict detection (you'd expand this based on your actual standards)
  const conflicts = [];
  
  if (category === 'Colors' && !isValidColorFormat(value)) {
    conflicts.push('Invalid color format detected');
  }
  
  if (category === 'Typography' && !isStandardFontFamily(value)) {
    conflicts.push('Non-standard font family may impact performance');
  }
  
  if (conflicts.length > 0) {
    const conflictId = Utilities.getUuid();
    const newConflict = [
      conflictId,
      clientName,
      category,
      value,
      'Company Standard',
      conflicts.join('; '),
      'Medium',
      'Propose alternative or validate with client',
      'Pending',
      'Identified',
      ''
    ];
    
    conflictsSheet.appendRow(newConflict);
  }
}

/**
 * Generate CSS variables based on mapping
 */
function generateCSSVariables(portalType, clientName) {
  const mappingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.COMPANY_MAPPING);
  const cssSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.CSS_VARIABLES);
  
  const mappingData = mappingSheet.getDataRange().getValues();
  const cssVariables = [];
  
  for (let i = 1; i < mappingData.length; i++) {
    const [clientVar, companyVar, cssVarName, status] = mappingData[i];
    
    if (status === 'Approved' && cssVarName) {
      // Generate CSS variable entry
      const cssVarRow = [
        portalType,
        getCategoryFromVariable(cssVarName),
        cssVarName,
        getValueFromMapping(clientVar, companyVar),
        'Generated',
        'figma-tokens generate --platform css',
        `/styles/${portalType.toLowerCase().replace(' ', '-')}-tokens.css`,
        new Date(),
        'Valid'
      ];
      
      cssVariables.push(cssVarRow);
    }
  }
  
  // Append all generated variables
  if (cssVariables.length > 0) {
    const startRow = cssSheet.getLastRow() + 1;
    const range = cssSheet.getRange(startRow, 1, cssVariables.length, 9);
    range.setValues(cssVariables);
  }
  
  return `Generated ${cssVariables.length} CSS variables for ${portalType}`;
}

/**
 * Create communication log entry
 */
function logCommunication(clientName, type, subject, participants, keyPoints, actionItems, followUpDate = null) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.COMMUNICATION_LOG);
  
  const logEntry = [
    new Date(),
    clientName,
    type,
    subject,
    participants,
    keyPoints,
    actionItems,
    followUpDate || '',
    'Active',
    ''
  ];
  
  sheet.appendRow(logEntry);
  
  // Update dashboard
  updateDashboard();
  
  return 'Communication logged successfully';
}

/**
 * Send automated communication using templates
 */
function sendTemplatedCommunication(templateName, clientEmail, variables = {}) {
  const templatesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.TEMPLATES);
  const templateData = templatesSheet.getDataRange().getValues();
  
  let template = null;
  for (let i = 1; i < templateData.length; i++) {
    if (templateData[i][0] === templateName) {
      template = templateData[i];
      break;
    }
  }
  
  if (!template) {
    throw new Error(`Template ${templateName} not found`);
  }
  
  let content = template[2];
  
  // Replace variables in template
  Object.keys(variables).forEach(key => {
    const placeholder = `{{${key}}}`;
    content = content.replace(new RegExp(placeholder, 'g'), variables[key]);
  });
  
  // Send email (you'd integrate with your email system)
  try {
    MailApp.sendEmail({
      to: clientEmail,
      subject: template[1],
      body: content,
      htmlBody: content.replace(/\n/g, '<br>')
    });
    
    // Log the communication
    logCommunication(
      variables.CLIENT_NAME || 'Unknown',
      'Email',
      template[1],
      clientEmail,
      `Sent template: ${templateName}`,
      'Awaiting client response',
      new Date(Date.now() + 7 * 24 * 60 * 60 * 1000) // 7 days from now
    );
    
    return 'Template communication sent successfully';
  } catch (error) {
    Logger.log('Email sending failed: ' + error.message);
    return 'Failed to send communication: ' + error.message;
  }
}

/**
 * Update management dashboard with latest metrics
 */
function updateDashboard() {
  const dashboard = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.DASHBOARD);
  
  // Update last updated timestamp
  dashboard.getRange('D4:D7').setValue(new Date());
  
  // Get recent communication activities
  const commLog = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.COMMUNICATION_LOG);
  const commData = commLog.getDataRange().getValues();
  
  // Sort by date and get latest 10
  const recentComm = commData.slice(1)
    .sort((a, b) => new Date(b[0]) - new Date(a[0]))
    .slice(0, 10);
  
  if (recentComm.length > 0) {
    const startRow = 11;
    dashboard.getRange(startRow, 1, recentComm.length, 6).setValues(
      recentComm.map(row => [row[0], row[1], row[2], row[8], row[6], row[4]])
    );
  }
}

/**
 * Helper Functions
 */
function isValidColorFormat(value) {
  const hexPattern = /^#[0-9A-Fa-f]{6}$/;
  const rgbPattern = /^rgb\(\s*\d+\s*,\s*\d+\s*,\s*\d+\s*\)$/;
  const hslPattern = /^hsl\(\s*\d+\s*,\s*\d+%\s*,\s*\d+%\s*\)$/;
  
  return hexPattern.test(value) || rgbPattern.test(value) || hslPattern.test(value);
}

function isStandardFontFamily(value) {
  const standardFonts = ['Inter', 'Roboto', 'Open Sans', 'Lato', 'Source Sans Pro'];
  return standardFonts.some(font => value.includes(font));
}

function getCategoryFromVariable(cssVarName) {
  if (cssVarName.includes('color')) return DESIGN_CATEGORIES.COLORS;
  if (cssVarName.includes('font')) return DESIGN_CATEGORIES.TYPOGRAPHY;
  if (cssVarName.includes('spacing') || cssVarName.includes('margin') || cssVarName.includes('padding')) return DESIGN_CATEGORIES.SPACING;
  return DESIGN_CATEGORIES.COMPONENTS;
}

function getValueFromMapping(clientVar, companyVar) {
  // This would integrate with your actual design system
  // For now, return a placeholder
  return `var(${companyVar})`;
}

/**
 * Menu creation for easy access
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Design System Manager')
    .addItem('Initialize System', 'initializeDesignSystemTool')
    .addSeparator()
    .addItem('Add Client Variable', 'showAddVariableDialog')
    .addItem('Generate CSS Variables', 'showGenerateCSSDialog')
    .addItem('Log Communication', 'showCommunicationDialog')
    .addSeparator()
    .addItem('Update Dashboard', 'updateDashboard')
    .addItem('Export Variables', 'exportCSSVariables')
    .addToUi();
}

/**
 * Dialog functions for user interaction
 */
function showAddVariableDialog() {
  const html = HtmlService.createHtmlOutputFromFile('add-variable-dialog')
    .setWidth(400)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add Client Design Variable');
}

function showGenerateCSSDialog() {
  const html = HtmlService.createHtmlOutputFromFile('generate-css-dialog')
    .setWidth(400)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Generate CSS Variables');
}

function showCommunicationDialog() {
  const html = HtmlService.createHtmlOutputFromFile('communication-dialog')
    .setWidth(500)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Log Communication');
}

/**
 * Export CSS variables to file
 */
function exportCSSVariables(portalType = null) {
  const cssSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.CSS_VARIABLES);
  const data = cssSheet.getDataRange().getValues();
  
  let cssContent = ':root {\n';
  
  for (let i = 1; i < data.length; i++) {
    const [portal, category, varName, value, status] = data[i];
    
    if (status === 'Generated' && (!portalType || portal === portalType)) {
      cssContent += `  ${varName}: ${value};\n`;
    }
  }
  
  cssContent += '}\n';
  
  // Create a blob and download (in actual implementation, you'd integrate with your deployment system)
  const blob = Utilities.newBlob(cssContent, 'text/css', `design-variables${portalType ? '-' + portalType.toLowerCase().replace(' ', '-') : ''}.css`);
  
  // Store in Drive for download
  const file = DriveApp.createFile(blob);
  
  return {
    fileId: file.getId(),
    fileName: file.getName(),
    downloadUrl: `https://drive.google.com/file/d/${file.getId()}/view`
  };
}
