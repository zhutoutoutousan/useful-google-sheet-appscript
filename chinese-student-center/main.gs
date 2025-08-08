// ========================================
// 🏛️ X.LAB CHINESE STUDENT CENTER
// Advanced Google Apps Script Management System
// ========================================

/**
 * @OnlyCurrentDoc
 * 
 * Required OAuth Scopes:
 * - https://www.googleapis.com/auth/spreadsheets.currentonly
 * - https://www.googleapis.com/auth/script.external_request
 */
// One-stop platform for Chinese students at XU Exponential University
// Features: Job posting, job seeking, places to visit, events, useful websites, business opportunities

// ========================================
// 🚀 MAIN SETUP & INITIALIZATION
// ========================================

function setupChineseStudentCenter() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create all necessary sheets with innovative structure
  createSheetIfNotExists(ss, '🎯 Dashboard');
  createSheetIfNotExists(ss, '💼 Job Postings');
  createSheetIfNotExists(ss, '📝 Job Applications');
  createSheetIfNotExists(ss, '🔍 Job Seeking');
  createSheetIfNotExists(ss, '🗺️ Places & Activities');
  createSheetIfNotExists(ss, '📅 Events');
  createSheetIfNotExists(ss, '🌐 Useful Websites');
  createSheetIfNotExists(ss, '💡 Business Opportunities');
  createSheetIfNotExists(ss, '👥 Community');
  createSheetIfNotExists(ss, '📊 Analytics');
  createSheetIfNotExists(ss, '⚙️ Settings');
  
  // Initialize all modules
  setupJobPostings();
  setupJobApplications();
  setupJobSeeking();
  setupPlacesAndActivities();
  setupEvents();
  setupUsefulWebsites();
  setupBusinessOpportunities();
  setupCommunity();
  setupAnalytics();
  setupSettings();
  
  // Create innovative dashboard
  createInnovativeDashboard();
  
  SpreadsheetApp.getUi().alert('🎉 X.lab Chinese Student Center initialized!\n\nYour comprehensive platform is ready for Chinese students at XU Exponential University.');
}

function createSheetIfNotExists(ss, sheetName) {
  if (!ss.getSheetByName(sheetName)) {
    ss.insertSheet(sheetName);
  }
}

// ========================================
// 🎯 INNOVATIVE DASHBOARD
// ========================================

function createInnovativeDashboard() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('🎯 Dashboard');
  sheet.clear();
  
  // Calculate real statistics
  const stats = calculateQuickStats();
  
  // Create hexagonal grid layout with dynamic stats
  const dashboardData = [
    ['🏛️ X.LAB CHINESE STUDENT CENTER', '', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', '', ''],
    ['💼 JOB POSTINGS', '📝 JOB APPLICATIONS', '🗺️ PLACES & ACTIVITIES', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', '', ''],
    ['📅 EVENTS', '🌐 USEFUL WEBSITES', '💡 BUSINESS OPPORTUNITIES', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', '', ''],
    ['👥 COMMUNITY', '📊 ANALYTICS', '⚙️ SETTINGS', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', '', ''],
    ['📈 QUICK STATS', '', '', '', '', '', '', '', '', ''],
    [
      `Active Job Postings: ${stats.activeJobPostings}`, 
      `Total Applications: ${stats.totalApplications}`, 
      `Upcoming Events: ${stats.upcomingEvents}`, 
      `Community Members: ${stats.communityMembers}`, 
      '', '', '', '', '', ''
    ],
    ['', '', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', '', ''],
    ['🌍 LOCATION SERVICES', '', '', '', '', '', '', '', '', ''],
    ['Potsdam Map', 'Berlin Map', 'Transportation', 'Student Discounts', '', '', '', '', '', '']
  ];
  
  const range = sheet.getRange(1, 1, dashboardData.length, 10);
  range.setValues(dashboardData);
  
  // Apply innovative styling
  applyInnovativeStyling(sheet);
  
  // Add interactive elements
  addDashboardInteractivity(sheet);
}

function applyInnovativeStyling(sheet) {
  // Main title styling
  sheet.getRange('A1:J1').merge();
  sheet.getRange('A1').setFontSize(24).setFontWeight('bold').setBackground('#FF6B6B').setFontColor('white');
  
  // Module boxes styling
  const moduleRanges = ['A3:C3', 'A5:C5', 'A7:C7'];
  moduleRanges.forEach(range => {
    sheet.getRange(range).setBackground('#4ECDC4').setFontColor('white').setFontWeight('bold');
  });
  
  // Stats styling
  sheet.getRange('A11:D11').setBackground('#45B7D1').setFontColor('white');
  
  // Quick Actions removed from cells - now only in native menu
  
  // Location styling - make them look like clickable buttons
  sheet.getRange('A17:D17').setBackground('#FFEAA7').setFontColor('#2D3436').setFontWeight('bold');
  // Add borders to make them look more button-like
  sheet.getRange('A17:D17').setBorder(true, true, true, true, true, true);
  
  // Set column widths
  sheet.setColumnWidths(1, 10, 120);
  sheet.setRowHeights(1, 18, 30);
}

function addDashboardInteractivity(sheet) {
  // Add hyperlinks to modules based on their actual positions in the dashboard
  const moduleLinks = [
    {cell: 'A3', sheet: '💼 Job Postings', text: '💼 JOB POSTINGS'},
    {cell: 'B3', sheet: '📝 Job Applications', text: '📝 JOB APPLICATIONS'},
    {cell: 'C3', sheet: '🗺️ Places & Activities', text: '🗺️ PLACES & ACTIVITIES'},
    {cell: 'A5', sheet: '📅 Events', text: '📅 EVENTS'},
    {cell: 'B5', sheet: '🌐 Useful Websites', text: '🌐 USEFUL WEBSITES'},
    {cell: 'C5', sheet: '💡 Business Opportunities', text: '💡 BUSINESS OPPORTUNITIES'},
    {cell: 'A7', sheet: '👥 Community', text: '👥 COMMUNITY'},
    {cell: 'B7', sheet: '📊 Analytics', text: '📊 ANALYTICS'},
    {cell: 'C7', sheet: '⚙️ Settings', text: '⚙️ SETTINGS'}
  ];
  
  moduleLinks.forEach(link => {
    try {
      const cell = sheet.getRange(link.cell);
      cell.setFormula(`=HYPERLINK("#gid=${getSheetId(link.sheet)}", "${link.text}")`);
    } catch (error) {
      console.log(`Could not create hyperlink for ${link.sheet}: ${error}`);
    }
  });
  
  // Quick Actions removed from cells - now only available in native menu
  
  // Add hyperlinks for Location Services buttons
  const locationServiceLinks = [
    {cell: 'A17', url: 'https://maps.google.com/?q=Potsdam,Germany', text: 'Potsdam Map'},
    {cell: 'B17', url: 'https://maps.google.com/?q=Berlin,Germany', text: 'Berlin Map'},
    {cell: 'C17', url: 'https://www.bvg.de/en', text: 'Transportation'},
    {cell: 'D17', url: 'https://www.studentenwerk-berlin.de/english/', text: 'Student Discounts'}
  ];
  
  locationServiceLinks.forEach(link => {
    try {
      const cell = sheet.getRange(link.cell);
      cell.setFormula(`=HYPERLINK("${link.url}", "${link.text}")`);
    } catch (error) {
      console.log(`Could not create location service link for ${link.text}: ${error}`);
    }
  });
}

function getSheetId(sheetName) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    return sheet ? sheet.getSheetId() : '';
  } catch (error) {
    console.log(`Could not get sheet ID for ${sheetName}: ${error}`);
    return '';
  }
}

// ========================================
// 💼 JOB POSTINGS MODULE
// ========================================

function setupJobPostings() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('💼 Job Postings');
  sheet.clear();
  
  const headers = [
    'Job ID', 'Title', 'Company', 'Location', 'Type', 'Salary Range', 'Requirements',
    'Description', 'Contact Email', 'Contact Phone', 'Posted Date', 'Deadline',
    'Status', 'Tags', 'Benefits', 'Application Link', 'Created By', 'Applications Count'
  ];
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  
  // Apply styling
  headerRange.setFontWeight('bold').setBackground('#FF6B6B').setFontColor('white');
  
  // Smart data validation with dynamic options
  const typeOptions = ['Full-time', 'Part-time', 'Internship', 'Freelance', 'Remote', 'Contract', 'Temporary'];
  const statusOptions = ['Active', 'Closed', 'Expired', 'Filled', 'Under Review', 'Interviewing'];
  const locationOptions = ['Berlin', 'Potsdam', 'Munich', 'Hamburg', 'Frankfurt', 'Cologne', 'Remote', 'Hybrid'];
  const salaryRanges = ['€20k-30k', '€30k-40k', '€40k-50k', '€50k-60k', '€60k-70k', '€70k+', 'Negotiable', 'Competitive'];
  
  sheet.getRange(2, 5, 1000, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(typeOptions, true).build()
  );
  
  sheet.getRange(2, 4, 1000, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(locationOptions, true).build()
  );
  
  sheet.getRange(2, 6, 1000, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(salaryRanges, true).build()
  );
  
  sheet.getRange(2, 13, 1000, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(statusOptions, true).build()
  );
  
  // Date formatting
  sheet.getRange('K:L').setNumberFormat('mm/dd/yyyy');
  
  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1, headers.length, 150);
}

function addJobPosting() {
  const htmlOutput = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }
        .container { max-width: 600px; margin: 0 auto; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
        h2 { color: #2c3e50; text-align: center; margin-bottom: 30px; }
        .form-group { margin-bottom: 20px; }
        label { display: block; margin-bottom: 5px; font-weight: bold; color: #34495e; }
        input, select, textarea { width: 100%; padding: 12px; border: 2px solid #ddd; border-radius: 5px; font-size: 14px; box-sizing: border-box; }
        input:focus, select:focus, textarea:focus { border-color: #3498db; outline: none; }
        textarea { height: 80px; resize: vertical; }
        .row { display: flex; gap: 15px; }
        .col { flex: 1; }
        .btn { background: #3498db; color: white; padding: 12px 24px; border: none; border-radius: 5px; cursor: pointer; font-size: 16px; margin: 5px; }
        .btn:hover { background: #2980b9; }
        .btn-secondary { background: #95a5a6; }
        .btn-secondary:hover { background: #7f8c8d; }
        .required { color: #e74c3c; }
        .section { background: #ecf0f1; padding: 20px; border-radius: 8px; margin-bottom: 20px; }
        .section h3 { margin-top: 0; color: #2c3e50; }
      </style>
    </head>
    <body>
      <div class="container">
        <h2>💼 Add New Job Posting</h2>
        <form id="jobForm">
          <div class="section">
            <h3>📋 Basic Information</h3>
            <div class="form-group">
              <label for="title">Job Title <span class="required">*</span></label>
              <input type="text" id="title" name="title" required placeholder="e.g., Software Developer">
            </div>
            <div class="row">
              <div class="col">
                <label for="company">Company Name <span class="required">*</span></label>
                <input type="text" id="company" name="company" required placeholder="e.g., TechCorp">
              </div>
              <div class="col">
                <label for="location">Location <span class="required">*</span></label>
                <input type="text" id="location" name="location" required placeholder="e.g., Berlin">
              </div>
            </div>
            <div class="row">
              <div class="col">
                <label for="type">Job Type <span class="required">*</span></label>
                <select id="type" name="type" required>
                  <option value="">Select Type</option>
                  <option value="Full-time">Full-time</option>
                  <option value="Part-time">Part-time</option>
                  <option value="Contract">Contract</option>
                  <option value="Internship">Internship</option>
                  <option value="Freelance">Freelance</option>
                  <option value="Remote">Remote</option>
                </select>
              </div>
              <div class="col">
                <label for="salary">Salary Range</label>
                <input type="text" id="salary" name="salary" placeholder="e.g., €50k-60k">
              </div>
            </div>
          </div>

          <div class="section">
            <h3>📝 Job Details</h3>
            <div class="form-group">
              <label for="requirements">Requirements & Skills</label>
              <textarea id="requirements" name="requirements" placeholder="e.g., Bachelor's degree in Computer Science, 3+ years experience, JavaScript, React, Node.js"></textarea>
            </div>
            <div class="form-group">
              <label for="description">Job Description</label>
              <textarea id="description" name="description" placeholder="Describe the role, responsibilities, and what the company offers"></textarea>
            </div>
          </div>

          <div class="section">
            <h3>📞 Contact Information</h3>
            <div class="row">
              <div class="col">
                <label for="email">Contact Email</label>
                <input type="email" id="email" name="email" placeholder="e.g., hr@company.com">
              </div>
              <div class="col">
                <label for="phone">Contact Phone</label>
                <input type="tel" id="phone" name="phone" placeholder="e.g., +49 30 123456">
              </div>
            </div>
            <div class="form-group">
              <label for="deadline">Application Deadline</label>
              <input type="date" id="deadline" name="deadline">
            </div>
          </div>

          <div class="section">
            <h3>🏷️ Additional Information</h3>
            <div class="row">
              <div class="col">
                <label for="tags">Tags</label>
                <input type="text" id="tags" name="tags" placeholder="e.g., Remote, Startup, Health Insurance">
              </div>
              <div class="col">
                <label for="benefits">Benefits</label>
                <input type="text" id="benefits" name="benefits" placeholder="e.g., Health Insurance, Flexible Hours">
              </div>
            </div>
            <div class="form-group">
              <label for="applicationLink">Application Link</label>
              <input type="url" id="applicationLink" name="applicationLink" placeholder="e.g., https://apply.company.com">
            </div>
          </div>

          <div style="text-align: center; margin-top: 30px;">
            <button type="button" class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
            <button type="submit" class="btn">💼 Add Job Posting</button>
          </div>
        </form>
      </div>

      <script>
        document.getElementById('jobForm').addEventListener('submit', function(e) {
          e.preventDefault();
          
          const formData = new FormData(this);
          const data = {};
          for (let [key, value] of formData.entries()) {
            data[key] = value;
          }
          
          // Validate required fields
          if (!data.title || !data.company || !data.location || !data.type) {
            alert('Please fill in all required fields (marked with *)');
            return;
          }
          
          google.script.run
            .withSuccessHandler(function(result) {
              alert('✅ Job posting added successfully!');
              google.script.host.close();
            })
            .withFailureHandler(function(error) {
              alert('❌ Error: ' + error);
            })
            .processJobPosting(data);
        });
      </script>
    </body>
    </html>
  `)
  .setWidth(650)
  .setHeight(800);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '💼 Add New Job Posting');
}

function processJobPosting(data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('💼 Job Postings');
    
    // Validate and process deadline
    let deadline = '';
    if (data.deadline) {
      try {
        deadline = new Date(data.deadline);
      } catch (e) {
        deadline = '';
      }
    }
    
    const newRow = [
      generateJobId(),
      data.title || '',
      data.company || '',
      data.location || '',
      data.type || 'Full-time',
      data.salary || '',
      data.requirements || '',
      data.description || '',
      data.email || '',
      data.phone || '',
      new Date(), // Posted Date
      deadline,
      'Active', // Status
      data.tags || '',
      data.benefits || '',
      data.applicationLink || '',
      Session.getActiveUser().getEmail(), // Created By
      0 // Applications Count
    ];
    
    sheet.appendRow(newRow);
    
    return true;
  } catch (error) {
    throw new Error('Failed to add job posting: ' + error.message);
  }
}

function generateJobId() {
  return 'JOB-' + Math.random().toString(36).substr(2, 9).toUpperCase();
}

// ========================================
// 📝 JOB APPLICATIONS MODULE
// ========================================

function setupJobApplications() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('📝 Job Applications');
  sheet.clear();
  
  const headers = [
    'Application ID', 'Job ID', 'Job Title', 'Company', 'Applicant Name', 'Applicant Email',
    'Phone', 'Cover Letter', 'Resume Link', 'Portfolio Link', 'LinkedIn', 'Skills',
    'Experience Level', 'Languages', 'Application Date', 'Status', 'Interview Date',
    'Notes', 'Created By'
  ];
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  
  // Apply styling
  headerRange.setFontWeight('bold').setBackground('#FF8E53').setFontColor('white');
  
  // Smart data validation
  const statusOptions = ['Applied', 'Under Review', 'Interview Scheduled', 'Interviewed', 'Rejected', 'Hired', 'Withdrawn'];
  const experienceLevels = ['Entry Level', 'Junior', 'Mid-level', 'Senior', 'Expert'];
  const languages = ['Chinese', 'English', 'German', 'French', 'Spanish', 'Italian', 'Multilingual'];
  
  sheet.getRange(2, 16, 1000, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(statusOptions, true).build()
  );
  
  sheet.getRange(2, 13, 1000, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(experienceLevels, true).build()
  );
  
  sheet.getRange(2, 14, 1000, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(languages, true).build()
  );
  
  // Date formatting
  sheet.getRange('O:P').setNumberFormat('mm/dd/yyyy');
  
  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1, headers.length, 150);
}

function applyForJob() {
  // Get active job postings for the dropdown
  const activeJobs = getActiveJobPostings();
  
  let jobOptionsHtml = '<option value="">-- Select a Job --</option>';
  activeJobs.forEach(job => {
    jobOptionsHtml += `<option value="${job.id}" data-title="${job.title}" data-company="${job.company}">${job.id} - ${job.title} at ${job.company}</option>`;
  });
  
  const htmlOutput = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }
        .container { max-width: 600px; margin: 0 auto; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
        h2 { color: #2c3e50; text-align: center; margin-bottom: 30px; }
        .form-group { margin-bottom: 20px; }
        label { display: block; margin-bottom: 5px; font-weight: bold; color: #34495e; }
        input, select, textarea { width: 100%; padding: 12px; border: 2px solid #ddd; border-radius: 5px; font-size: 14px; box-sizing: border-box; }
        input:focus, select:focus, textarea:focus { border-color: #3498db; outline: none; }
        textarea { height: 80px; resize: vertical; }
        .row { display: flex; gap: 15px; }
        .col { flex: 1; }
        .btn { background: #3498db; color: white; padding: 12px 24px; border: none; border-radius: 5px; cursor: pointer; font-size: 16px; margin: 5px; }
        .btn:hover { background: #2980b9; }
        .btn-secondary { background: #95a5a6; }
        .btn-secondary:hover { background: #7f8c8d; }
        .required { color: #e74c3c; }
        .section { background: #ecf0f1; padding: 20px; border-radius: 8px; margin-bottom: 20px; }
        .section h3 { margin-top: 0; color: #2c3e50; }
        .auto-filled { background-color: #e8f5e8; border-color: #4caf50; }
        .info-text { font-size: 12px; color: #666; margin-top: 5px; }
      </style>
    </head>
    <body>
      <div class="container">
        <h2>📝 Apply for Job</h2>
        <form id="applicationForm">
          <div class="section">
            <h3>📋 Job Information</h3>
            <div class="form-group">
              <label for="jobSelection">Select Job Position <span class="required">*</span></label>
              <select id="jobSelection" name="jobSelection" required>
                ${jobOptionsHtml}
              </select>
              <div class="info-text">Choose from available active job postings</div>
            </div>
            <div class="row">
              <div class="col">
                <label for="jobId">Job ID <span class="required">*</span></label>
                <input type="text" id="jobId" name="jobId" required readonly placeholder="Will be auto-filled">
                <div class="info-text">Auto-filled when you select a job above</div>
              </div>
              <div class="col">
                <label for="jobTitle">Job Title <span class="required">*</span></label>
                <input type="text" id="jobTitle" name="jobTitle" required readonly placeholder="Will be auto-filled">
                <div class="info-text">Auto-filled when you select a job above</div>
              </div>
            </div>
            <div class="form-group">
              <label for="company">Company <span class="required">*</span></label>
              <input type="text" id="company" name="company" required readonly placeholder="Will be auto-filled">
              <div class="info-text">Auto-filled when you select a job above</div>
            </div>
          </div>

          <div class="section">
            <h3>👤 Personal Information</h3>
            <div class="row">
              <div class="col">
                <label for="applicantName">Your Name <span class="required">*</span></label>
                <input type="text" id="applicantName" name="applicantName" required placeholder="e.g., John Doe">
              </div>
              <div class="col">
                <label for="applicantEmail">Your Email <span class="required">*</span></label>
                <input type="email" id="applicantEmail" name="applicantEmail" required placeholder="e.g., john@email.com">
              </div>
            </div>
            <div class="form-group">
              <label for="phone">Phone Number</label>
              <input type="tel" id="phone" name="phone" placeholder="e.g., +49 30 123456">
            </div>
          </div>

          <div class="section">
            <h3>💼 Professional Details</h3>
            <div class="form-group">
              <label for="skills">Skills & Experience</label>
              <textarea id="skills" name="skills" placeholder="e.g., JavaScript, React, Node.js, 3+ years experience"></textarea>
            </div>
            <div class="row">
              <div class="col">
                <label for="experienceLevel">Experience Level</label>
                <select id="experienceLevel" name="experienceLevel">
                  <option value="Entry Level">Entry Level</option>
                  <option value="Junior">Junior</option>
                  <option value="Mid-level">Mid-level</option>
                  <option value="Senior">Senior</option>
                  <option value="Expert">Expert</option>
                </select>
              </div>
              <div class="col">
                <label for="languages">Languages</label>
                <input type="text" id="languages" name="languages" placeholder="e.g., Chinese, English, German">
              </div>
            </div>
          </div>

          <div class="section">
            <h3>📄 Additional Information</h3>
            <div class="form-group">
              <label for="coverLetter">Cover Letter</label>
              <textarea id="coverLetter" name="coverLetter" placeholder="Brief introduction and why you're interested in this position"></textarea>
            </div>
            <div class="row">
              <div class="col">
                <label for="resumeLink">Resume Link</label>
                <input type="url" id="resumeLink" name="resumeLink" placeholder="e.g., https://drive.google.com/...">
              </div>
              <div class="col">
                <label for="portfolioLink">Portfolio Link</label>
                <input type="url" id="portfolioLink" name="portfolioLink" placeholder="e.g., https://portfolio.com">
              </div>
            </div>
            <div class="form-group">
              <label for="linkedin">LinkedIn Profile</label>
              <input type="url" id="linkedin" name="linkedin" placeholder="e.g., https://linkedin.com/in/...">
            </div>
          </div>

          <div style="text-align: center; margin-top: 30px;">
            <button type="button" class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
            <button type="submit" class="btn">📝 Submit Application</button>
          </div>
        </form>
      </div>

      <script>
        // Auto-populate job details when job is selected
        document.getElementById('jobSelection').addEventListener('change', function() {
          const selectedOption = this.options[this.selectedIndex];
          const jobIdField = document.getElementById('jobId');
          const jobTitleField = document.getElementById('jobTitle');
          const companyField = document.getElementById('company');
          
          if (selectedOption.value) {
            // Auto-fill the fields
            jobIdField.value = selectedOption.value;
            jobTitleField.value = selectedOption.getAttribute('data-title');
            companyField.value = selectedOption.getAttribute('data-company');
            
            // Add visual feedback that fields are auto-filled
            jobIdField.classList.add('auto-filled');
            jobTitleField.classList.add('auto-filled');
            companyField.classList.add('auto-filled');
          } else {
            // Clear fields if no job selected
            jobIdField.value = '';
            jobTitleField.value = '';
            companyField.value = '';
            
            // Remove auto-filled styling
            jobIdField.classList.remove('auto-filled');
            jobTitleField.classList.remove('auto-filled');
            companyField.classList.remove('auto-filled');
          }
        });
        
        document.getElementById('applicationForm').addEventListener('submit', function(e) {
          e.preventDefault();
          
          const formData = new FormData(this);
          const data = {};
          for (let [key, value] of formData.entries()) {
            data[key] = value;
          }
          
          // Validate that a job was selected
          if (!data.jobSelection) {
            alert('Please select a job position from the dropdown');
            return;
          }
          
          // Validate required fields
          if (!data.jobId || !data.jobTitle || !data.company || !data.applicantName || !data.applicantEmail) {
            alert('Please fill in all required fields (marked with *)');
            return;
          }
          
          google.script.run
            .withSuccessHandler(function(result) {
              alert('✅ Job application submitted successfully!');
              google.script.host.close();
            })
            .withFailureHandler(function(error) {
              alert('❌ Error: ' + error);
            })
            .processJobApplication(data);
        });
      </script>
    </body>
    </html>
  `)
  .setWidth(650)
  .setHeight(800);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '📝 Apply for Job');
}

function processJobApplication(data) {
  return safeExecute('processJobApplication', 'submit job application', () => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('📝 Job Applications');
    
    const newRow = [
      generateApplicationId(),
      data.jobId || '',
      data.jobTitle || '',
      data.company || '',
      data.applicantName || '',
      data.applicantEmail || '',
      data.phone || '',
      data.coverLetter || '',
      data.resumeLink || '',
      data.portfolioLink || '',
      data.linkedin || '',
      data.skills || '',
      data.experienceLevel || 'Entry Level',
      data.languages || 'Chinese,English',
      new Date(), // Application Date
      'Applied', // Status
      '', // Interview Date
      '', // Notes
      Session.getActiveUser().getEmail() // Created By
    ];
    
    sheet.appendRow(newRow);
    
    // Update job applications count
    updateJobApplicationsCount(data.jobId);
    
    return true;
  });
}

function generateApplicationId() {
  return 'APP-' + Math.random().toString(36).substr(2, 9).toUpperCase();
}

function updateJobApplicationsCount(jobId) {
  const jobSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('💼 Job Postings');
  const appSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('📝 Job Applications');
  
  // Count applications for this job
  const appData = appSheet.getDataRange().getValues();
  const appCount = appData.filter(row => row[1] === jobId).length;
  
  // Update the job posting with new count
  const jobData = jobSheet.getDataRange().getValues();
  for (let i = 1; i < jobData.length; i++) {
    if (jobData[i][0] === jobId) {
      jobSheet.getRange(i + 1, 18).setValue(appCount); // Applications Count column
      break;
    }
  }
}

function getActiveJobPostings() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('💼 Job Postings');
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) return []; // No data beyond headers
    
    const activeJobs = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const status = row[12]; // Status column
      
      // Only include active jobs
      if (status === 'Active') {
        activeJobs.push({
          id: row[0],     // Job ID
          title: row[1],  // Title
          company: row[2] // Company
        });
      }
    }
    
    return activeJobs;
  } catch (error) {
    console.log('Error getting active job postings: ' + error);
    return [];
  }
}

// ========================================
// 🔍 JOB SEEKING MODULE
// ========================================

function setupJobSeeking() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('🔍 Job Seeking');
  sheet.clear();
  
  const headers = [
    'Seeker ID', 'Name', 'Email', 'Phone', 'Skills', 'Experience Level',
    'Preferred Location', 'Salary Expectation', 'Job Type Preference',
    'Languages', 'Portfolio Link', 'Resume Link', 'LinkedIn',
    'Availability', 'Notes', 'Status', 'Created Date'
  ];
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  
  // Apply styling
  headerRange.setFontWeight('bold').setBackground('#4ECDC4').setFontColor('white');
  
  // Data validation
  const levelOptions = ['Entry Level', 'Junior', 'Mid-level', 'Senior', 'Expert'];
  const typeOptions = ['Full-time', 'Part-time', 'Internship', 'Freelance', 'Remote'];
  const statusOptions = ['Active', 'Employed', 'Not Available'];
  
  sheet.getRange(2, 6, 1000, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(levelOptions, true).build()
  );
  
  sheet.getRange(2, 9, 1000, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(typeOptions, true).build()
  );
  
  sheet.getRange(2, 16, 1000, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(statusOptions, true).build()
  );
  
  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1, headers.length, 150);
}

// ========================================
// 🗺️ PLACES & ACTIVITIES MODULE
// ========================================

function setupPlacesAndActivities() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('🗺️ Places & Activities');
  sheet.clear();
  
  const headers = [
    'Place ID', 'Name', 'Category', 'Location', 'Address', 'Coordinates',
    'Description', 'Student Discount', 'Opening Hours', 'Contact',
    'Website', 'Rating', 'Price Range', 'Tags', 'Map Link', 'Created Date'
  ];
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  
  // Apply styling
  headerRange.setFontWeight('bold').setBackground('#45B7D1').setFontColor('white');
  
  // Data validation
  const categoryOptions = ['Restaurant', 'Cafe', 'Bar', 'Shopping', 'Entertainment', 'Sports', 'Study', 'Transport'];
  const priceOptions = ['$', '$$', '$$$', '$$$$'];
  
  sheet.getRange(2, 3, 1000, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(categoryOptions, true).build()
  );
  
  sheet.getRange(2, 13, 1000, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(priceOptions, true).build()
  );
  
  // Pre-populate with popular places
  populatePopularPlaces(sheet);
  
  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1, headers.length, 150);
}

function populatePopularPlaces(sheet) {
  const popularPlaces = [
    ['PLACE-001', 'Brandenburger Tor', 'Entertainment', 'Berlin', 'Pariser Platz, 10117 Berlin', '52.5163,13.3777', 'Famous landmark and symbol of Berlin', 'Yes', '24/7', '', 'https://www.berlin.de', '4.5', 'Free', 'Landmark,History,Photo', '', new Date()],
    ['PLACE-002', 'Potsdam Palace', 'Entertainment', 'Potsdam', 'Am Neuen Markt, 14467 Potsdam', '52.3958,13.0617', 'UNESCO World Heritage site', 'Yes', '10:00-18:00', '', 'https://www.spsg.de', '4.8', '$$', 'Palace,History,Culture', '', new Date()],
    ['PLACE-003', 'XU Campus Cafe', 'Cafe', 'Potsdam', 'Campus Location', '52.3958,13.0617', 'Student-friendly cafe on campus', 'Yes', '08:00-20:00', '', '', '4.2', '$', 'Cafe,Study,Student', '', new Date()]
  ];
  
  popularPlaces.forEach(place => {
    sheet.appendRow(place);
  });
}

// ========================================
// 📅 EVENTS MODULE
// ========================================

function setupEvents() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('📅 Events');
  sheet.clear();
  
  const headers = [
    'Event ID', 'Title', 'Category', 'Date', 'Time', 'Location', 'Address',
    'Description', 'Organizer', 'Contact', 'Registration Link', 'Cost',
    'Max Participants', 'Current Participants', 'Status', 'Tags', 'Created Date',
    'Attendees List', 'Waitlist Count'
  ];
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  
  // Apply styling
  headerRange.setFontWeight('bold').setBackground('#96CEB4').setFontColor('white');
  
  // Smart data validation with enhanced options
  const categoryOptions = ['Academic', 'Social', 'Cultural', 'Professional', 'Sports', 'Workshop', 'Conference', 'Networking', 'Career Fair', 'Language Exchange', 'Study Group', 'Cultural Festival'];
  const statusOptions = ['Upcoming', 'Ongoing', 'Completed', 'Cancelled', 'Postponed', 'Registration Closed'];
  const locationOptions = ['XU Campus', 'Potsdam City Center', 'Berlin', 'Online', 'Hybrid', 'Other'];
  const costOptions = ['Free', '€5-10', '€10-20', '€20-50', '€50+', 'Student Discount', 'Donation'];
  
  sheet.getRange(2, 3, 1000, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(categoryOptions, true).build()
  );
  
  sheet.getRange(2, 6, 1000, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(locationOptions, true).build()
  );
  
  sheet.getRange(2, 12, 1000, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(costOptions, true).build()
  );
  
  sheet.getRange(2, 15, 1000, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(statusOptions, true).build()
  );
  
  // Date formatting
  sheet.getRange('D:D').setNumberFormat('mm/dd/yyyy');
  
  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1, headers.length, 150);
}

// ========================================
// 🌐 USEFUL WEBSITES MODULE
// ========================================

function setupUsefulWebsites() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('🌐 Useful Websites');
  sheet.clear();
  
  const headers = [
    'Website ID', 'Name', 'Category', 'URL', 'Description', 'Language',
    'Student Specific', 'Rating', 'Tags', 'Last Updated', 'Created Date'
  ];
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  
  // Apply styling
  headerRange.setFontWeight('bold').setBackground('#FFEAA7').setFontColor('#2D3436');
  
  // Data validation
  const categoryOptions = ['Academic', 'Job Search', 'Transportation', 'Shopping', 'Entertainment', 'News', 'Tools', 'Social'];
  const languageOptions = ['Chinese', 'English', 'German', 'Multilingual'];
  
  sheet.getRange(2, 3, 1000, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(categoryOptions, true).build()
  );
  
  sheet.getRange(2, 6, 1000, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(languageOptions, true).build()
  );
  
  // Pre-populate with useful websites
  populateUsefulWebsites(sheet);
  
  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1, headers.length, 150);
}

function populateUsefulWebsites(sheet) {
  const websites = [
    ['WEB-001', 'Google Translate', 'Tools', 'https://translate.google.com', 'Free translation service', 'Multilingual', 'Yes', '4.8', 'Translation,Free', new Date(), new Date()],
    ['WEB-002', 'Deutsche Bahn', 'Transportation', 'https://www.bahn.de', 'German railway system', 'German', 'Yes', '4.5', 'Transport,Train', new Date(), new Date()],
    ['WEB-003', 'Indeed Germany', 'Job Search', 'https://de.indeed.com', 'Job search platform', 'German', 'Yes', '4.6', 'Jobs,Employment', new Date(), new Date()],
    ['WEB-004', 'XU Student Portal', 'Academic', 'https://xu-university.de', 'University student portal', 'German', 'Yes', '4.7', 'Academic,Student', new Date(), new Date()]
  ];
  
  websites.forEach(website => {
    sheet.appendRow(website);
  });
}

// ========================================
// 💡 BUSINESS OPPORTUNITIES MODULE
// ========================================

function setupBusinessOpportunities() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('💡 Business Opportunities');
  sheet.clear();
  
  const headers = [
    'Opportunity ID', 'Title', 'Category', 'Description', 'Investment Required',
    'Expected Return', 'Timeline', 'Location', 'Contact Person', 'Contact Email',
    'Status', 'Tags', 'Created Date', 'Deadline'
  ];
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  
  // Apply styling
  headerRange.setFontWeight('bold').setBackground('#DDA0DD').setFontColor('white');
  
  // Data validation
  const categoryOptions = ['Startup', 'Partnership', 'Investment', 'Consulting', 'Freelance', 'E-commerce', 'Service'];
  const statusOptions = ['Open', 'In Progress', 'Closed', 'Expired'];
  
  sheet.getRange(2, 3, 1000, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(categoryOptions, true).build()
  );
  
  sheet.getRange(2, 11, 1000, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(statusOptions, true).build()
  );
  
  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1, headers.length, 150);
}

// ========================================
// 👥 COMMUNITY MODULE
// ========================================

function setupCommunity() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('👥 Community');
  sheet.clear();
  
  const headers = [
    'Member ID', 'Name', 'Email', 'Phone', 'Study Program', 'Year',
    'Skills', 'Interests', 'Languages', 'Mentor Status', 'Mentee Status',
    'Social Links', 'Availability', 'Notes', 'Join Date'
  ];
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  
  // Apply styling
  headerRange.setFontWeight('bold').setBackground('#98D8C8').setFontColor('white');
  
  // Data validation
  const mentorOptions = ['Available', 'Busy', 'Not Available'];
  const yearOptions = ['1st Year', '2nd Year', '3rd Year', '4th Year', 'Graduate'];
  
  sheet.getRange(2, 10, 1000, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(mentorOptions, true).build()
  );
  
  sheet.getRange(2, 11, 1000, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(mentorOptions, true).build()
  );
  
  sheet.getRange(2, 6, 1000, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(yearOptions, true).build()
  );
  
  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1, headers.length, 150);
}

// ========================================
// 📊 ANALYTICS MODULE
// ========================================

function setupAnalytics() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('📊 Analytics');
  sheet.clear();
  
  // Create analytics dashboard
  const analyticsData = [
    ['📊 X.LAB CHINESE STUDENT CENTER ANALYTICS', '', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', '', ''],
    ['📈 KEY METRICS', '', '', '', '', '', '', '', '', ''],
    ['Total Job Postings', '=COUNTA(\'💼 Job Postings\'!A:A)-1', '', '', '', '', '', '', '', ''],
    ['Total Applications', '=COUNTA(\'📝 Job Applications\'!A:A)-1', '', '', '', '', '', '', '', ''],
    ['Active Job Seekers', '=COUNTIF(\'🔍 Job Seeking\'!P:P, "Active")', '', '', '', '', '', '', '', ''],
    ['Upcoming Events', '=COUNTIF(\'📅 Events\'!O:O, "Upcoming")', '', '', '', '', '', '', '', ''],
    ['Community Members', '=COUNTA(\'👥 Community\'!A:A)-1', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', '', ''],
    ['🎯 TOP CATEGORIES', '', '', '', '', '', '', '', '', ''],
    ['Most Popular Job Type', '=MODE(\'💼 Job Postings\'!E:E)', '', '', '', '', '', '', '', ''],
    ['Most Sought Skills', '=MODE(\'🔍 Job Seeking\'!E:E)', '', '', '', '', '', '', '', ''],
    ['Popular Event Category', '=MODE(\'📅 Events\'!C:C)', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', '', ''],
    ['🗺️ LOCATION INSIGHTS', '', '', '', '', '', '', '', '', ''],
    ['Jobs in Berlin', '=COUNTIF(\'💼 Job Postings\'!D:D, "*Berlin*")', '', '', '', '', '', '', '', ''],
    ['Jobs in Potsdam', '=COUNTIF(\'💼 Job Postings\'!D:D, "*Potsdam*")', '', '', '', '', '', '', '', ''],
    ['Places in Berlin', '=COUNTIF(\'🗺️ Places & Activities\'!D:D, "*Berlin*")', '', '', '', '', '', '', '', ''],
    ['Places in Potsdam', '=COUNTIF(\'🗺️ Places & Activities\'!D:D, "*Potsdam*")', '', '', '', '', '', '', '', '']
  ];
  
  const range = sheet.getRange(1, 1, analyticsData.length, 10);
  range.setValues(analyticsData);
  
  // Apply styling
  sheet.getRange('A1:J1').merge();
  sheet.getRange('A1').setFontSize(20).setFontWeight('bold').setBackground('#6C5CE7').setFontColor('white');
  
  sheet.getRange('A3:J3').merge();
  sheet.getRange('A3').setFontSize(16).setFontWeight('bold').setBackground('#A29BFE');
  
  sheet.getRange('A9:J9').merge();
  sheet.getRange('A9').setFontSize(16).setFontWeight('bold').setBackground('#FD79A8');
  
  sheet.getRange('A13:J13').merge();
  sheet.getRange('A13').setFontSize(16).setFontWeight('bold').setBackground('#FDCB6E');
  
  sheet.setColumnWidths(1, 10, 200);
}

// ========================================
// ⚙️ SETTINGS MODULE
// ========================================

function setupSettings() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('⚙️ Settings');
  sheet.clear();
  
  const settingsData = [
    ['⚙️ SYSTEM SETTINGS', '', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', '', ''],
    ['🔧 GENERAL SETTINGS', '', '', '', '', '', '', '', '', ''],
    ['System Name', 'X.lab Chinese Student Center', '', '', '', '', '', '', '', ''],
    ['University', 'XU Exponential University', '', '', '', '', '', '', '', ''],
    ['Location', 'Potsdam, Germany', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', '', ''],
    ['🌍 LOCATION SETTINGS', '', '', '', '', '', '', '', '', ''],
    ['Default City', 'Potsdam', '', '', '', '', '', '', '', ''],
    ['Secondary City', 'Berlin', '', '', '', '', '', '', '', ''],
    ['Map API Key', '[Configure in Google Cloud Console]', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', '', ''],
    ['📧 NOTIFICATION SETTINGS', '', '', '', '', '', '', '', '', ''],
    ['Email Notifications', 'Enabled', '', '', '', '', '', '', '', ''],
    ['Event Reminders', 'Enabled', '', '', '', '', '', '', '', ''],
    ['Job Alerts', 'Enabled', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', '', ''],
    ['🔐 SECURITY SETTINGS', '', '', '', '', '', '', '', '', ''],
    ['Data Encryption', 'Enabled', '', '', '', '', '', '', '', ''],
    ['Access Control', 'Student Only', '', '', '', '', '', '', '', ''],
    ['Backup Frequency', 'Daily', '', '', '', '', '', '', '', '']
  ];
  
  const range = sheet.getRange(1, 1, settingsData.length, 10);
  range.setValues(settingsData);
  
  // Apply styling
  sheet.getRange('A1:J1').merge();
  sheet.getRange('A1').setFontSize(20).setFontWeight('bold').setBackground('#2D3436').setFontColor('white');
  
  sheet.getRange('A3:J3').merge();
  sheet.getRange('A3').setFontSize(16).setFontWeight('bold').setBackground('#636E72');
  
  sheet.getRange('A8:J8').merge();
  sheet.getRange('A8').setFontSize(16).setFontWeight('bold').setBackground('#74B9FF');
  
  sheet.getRange('A13:J13').merge();
  sheet.getRange('A13').setFontSize(16).setFontWeight('bold').setBackground('#55A3FF');
  
  sheet.getRange('A18:J18').merge();
  sheet.getRange('A18').setFontSize(16).setFontWeight('bold').setBackground('#A29BFE');
  
  sheet.setColumnWidths(1, 10, 200);
}

// ========================================
// 🗺️ MAP INTEGRATION FUNCTIONS
// ========================================

function getMapUrl(location) {
  return `https://www.google.com/maps/search/${encodeURIComponent(location)}`;
}

function createInteractiveMap() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt('Interactive Map', 'Enter location to view on map:', ui.ButtonSet.OK_CANCEL);
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const location = result.getResponseText();
    const mapUrl = getMapUrl(location);
    
    // Open map in new tab
    const htmlOutput = HtmlService.createHtmlOutput(`
      <script>
        window.open('${mapUrl}', '_blank');
        google.script.host.close();
      </script>
    `);
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Opening Map...');
  }
}

// ========================================
// 🔍 SEARCH & FILTER FUNCTIONS
// ========================================

function searchJobs(keyword) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('💼 Job Postings');
  const data = sheet.getDataRange().getValues();
  
  const results = data.filter(row => 
    row[1] && row[1].toLowerCase().includes(keyword.toLowerCase()) ||
    row[2] && row[2].toLowerCase().includes(keyword.toLowerCase()) ||
    row[3] && row[3].toLowerCase().includes(keyword.toLowerCase())
  );
  
  return results;
}

function filterEventsByDate(startDate, endDate) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('📅 Events');
  const data = sheet.getDataRange().getValues();
  
  const results = data.filter(row => {
    const eventDate = new Date(row[3]);
    return eventDate >= startDate && eventDate <= endDate;
  });
  
  return results;
}

// ========================================
// 📧 NOTIFICATION FUNCTIONS
// ========================================

function sendEventReminder() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('📅 Events');
  const data = sheet.getDataRange().getValues();
  const today = new Date();
  
  const upcomingEvents = data.filter(row => {
    const eventDate = new Date(row[3]);
    const diffTime = eventDate.getTime() - today.getTime();
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    return diffDays >= 0 && diffDays <= 7;
  });
  
  if (upcomingEvents.length > 0) {
    const message = `📅 Upcoming Events:\n\n${upcomingEvents.map(event => 
      `• ${event[1]} - ${event[3]} at ${event[4]}`
    ).join('\n')}`;
    
    SpreadsheetApp.getUi().alert('Event Reminders', message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ========================================
// 🎯 QUICK ACTIONS MENU
// ========================================

function onOpen() {
  // Novel approach: Multiple menu creation attempts with different strategies
  const ui = SpreadsheetApp.getUi();
  
  // Strategy 1: Try creating the full menu
  try {
    const mainMenu = ui.createMenu('🏛️ X.Lab Center');
    
    // Add main items
    mainMenu.addItem('📊 View Dashboard', 'showDashboard');
    mainMenu.addItem('🔄 Refresh Dashboard Stats', 'refreshDashboardStats');
    mainMenu.addItem('📈 Show Detailed Stats', 'showDetailedStats');
    mainMenu.addSeparator();
    
    // Add Jobs submenu
    const jobsMenu = ui.createMenu('💼 Jobs');
    jobsMenu.addItem('Add Job Posting', 'addJobPosting');
    jobsMenu.addItem('Apply for Job', 'applyForJob');
    jobsMenu.addItem('Search Jobs', 'searchJobsMenu');
    jobsMenu.addItem('View All Jobs', 'showJobPostings');
    jobsMenu.addItem('View Applications', 'showJobApplications');
    mainMenu.addSubMenu(jobsMenu);
    
    // Add Events submenu
    const eventsMenu = ui.createMenu('📅 Events');
    eventsMenu.addItem('Add New Event', 'addNewEvent');
    eventsMenu.addItem('Register for Event', 'registerForEvent');
    eventsMenu.addItem('View Upcoming Events', 'viewUpcomingEvents');
    eventsMenu.addItem('Event Reminders', 'sendEventReminder');
    mainMenu.addSubMenu(eventsMenu);
    
    // Add Places submenu
    const placesMenu = ui.createMenu('🗺️ Places');
    placesMenu.addItem('Add New Place', 'addNewPlace');
    placesMenu.addItem('View Interactive Map', 'createInteractiveMap');
    placesMenu.addItem('Find Places Near Me', 'findNearbyPlaces');
    mainMenu.addSubMenu(placesMenu);
    
    // Add Community submenu
    const communityMenu = ui.createMenu('👥 Community');
    communityMenu.addItem('Join Community', 'joinCommunity');
    communityMenu.addItem('Find Mentor', 'findMentor');
    communityMenu.addItem('Become Mentor', 'becomeMentor');
    mainMenu.addSubMenu(communityMenu);
    
    // Add Resources submenu
    const resourcesMenu = ui.createMenu('🌐 Resources');
    resourcesMenu.addItem('Add Useful Website', 'addUsefulWebsite');
    resourcesMenu.addItem('Add Business Opportunity', 'addBusinessOpportunity');
    mainMenu.addSubMenu(resourcesMenu);
    
    mainMenu.addSeparator();
    
    // Add Quick Actions submenu (NOW ONLY IN MENU)
    const quickActionsMenu = ui.createMenu('⚡ Quick Actions');
    quickActionsMenu.addItem('🚀 Quick Job Search', 'quickJobSearch');
    quickActionsMenu.addItem('📅 Quick Event Registration', 'quickEventRegistration');
    quickActionsMenu.addItem('🗺️ Quick Place Finder', 'quickPlaceFinder');
    quickActionsMenu.addItem('👥 Quick Community Join', 'quickCommunityJoin');
    quickActionsMenu.addItem('📊 Quick Analytics', 'quickAnalytics');
    mainMenu.addSubMenu(quickActionsMenu);
    
    mainMenu.addSeparator();
    mainMenu.addItem('📊 View Analytics', 'showAnalytics');
    mainMenu.addItem('⚙️ Settings', 'showSettings');
    
    mainMenu.addToUi();
    console.log('✅ Full menu created successfully');
    return;
    
  } catch (error) {
    console.error('❌ Strategy 1 failed:', error);
  }
  
  // Strategy 2: Create a simpler menu
  try {
    ui.createMenu('🏛️ X.Lab Center')
      .addItem('📊 Dashboard', 'showDashboard')
      .addSeparator()
      .addItem('⚡ Quick Actions', 'showQuickActionsMenu')
      .addItem('💼 Jobs', 'showJobPostings')
      .addItem('📅 Events', 'showEvents')
      .addItem('🗺️ Places', 'showPlaces')
      .addToUi();
    console.log('✅ Simple menu created successfully');
    return;
    
  } catch (error) {
    console.error('❌ Strategy 2 failed:', error);
  }
  
  // Strategy 3: Create minimal menu
  try {
    ui.createMenu('🏛️ X.Lab Center')
      .addItem('📊 Dashboard', 'showDashboard')
      .addItem('⚡ Quick Actions', 'showQuickActionsMenu')
      .addToUi();
    console.log('✅ Minimal menu created successfully');
    
  } catch (error) {
    console.error('❌ All menu creation strategies failed:', error);
    // Last resort: show alert to user
    ui.alert('⚠️ Menu Creation Issue', 
      'The X.Lab Center menu could not be created automatically.\n\n' +
      'Please run the "forceCreateMenu" function manually to create the menu.', 
      ui.ButtonSet.OK);
  }
}

// Manual menu creation function (can be called if onOpen doesn't work)
function createMenu() {
  onOpen();
}

// Function to manually create the menu (run this if menu doesn't appear)
function forceCreateMenu() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Create comprehensive top menu with quick actions
    const mainMenu = ui.createMenu('🏛️ X.Lab Center');
    
    // Add main items
    mainMenu.addItem('📊 View Dashboard', 'showDashboard');
    mainMenu.addItem('🔄 Refresh Dashboard Stats', 'refreshDashboardStats');
    mainMenu.addItem('📈 Show Detailed Stats', 'showDetailedStats');
    mainMenu.addSeparator();
    
    // Add Jobs submenu
    const jobsMenu = ui.createMenu('💼 Jobs');
    jobsMenu.addItem('Add Job Posting', 'addJobPosting');
    jobsMenu.addItem('Apply for Job', 'applyForJob');
    jobsMenu.addItem('Search Jobs', 'searchJobsMenu');
    jobsMenu.addItem('View All Jobs', 'showJobPostings');
    jobsMenu.addItem('View Applications', 'showJobApplications');
    mainMenu.addSubMenu(jobsMenu);
    
    // Add Events submenu
    const eventsMenu = ui.createMenu('📅 Events');
    eventsMenu.addItem('Add New Event', 'addNewEvent');
    eventsMenu.addItem('Register for Event', 'registerForEvent');
    eventsMenu.addItem('View Upcoming Events', 'viewUpcomingEvents');
    eventsMenu.addItem('Event Reminders', 'sendEventReminder');
    mainMenu.addSubMenu(eventsMenu);
    
    // Add Places submenu
    const placesMenu = ui.createMenu('🗺️ Places');
    placesMenu.addItem('Add New Place', 'addNewPlace');
    placesMenu.addItem('View Interactive Map', 'createInteractiveMap');
    placesMenu.addItem('Find Places Near Me', 'findNearbyPlaces');
    mainMenu.addSubMenu(placesMenu);
    
    // Add Community submenu
    const communityMenu = ui.createMenu('👥 Community');
    communityMenu.addItem('Join Community', 'joinCommunity');
    communityMenu.addItem('Find Mentor', 'findMentor');
    communityMenu.addItem('Become Mentor', 'becomeMentor');
    mainMenu.addSubMenu(communityMenu);
    
    // Add Resources submenu
    const resourcesMenu = ui.createMenu('🌐 Resources');
    resourcesMenu.addItem('Add Useful Website', 'addUsefulWebsite');
    resourcesMenu.addItem('Add Business Opportunity', 'addBusinessOpportunity');
    mainMenu.addSubMenu(resourcesMenu);
    
    mainMenu.addSeparator();
    
    // Add Quick Actions submenu (NOW ONLY IN MENU)
    const quickActionsMenu = ui.createMenu('⚡ Quick Actions');
    quickActionsMenu.addItem('🚀 Quick Job Search', 'quickJobSearch');
    quickActionsMenu.addItem('📅 Quick Event Registration', 'quickEventRegistration');
    quickActionsMenu.addItem('🗺️ Quick Place Finder', 'quickPlaceFinder');
    quickActionsMenu.addItem('👥 Quick Community Join', 'quickCommunityJoin');
    quickActionsMenu.addItem('📊 Quick Analytics', 'quickAnalytics');
    mainMenu.addSubMenu(quickActionsMenu);
    
    mainMenu.addSeparator();
    mainMenu.addItem('📊 View Analytics', 'showAnalytics');
    mainMenu.addItem('⚙️ Settings', 'showSettings');
    
    mainMenu.addToUi();
    console.log('✅ Menu created successfully via forceCreateMenu');
    return true;
  } catch (error) {
    console.error('❌ Error in forceCreateMenu:', error);
    return false;
  }
}

// Fallback quick actions menu
function showQuickActionsMenu() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert('⚡ Quick Actions', 
    'Choose an action:\n\n' +
    '1. Quick Job Search\n' +
    '2. Quick Event Registration\n' +
    '3. Quick Place Finder\n' +
    '4. Quick Community Join\n' +
    '5. Quick Analytics\n\n' +
    'Click OK for Job Search, Cancel for Event Registration, or No for Place Finder',
    ui.ButtonSet.YES_NO_CANCEL);
    
  switch(result) {
    case ui.Button.YES:
      quickJobSearch();
      break;
    case ui.Button.NO:
      quickEventRegistration();
      break;
    case ui.Button.CANCEL:
      quickPlaceFinder();
      break;
  }
}

// Function to help users create the menu manually
function createMenuManually() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert('🏛️ Menu Creation', 
    'The X.Lab Center menu should appear automatically.\n\n' +
    'If you don\'t see it, click "OK" to create it manually.\n\n' +
    'After creation, refresh the page to see the menu.',
    ui.ButtonSet.OK_CANCEL);
    
  if (result === ui.Button.OK) {
    const success = forceCreateMenu();
    if (success) {
      ui.alert('✅ Menu Created', 
        'The X.Lab Center menu has been created successfully!\n\n' +
        'Please refresh the page to see the menu in the native Google Sheets menu bar.',
        ui.ButtonSet.OK);
    } else {
      ui.alert('❌ Menu Creation Failed', 
        'The menu could not be created. Please try again or contact support.',
        ui.ButtonSet.OK);
    }
  }
}

function showDashboard() {
  // Refresh the dashboard to ensure Quick Actions are removed
  createInnovativeDashboard();
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('🎯 Dashboard')
  );
  
  // Show confirmation that Quick Actions are now in the menu only
  const ui = SpreadsheetApp.getUi();
  ui.alert('🎯 Dashboard Updated', 
    'Quick Actions have been removed from the dashboard cells.\n\n' +
    'All Quick Actions are now available in the "🏛️ X.Lab Center" menu under "⚡ Quick Actions".\n\n' +
    'Use the native Google Sheets menu bar to access all features!', 
    ui.ButtonSet.OK);
}

function showJobPostings() {
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('💼 Job Postings')
  );
}

function showJobApplications() {
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('📝 Job Applications')
  );
}

// ========================================
// ⚡ QUICK ACTION FUNCTIONS
// ========================================

function quickJobSearch() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt('Quick Job Search', 
    'Enter keywords (title, company, location, skills):', 
    ui.ButtonSet.OK_CANCEL);
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const keyword = result.getResponseText();
    const results = searchJobs(keyword);
    
    if (results.length > 0) {
      let message = `Found ${results.length} job(s) matching "${keyword}":\n\n`;
      results.forEach((job, index) => {
        message += `${index + 1}. ${job[1]} at ${job[2]} (${job[3]})\n   Type: ${job[4]} | Salary: ${job[5]}\n   Applications: ${job[17] || 0}\n\n`;
      });
      message += '\nUse "Apply for Job" from the menu to apply!';
      ui.alert('Quick Job Search Results', message, ui.ButtonSet.OK);
    } else {
      ui.alert('No Results', `No jobs found matching "${keyword}"`, ui.ButtonSet.OK);
    }
  }
}

function quickEventRegistration() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('📅 Events');
  const eventData = sheet.getDataRange().getValues();
  
  // Show available events
  const upcomingEvents = eventData.filter(row => 
    row[15] === 'Upcoming' && row[1] // Status is Upcoming and has title
  );
  
  if (upcomingEvents.length > 0) {
    let message = 'Available Events:\n\n';
    upcomingEvents.forEach((event, index) => {
      message += `${index + 1}. ${event[1]} - ${event[3]} at ${event[4]}\n   Location: ${event[5]} | Cost: ${event[11]}\n   ID: ${event[0]}\n\n`;
    });
    message += '\nUse "Register for Event" from the menu to register!';
    ui.alert('Quick Event Registration', message, ui.ButtonSet.OK);
  } else {
    ui.alert('No Events', 'No upcoming events available for registration.', ui.ButtonSet.OK);
  }
}

function quickPlaceFinder() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt('Quick Place Finder', 
    'Enter location or category (e.g., Berlin, Cafe, Restaurant, Study):', 
    ui.ButtonSet.OK_CANCEL);
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const searchTerm = result.getResponseText();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('🗺️ Places & Activities');
    const data = sheet.getDataRange().getValues();
    
    const matchingPlaces = data.filter(row => 
      (row[3] && row[3].toLowerCase().includes(searchTerm.toLowerCase())) ||
      (row[2] && row[2].toLowerCase().includes(searchTerm.toLowerCase())) ||
      (row[1] && row[1].toLowerCase().includes(searchTerm.toLowerCase()))
    );
    
    if (matchingPlaces.length > 0) {
      let message = `Found ${matchingPlaces.length} place(s) matching "${searchTerm}":\n\n`;
      matchingPlaces.forEach((place, index) => {
        message += `${index + 1}. ${place[1]} (${place[2]})\n   Location: ${place[3]} | Rating: ${place[11] || 'N/A'}\n   Price: ${place[12] || 'N/A'}\n\n`;
      });
      ui.alert('Quick Place Finder Results', message, ui.ButtonSet.OK);
    } else {
      ui.alert('No Places Found', `No places found matching "${searchTerm}"`, ui.ButtonSet.OK);
    }
  }
}

function quickCommunityJoin() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt('Quick Community Join', 
    'Enter your details (Name, Email, Study Program, Skills):', 
    ui.ButtonSet.OK_CANCEL);
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const memberData = result.getResponseText().split(',');
    
    if (memberData.length >= 3) {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('👥 Community');
      const newRow = [
        generateMemberId(),
        memberData[0] || '', // Name
        memberData[1] || '', // Email
        '', // Phone
        memberData[2] || '', // Study Program
        '1st Year', // Year
        memberData[3] || '', // Skills
        '', // Interests
        'Chinese,English', // Languages
        'Available', // Mentor Status
        'Available', // Mentee Status
        '', // Social Links
        'Available', // Availability
        '', // Notes
        new Date() // Join Date
      ];
      
      sheet.appendRow(newRow);
      ui.alert('🎉 Welcome to the X.lab Community! You can now connect with other students and mentors.');
    }
  }
}

function quickAnalytics() {
  const ui = SpreadsheetApp.getUi();
  
  // Get quick stats
  const jobSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('💼 Job Postings');
  const appSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('📝 Job Applications');
  const eventSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('📅 Events');
  const communitySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('👥 Community');
  
  const jobData = jobSheet.getDataRange().getValues();
  const appData = appSheet.getDataRange().getValues();
  const eventData = eventSheet.getDataRange().getValues();
  const communityData = communitySheet.getDataRange().getValues();
  
  const activeJobs = jobData.filter(row => row[12] === 'Active').length;
  const totalApplications = appData.length - 1; // Subtract header
  const upcomingEvents = eventData.filter(row => row[15] === 'Upcoming').length;
  const communityMembers = communityData.length - 1; // Subtract header
  
  const message = `📊 Quick Analytics Summary:\n\n` +
    `💼 Active Job Postings: ${activeJobs}\n` +
    `📝 Total Applications: ${totalApplications}\n` +
    `📅 Upcoming Events: ${upcomingEvents}\n` +
    `👥 Community Members: ${communityMembers}\n\n` +
    `Use "View Analytics" for detailed insights!`;
  
  ui.alert('Quick Analytics', message, ui.ButtonSet.OK);
}

function showAnalytics() {
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('📊 Analytics')
  );
}

function showSettings() {
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('⚙️ Settings')
  );
}

// ========================================
// 🔍 ENHANCED SEARCH & MENU FUNCTIONS
// ========================================

function searchJobsMenu() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt('Search Jobs', 'Enter keywords (title, company, location):', ui.ButtonSet.OK_CANCEL);
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const keyword = result.getResponseText();
    const results = searchJobs(keyword);
    
    if (results.length > 0) {
      let message = `Found ${results.length} job(s) matching "${keyword}":\n\n`;
      results.forEach((job, index) => {
        message += `${index + 1}. ${job[1]} at ${job[2]} (${job[3]})\n`;
      });
      ui.alert('Search Results', message, ui.ButtonSet.OK);
    } else {
      ui.alert('No Results', `No jobs found matching "${keyword}"`, ui.ButtonSet.OK);
    }
  }
}

function addNewPlace() {
  const htmlOutput = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }
        .container { max-width: 600px; margin: 0 auto; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
        h2 { color: #2c3e50; text-align: center; margin-bottom: 30px; }
        .form-group { margin-bottom: 20px; }
        label { display: block; margin-bottom: 5px; font-weight: bold; color: #34495e; }
        input, select, textarea { width: 100%; padding: 12px; border: 2px solid #ddd; border-radius: 5px; font-size: 14px; box-sizing: border-box; }
        input:focus, select:focus, textarea:focus { border-color: #3498db; outline: none; }
        textarea { height: 80px; resize: vertical; }
        .row { display: flex; gap: 15px; }
        .col { flex: 1; }
        .btn { background: #3498db; color: white; padding: 12px 24px; border: none; border-radius: 5px; cursor: pointer; font-size: 16px; margin: 5px; }
        .btn:hover { background: #2980b9; }
        .btn-secondary { background: #95a5a6; }
        .btn-secondary:hover { background: #7f8c8d; }
        .required { color: #e74c3c; }
        .section { background: #ecf0f1; padding: 20px; border-radius: 8px; margin-bottom: 20px; }
        .section h3 { margin-top: 0; color: #2c3e50; }
      </style>
    </head>
    <body>
      <div class="container">
        <h2>🗺️ Add New Place</h2>
        <form id="placeForm">
          <div class="section">
            <h3>📋 Basic Information</h3>
            <div class="form-group">
              <label for="name">Place Name <span class="required">*</span></label>
              <input type="text" id="name" name="name" required placeholder="e.g., Brandenburg Gate">
            </div>
            <div class="row">
              <div class="col">
                <label for="category">Category <span class="required">*</span></label>
                <select id="category" name="category" required>
                  <option value="">Select Category</option>
                  <option value="Entertainment">Entertainment</option>
                  <option value="Restaurant">Restaurant</option>
                  <option value="Shopping">Shopping</option>
                  <option value="Cultural">Cultural</option>
                  <option value="Outdoor">Outdoor</option>
                  <option value="Sports">Sports</option>
                  <option value="Transportation">Transportation</option>
                  <option value="Education">Education</option>
                  <option value="Healthcare">Healthcare</option>
                  <option value="Other">Other</option>
                </select>
              </div>
              <div class="col">
                <label for="location">Location <span class="required">*</span></label>
                <input type="text" id="location" name="location" required placeholder="e.g., Berlin">
              </div>
            </div>
          </div>

          <div class="section">
            <h3>📍 Location Details</h3>
            <div class="form-group">
              <label for="address">Address</label>
              <input type="text" id="address" name="address" placeholder="e.g., Unter den Linden 77, 10117 Berlin">
            </div>
            <div class="row">
              <div class="col">
                <label for="coordinates">Coordinates</label>
                <input type="text" id="coordinates" name="coordinates" placeholder="e.g., 52.5163, 13.3777">
              </div>
              <div class="col">
                <label for="mapLink">Map Link</label>
                <input type="url" id="mapLink" name="mapLink" placeholder="e.g., https://maps.google.com/...">
              </div>
            </div>
          </div>

          <div class="section">
            <h3>📝 Description & Details</h3>
            <div class="form-group">
              <label for="description">Description</label>
              <textarea id="description" name="description" placeholder="Describe the place, what makes it special, and why students should visit"></textarea>
            </div>
            <div class="row">
              <div class="col">
                <label for="openingHours">Opening Hours</label>
                <input type="text" id="openingHours" name="openingHours" placeholder="e.g., Mon-Fri 9AM-6PM, Sat-Sun 10AM-8PM">
              </div>
              <div class="col">
                <label for="contact">Contact Information</label>
                <input type="text" id="contact" name="contact" placeholder="e.g., +49 30 123456">
              </div>
            </div>
          </div>

          <div class="section">
            <h3>💰 Pricing & Features</h3>
            <div class="row">
              <div class="col">
                <label for="priceRange">Price Range</label>
                <select id="priceRange" name="priceRange">
                  <option value="$">$ (Budget-friendly)</option>
                  <option value="$$">$$ (Moderate)</option>
                  <option value="$$$">$$$ (Expensive)</option>
                  <option value="$$$$">$$$$ (Luxury)</option>
                  <option value="Free">Free</option>
                </select>
              </div>
              <div class="col">
                <label for="studentDiscount">Student Discount</label>
                <select id="studentDiscount" name="studentDiscount">
                  <option value="Yes">Yes</option>
                  <option value="No">No</option>
                  <option value="Unknown">Unknown</option>
                </select>
              </div>
            </div>
            <div class="form-group">
              <label for="website">Website</label>
              <input type="url" id="website" name="website" placeholder="e.g., https://example.com">
            </div>
            <div class="form-group">
              <label for="tags">Tags</label>
              <input type="text" id="tags" name="tags" placeholder="e.g., Historical, Tourist, Photo Spot, Free Entry">
            </div>
          </div>

          <div style="text-align: center; margin-top: 30px;">
            <button type="button" class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
            <button type="submit" class="btn">🗺️ Add Place</button>
          </div>
        </form>
      </div>

      <script>
        document.getElementById('placeForm').addEventListener('submit', function(e) {
          e.preventDefault();
          
          const formData = new FormData(this);
          const data = {};
          for (let [key, value] of formData.entries()) {
            data[key] = value;
          }
          
          // Validate required fields
          if (!data.name || !data.category || !data.location) {
            alert('Please fill in all required fields (marked with *)');
            return;
          }
          
          google.script.run
            .withSuccessHandler(function(result) {
              alert('✅ Place added successfully!');
              google.script.host.close();
            })
            .withFailureHandler(function(error) {
              alert('❌ Error: ' + error);
            })
            .processNewPlace(data);
        });
      </script>
    </body>
    </html>
  `)
  .setWidth(650)
  .setHeight(800);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '🗺️ Add New Place');
}

function processNewPlace(data) {
  return safeExecute('processNewPlace', 'add new place', () => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('🗺️ Places & Activities');
    
    const newRow = [
      generatePlaceId(),
      data.name || '',
      data.category || 'Entertainment',
      data.location || '',
      data.address || '',
      data.coordinates || '',
      data.description || '',
      data.studentDiscount || 'Yes',
      data.openingHours || '',
      data.contact || '',
      data.website || '',
      '', // Rating
      data.priceRange || '$',
      data.tags || '',
      data.mapLink || '',
      new Date() // Created Date
    ];
    
    sheet.appendRow(newRow);
    
    return true;
  });
}

function findNearbyPlaces() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt('Find Places Near Me', 'Enter your location (e.g., Potsdam, Berlin):', ui.ButtonSet.OK_CANCEL);
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const location = result.getResponseText();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('🗺️ Places & Activities');
    const data = sheet.getDataRange().getValues();
    
    const nearbyPlaces = data.filter(row => 
      row[3] && row[3].toLowerCase().includes(location.toLowerCase())
    );
    
    if (nearbyPlaces.length > 0) {
      let message = `Places in ${location}:\n\n`;
      nearbyPlaces.forEach((place, index) => {
        message += `${index + 1}. ${place[1]} (${place[2]})\n`;
      });
      ui.alert('Nearby Places', message, ui.ButtonSet.OK);
    } else {
      ui.alert('No Places Found', `No places found in ${location}`, ui.ButtonSet.OK);
    }
  }
}

function addNewEvent() {
  const htmlOutput = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }
        .container { max-width: 600px; margin: 0 auto; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
        h2 { color: #2c3e50; text-align: center; margin-bottom: 30px; }
        .form-group { margin-bottom: 20px; }
        label { display: block; margin-bottom: 5px; font-weight: bold; color: #34495e; }
        input, select, textarea { width: 100%; padding: 12px; border: 2px solid #ddd; border-radius: 5px; font-size: 14px; box-sizing: border-box; }
        input:focus, select:focus, textarea:focus { border-color: #3498db; outline: none; }
        textarea { height: 80px; resize: vertical; }
        .row { display: flex; gap: 15px; }
        .col { flex: 1; }
        .btn { background: #3498db; color: white; padding: 12px 24px; border: none; border-radius: 5px; cursor: pointer; font-size: 16px; margin: 5px; }
        .btn:hover { background: #2980b9; }
        .btn-secondary { background: #95a5a6; }
        .btn-secondary:hover { background: #7f8c8d; }
        .required { color: #e74c3c; }
        .section { background: #ecf0f1; padding: 20px; border-radius: 8px; margin-bottom: 20px; }
        .section h3 { margin-top: 0; color: #2c3e50; }
      </style>
    </head>
    <body>
      <div class="container">
        <h2>📅 Add New Event</h2>
        <form id="eventForm">
          <div class="section">
            <h3>📋 Basic Information</h3>
            <div class="form-group">
              <label for="title">Event Title <span class="required">*</span></label>
              <input type="text" id="title" name="title" required placeholder="e.g., Chinese New Year Celebration">
            </div>
            <div class="row">
              <div class="col">
                <label for="category">Category <span class="required">*</span></label>
                <select id="category" name="category" required>
                  <option value="">Select Category</option>
                  <option value="Social">Social</option>
                  <option value="Academic">Academic</option>
                  <option value="Cultural">Cultural</option>
                  <option value="Professional">Professional</option>
                  <option value="Sports">Sports</option>
                  <option value="Entertainment">Entertainment</option>
                  <option value="Workshop">Workshop</option>
                  <option value="Networking">Networking</option>
                  <option value="Other">Other</option>
                </select>
              </div>
              <div class="col">
                <label for="cost">Cost</label>
                <select id="cost" name="cost">
                  <option value="Free">Free</option>
                  <option value="Low">Low Cost</option>
                  <option value="Moderate">Moderate</option>
                  <option value="High">High Cost</option>
                </select>
              </div>
            </div>
          </div>

          <div class="section">
            <h3>📅 Date & Time</h3>
            <div class="row">
              <div class="col">
                <label for="date">Event Date <span class="required">*</span></label>
                <input type="date" id="date" name="date" required>
              </div>
              <div class="col">
                <label for="time">Event Time <span class="required">*</span></label>
                <input type="time" id="time" name="time" required>
              </div>
            </div>
          </div>

          <div class="section">
            <h3>📍 Location</h3>
            <div class="form-group">
              <label for="location">Location <span class="required">*</span></label>
              <input type="text" id="location" name="location" required placeholder="e.g., XU Campus, Room 101">
            </div>
            <div class="form-group">
              <label for="address">Address</label>
              <input type="text" id="address" name="address" placeholder="e.g., 123 Main Street, Potsdam">
            </div>
          </div>

          <div class="section">
            <h3>📝 Event Details</h3>
            <div class="form-group">
              <label for="description">Description</label>
              <textarea id="description" name="description" placeholder="Describe the event, what attendees can expect, and any special requirements"></textarea>
            </div>
            <div class="row">
              <div class="col">
                <label for="maxParticipants">Max Participants</label>
                <input type="number" id="maxParticipants" name="maxParticipants" placeholder="e.g., 50">
              </div>
              <div class="col">
                <label for="tags">Tags</label>
                <input type="text" id="tags" name="tags" placeholder="e.g., Chinese, Networking, Free Food">
              </div>
            </div>
          </div>

          <div class="section">
            <h3>📞 Contact & Registration</h3>
            <div class="form-group">
              <label for="contact">Contact Information</label>
              <input type="text" id="contact" name="contact" placeholder="e.g., event@xu.edu">
            </div>
            <div class="form-group">
              <label for="registrationLink">Registration Link</label>
              <input type="url" id="registrationLink" name="registrationLink" placeholder="e.g., https://forms.google.com/...">
            </div>
          </div>

          <div style="text-align: center; margin-top: 30px;">
            <button type="button" class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
            <button type="submit" class="btn">📅 Add Event</button>
          </div>
        </form>
      </div>

      <script>
        document.getElementById('eventForm').addEventListener('submit', function(e) {
          e.preventDefault();
          
          const formData = new FormData(this);
          const data = {};
          for (let [key, value] of formData.entries()) {
            data[key] = value;
          }
          
          // Validate required fields
          if (!data.title || !data.category || !data.date || !data.time || !data.location) {
            alert('Please fill in all required fields (marked with *)');
            return;
          }
          
          google.script.run
            .withSuccessHandler(function(result) {
              alert('✅ Event added successfully!');
              google.script.host.close();
            })
            .withFailureHandler(function(error) {
              alert('❌ Error: ' + error);
            })
            .processNewEvent(data);
        });
      </script>
    </body>
    </html>
  `)
  .setWidth(650)
  .setHeight(800);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '📅 Add New Event');
}

function processNewEvent(data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('📅 Events');
    
    const newRow = [
      generateEventId(),
      data.title || '',
      data.category || 'Social',
      data.date || '',
      data.time || '',
      data.location || '',
      data.address || '',
      data.description || '',
      Session.getActiveUser().getEmail(), // Organizer
      data.contact || '',
      data.registrationLink || '',
      data.cost || 'Free',
      data.maxParticipants || '',
      '0', // Current Participants
      'Upcoming', // Status
      data.tags || '',
      new Date(), // Created Date
      '', // Attendees List
      '0' // Waitlist Count
    ];
    
    sheet.appendRow(newRow);
    
    return true;
  } catch (error) {
    throw new Error('Failed to add event: ' + error.message);
  }
}

function registerForEvent() {
  const htmlOutput = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }
        .container { max-width: 600px; margin: 0 auto; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
        h2 { color: #2c3e50; text-align: center; margin-bottom: 30px; }
        .form-group { margin-bottom: 20px; }
        label { display: block; margin-bottom: 5px; font-weight: bold; color: #34495e; }
        input, select, textarea { width: 100%; padding: 12px; border: 2px solid #ddd; border-radius: 5px; font-size: 14px; box-sizing: border-box; }
        input:focus, select:focus, textarea:focus { border-color: #3498db; outline: none; }
        textarea { height: 80px; resize: vertical; }
        .row { display: flex; gap: 15px; }
        .col { flex: 1; }
        .btn { background: #3498db; color: white; padding: 12px 24px; border: none; border-radius: 5px; cursor: pointer; font-size: 16px; margin: 5px; }
        .btn:hover { background: #2980b9; }
        .btn-secondary { background: #95a5a6; }
        .btn-secondary:hover { background: #7f8c8d; }
        .required { color: #e74c3c; }
        .section { background: #ecf0f1; padding: 20px; border-radius: 8px; margin-bottom: 20px; }
        .section h3 { margin-top: 0; color: #2c3e50; }
        .info-box { background: #e8f4fd; border-left: 4px solid #3498db; padding: 15px; margin-bottom: 20px; border-radius: 5px; }
      </style>
    </head>
    <body>
      <div class="container">
        <h2>📅 Register for Event</h2>
        
        <div class="info-box">
          <strong>ℹ️ How to register:</strong><br>
          1. Find the Event ID from the Events sheet<br>
          2. Fill in your personal information<br>
          3. Submit your registration
        </div>
        
        <form id="registrationForm">
          <div class="section">
            <h3>📋 Event Information</h3>
            <div class="row">
              <div class="col">
                <label for="eventId">Event ID <span class="required">*</span></label>
                <input type="text" id="eventId" name="eventId" required placeholder="e.g., EVT-ABC123">
              </div>
              <div class="col">
                <label for="eventTitle">Event Title <span class="required">*</span></label>
                <input type="text" id="eventTitle" name="eventTitle" required placeholder="e.g., Chinese New Year Celebration">
              </div>
            </div>
          </div>

          <div class="section">
            <h3>👤 Personal Information</h3>
            <div class="row">
              <div class="col">
                <label for="attendeeName">Your Name <span class="required">*</span></label>
                <input type="text" id="attendeeName" name="attendeeName" required placeholder="e.g., John Doe">
              </div>
              <div class="col">
                <label for="attendeeEmail">Your Email <span class="required">*</span></label>
                <input type="email" id="attendeeEmail" name="attendeeEmail" required placeholder="e.g., john@email.com">
              </div>
            </div>
            <div class="form-group">
              <label for="attendeePhone">Phone Number</label>
              <input type="tel" id="attendeePhone" name="attendeePhone" placeholder="e.g., +49 30 123456">
            </div>
          </div>

          <div class="section">
            <h3>📝 Additional Information</h3>
            <div class="form-group">
              <label for="specialRequirements">Special Requirements</label>
              <textarea id="specialRequirements" name="specialRequirements" placeholder="Any dietary restrictions, accessibility needs, or other special requirements"></textarea>
            </div>
            <div class="row">
              <div class="col">
                <label for="howDidYouHear">How did you hear about this event?</label>
                <select id="howDidYouHear" name="howDidYouHear">
                  <option value="">Select option</option>
                  <option value="X.Lab Center">X.Lab Center</option>
                  <option value="Social Media">Social Media</option>
                  <option value="Friend">Friend</option>
                  <option value="Email">Email</option>
                  <option value="Other">Other</option>
                </select>
              </div>
              <div class="col">
                <label for="willBringGuests">Will you bring guests?</label>
                <select id="willBringGuests" name="willBringGuests">
                  <option value="No">No</option>
                  <option value="Yes">Yes</option>
                </select>
              </div>
            </div>
          </div>

          <div style="text-align: center; margin-top: 30px;">
            <button type="button" class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
            <button type="submit" class="btn">📅 Register for Event</button>
          </div>
        </form>
      </div>

      <script>
        document.getElementById('registrationForm').addEventListener('submit', function(e) {
          e.preventDefault();
          
          const formData = new FormData(this);
          const data = {};
          for (let [key, value] of formData.entries()) {
            data[key] = value;
          }
          
          // Validate required fields
          if (!data.eventId || !data.eventTitle || !data.attendeeName || !data.attendeeEmail) {
            alert('Please fill in all required fields (marked with *)');
            return;
          }
          
          google.script.run
            .withSuccessHandler(function(result) {
              if (result.success) {
                alert('✅ Successfully registered for the event!');
              } else {
                alert('⚠️ ' + result.message);
              }
              google.script.host.close();
            })
            .withFailureHandler(function(error) {
              alert('❌ Error: ' + error);
            })
            .processEventRegistration(data);
        });
      </script>
    </body>
    </html>
  `)
  .setWidth(650)
  .setHeight(800);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '📅 Register for Event');
}

function processEventRegistration(data) {
  try {
    const eventId = data.eventId;
    const eventTitle = data.eventTitle;
    const attendeeName = data.attendeeName;
    const attendeeEmail = data.attendeeEmail;
    const attendeePhone = data.attendeePhone || '';
    
    // Update event with new attendee
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('📅 Events');
    const eventData = sheet.getDataRange().getValues();
    
    for (let i = 1; i < eventData.length; i++) {
      if (eventData[i][0] === eventId) {
        const currentParticipants = parseInt(eventData[i][13]) || 0;
        const maxParticipants = parseInt(eventData[i][12]) || 999;
        const attendeesList = eventData[i][17] || '';
        
        if (currentParticipants < maxParticipants) {
          // Add to attendees list
          const newAttendeesList = attendeesList ? attendeesList + '; ' + attendeeName : attendeeName;
          
          sheet.getRange(i + 1, 14).setValue(currentParticipants + 1); // Current Participants
          sheet.getRange(i + 1, 18).setValue(newAttendeesList); // Attendees List
          
          return { success: true, message: 'Successfully registered for the event!' };
        } else {
          // Add to waitlist
          const waitlistCount = parseInt(eventData[i][19]) || 0;
          sheet.getRange(i + 1, 20).setValue(waitlistCount + 1); // Waitlist Count
          
          return { success: false, message: 'Event is full. You have been added to the waitlist.' };
        }
      }
    }
    
    return { success: false, message: 'Event not found. Please check the Event ID.' };
  } catch (error) {
    throw new Error('Failed to register for event: ' + error.message);
  }
}

function viewUpcomingEvents() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('📅 Events');
  const data = sheet.getDataRange().getValues();
  const today = new Date();
  
  const upcomingEvents = data.filter(row => {
    if (row[3] && row[15] === 'Upcoming') {
      const eventDate = new Date(row[3]);
      return eventDate >= today;
    }
    return false;
  });
  
  if (upcomingEvents.length > 0) {
    let message = `Upcoming Events:\n\n`;
    upcomingEvents.forEach((event, index) => {
      message += `${index + 1}. ${event[1]} - ${event[3]} at ${event[4]}\n   Location: ${event[5]}\n\n`;
    });
    SpreadsheetApp.getUi().alert('Upcoming Events', message, SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    SpreadsheetApp.getUi().alert('No Upcoming Events', 'No upcoming events found.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function joinCommunity() {
  const htmlOutput = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }
        .container { max-width: 600px; margin: 0 auto; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
        h2 { color: #2c3e50; text-align: center; margin-bottom: 30px; }
        .form-group { margin-bottom: 20px; }
        label { display: block; margin-bottom: 5px; font-weight: bold; color: #34495e; }
        input, select, textarea { width: 100%; padding: 12px; border: 2px solid #ddd; border-radius: 5px; font-size: 14px; box-sizing: border-box; }
        input:focus, select:focus, textarea:focus { border-color: #3498db; outline: none; }
        textarea { height: 80px; resize: vertical; }
        .row { display: flex; gap: 15px; }
        .col { flex: 1; }
        .btn { background: #3498db; color: white; padding: 12px 24px; border: none; border-radius: 5px; cursor: pointer; font-size: 16px; margin: 5px; }
        .btn:hover { background: #2980b9; }
        .btn-secondary { background: #95a5a6; }
        .btn-secondary:hover { background: #7f8c8d; }
        .required { color: #e74c3c; }
        .section { background: #ecf0f1; padding: 20px; border-radius: 8px; margin-bottom: 20px; }
        .section h3 { margin-top: 0; color: #2c3e50; }
        .welcome-box { background: #e8f4fd; border-left: 4px solid #3498db; padding: 15px; margin-bottom: 20px; border-radius: 5px; }
      </style>
    </head>
    <body>
      <div class="container">
        <h2>👥 Join X.Lab Community</h2>
        
        <div class="welcome-box">
          <strong>🎉 Welcome to X.Lab Chinese Student Community!</strong><br>
          Connect with fellow Chinese students, share experiences, and build your network.
        </div>
        
        <form id="communityForm">
          <div class="section">
            <h3>👤 Personal Information</h3>
            <div class="row">
              <div class="col">
                <label for="name">Full Name <span class="required">*</span></label>
                <input type="text" id="name" name="name" required placeholder="e.g., John Doe">
              </div>
              <div class="col">
                <label for="email">Email <span class="required">*</span></label>
                <input type="email" id="email" name="email" required placeholder="e.g., john@email.com">
              </div>
            </div>
            <div class="form-group">
              <label for="phone">Phone Number</label>
              <input type="tel" id="phone" name="phone" placeholder="e.g., +49 30 123456">
            </div>
          </div>

          <div class="section">
            <h3>🎓 Academic Information</h3>
            <div class="row">
              <div class="col">
                <label for="studyProgram">Study Program <span class="required">*</span></label>
                <input type="text" id="studyProgram" name="studyProgram" required placeholder="e.g., Computer Science">
              </div>
              <div class="col">
                <label for="year">Year of Study</label>
                <select id="year" name="year">
                  <option value="1st Year">1st Year</option>
                  <option value="2nd Year">2nd Year</option>
                  <option value="3rd Year">3rd Year</option>
                  <option value="4th Year">4th Year</option>
                  <option value="Graduate">Graduate</option>
                  <option value="PhD">PhD</option>
                </select>
              </div>
            </div>
          </div>

          <div class="section">
            <h3>💼 Skills & Interests</h3>
            <div class="form-group">
              <label for="skills">Skills & Expertise</label>
              <textarea id="skills" name="skills" placeholder="e.g., Programming, Design, Languages, Leadership, etc."></textarea>
            </div>
            <div class="form-group">
              <label for="interests">Interests & Hobbies</label>
              <textarea id="interests" name="interests" placeholder="e.g., Photography, Sports, Music, Travel, etc."></textarea>
            </div>
          </div>

          <div class="section">
            <h3>🌍 Languages & Background</h3>
            <div class="row">
              <div class="col">
                <label for="languages">Languages Spoken</label>
                <input type="text" id="languages" name="languages" placeholder="e.g., Chinese, English, German">
              </div>
              <div class="col">
                <label for="hometown">Hometown</label>
                <input type="text" id="hometown" name="hometown" placeholder="e.g., Beijing, China">
              </div>
            </div>
          </div>

          <div class="section">
            <h3>🤝 Mentorship</h3>
            <div class="row">
              <div class="col">
                <label for="mentorStatus">Available as Mentor</label>
                <select id="mentorStatus" name="mentorStatus">
                  <option value="Available">Available</option>
                  <option value="Not Available">Not Available</option>
                  <option value="Maybe">Maybe</option>
                </select>
              </div>
              <div class="col">
                <label for="menteeStatus">Looking for Mentor</label>
                <select id="menteeStatus" name="menteeStatus">
                  <option value="Available">Available</option>
                  <option value="Not Available">Not Available</option>
                  <option value="Maybe">Maybe</option>
                </select>
              </div>
            </div>
            <div class="form-group">
              <label for="availability">Availability</label>
              <select id="availability" name="availability">
                <option value="Available">Available for activities</option>
                <option value="Busy">Currently busy</option>
                <option value="Selective">Selective participation</option>
              </select>
            </div>
          </div>

          <div class="section">
            <h3>📱 Social Links</h3>
            <div class="form-group">
              <label for="socialLinks">Social Media Links</label>
              <input type="text" id="socialLinks" name="socialLinks" placeholder="e.g., LinkedIn, WeChat, Instagram">
            </div>
            <div class="form-group">
              <label for="notes">Additional Notes</label>
              <textarea id="notes" name="notes" placeholder="Any additional information you'd like to share with the community"></textarea>
            </div>
          </div>

          <div style="text-align: center; margin-top: 30px;">
            <button type="button" class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
            <button type="submit" class="btn">👥 Join Community</button>
          </div>
        </form>
      </div>

      <script>
        document.getElementById('communityForm').addEventListener('submit', function(e) {
          e.preventDefault();
          
          const formData = new FormData(this);
          const data = {};
          for (let [key, value] of formData.entries()) {
            data[key] = value;
          }
          
          // Validate required fields
          if (!data.name || !data.email || !data.studyProgram) {
            alert('Please fill in all required fields (marked with *)');
            return;
          }
          
          google.script.run
            .withSuccessHandler(function(result) {
              alert('✅ Welcome to the X.Lab Community! You have been successfully added.');
              google.script.host.close();
            })
            .withFailureHandler(function(error) {
              alert('❌ Error: ' + error);
            })
            .processCommunityJoin(data);
        });
      </script>
    </body>
    </html>
  `)
  .setWidth(650)
  .setHeight(800);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '👥 Join Community');
}

function processCommunityJoin(data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('👥 Community');
    
    const newRow = [
      generateMemberId(),
      data.name || '',
      data.email || '',
      data.phone || '',
      data.studyProgram || '',
      data.year || '1st Year',
      data.skills || '',
      data.interests || '',
      data.languages || 'Chinese,English',
      data.mentorStatus || 'Available',
      data.menteeStatus || 'Available',
      data.socialLinks || '',
      data.availability || 'Available',
      data.notes || '',
      new Date() // Join Date
    ];
    
    sheet.appendRow(newRow);
    
    return true;
  } catch (error) {
    throw new Error('Failed to join community: ' + error.message);
  }
}

// ========================================
// 🆔 ID GENERATION FUNCTIONS
// ========================================

function generatePlaceId() {
  return 'PLACE-' + Math.random().toString(36).substr(2, 9).toUpperCase();
}

function generateEventId() {
  return 'EVENT-' + Math.random().toString(36).substr(2, 9).toUpperCase();
}

function generateMemberId() {
  return 'MEMBER-' + Math.random().toString(36).substr(2, 9).toUpperCase();
}

// ========================================
// 📊 STATISTICS CALCULATION FUNCTIONS
// ========================================

function calculateQuickStats() {
  try {
    const stats = {
      activeJobPostings: getActiveJobPostingsCount(),
      totalApplications: getTotalApplicationsCount(),
      upcomingEvents: getUpcomingEventsCount(),
      communityMembers: getCommunityMembersCount()
    };
    
    return stats;
  } catch (error) {
    console.log('Error calculating quick stats: ' + error);
    return {
      activeJobPostings: 0,
      totalApplications: 0,
      upcomingEvents: 0,
      communityMembers: 0
    };
  }
}

function getActiveJobPostingsCount() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('💼 Job Postings');
    if (!sheet) return 0;
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return 0; // No data beyond headers
    
    let activeCount = 0;
    for (let i = 1; i < data.length; i++) {
      const status = data[i][12]; // Status column
      if (status === 'Active') {
        activeCount++;
      }
    }
    
    return activeCount;
  } catch (error) {
    console.log('Error getting active job postings count: ' + error);
    return 0;
  }
}

function getTotalApplicationsCount() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('📝 Job Applications');
    if (!sheet) return 0;
    
    const data = sheet.getDataRange().getValues();
    return Math.max(0, data.length - 1); // Subtract header row
  } catch (error) {
    console.log('Error getting total applications count: ' + error);
    return 0;
  }
}

function getUpcomingEventsCount() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('📅 Events');
    if (!sheet) return 0;
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return 0; // No data beyond headers
    
    const today = new Date();
    today.setHours(0, 0, 0, 0); // Set to start of today
    
    let upcomingCount = 0;
    for (let i = 1; i < data.length; i++) {
      const eventDate = data[i][2]; // Date column
      if (eventDate instanceof Date && eventDate >= today) {
        upcomingCount++;
      }
    }
    
    return upcomingCount;
  } catch (error) {
    console.log('Error getting upcoming events count: ' + error);
    return 0;
  }
}

function getCommunityMembersCount() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('👥 Community');
    if (!sheet) return 0;
    
    const data = sheet.getDataRange().getValues();
    return Math.max(0, data.length - 1); // Subtract header row
  } catch (error) {
    console.log('Error getting community members count: ' + error);
    return 0;
  }
}

function generateOpportunityId() {
  return 'OPP-' + Math.random().toString(36).substr(2, 9).toUpperCase();
}

// Function to refresh dashboard statistics
function refreshDashboardStats() {
  try {
    createInnovativeDashboard(); // This will recalculate and display fresh stats
    SpreadsheetApp.getUi().alert('✅ Dashboard Refreshed', 'Statistics have been updated with the latest data!', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Refresh Error', 'Failed to refresh dashboard: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// Function to show current statistics in a dialog
function showDetailedStats() {
  const stats = calculateQuickStats();
  
  const message = `📊 Detailed Statistics:

💼 Job Market:
• Active Job Postings: ${stats.activeJobPostings}
• Total Applications: ${stats.totalApplications}

📅 Events:
• Upcoming Events: ${stats.upcomingEvents}

👥 Community:
• Total Members: ${stats.communityMembers}

📈 Last Updated: ${new Date().toLocaleString()}`;

  SpreadsheetApp.getUi().alert('📊 System Statistics', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ========================================
// 🔐 PERMISSION AND ACCESS MANAGEMENT
// ========================================

/**
 * Check if the current user has the necessary permissions
 */
function checkPermissions() {
  try {
    // Try to access the spreadsheet
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName('🎯 Dashboard');
    
    if (!sheet) {
      throw new Error('Dashboard sheet not found');
    }
    
    // Try to read a cell (minimal operation)
    sheet.getRange('A1').getValue();
    
    return { hasPermission: true, message: 'Permissions OK' };
  } catch (error) {
    return { 
      hasPermission: false, 
      message: error.message,
      isPermissionError: error.message.includes('permission') || error.message.includes('auth')
    };
  }
}

/**
 * Handle permission errors with user-friendly messages
 */
function handlePermissionError(operation = 'perform this action') {
  const errorMessage = `🔐 Permission Required

To ${operation}, you need access to this spreadsheet.

Steps to fix:
1. Ask the spreadsheet owner to share it with you
2. Make sure you have "Editor" access (not just "Viewer")
3. If you're the owner, try refreshing the page
4. Check that you're signed in to the correct Google account

Contact the system administrator if the problem persists.`;

  SpreadsheetApp.getUi().alert('🔐 Access Denied', errorMessage, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Wrapper function to safely execute operations with permission checking
 */
function safeExecute(operation, operationName, callback) {
  try {
    const permissionCheck = checkPermissions();
    
    if (!permissionCheck.hasPermission) {
      if (permissionCheck.isPermissionError) {
        handlePermissionError(operationName);
      } else {
        throw new Error(permissionCheck.message);
      }
      return false;
    }
    
    return callback();
  } catch (error) {
    console.error(`Error in ${operationName}:`, error);
    
    if (error.message.includes('permission') || error.message.includes('auth')) {
      handlePermissionError(operationName);
    } else {
      SpreadsheetApp.getUi().alert(
        '❌ Error', 
        `Failed to ${operationName}:\n${error.message}`, 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return false;
  }
}

// ========================================
// 📊 ENHANCED ANALYTICS FUNCTIONS
// ========================================

function updateAnalytics() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('📊 Analytics');
  
  // Update real-time statistics
  const jobSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('💼 Job Postings');
  const seekerSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('🔍 Job Seeking');
  const eventSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('📅 Events');
  const communitySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('👥 Community');
  
  // Count active job postings
  const jobData = jobSheet.getDataRange().getValues();
  const activeJobs = jobData.filter(row => row[12] === 'Active').length;
  
  // Count active job seekers
  const seekerData = seekerSheet.getDataRange().getValues();
  const activeSeekers = seekerData.filter(row => row[15] === 'Active').length;
  
  // Count upcoming events
  const eventData = eventSheet.getDataRange().getValues();
  const today = new Date();
  const upcomingEvents = eventData.filter(row => {
    if (row[3] && row[14] === 'Upcoming') {
      const eventDate = new Date(row[3]);
      return eventDate >= today;
    }
    return false;
  }).length;
  
  // Count community members
  const communityData = communitySheet.getDataRange().getValues();
  const communityMembers = communityData.length - 1; // Subtract header row
  
  // Update dashboard stats
  sheet.getRange('B4').setValue(activeJobs);
  sheet.getRange('B5').setValue(activeSeekers);
  sheet.getRange('B6').setValue(upcomingEvents);
  sheet.getRange('B7').setValue(communityMembers);
}

// ========================================
// 🔄 AUTOMATION FUNCTIONS
// ========================================

function setupTriggers() {
  // Delete existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  
  // Create new triggers
  ScriptApp.newTrigger('updateAnalytics')
    .timeBased()
    .everyMinutes(30)
    .create();
  
  ScriptApp.newTrigger('sendEventReminder')
    .timeBased()
    .dailyAt(9, 0) // 9 AM daily
    .create();
  
  ScriptApp.newTrigger('onOpen')
    .onOpen()
    .create();
  
  // Set up onSelectionChange trigger for Quick Actions
  ScriptApp.newTrigger('onSelectionChange')
    .onSelectionChange()
    .create();
}

// ========================================
// 🎯 QUICK ACTIONS HANDLER
// ========================================

function onSelectionChange(e) {
  try {
    const range = e.range;
    const sheet = range.getSheet();
    
    // Only handle selections in the Dashboard sheet
    if (sheet.getName() !== '🎯 Dashboard') {
      return;
    }
    
    const row = range.getRow();
    const col = range.getColumn();
    
    // Quick Actions removed from cells - no longer needed
    // All quick actions are now available in the native menu only
    
  } catch (error) {
    console.log('Error in onSelectionChange:', error);
  }
}

// ========================================
// 📧 EMAIL NOTIFICATION FUNCTIONS
// ========================================

function sendEmailNotification(recipient, subject, message) {
  try {
    GmailApp.sendEmail(recipient, subject, message);
    return true;
  } catch (error) {
    console.error('Email sending failed:', error);
    return false;
  }
}

function notifyNewJobPosting(jobData) {
  const subject = '🎯 New Job Opportunity at X.lab Chinese Student Center';
  const message = `
    Hello Chinese Students!
    
    A new job opportunity has been posted:
    
    Position: ${jobData[1]}
    Company: ${jobData[2]}
    Location: ${jobData[3]}
    Type: ${jobData[4]}
    
    Check it out in the X.lab Chinese Student Center!
    
    Best regards,
    X.lab Team
  `;
  
  // Get all community members
  const communitySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('👥 Community');
  const members = communitySheet.getDataRange().getValues();
  
  members.forEach(member => {
    if (member[2]) { // Email exists
      sendEmailNotification(member[2], subject, message);
    }
  });
}

// ========================================
// 🌐 WEBSITE INTEGRATION FUNCTIONS
// ========================================

function addUsefulWebsite() {
  const htmlOutput = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }
        .container { max-width: 600px; margin: 0 auto; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
        h2 { color: #2c3e50; text-align: center; margin-bottom: 30px; }
        .form-group { margin-bottom: 20px; }
        label { display: block; margin-bottom: 5px; font-weight: bold; color: #34495e; }
        input, select, textarea { width: 100%; padding: 12px; border: 2px solid #ddd; border-radius: 5px; font-size: 14px; box-sizing: border-box; }
        input:focus, select:focus, textarea:focus { border-color: #3498db; outline: none; }
        textarea { height: 80px; resize: vertical; }
        .row { display: flex; gap: 15px; }
        .col { flex: 1; }
        .btn { background: #3498db; color: white; padding: 12px 24px; border: none; border-radius: 5px; cursor: pointer; font-size: 16px; margin: 5px; }
        .btn:hover { background: #2980b9; }
        .btn-secondary { background: #95a5a6; }
        .btn-secondary:hover { background: #7f8c8d; }
        .required { color: #e74c3c; }
        .section { background: #ecf0f1; padding: 20px; border-radius: 8px; margin-bottom: 20px; }
        .section h3 { margin-top: 0; color: #2c3e50; }
      </style>
    </head>
    <body>
      <div class="container">
        <h2>🌐 Add Useful Website</h2>
        <form id="websiteForm">
          <div class="section">
            <h3>📋 Basic Information</h3>
            <div class="form-group">
              <label for="name">Website Name <span class="required">*</span></label>
              <input type="text" id="name" name="name" required placeholder="e.g., Google Translate">
            </div>
            <div class="row">
              <div class="col">
                <label for="category">Category <span class="required">*</span></label>
                <select id="category" name="category" required>
                  <option value="">Select Category</option>
                  <option value="Tools">Tools</option>
                  <option value="Learning">Learning</option>
                  <option value="Translation">Translation</option>
                  <option value="Transportation">Transportation</option>
                  <option value="Shopping">Shopping</option>
                  <option value="Entertainment">Entertainment</option>
                  <option value="News">News</option>
                  <option value="Social">Social</option>
                  <option value="Finance">Finance</option>
                  <option value="Health">Health</option>
                  <option value="Other">Other</option>
                </select>
              </div>
              <div class="col">
                <label for="language">Language</label>
                <select id="language" name="language">
                  <option value="Multilingual">Multilingual</option>
                  <option value="Chinese">Chinese</option>
                  <option value="English">English</option>
                  <option value="German">German</option>
                  <option value="Other">Other</option>
                </select>
              </div>
            </div>
          </div>

          <div class="section">
            <h3>🔗 Website Details</h3>
            <div class="form-group">
              <label for="url">Website URL <span class="required">*</span></label>
              <input type="url" id="url" name="url" required placeholder="e.g., https://translate.google.com">
            </div>
            <div class="form-group">
              <label for="description">Description</label>
              <textarea id="description" name="description" placeholder="Describe what this website offers and why it's useful for Chinese students"></textarea>
            </div>
          </div>

          <div class="section">
            <h3>🏷️ Additional Information</h3>
            <div class="row">
              <div class="col">
                <label for="studentSpecific">Student Specific</label>
                <select id="studentSpecific" name="studentSpecific">
                  <option value="Yes">Yes</option>
                  <option value="No">No</option>
                  <option value="Maybe">Maybe</option>
                </select>
              </div>
              <div class="col">
                <label for="rating">Rating</label>
                <select id="rating" name="rating">
                  <option value="">No rating</option>
                  <option value="5">⭐⭐⭐⭐⭐ (5/5)</option>
                  <option value="4">⭐⭐⭐⭐ (4/5)</option>
                  <option value="3">⭐⭐⭐ (3/5)</option>
                  <option value="2">⭐⭐ (2/5)</option>
                  <option value="1">⭐ (1/5)</option>
                </select>
              </div>
            </div>
            <div class="form-group">
              <label for="tags">Tags</label>
              <input type="text" id="tags" name="tags" placeholder="e.g., Free, Translation, Mobile App, Offline">
            </div>
          </div>

          <div style="text-align: center; margin-top: 30px;">
            <button type="button" class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
            <button type="submit" class="btn">🌐 Add Website</button>
          </div>
        </form>
      </div>

      <script>
        document.getElementById('websiteForm').addEventListener('submit', function(e) {
          e.preventDefault();
          
          const formData = new FormData(this);
          const data = {};
          for (let [key, value] of formData.entries()) {
            data[key] = value;
          }
          
          // Validate required fields
          if (!data.name || !data.category || !data.url) {
            alert('Please fill in all required fields (marked with *)');
            return;
          }
          
          google.script.run
            .withSuccessHandler(function(result) {
              alert('✅ Website added successfully!');
              google.script.host.close();
            })
            .withFailureHandler(function(error) {
              alert('❌ Error: ' + error);
            })
            .processNewWebsite(data);
        });
      </script>
    </body>
    </html>
  `)
  .setWidth(650)
  .setHeight(700);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '🌐 Add Useful Website');
}

function processNewWebsite(data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('🌐 Useful Websites');
    
    const newRow = [
      generateWebsiteId(),
      data.name || '',
      data.category || 'Tools',
      data.url || '',
      data.description || '',
      data.language || 'Multilingual',
      data.studentSpecific || 'Yes',
      data.rating || '',
      data.tags || '',
      new Date(), // Last Updated
      new Date() // Created Date
    ];
    
    sheet.appendRow(newRow);
    
    return true;
  } catch (error) {
    throw new Error('Failed to add website: ' + error.message);
  }
}

function generateWebsiteId() {
  return 'WEB-' + Math.random().toString(36).substr(2, 9).toUpperCase();
}

// ========================================
// 💡 BUSINESS OPPORTUNITY FUNCTIONS
// ========================================

function addBusinessOpportunity() {
  const htmlOutput = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }
        .container { max-width: 600px; margin: 0 auto; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
        h2 { color: #2c3e50; text-align: center; margin-bottom: 30px; }
        .form-group { margin-bottom: 20px; }
        label { display: block; margin-bottom: 5px; font-weight: bold; color: #34495e; }
        input, select, textarea { width: 100%; padding: 12px; border: 2px solid #ddd; border-radius: 5px; font-size: 14px; box-sizing: border-box; }
        input:focus, select:focus, textarea:focus { border-color: #3498db; outline: none; }
        textarea { height: 80px; resize: vertical; }
        .row { display: flex; gap: 15px; }
        .col { flex: 1; }
        .btn { background: #3498db; color: white; padding: 12px 24px; border: none; border-radius: 5px; cursor: pointer; font-size: 16px; margin: 5px; }
        .btn:hover { background: #2980b9; }
        .btn-secondary { background: #95a5a6; }
        .btn-secondary:hover { background: #7f8c8d; }
        .required { color: #e74c3c; }
        .section { background: #ecf0f1; padding: 20px; border-radius: 8px; margin-bottom: 20px; }
        .section h3 { margin-top: 0; color: #2c3e50; }
      </style>
    </head>
    <body>
      <div class="container">
        <h2>💡 Add Business Opportunity</h2>
        <form id="opportunityForm">
          <div class="section">
            <h3>📋 Basic Information</h3>
            <div class="form-group">
              <label for="title">Opportunity Title <span class="required">*</span></label>
              <input type="text" id="title" name="title" required placeholder="e.g., Chinese Food Delivery Service">
            </div>
            <div class="row">
              <div class="col">
                <label for="category">Category <span class="required">*</span></label>
                <select id="category" name="category" required>
                  <option value="">Select Category</option>
                  <option value="Startup">Startup</option>
                  <option value="Partnership">Partnership</option>
                  <option value="Investment">Investment</option>
                  <option value="Franchise">Franchise</option>
                  <option value="Consulting">Consulting</option>
                  <option value="Import/Export">Import/Export</option>
                  <option value="Technology">Technology</option>
                  <option value="Education">Education</option>
                  <option value="Other">Other</option>
                </select>
              </div>
              <div class="col">
                <label for="status">Status</label>
                <select id="status" name="status">
                  <option value="Open">Open</option>
                  <option value="In Progress">In Progress</option>
                  <option value="Closed">Closed</option>
                  <option value="Expired">Expired</option>
                </select>
              </div>
            </div>
          </div>

          <div class="section">
            <h3>📝 Opportunity Details</h3>
            <div class="form-group">
              <label for="description">Description <span class="required">*</span></label>
              <textarea id="description" name="description" required placeholder="Describe the business opportunity, market potential, and requirements"></textarea>
            </div>
            <div class="row">
              <div class="col">
                <label for="investmentRequired">Investment Required</label>
                <input type="text" id="investmentRequired" name="investmentRequired" placeholder="e.g., €10,000 - €50,000">
              </div>
              <div class="col">
                <label for="expectedReturn">Expected Return</label>
                <input type="text" id="expectedReturn" name="expectedReturn" placeholder="e.g., 15-25% annually">
              </div>
            </div>
          </div>

          <div class="section">
            <h3>📍 Location & Timeline</h3>
            <div class="row">
              <div class="col">
                <label for="location">Location</label>
                <input type="text" id="location" name="location" placeholder="e.g., Berlin, Germany">
              </div>
              <div class="col">
                <label for="timeline">Timeline</label>
                <input type="text" id="timeline" name="timeline" placeholder="e.g., 6-12 months">
              </div>
            </div>
            <div class="form-group">
              <label for="deadline">Application Deadline</label>
              <input type="date" id="deadline" name="deadline">
            </div>
          </div>

          <div class="section">
            <h3>📞 Contact Information</h3>
            <div class="row">
              <div class="col">
                <label for="contactPerson">Contact Person</label>
                <input type="text" id="contactPerson" name="contactPerson" placeholder="e.g., John Doe">
              </div>
              <div class="col">
                <label for="contactEmail">Contact Email</label>
                <input type="email" id="contactEmail" name="contactEmail" placeholder="e.g., contact@business.com">
              </div>
            </div>
          </div>

          <div class="section">
            <h3>🏷️ Additional Information</h3>
            <div class="form-group">
              <label for="tags">Tags</label>
              <input type="text" id="tags" name="tags" placeholder="e.g., Chinese Market, Food, Technology, Student-Friendly">
            </div>
          </div>

          <div style="text-align: center; margin-top: 30px;">
            <button type="button" class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
            <button type="submit" class="btn">💡 Add Opportunity</button>
          </div>
        </form>
      </div>

      <script>
        document.getElementById('opportunityForm').addEventListener('submit', function(e) {
          e.preventDefault();
          
          const formData = new FormData(this);
          const data = {};
          for (let [key, value] of formData.entries()) {
            data[key] = value;
          }
          
          // Validate required fields
          if (!data.title || !data.category || !data.description) {
            alert('Please fill in all required fields (marked with *)');
            return;
          }
          
          google.script.run
            .withSuccessHandler(function(result) {
              alert('✅ Business opportunity added successfully!');
              google.script.host.close();
            })
            .withFailureHandler(function(error) {
              alert('❌ Error: ' + error);
            })
            .processNewOpportunity(data);
        });
      </script>
    </body>
    </html>
  `)
  .setWidth(650)
  .setHeight(800);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '💡 Add Business Opportunity');
}

function processNewOpportunity(data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('💡 Business Opportunities');
    
    // Validate and process deadline
    let deadline = '';
    if (data.deadline) {
      try {
        deadline = new Date(data.deadline);
      } catch (e) {
        deadline = '';
      }
    }
    
    const newRow = [
      generateOpportunityId(),
      data.title || '',
      data.category || 'Startup',
      data.description || '',
      data.investmentRequired || '',
      data.expectedReturn || '',
      data.timeline || '',
      data.location || '',
      data.contactPerson || Session.getActiveUser().getEmail(),
      data.contactEmail || '',
      data.status || 'Open',
      data.tags || '',
      new Date(), // Created Date
      deadline // Deadline
    ];
    
    sheet.appendRow(newRow);
    
    return true;
  } catch (error) {
    throw new Error('Failed to add business opportunity: ' + error.message);
  }
}

function findMentor() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('👥 Community');
  const data = sheet.getDataRange().getValues();
  
  const availableMentors = data.filter(row => 
    row[9] === 'Available' && row[1] // Mentor Status is Available and has a name
  );
  
  if (availableMentors.length > 0) {
    let message = `Available Mentors:\n\n`;
    availableMentors.forEach((mentor, index) => {
      message += `${index + 1}. ${mentor[1]} (${mentor[4]}, ${mentor[5]})\n   Skills: ${mentor[6]}\n   Languages: ${mentor[8]}\n\n`;
    });
    SpreadsheetApp.getUi().alert('Available Mentors', message, SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    SpreadsheetApp.getUi().alert('No Mentors Available', 'No mentors are currently available.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function becomeMentor() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt('Become a Mentor', 
    'Are you sure you want to become a mentor? (yes/no):', 
    ui.ButtonSet.OK_CANCEL);
  
  if (result.getSelectedButton() === ui.Button.OK && result.getResponseText().toLowerCase() === 'yes') {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('👥 Community');
    const userEmail = Session.getActiveUser().getEmail();
    
    // Find the user's row and update mentor status
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][2] === userEmail) { // Email column
        sheet.getRange(i + 1, 10).setValue('Available'); // Mentor Status column
        ui.alert('✅ You are now a mentor! Students can find you in the mentor directory.');
        return;
      }
    }
    
    ui.alert('❌ User not found. Please join the community first.');
  }
}

// ========================================
// 🚀 INITIALIZATION TRIGGER
// ========================================

function initializeChineseStudentCenter() {
  setupChineseStudentCenter();
  setupTriggers();
  
  // Force menu creation with multiple attempts
  let menuCreated = false;
  
  // Try automatic menu creation
  try {
    onOpen();
    menuCreated = true;
  } catch (error) {
    console.error('Automatic menu creation failed:', error);
  }
  
  // If automatic failed, try manual creation
  if (!menuCreated) {
    try {
      menuCreated = forceCreateMenu();
    } catch (error) {
      console.error('Manual menu creation failed:', error);
    }
  }
  
  const ui = SpreadsheetApp.getUi();
  
  if (menuCreated) {
    ui.alert(
      '🎉 X.lab Chinese Student Center Successfully Initialized!',
      'Your comprehensive platform is now ready with:\n\n' +
      '• 💼 Job posting and seeking\n' +
      '• 🗺️ Places and activities in Potsdam/Berlin\n' +
      '• 📅 Event management\n' +
      '• 🌐 Useful websites directory\n' +
      '• 💡 Business opportunities\n' +
      '• 👥 Community features\n' +
      '• 📊 Analytics dashboard\n' +
      '• ⚙️ Customizable settings\n\n' +
      '✅ The "🏛️ X.Lab Center" menu has been created!\n' +
      'Quick Actions are now available in the native menu only.\n\n' +
      'Use the native Google Sheets menu bar to navigate!',
      ui.ButtonSet.OK
    );
  } else {
    ui.alert(
      '⚠️ X.lab Chinese Student Center Initialized with Issues',
      'Your platform is ready, but the menu could not be created automatically.\n\n' +
      'Please run the "createMenuManually" function to create the menu.\n\n' +
      'All features are available through individual function calls.',
      ui.ButtonSet.OK
    );
  }
}
