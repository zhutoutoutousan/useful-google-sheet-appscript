/**
 * Client Onboarding and Style Guide Analysis Tools
 * 
 * Sophisticated tools for analyzing client style guides, extracting design tokens,
 * and automating the onboarding process for new design system integrations.
 */

/**
 * Client Onboarding Manager
 */
class ClientOnboardingManager {
  constructor() {
    this.ss = SpreadsheetApp.getActiveSpreadsheet();
    this.clientSheet = this.ss.getSheetByName(SHEET_NAMES.CLIENT_VARIABLES);
    this.projectSheet = this.ss.getSheetByName(SHEET_NAMES.PROJECT_TRACKER);
    this.protocol = new CommunicationProtocol();
    
    // Style guide analysis patterns
    this.analysisPatterns = {
      colors: {
        hex: /#[A-Fa-f0-9]{6}|#[A-Fa-f0-9]{3}/g,
        rgb: /rgb\(\s*\d+\s*,\s*\d+\s*,\s*\d+\s*\)/g,
        hsl: /hsl\(\s*\d+\s*,\s*\d+%\s*,\s*\d+%\s*\)/g,
        variables: /--[\w-]*color[\w-]*|[\w-]*color[\w-]*:|color\s*[:=]/gi
      },
      typography: {
        fontFamily: /font-family\s*[:=]\s*[^;,}]+/gi,
        fontSize: /font-size\s*[:=]\s*[\d.]+(?:px|em|rem|pt)/gi,
        fontWeight: /font-weight\s*[:=]\s*(?:\d+|normal|bold|lighter|bolder)/gi,
        lineHeight: /line-height\s*[:=]\s*[\d.]+/gi
      },
      spacing: {
        margin: /margin\s*[:=]\s*[\d.]+(?:px|em|rem)/gi,
        padding: /padding\s*[:=]\s*[\d.]+(?:px|em|rem)/gi,
        gap: /gap\s*[:=]\s*[\d.]+(?:px|em|rem)/gi,
        spacing: /--[\w-]*(?:spacing|space|gap|margin|padding)[\w-]*/gi
      }
    };
    
    // Component detection patterns
    this.componentPatterns = {
      button: /\.btn|\.button|button\s*{|\[data-.*button.*\]/gi,
      modal: /\.modal|\.dialog|\.popup|\[role="dialog"\]/gi,
      card: /\.card|\.panel|\.tile/gi,
      navigation: /\.nav|\.menu|\.header|nav\s*{/gi,
      table: /\.table|table\s*{|\.data-table/gi,
      form: /\.form|\.input|input\s*{|\.field/gi
    };
  }
  
  /**
   * Complete client onboarding process
   */
  async onboardNewClient(clientInfo) {
    const {
      clientName,
      clientEmail,
      contactPerson,
      projectManager,
      portalTypes,
      styleGuideUrl,
      figmaUrl,
      brandAssets,
      timeline,
      priority = 'MEDIUM'
    } = clientInfo;
    
    // Step 1: Create project and initialize tracking
    const project = await this.initializeClientProject(clientInfo);
    
    // Step 2: Send initial communication
    const communicationResult = await this.protocol.startClientOnboarding(
      clientName, 
      clientEmail, 
      projectManager, 
      portalTypes, 
      priority
    );
    
    // Step 3: If style guide URL provided, analyze it immediately
    let analysisResult = null;
    if (styleGuideUrl) {
      analysisResult = await this.analyzeStyleGuideFromUrl(styleGuideUrl, clientName);
    }
    
    // Step 4: Set up project milestones and tracking
    const milestones = this.createProjectMilestones(project.id, timeline);
    
    // Step 5: Generate onboarding checklist
    const checklist = this.generateOnboardingChecklist(clientInfo, analysisResult);
    
    return {
      projectId: project.id,
      communicationSent: communicationResult.emailSent,
      analysisCompleted: !!analysisResult,
      variablesFound: analysisResult ? analysisResult.totalVariables : 0,
      conflictsDetected: analysisResult ? analysisResult.conflicts.length : 0,
      nextSteps: checklist.nextSteps,
      checklist: checklist.items
    };
  }
  
  /**
   * Analyze style guide from URL or document
   */
  async analyzeStyleGuideFromUrl(url, clientName) {
    try {
      // In a real implementation, you'd fetch the URL content
      // For now, we'll simulate analysis
      const mockContent = this.getMockStyleGuideContent();
      return this.analyzeStyleGuideContent(mockContent, clientName, 'URL Analysis');
    } catch (error) {
      Logger.log(`Failed to analyze style guide from URL ${url}: ${error.message}`);
      return null;
    }
  }
  
  /**
   * Analyze uploaded style guide document
   */
  analyzeStyleGuideDocument(fileId, clientName) {
    try {
      const file = DriveApp.getFileById(fileId);
      const content = file.getBlob().getDataAsString();
      
      return this.analyzeStyleGuideContent(content, clientName, 'Document Analysis');
    } catch (error) {
      Logger.log(`Failed to analyze style guide document: ${error.message}`);
      return null;
    }
  }
  
  /**
   * Core style guide content analysis
   */
  analyzeStyleGuideContent(content, clientName, source = 'Manual Analysis') {
    const analysis = {
      source,
      timestamp: new Date(),
      variables: {
        colors: [],
        typography: [],
        spacing: [],
        components: []
      },
      conflicts: [],
      recommendations: [],
      confidence: 0
    };
    
    // Extract colors
    analysis.variables.colors = this.extractColors(content);
    
    // Extract typography
    analysis.variables.typography = this.extractTypography(content);
    
    // Extract spacing
    analysis.variables.spacing = this.extractSpacing(content);
    
    // Detect components
    analysis.variables.components = this.detectComponents(content);
    
    // Analyze potential conflicts
    analysis.conflicts = this.analyzeConflicts(analysis.variables, clientName);
    
    // Generate recommendations
    analysis.recommendations = this.generateRecommendations(analysis.variables, analysis.conflicts);
    
    // Calculate confidence score
    analysis.confidence = this.calculateAnalysisConfidence(analysis.variables);
    
    // Store extracted variables
    this.storeExtractedVariables(clientName, analysis.variables, source);
    
    // Store conflicts
    if (analysis.conflicts.length > 0) {
      this.storeConflicts(clientName, analysis.conflicts);
    }
    
    return {
      totalVariables: this.countTotalVariables(analysis.variables),
      conflicts: analysis.conflicts,
      recommendations: analysis.recommendations,
      confidence: analysis.confidence,
      analysis: analysis
    };
  }
  
  /**
   * Extract color variables from content
   */
  extractColors(content) {
    const colors = [];
    const colorPatterns = this.analysisPatterns.colors;
    
    // Extract hex colors
    const hexMatches = content.match(colorPatterns.hex) || [];
    hexMatches.forEach((hex, index) => {
      colors.push({
        name: `color-extracted-${index + 1}`,
        value: hex,
        type: 'hex',
        context: this.extractColorContext(content, hex),
        confidence: 0.8
      });
    });
    
    // Extract RGB colors
    const rgbMatches = content.match(colorPatterns.rgb) || [];
    rgbMatches.forEach((rgb, index) => {
      colors.push({
        name: `color-rgb-${index + 1}`,
        value: rgb,
        type: 'rgb',
        context: this.extractColorContext(content, rgb),
        confidence: 0.8
      });
    });
    
    // Extract color variable names
    const variableMatches = content.match(colorPatterns.variables) || [];
    variableMatches.forEach((variable, index) => {
      colors.push({
        name: this.cleanVariableName(variable),
        value: this.extractVariableValue(content, variable),
        type: 'variable',
        context: 'CSS variable',
        confidence: 0.9
      });
    });
    
    return this.deduplicateColors(colors);
  }
  
  /**
   * Extract typography variables
   */
  extractTypography(content) {
    const typography = [];
    const typoPatterns = this.analysisPatterns.typography;
    
    // Extract font families
    const fontFamilyMatches = content.match(typoPatterns.fontFamily) || [];
    fontFamilyMatches.forEach((match, index) => {
      const value = match.split(/[:=]/)[1]?.trim().replace(/[;"']/g, '');
      if (value) {
        typography.push({
          name: `font-family-${index + 1}`,
          value: value,
          type: 'font-family',
          context: 'Typography system',
          confidence: 0.9
        });
      }
    });
    
    // Extract font sizes
    const fontSizeMatches = content.match(typoPatterns.fontSize) || [];
    fontSizeMatches.forEach((match, index) => {
      const value = match.split(/[:=]/)[1]?.trim().replace(/[;"]/g, '');
      if (value) {
        typography.push({
          name: `font-size-${index + 1}`,
          value: value,
          type: 'font-size',
          context: 'Typography scale',
          confidence: 0.8
        });
      }
    });
    
    // Extract font weights
    const fontWeightMatches = content.match(typoPatterns.fontWeight) || [];
    fontWeightMatches.forEach((match, index) => {
      const value = match.split(/[:=]/)[1]?.trim().replace(/[;"]/g, '');
      if (value) {
        typography.push({
          name: `font-weight-${index + 1}`,
          value: value,
          type: 'font-weight',
          context: 'Typography system',
          confidence: 0.8
        });
      }
    });
    
    return typography;
  }
  
  /**
   * Extract spacing variables
   */
  extractSpacing(content) {
    const spacing = [];
    const spacingPatterns = this.analysisPatterns.spacing;
    
    // Extract margin values
    const marginMatches = content.match(spacingPatterns.margin) || [];
    marginMatches.forEach((match, index) => {
      const value = match.split(/[:=]/)[1]?.trim().replace(/[;"]/g, '');
      if (value) {
        spacing.push({
          name: `margin-${index + 1}`,
          value: value,
          type: 'margin',
          context: 'Spacing system',
          confidence: 0.7
        });
      }
    });
    
    // Extract padding values
    const paddingMatches = content.match(spacingPatterns.padding) || [];
    paddingMatches.forEach((match, index) => {
      const value = match.split(/[:=]/)[1]?.trim().replace(/[;"]/g, '');
      if (value) {
        spacing.push({
          name: `padding-${index + 1}`,
          value: value,
          type: 'padding',
          context: 'Spacing system',
          confidence: 0.7
        });
      }
    });
    
    // Extract spacing variables
    const spacingVarMatches = content.match(spacingPatterns.spacing) || [];
    spacingVarMatches.forEach((variable, index) => {
      spacing.push({
        name: this.cleanVariableName(variable),
        value: this.extractVariableValue(content, variable),
        type: 'spacing-variable',
        context: 'Spacing system',
        confidence: 0.9
      });
    });
    
    return spacing;
  }
  
  /**
   * Detect component patterns
   */
  detectComponents(content) {
    const components = [];
    
    Object.entries(this.componentPatterns).forEach(([componentType, pattern]) => {
      const matches = content.match(pattern) || [];
      if (matches.length > 0) {
        components.push({
          name: componentType,
          occurrences: matches.length,
          type: 'component',
          context: `Found ${matches.length} instances`,
          confidence: 0.6
        });
      }
    });
    
    return components;
  }
  
  /**
   * Store extracted variables in spreadsheet
   */
  storeExtractedVariables(clientName, variables, source) {
    const allVariables = [
      ...variables.colors.map(v => ({ ...v, category: 'Colors' })),
      ...variables.typography.map(v => ({ ...v, category: 'Typography' })),
      ...variables.spacing.map(v => ({ ...v, category: 'Spacing' })),
      ...variables.components.map(v => ({ ...v, category: 'Components' }))
    ];
    
    allVariables.forEach(variable => {
      const priority = variable.confidence > 0.8 ? 'High' : 
                     variable.confidence > 0.6 ? 'Medium' : 'Low';
      
      addClientVariable(
        clientName,
        'All Portals',
        variable.category,
        variable.name,
        variable.value || 'TBD',
        variable.type || 'Token',
        variable.context || '',
        priority,
        `Auto-extracted from ${source} with ${Math.round(variable.confidence * 100)}% confidence`
      );
    });
  }
  
  /**
   * Generate onboarding checklist
   */
  generateOnboardingChecklist(clientInfo, analysisResult = null) {
    const checklist = {
      items: [
        {
          id: 'initial_contact',
          title: 'Initial Contact Established',
          completed: true,
          dueDate: new Date(),
          assignee: clientInfo.projectManager
        },
        {
          id: 'style_guide_request',
          title: 'Style Guide Materials Requested',
          completed: true,
          dueDate: new Date(),
          assignee: clientInfo.projectManager
        },
        {
          id: 'materials_received',
          title: 'Client Style Guide Materials Received',
          completed: !!analysisResult,
          dueDate: new Date(Date.now() + 7 * 24 * 60 * 60 * 1000),
          assignee: clientInfo.clientName
        },
        {
          id: 'analysis_complete',
          title: 'Style Guide Analysis Completed',
          completed: !!analysisResult,
          dueDate: new Date(Date.now() + 10 * 24 * 60 * 60 * 1000),
          assignee: 'Design Team'
        },
        {
          id: 'conflicts_identified',
          title: 'Design Conflicts Identified and Documented',
          completed: analysisResult && analysisResult.conflicts.length > 0,
          dueDate: new Date(Date.now() + 12 * 24 * 60 * 60 * 1000),
          assignee: 'Design Team'
        },
        {
          id: 'mappings_created',
          title: 'Variable Mappings Created',
          completed: false,
          dueDate: new Date(Date.now() + 14 * 24 * 60 * 60 * 1000),
          assignee: 'Development Team'
        },
        {
          id: 'client_approval',
          title: 'Client Approval Received',
          completed: false,
          dueDate: new Date(Date.now() + 21 * 24 * 60 * 60 * 1000),
          assignee: clientInfo.clientName
        },
        {
          id: 'css_generation',
          title: 'CSS Variables Generated',
          completed: false,
          dueDate: new Date(Date.now() + 24 * 24 * 60 * 60 * 1000),
          assignee: 'Development Team'
        },
        {
          id: 'implementation',
          title: 'Portal Implementation Completed',
          completed: false,
          dueDate: new Date(Date.now() + 35 * 24 * 60 * 60 * 1000),
          assignee: 'Development Team'
        }
      ],
      nextSteps: this.generateNextSteps(clientInfo, analysisResult)
    };
    
    return checklist;
  }
  
  /**
   * Generate next steps based on current state
   */
  generateNextSteps(clientInfo, analysisResult) {
    const steps = [];
    
    if (!analysisResult) {
      steps.push('Wait for client to provide style guide materials');
      steps.push('Follow up if materials not received within 7 days');
    } else {
      if (analysisResult.conflicts.length > 0) {
        steps.push(`Address ${analysisResult.conflicts.length} design conflicts`);
        steps.push('Schedule conflict resolution call with client');
      }
      
      steps.push('Create variable mappings for approved elements');
      steps.push('Generate CSS variables for portal implementation');
    }
    
    steps.push('Schedule weekly status updates with client');
    
    return steps;
  }
  
  /**
   * Utility functions
   */
  initializeClientProject(clientInfo) {
    const projectId = `PRJ-${clientInfo.clientName.toUpperCase().replace(/\s+/g, '')}-${Date.now().toString(36)}`;
    
    // Create project entry
    const project = {
      id: projectId,
      clientName: clientInfo.clientName,
      portalTypes: clientInfo.portalTypes,
      projectManager: clientInfo.projectManager,
      timeline: clientInfo.timeline || '30 days',
      status: 'Initiated'
    };
    
    return project;
  }
  
  createProjectMilestones(projectId, timeline) {
    const timelineDays = parseInt(timeline) || 30;
    const milestones = [
      { name: 'Style Guide Collection', days: 7 },
      { name: 'Analysis & Mapping', days: 14 },
      { name: 'Client Approval', days: 21 },
      { name: 'Implementation', days: timelineDays - 3 },
      { name: 'Testing & Deployment', days: timelineDays }
    ];
    
    return milestones.map(milestone => ({
      ...milestone,
      dueDate: new Date(Date.now() + milestone.days * 24 * 60 * 60 * 1000),
      projectId
    }));
  }
  
  extractColorContext(content, color) {
    const lines = content.split('\n');
    for (const line of lines) {
      if (line.includes(color)) {
        // Extract meaningful context from the line
        const context = line.trim().substring(0, 100);
        return context;
      }
    }
    return 'Color usage context';
  }
  
  cleanVariableName(variable) {
    return variable.replace(/[^a-zA-Z0-9-_]/g, '').toLowerCase();
  }
  
  extractVariableValue(content, variable) {
    // Simple pattern to extract variable values
    const pattern = new RegExp(`${variable}\\s*[:=]\\s*([^;,}\\n]+)`, 'i');
    const match = content.match(pattern);
    return match ? match[1].trim() : 'TBD';
  }
  
  deduplicateColors(colors) {
    const unique = [];
    const seen = new Set();
    
    colors.forEach(color => {
      const key = `${color.value}-${color.type}`;
      if (!seen.has(key)) {
        seen.add(key);
        unique.push(color);
      }
    });
    
    return unique;
  }
  
  analyzeConflicts(variables, clientName) {
    const mapper = new DesignSystemMapper();
    const conflicts = [];
    
    // Check color conflicts
    variables.colors.forEach(color => {
      const mapping = mapper.findBestMapping({
        name: color.name,
        value: color.value,
        category: 'colors'
      });
      
      if (mapping && mapping.conflictLevel !== 'None') {
        conflicts.push({
          type: 'Color Conflict',
          clientValue: color.value,
          companyValue: mapping.companyVariable,
          severity: mapping.conflictLevel,
          description: `Color ${color.name} differs from company standard`
        });
      }
    });
    
    return conflicts;
  }
  
  generateRecommendations(variables, conflicts) {
    const recommendations = [];
    
    if (variables.colors.length === 0) {
      recommendations.push('No color variables detected - manual color specification needed');
    }
    
    if (variables.typography.length === 0) {
      recommendations.push('No typography system detected - review font specifications');
    }
    
    if (conflicts.length > 5) {
      recommendations.push('High number of conflicts detected - consider design system alignment discussion');
    }
    
    if (variables.components.length === 0) {
      recommendations.push('No component patterns detected - may need manual component analysis');
    }
    
    return recommendations;
  }
  
  calculateAnalysisConfidence(variables) {
    const totalVariables = this.countTotalVariables(variables);
    if (totalVariables === 0) return 0;
    
    const confidenceSum = [
      ...variables.colors,
      ...variables.typography,
      ...variables.spacing,
      ...variables.components
    ].reduce((sum, variable) => sum + (variable.confidence || 0.5), 0);
    
    return confidenceSum / totalVariables;
  }
  
  countTotalVariables(variables) {
    return variables.colors.length + 
           variables.typography.length + 
           variables.spacing.length + 
           variables.components.length;
  }
  
  storeConflicts(clientName, conflicts) {
    const conflictsSheet = this.ss.getSheetByName(SHEET_NAMES.STYLE_CONFLICTS);
    
    conflicts.forEach(conflict => {
      const conflictId = Utilities.getUuid();
      const row = [
        conflictId,
        clientName,
        conflict.type,
        conflict.clientValue,
        conflict.companyValue,
        conflict.description,
        conflict.severity,
        'Auto-detected during analysis',
        'Pending',
        'Identified',
        ''
      ];
      
      conflictsSheet.appendRow(row);
    });
  }
  
  getMockStyleGuideContent() {
    // Mock style guide content for demonstration
    return `
/* Color System */
:root {
  --primary-color: #2563eb;
  --secondary-color: #64748b;
  --success-color: #10b981;
  --warning-color: #f59e0b;
  --error-color: #ef4444;
  --background: #ffffff;
  --surface: #f8fafc;
}

/* Typography */
body {
  font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
  font-size: 16px;
  line-height: 1.5;
}

h1 { font-size: 2.5rem; font-weight: 700; }
h2 { font-size: 2rem; font-weight: 600; }
h3 { font-size: 1.5rem; font-weight: 500; }

/* Spacing */
.container { padding: 24px; margin: 16px; }
.section { margin-bottom: 32px; }
.button { padding: 12px 24px; margin: 8px; }

/* Components */
.btn-primary { background-color: var(--primary-color); }
.card { background: var(--surface); border-radius: 8px; }
.modal { background: var(--background); box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
    `;
  }
}

/**
 * Public API functions for client onboarding
 */
function startClientOnboarding(clientInfo) {
  const manager = new ClientOnboardingManager();
  return manager.onboardNewClient(clientInfo);
}

function analyzeClientStyleGuide(source, identifier, clientName) {
  const manager = new ClientOnboardingManager();
  
  if (source === 'url') {
    return manager.analyzeStyleGuideFromUrl(identifier, clientName);
  } else if (source === 'file') {
    return manager.analyzeStyleGuideDocument(identifier, clientName);
  } else {
    throw new Error('Invalid source type. Use "url" or "file".');
  }
}

function getClientOnboardingStatus(clientName) {
  const manager = new ClientOnboardingManager();
  const project = manager.getClientProject(clientName);
  
  if (!project) {
    return { status: 'Not Found', message: 'No project found for this client' };
  }
  
  return {
    status: project.status,
    phase: project.currentPhase,
    progress: project.progress,
    nextMilestone: project.nextMilestone,
    blockers: project.blockers
  };
}
