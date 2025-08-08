/**
 * Communication Protocol System for Design System Management
 * 
 * Handles automated communication, template management, and client interaction protocols
 * for efficient design system onboarding and maintenance.
 */

/**
 * Communication Protocol Manager
 */
class CommunicationProtocol {
  constructor() {
    this.ss = SpreadsheetApp.getActiveSpreadsheet();
    this.templatesSheet = this.ss.getSheetByName(SHEET_NAMES.TEMPLATES);
    this.communicationSheet = this.ss.getSheetByName(SHEET_NAMES.COMMUNICATION_LOG);
    this.projectSheet = this.ss.getSheetByName(SHEET_NAMES.PROJECT_TRACKER);
    
    // Communication workflows and timing
    this.workflows = {
      CLIENT_ONBOARDING: 'client_onboarding',
      STYLE_CONFLICT_RESOLUTION: 'style_conflict_resolution',
      APPROVAL_PROCESS: 'approval_process',
      IMPLEMENTATION_UPDATES: 'implementation_updates',
      PROJECT_STATUS: 'project_status'
    };
    
    // Default response timeframes
    this.timeframes = {
      URGENT: 1, // 1 day
      HIGH: 3,   // 3 days  
      MEDIUM: 7, // 1 week
      LOW: 14    // 2 weeks
    };
  }
  
  /**
   * Initialize client onboarding workflow
   */
  startClientOnboarding(clientName, clientEmail, projectManager, portalTypes = [], urgency = 'MEDIUM') {
    const projectId = this.createProject(clientName, portalTypes, projectManager);
    
    // Send initial style guide request
    const requestResult = this.sendStyleGuideRequest(clientName, clientEmail, {
      PROJECT_ID: projectId,
      PORTAL_TYPES: portalTypes.join(', '),
      PROJECT_MANAGER: projectManager,
      TIMELINE: `${this.timeframes[urgency]} business days`,
      URGENCY: urgency
    });
    
    // Schedule follow-up communications
    this.scheduleFollowUps(clientName, projectId, urgency);
    
    // Log the onboarding initiation
    this.logCommunication(
      clientName,
      'Email',
      'Design System Onboarding - Style Guide Request',
      clientEmail,
      `Initiated onboarding process for portals: ${portalTypes.join(', ')}`,
      `Client to provide style guide materials within ${this.timeframes[urgency]} days`,
      new Date(Date.now() + this.timeframes[urgency] * 24 * 60 * 60 * 1000)
    );
    
    return {
      projectId,
      emailSent: requestResult.success,
      followUpsScheduled: true,
      nextAction: `Await style guide materials from ${clientName}`
    };
  }
  
  /**
   * Send style guide request using template
   */
  sendStyleGuideRequest(clientName, clientEmail, variables = {}) {
    const template = this.getTemplate('Style Guide Request');
    if (!template) {
      throw new Error('Style Guide Request template not found');
    }
    
    const emailContent = this.processTemplate(template.content, {
      CLIENT_NAME: clientName,
      ...variables
    });
    
    try {
      MailApp.sendEmail({
        to: clientEmail,
        subject: `Design System Onboarding - ${clientName}`,
        body: emailContent,
        htmlBody: emailContent.replace(/\n/g, '<br>'),
        attachments: this.getOnboardingAttachments()
      });
      
      return { success: true, templateUsed: 'Style Guide Request' };
    } catch (error) {
      Logger.log(`Failed to send email to ${clientEmail}: ${error.message}`);
      return { success: false, error: error.message };
    }
  }
  
  /**
   * Handle style conflict communication
   */
  handleStyleConflicts(clientName, clientEmail, conflicts, urgency = 'HIGH') {
    const conflictSummary = this.generateConflictSummary(conflicts);
    const resolutionDoc = this.createConflictResolutionDocument(clientName, conflicts);
    
    const variables = {
      CLIENT_NAME: clientName,
      CONFLICT_LIST: conflictSummary.list,
      CONFLICT_COUNT: conflicts.length,
      RECOMMENDATIONS: conflictSummary.recommendations,
      TIMEFRAME: `${this.timeframes[urgency]} business days`,
      RESOLUTION_DOC_URL: resolutionDoc.url
    };
    
    const template = this.getTemplate('Style Conflict Resolution');
    const emailContent = this.processTemplate(template.content, variables);
    
    try {
      MailApp.sendEmail({
        to: clientEmail,
        subject: `Design System Conflicts Require Resolution - ${clientName}`,
        body: emailContent,
        htmlBody: emailContent.replace(/\n/g, '<br>'),
        attachments: [resolutionDoc.file]
      });
      
      // Log communication
      this.logCommunication(
        clientName,
        'Email',
        'Style Conflict Resolution Request',
        clientEmail,
        `Sent conflict resolution document with ${conflicts.length} issues`,
        `Client review and response needed within ${this.timeframes[urgency]} days`,
        new Date(Date.now() + this.timeframes[urgency] * 24 * 60 * 60 * 1000)
      );
      
      // Schedule escalation if needed
      if (urgency === 'HIGH') {
        this.scheduleEscalation(clientName, 'Conflict Resolution', 2);
      }
      
      return { success: true, conflictsDocumented: conflicts.length };
    } catch (error) {
      Logger.log(`Failed to send conflict resolution email: ${error.message}`);
      return { success: false, error: error.message };
    }
  }
  
  /**
   * Send mapping approval request
   */
  requestMappingApproval(clientName, clientEmail, mappingResults, portalType = null) {
    const approvalDoc = this.createMappingApprovalDocument(clientName, mappingResults, portalType);
    
    const variables = {
      CLIENT_NAME: clientName,
      TOTAL_MAPPINGS: mappingResults.mappedVariables,
      PORTAL_TYPE: portalType || 'All Portals',
      APPROVAL_DOC_URL: approvalDoc.url,
      IMPLEMENTATION_TIMELINE: '5-7 business days post-approval'
    };
    
    const template = this.getTemplate('Mapping Approval Request');
    const emailContent = this.processTemplate(template.content, variables);
    
    try {
      MailApp.sendEmail({
        to: clientEmail,
        subject: `Design System Mapping Approval Required - ${clientName}`,
        body: emailContent,
        htmlBody: emailContent.replace(/\n/g, '<br>'),
        attachments: [approvalDoc.file]
      });
      
      this.logCommunication(
        clientName,
        'Email',
        'Mapping Approval Request',
        clientEmail,
        `Sent mapping approval document for ${mappingResults.mappedVariables} variables`,
        'Client approval required before CSS generation',
        new Date(Date.now() + 5 * 24 * 60 * 60 * 1000)
      );
      
      return { success: true, mappingsToApprove: mappingResults.mappedVariables };
    } catch (error) {
      return { success: false, error: error.message };
    }
  }
  
  /**
   * Send project status updates
   */
  sendProjectStatusUpdate(clientName, clientEmail, projectId, customMessage = null) {
    const project = this.getProject(projectId);
    if (!project) {
      throw new Error(`Project ${projectId} not found`);
    }
    
    const statusSummary = this.generateProjectStatusSummary(project);
    
    const variables = {
      CLIENT_NAME: clientName,
      PROJECT_ID: projectId,
      CURRENT_PHASE: project.currentPhase,
      PROGRESS_PERCENT: project.progress,
      NEXT_MILESTONE: project.nextMilestone,
      TIMELINE_STATUS: statusSummary.timelineStatus,
      COMPLETED_TASKS: statusSummary.completedTasks,
      UPCOMING_TASKS: statusSummary.upcomingTasks,
      BLOCKERS: project.blockers || 'None',
      CUSTOM_MESSAGE: customMessage || ''
    };
    
    const template = this.getTemplate('Project Status Update');
    const emailContent = this.processTemplate(template.content, variables);
    
    try {
      MailApp.sendEmail({
        to: clientEmail,
        subject: `Project Status Update - ${clientName} Design System Implementation`,
        body: emailContent,
        htmlBody: emailContent.replace(/\n/g, '<br>')
      });
      
      this.logCommunication(
        clientName,
        'Email',
        'Project Status Update',
        clientEmail,
        `Progress: ${project.progress}%, Phase: ${project.currentPhase}`,
        statusSummary.nextActions,
        new Date(Date.now() + 7 * 24 * 60 * 60 * 1000)
      );
      
      return { success: true, phase: project.currentPhase };
    } catch (error) {
      return { success: false, error: error.message };
    }
  }
  
  /**
   * Automated follow-up system
   */
  scheduleFollowUps(clientName, projectId, urgency = 'MEDIUM') {
    const followUps = [
      {
        days: Math.floor(this.timeframes[urgency] / 2),
        type: 'reminder',
        action: 'sendStyleGuideReminder'
      },
      {
        days: this.timeframes[urgency],
        type: 'check_response',
        action: 'checkStyleGuideResponse'
      },
      {
        days: this.timeframes[urgency] + 2,
        type: 'escalation',
        action: 'escalateNoResponse'
      }
    ];
    
    // In a production environment, you'd use triggers or external scheduling
    // For now, we'll log the intended follow-ups
    followUps.forEach(followUp => {
      const followUpDate = new Date(Date.now() + followUp.days * 24 * 60 * 60 * 1000);
      this.logScheduledAction(clientName, projectId, followUp.type, followUp.action, followUpDate);
    });
    
    return followUps;
  }
  
  /**
   * Process communication templates with variable substitution
   */
  processTemplate(templateContent, variables) {
    let processedContent = templateContent;
    
    Object.entries(variables).forEach(([key, value]) => {
      const placeholder = new RegExp(`{{${key}}}`, 'g');
      processedContent = processedContent.replace(placeholder, value || '');
    });
    
    // Clean up any remaining placeholders
    processedContent = processedContent.replace(/{{[^}]+}}/g, '[NOT PROVIDED]');
    
    return processedContent;
  }
  
  /**
   * Get template by name
   */
  getTemplate(templateName) {
    const data = this.templatesSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === templateName) {
        return {
          name: data[i][0],
          purpose: data[i][1],
          content: data[i][2],
          variables: data[i][3],
          instructions: data[i][4]
        };
      }
    }
    
    return null;
  }
  
  /**
   * Generate conflict summary for communication
   */
  generateConflictSummary(conflicts) {
    const groupedConflicts = conflicts.reduce((groups, conflict) => {
      const severity = conflict.severity || conflict.conflictLevel;
      if (!groups[severity]) groups[severity] = [];
      groups[severity].push(conflict);
      return groups;
    }, {});
    
    const list = Object.entries(groupedConflicts)
      .map(([severity, items]) => 
        `**${severity} Priority (${items.length} items):**\n${items.map(item => `  - ${item.description || item.elementType}: ${item.clientSpecification}`).join('\n')}`
      ).join('\n\n');
    
    const recommendations = Object.entries(groupedConflicts)
      .map(([severity, items]) => {
        if (severity === 'High') {
          return `• **High Priority:** Require immediate discussion - may impact brand integrity or technical feasibility`;
        } else if (severity === 'Medium') {
          return `• **Medium Priority:** Need client preference confirmation - alternative solutions available`;
        } else {
          return `• **Low Priority:** Minor adjustments recommended for optimal user experience`;
        }
      }).join('\n');
    
    return { list, recommendations };
  }
  
  /**
   * Create conflict resolution document
   */
  createConflictResolutionDocument(clientName, conflicts) {
    const docContent = this.generateConflictResolutionDoc(clientName, conflicts);
    const fileName = `${clientName}_Design_Conflicts_${new Date().toISOString().split('T')[0]}.txt`;
    
    const blob = Utilities.newBlob(docContent, 'text/plain', fileName);
    const file = DriveApp.createFile(blob);
    
    return {
      file: file,
      url: `https://drive.google.com/file/d/${file.getId()}/view`,
      id: file.getId()
    };
  }
  
  /**
   * Create mapping approval document  
   */
  createMappingApprovalDocument(clientName, mappingResults, portalType) {
    const docContent = this.generateMappingApprovalDoc(clientName, mappingResults, portalType);
    const fileName = `${clientName}_Mapping_Approval_${portalType || 'All'}_${new Date().toISOString().split('T')[0]}.txt`;
    
    const blob = Utilities.newBlob(docContent, 'text/plain', fileName);
    const file = DriveApp.createFile(blob);
    
    return {
      file: file,
      url: `https://drive.google.com/file/d/${file.getId()}/view`,
      id: file.getId()
    };
  }
  
  /**
   * Generate detailed conflict resolution document
   */
  generateConflictResolutionDoc(clientName, conflicts) {
    return `
DESIGN SYSTEM CONFLICT RESOLUTION DOCUMENT
Client: ${clientName}
Date: ${new Date().toLocaleDateString()}
Project: Design System Integration

EXECUTIVE SUMMARY
We have identified ${conflicts.length} potential conflicts between your design requirements and our platform standards. This document outlines each conflict and provides recommended solutions.

CONFLICT DETAILS
${conflicts.map((conflict, index) => `
${index + 1}. ${conflict.elementType || conflict.description}
   Client Requirement: ${conflict.clientSpecification}
   Platform Standard: ${conflict.companyStandard || 'Our Standard Approach'}
   Conflict Level: ${conflict.severity || conflict.conflictLevel}
   Impact: ${this.getConflictImpact(conflict)}
   Recommended Solution: ${conflict.proposedSolution || conflict.resolution}
   
   Options for Resolution:
   A) Accept our recommendation (maintains platform consistency)
   B) Implement client specification (may require additional development)
   C) Compromise solution (blend of both approaches)
`).join('\n')}

NEXT STEPS
1. Review each conflict and select your preferred resolution approach
2. Schedule a design review call to discuss high-priority items
3. Provide written approval for selected solutions
4. We will implement approved solutions and proceed with development

Please respond within the specified timeframe to maintain project schedule.

Contact: [Your Project Manager Contact Information]
`;
  }
  
  /**
   * Generate mapping approval document
   */
  generateMappingApprovalDoc(clientName, mappingResults, portalType) {
    return `
DESIGN SYSTEM MAPPING APPROVAL DOCUMENT
Client: ${clientName}
Portal Type: ${portalType || 'All Portals'}
Date: ${new Date().toLocaleDateString()}

MAPPING SUMMARY
Total Variables Processed: ${mappingResults.totalVariables}
Successfully Mapped: ${mappingResults.mappedVariables}
Requiring Manual Review: ${mappingResults.unmappedVariables}

VARIABLE MAPPINGS
${mappingResults.mappings ? mappingResults.mappings.map((mapping, index) => `
${index + 1}. ${mapping.clientVariable}
   → CSS Variable: ${mapping.cssVariableName}
   → Company Standard: ${mapping.companyVariable}
   Confidence: ${Math.round((mapping.confidence || 0.8) * 100)}%
   Status: ${mapping.mappingStatus}
`).join('\n') : 'No specific mappings provided'}

CSS VARIABLE NAMING CONVENTION
All variables follow our standard format: --t-gs-{category}-{name}
Examples:
- --t-gs-color-primary
- --t-gs-font-size-large  
- --t-gs-space-medium

APPROVAL REQUIRED
Please review the above mappings and confirm your approval by replying to this email with:
1. "APPROVED" for mappings you accept as-is
2. Specific feedback for any mappings requiring changes
3. Priority level for implementation (Standard/Rush)

Upon approval, CSS variables will be generated and integrated into your portal development.

Timeline: Implementation begins within 24 hours of approval
`;
  }
  
  /**
   * Utility and helper functions
   */
  createProject(clientName, portalTypes, projectManager) {
    const projectId = `PRJ-${clientName.toUpperCase().replace(/\s+/g, '')}-${Date.now().toString(36)}`;
    
    const projectRow = [
      projectId,
      clientName,
      portalTypes.join(', '),
      new Date(),
      new Date(Date.now() + 30 * 24 * 60 * 60 * 1000), // 30 days default
      'Style Guide Collection',
      '10%',
      projectManager,
      '',
      'Receive client style guide materials',
      'Active'
    ];
    
    this.projectSheet.appendRow(projectRow);
    return projectId;
  }
  
  getProject(projectId) {
    const data = this.projectSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === projectId) {
        return {
          id: data[i][0],
          clientName: data[i][1],
          portalTypes: data[i][2],
          startDate: data[i][3],
          expectedCompletion: data[i][4],
          currentPhase: data[i][5],
          progress: data[i][6],
          teamMembers: data[i][7],
          blockers: data[i][8],
          nextMilestone: data[i][9],
          status: data[i][10],
          row: i + 1
        };
      }
    }
    
    return null;
  }
  
  logCommunication(clientName, type, subject, participants, keyPoints, actionItems, followUpDate) {
    const row = [
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
    
    this.communicationSheet.appendRow(row);
  }
  
  logScheduledAction(clientName, projectId, type, action, scheduledDate) {
    // This would integrate with your scheduling system
    Logger.log(`Scheduled ${type} for ${clientName} (${projectId}): ${action} on ${scheduledDate}`);
  }
  
  getConflictImpact(conflict) {
    const severity = conflict.severity || conflict.conflictLevel;
    
    switch (severity) {
      case 'High':
        return 'May require significant development changes or compromise brand integrity';
      case 'Medium':
        return 'Could affect user experience consistency or require additional QA';
      case 'Low':
        return 'Minor impact on design consistency, easily addressed';
      default:
        return 'Impact assessment pending';
    }
  }
  
  generateProjectStatusSummary(project) {
    const progress = parseInt(project.progress.replace('%', ''));
    const today = new Date();
    const expectedCompletion = new Date(project.expectedCompletion);
    const daysRemaining = Math.ceil((expectedCompletion - today) / (1000 * 60 * 60 * 24));
    
    return {
      timelineStatus: daysRemaining > 0 ? `${daysRemaining} days remaining` : `${Math.abs(daysRemaining)} days overdue`,
      completedTasks: this.getCompletedTasksForPhase(project.currentPhase),
      upcomingTasks: this.getUpcomingTasksForPhase(project.currentPhase),
      nextActions: progress < 50 ? 'Awaiting client inputs and approvals' : 'Proceeding with implementation'
    };
  }
  
  getCompletedTasksForPhase(phase) {
    const phaseTasks = {
      'Style Guide Collection': ['Project initiated', 'Initial communication sent'],
      'Analysis & Mapping': ['Variables analyzed', 'Conflicts identified', 'Mapping completed'],
      'Approval': ['Documentation prepared', 'Client review initiated'],
      'Implementation': ['CSS variables generated', 'Portal integration started'],
      'Testing': ['Implementation completed', 'QA testing in progress'],
      'Deployment': ['Testing completed', 'Production deployment']
    };
    
    return phaseTasks[phase] || [];
  }
  
  getUpcomingTasksForPhase(phase) {
    const upcomingTasks = {
      'Style Guide Collection': ['Receive client materials', 'Begin analysis'],
      'Analysis & Mapping': ['Generate mapping report', 'Client approval'],
      'Approval': ['Finalize mappings', 'Begin CSS generation'],
      'Implementation': ['Complete portal integration', 'Internal testing'],
      'Testing': ['Client UAT', 'Bug fixes'],
      'Deployment': ['Production deployment', 'Project closure']
    };
    
    return upcomingTasks[phase] || [];
  }
  
  getOnboardingAttachments() {
    // Return standard onboarding documents
    // In production, these would be actual files from your Drive
    return [];
  }
}

/**
 * Public API functions for communication workflows
 */
function initiateClientOnboarding(clientName, clientEmail, projectManager, portalTypes, urgency = 'MEDIUM') {
  const protocol = new CommunicationProtocol();
  return protocol.startClientOnboarding(clientName, clientEmail, projectManager, portalTypes, urgency);
}

function sendConflictResolution(clientName, clientEmail, conflicts, urgency = 'HIGH') {
  const protocol = new CommunicationProtocol();
  return protocol.handleStyleConflicts(clientName, clientEmail, conflicts, urgency);
}

function requestApprovalForMappings(clientName, clientEmail, mappingResults, portalType = null) {
  const protocol = new CommunicationProtocol();
  return protocol.requestMappingApproval(clientName, clientEmail, mappingResults, portalType);
}

function sendStatusUpdate(clientName, clientEmail, projectId, customMessage = null) {
  const protocol = new CommunicationProtocol();
  return protocol.sendProjectStatusUpdate(clientName, clientEmail, projectId, customMessage);
}

/**
 * Automated workflow triggers (would be set up with time-based triggers)
 */
function checkPendingFollowUps() {
  const protocol = new CommunicationProtocol();
  const today = new Date();
  
  // Check for overdue communications and send reminders
  const commData = protocol.communicationSheet.getDataRange().getValues();
  
  for (let i = 1; i < commData.length; i++) {
    const [date, client, type, subject, participants, keyPoints, actionItems, followUpDate, status] = commData[i];
    
    if (status === 'Active' && followUpDate && new Date(followUpDate) <= today) {
      // Send follow-up reminder
      Logger.log(`Follow-up needed for ${client}: ${subject}`);
      // Implement follow-up logic here
    }
  }
}
