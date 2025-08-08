/**
 * Advanced Design System Mapping and Analysis Functions
 * 
 * This file contains sophisticated mapping logic between client design systems,
 * company Figma variables, and CSS custom properties with conflict detection.
 */

/**
 * Advanced mapping system that creates intelligent connections between design systems
 */
class DesignSystemMapper {
  constructor() {
    this.ss = SpreadsheetApp.getActiveSpreadsheet();
    this.clientSheet = this.ss.getSheetByName(SHEET_NAMES.CLIENT_VARIABLES);
    this.mappingSheet = this.ss.getSheetByName(SHEET_NAMES.COMPANY_MAPPING);
    this.conflictsSheet = this.ss.getSheetByName(SHEET_NAMES.STYLE_CONFLICTS);
    
    // Company design system standards (would come from your Figma API in production)
    this.companyStandards = {
      colors: {
        'primary': '#1976D2',
        'secondary': '#424242', 
        'success': '#388E3C',
        'warning': '#F57C00',
        'error': '#D32F2F',
        'surface': '#FFFFFF',
        'background': '#FAFAFA'
      },
      typography: {
        'font-family-primary': 'Inter, sans-serif',
        'font-family-secondary': 'Roboto, sans-serif',
        'font-size-xs': '12px',
        'font-size-sm': '14px',
        'font-size-md': '16px',
        'font-size-lg': '18px',
        'font-size-xl': '24px',
        'font-weight-normal': '400',
        'font-weight-medium': '500',
        'font-weight-bold': '700'
      },
      spacing: {
        'spacing-xs': '4px',
        'spacing-sm': '8px',
        'spacing-md': '16px',
        'spacing-lg': '24px',
        'spacing-xl': '32px',
        'spacing-2xl': '48px'
      }
    };
  }
  
  /**
   * Automatically map client variables to company standards using intelligent matching
   */
  autoMapClientVariables(clientName, portalType = null) {
    const clientVars = this.getClientVariables(clientName, portalType);
    const mappingResults = [];
    
    clientVars.forEach(clientVar => {
      const mapping = this.findBestMapping(clientVar);
      if (mapping) {
        mappingResults.push(mapping);
        this.addMapping(mapping);
      }
    });
    
    return {
      totalVariables: clientVars.length,
      mappedVariables: mappingResults.length,
      unmappedVariables: clientVars.length - mappingResults.length,
      mappings: mappingResults
    };
  }
  
  /**
   * Get client variables with filtering
   */
  getClientVariables(clientName, portalType = null, category = null) {
    const data = this.clientSheet.getDataRange().getValues();
    const variables = [];
    
    for (let i = 1; i < data.length; i++) {
      const [client, portal, cat, varName, value, type, context, priority, status, notes] = data[i];
      
      if (client === clientName && 
          (!portalType || portal === portalType) &&
          (!category || cat === category)) {
        variables.push({
          client,
          portal,
          category: cat,
          name: varName,
          value,
          type,
          context,
          priority,
          status,
          notes,
          row: i + 1
        });
      }
    }
    
    return variables;
  }
  
  /**
   * Find the best mapping for a client variable using intelligent matching
   */
  findBestMapping(clientVar) {
    const category = clientVar.category.toLowerCase();
    const name = clientVar.name.toLowerCase();
    const value = clientVar.value;
    
    // Get company standards for this category
    const standards = this.companyStandards[category];
    if (!standards) {
      return null;
    }
    
    // Try exact name matching first
    let bestMatch = this.findExactNameMatch(name, standards);
    
    // If no exact match, try semantic matching
    if (!bestMatch) {
      bestMatch = this.findSemanticMatch(name, value, standards);
    }
    
    // If still no match, try value-based matching for colors
    if (!bestMatch && category === 'colors') {
      bestMatch = this.findColorValueMatch(value, standards);
    }
    
    if (bestMatch) {
      const cssVarName = this.generateCSSVariableName(category, bestMatch.key);
      const conflictLevel = this.assessConflictLevel(clientVar, bestMatch);
      
      return {
        clientVariable: `${clientVar.name}=${clientVar.value}`,
        companyVariable: `${bestMatch.key}=${bestMatch.value}`,
        cssVariableName: cssVarName,
        mappingStatus: conflictLevel === 'None' ? 'Auto-Mapped' : 'Needs Review',
        conflictLevel,
        resolution: this.suggestResolution(clientVar, bestMatch, conflictLevel),
        confidence: bestMatch.confidence || 0.8
      };
    }
    
    return null;
  }
  
  /**
   * Find exact name matches
   */
  findExactNameMatch(clientName, standards) {
    const normalizedName = this.normalizeVariableName(clientName);
    
    for (const [key, value] of Object.entries(standards)) {
      const normalizedKey = this.normalizeVariableName(key);
      if (normalizedName === normalizedKey) {
        return { key, value, confidence: 1.0 };
      }
    }
    
    return null;
  }
  
  /**
   * Find semantic matches using keyword analysis
   */
  findSemanticMatch(clientName, clientValue, standards) {
    const semanticMappings = {
      colors: {
        'primary': ['main', 'brand', 'accent', 'theme'],
        'secondary': ['secondary', 'alt', 'alternative'],
        'success': ['success', 'green', 'positive', 'ok'],
        'warning': ['warning', 'yellow', 'caution', 'alert'],
        'error': ['error', 'red', 'danger', 'negative'],
        'surface': ['surface', 'card', 'container'],
        'background': ['background', 'bg', 'backdrop']
      },
      typography: {
        'font-family-primary': ['primary', 'main', 'body', 'text'],
        'font-family-secondary': ['secondary', 'heading', 'display'],
        'font-size-xs': ['xs', 'tiny', 'small'],
        'font-size-sm': ['sm', 'small'],
        'font-size-md': ['md', 'medium', 'base', 'normal'],
        'font-size-lg': ['lg', 'large'],
        'font-size-xl': ['xl', 'huge', 'big']
      },
      spacing: {
        'spacing-xs': ['xs', 'tiny', '4'],
        'spacing-sm': ['sm', 'small', '8'],
        'spacing-md': ['md', 'medium', 'base', '16'],
        'spacing-lg': ['lg', 'large', '24'],
        'spacing-xl': ['xl', 'huge', '32']
      }
    };
    
    const category = Object.keys(this.companyStandards).find(cat => 
      Object.keys(standards).some(key => key.includes(cat.slice(0, -1)))
    );
    
    if (!category || !semanticMappings[category]) {
      return null;
    }
    
    let bestMatch = null;
    let highestScore = 0;
    
    for (const [standardKey, keywords] of Object.entries(semanticMappings[category])) {
      if (!standards[standardKey]) continue;
      
      const score = this.calculateSemanticScore(clientName, keywords);
      if (score > highestScore && score > 0.6) {
        highestScore = score;
        bestMatch = {
          key: standardKey,
          value: standards[standardKey],
          confidence: score
        };
      }
    }
    
    return bestMatch;
  }
  
  /**
   * Find color matches based on value similarity
   */
  findColorValueMatch(clientValue, colorStandards) {
    const clientColor = this.parseColor(clientValue);
    if (!clientColor) return null;
    
    let bestMatch = null;
    let smallestDistance = Infinity;
    
    for (const [key, value] of Object.entries(colorStandards)) {
      const standardColor = this.parseColor(value);
      if (!standardColor) continue;
      
      const distance = this.calculateColorDistance(clientColor, standardColor);
      if (distance < smallestDistance) {
        smallestDistance = distance;
        bestMatch = {
          key,
          value,
          confidence: Math.max(0, 1 - (distance / 100))
        };
      }
    }
    
    // Only return if colors are reasonably similar
    return smallestDistance < 50 ? bestMatch : null;
  }
  
  /**
   * Assess conflict level between client and company variables
   */
  assessConflictLevel(clientVar, companyMatch) {
    const category = clientVar.category.toLowerCase();
    
    if (category === 'colors') {
      return this.assessColorConflict(clientVar.value, companyMatch.value);
    } else if (category === 'typography') {
      return this.assessTypographyConflict(clientVar, companyMatch);
    } else if (category === 'spacing') {
      return this.assessSpacingConflict(clientVar.value, companyMatch.value);
    }
    
    return 'Low';
  }
  
  /**
   * Assess color conflicts
   */
  assessColorConflict(clientColor, companyColor) {
    const client = this.parseColor(clientColor);
    const company = this.parseColor(companyColor);
    
    if (!client || !company) return 'High';
    
    const distance = this.calculateColorDistance(client, company);
    
    if (distance < 20) return 'None';
    if (distance < 40) return 'Low';
    if (distance < 70) return 'Medium';
    return 'High';
  }
  
  /**
   * Generate CSS variable name following company conventions
   */
  generateCSSVariableName(category, standardKey) {
    const prefix = '--t-gs';
    const categoryShort = {
      'colors': 'color',
      'typography': 'font',
      'spacing': 'space',
      'components': 'comp',
      'layout': 'layout',
      'animations': 'anim'
    };
    
    const catPrefix = categoryShort[category] || category;
    const cleanKey = standardKey.replace(/^(font-|spacing-|color-)/, '');
    
    return `${prefix}-${catPrefix}-${cleanKey}`;
  }
  
  /**
   * Add mapping to spreadsheet
   */
  addMapping(mapping) {
    const row = [
      mapping.clientVariable,
      mapping.companyVariable,
      mapping.cssVariableName,
      mapping.mappingStatus,
      mapping.conflictLevel,
      mapping.resolution,
      '', // Approved By
      '', // Implementation Date
      `Auto-generated with ${Math.round(mapping.confidence * 100)}% confidence`
    ];
    
    this.mappingSheet.appendRow(row);
  }
  
  /**
   * Utility functions
   */
  normalizeVariableName(name) {
    return name.toLowerCase()
      .replace(/[-_\s]+/g, '-')
      .replace(/^-+|-+$/g, '');
  }
  
  calculateSemanticScore(clientName, keywords) {
    const normalizedClient = this.normalizeVariableName(clientName);
    let score = 0;
    
    keywords.forEach(keyword => {
      if (normalizedClient.includes(keyword)) {
        score += 1 / keywords.length;
      }
    });
    
    return score;
  }
  
  parseColor(colorString) {
    if (!colorString) return null;
    
    // Handle hex colors
    const hexMatch = colorString.match(/^#([A-Fa-f0-9]{6}|[A-Fa-f0-9]{3})$/);
    if (hexMatch) {
      const hex = hexMatch[1];
      const fullHex = hex.length === 3 ? 
        hex.split('').map(c => c + c).join('') : hex;
      
      return {
        r: parseInt(fullHex.substr(0, 2), 16),
        g: parseInt(fullHex.substr(2, 2), 16),
        b: parseInt(fullHex.substr(4, 2), 16)
      };
    }
    
    // Handle rgb colors
    const rgbMatch = colorString.match(/rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)/);
    if (rgbMatch) {
      return {
        r: parseInt(rgbMatch[1]),
        g: parseInt(rgbMatch[2]),
        b: parseInt(rgbMatch[3])
      };
    }
    
    return null;
  }
  
  calculateColorDistance(color1, color2) {
    // Calculate Euclidean distance in RGB space
    const dr = color1.r - color2.r;
    const dg = color1.g - color2.g;
    const db = color1.b - color2.b;
    
    return Math.sqrt(dr * dr + dg * dg + db * db);
  }
  
  assessTypographyConflict(clientVar, companyMatch) {
    // Simplified typography conflict assessment
    if (clientVar.name.includes('family') && companyMatch.key.includes('family')) {
      return clientVar.value === companyMatch.value ? 'None' : 'Medium';
    }
    return 'Low';
  }
  
  assessSpacingConflict(clientValue, companyValue) {
    const clientNum = parseFloat(clientValue);
    const companyNum = parseFloat(companyValue);
    
    if (isNaN(clientNum) || isNaN(companyNum)) return 'High';
    
    const diff = Math.abs(clientNum - companyNum);
    const avgValue = (clientNum + companyNum) / 2;
    const percentDiff = (diff / avgValue) * 100;
    
    if (percentDiff < 10) return 'None';
    if (percentDiff < 25) return 'Low';
    if (percentDiff < 50) return 'Medium';
    return 'High';
  }
  
  suggestResolution(clientVar, companyMatch, conflictLevel) {
    if (conflictLevel === 'None') {
      return 'Direct mapping - no conflicts detected';
    }
    
    const category = clientVar.category.toLowerCase();
    
    if (category === 'colors' && conflictLevel === 'High') {
      return 'Consider creating client-specific color variant or discuss brand flexibility';
    }
    
    if (category === 'typography' && conflictLevel === 'Medium') {
      return 'Evaluate font loading impact and brand requirements';
    }
    
    if (category === 'spacing' && conflictLevel === 'Medium') {
      return 'Assess impact on layout consistency and responsive behavior';
    }
    
    return 'Review with design team and client for approval';
  }
}

/**
 * Bulk processing functions
 */
function processBulkClientVariables(clientName, filePath = null) {
  const mapper = new DesignSystemMapper();
  
  if (filePath) {
    // If file path provided, parse variables from file
    const variables = parseVariablesFromFile(filePath);
    variables.forEach(variable => {
      addClientVariable(
        clientName,
        variable.portal || 'All Portals',
        variable.category,
        variable.name,
        variable.value,
        variable.type || 'Token',
        variable.context || '',
        variable.priority || 'Medium',
        variable.notes || 'Imported from file'
      );
    });
  }
  
  // Auto-map all variables for this client
  const results = mapper.autoMapClientVariables(clientName);
  
  // Generate report
  const report = generateMappingReport(clientName, results);
  
  return {
    processed: results.totalVariables,
    mapped: results.mappedVariables,
    conflicts: results.mappings.filter(m => m.conflictLevel !== 'None').length,
    report: report
  };
}

/**
 * Generate comprehensive mapping report
 */
function generateMappingReport(clientName, mappingResults) {
  const template = `
# Design System Mapping Report
**Client:** ${clientName}
**Generated:** ${new Date().toLocaleDateString()}

## Summary
- **Total Variables:** ${mappingResults.totalVariables}
- **Successfully Mapped:** ${mappingResults.mappedVariables}
- **Unmapped Variables:** ${mappingResults.unmappedVariables}
- **Conflicts Detected:** ${mappingResults.mappings.filter(m => m.conflictLevel !== 'None').length}

## Mapping Details

### Successful Mappings
${mappingResults.mappings.filter(m => m.conflictLevel === 'None').map(m => 
  `- ${m.clientVariable} → ${m.cssVariableName}`
).join('\n')}

### Conflicts Requiring Review
${mappingResults.mappings.filter(m => m.conflictLevel !== 'None').map(m => 
  `- **${m.clientVariable}** (${m.conflictLevel} conflict)
    - Maps to: ${m.cssVariableName}
    - Resolution: ${m.resolution}`
).join('\n')}

## Next Steps
1. Review high-priority conflicts with design team
2. Schedule client discussion for medium conflicts
3. Approve automatic mappings
4. Generate CSS variables for approved mappings
`;

  return template;
}

/**
 * Export functions for external use
 */
function autoMapClientDesignSystem(clientName, portalType = null) {
  const mapper = new DesignSystemMapper();
  return mapper.autoMapClientVariables(clientName, portalType);
}

function generateMappingReportForClient(clientName) {
  const mapper = new DesignSystemMapper();
  const results = mapper.autoMapClientVariables(clientName);
  return generateMappingReport(clientName, results);
}

function validateAllMappings() {
  const mapper = new DesignSystemMapper();
  const mappingSheet = mapper.mappingSheet;
  const data = mappingSheet.getDataRange().getValues();
  
  const validationResults = [];
  
  for (let i = 1; i < data.length; i++) {
    const [clientVar, companyVar, cssVar, status, conflictLevel] = data[i];
    
    const validation = {
      row: i + 1,
      clientVariable: clientVar,
      cssVariable: cssVar,
      isValid: true,
      issues: []
    };
    
    // Validate CSS variable naming
    if (!cssVar.startsWith('--t-gs-')) {
      validation.isValid = false;
      validation.issues.push('CSS variable does not follow naming convention');
    }
    
    // Validate conflict resolution
    if (conflictLevel === 'High' && status === 'Auto-Mapped') {
      validation.isValid = false;
      validation.issues.push('High conflict should not be auto-mapped');
    }
    
    validationResults.push(validation);
  }
  
  return validationResults;
}
