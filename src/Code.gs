/**
 * Yolwise Lead Scoring Add-on
 * Main entry point for Google Sheets Add-on
 * Integrates Claude 4 Sonnet API with Yolwise scoring system for Turkish B2B market
 * ADAPTED FROM SMARTWAY: Turkish business context, Yolwise API integration
 * Robust state persistence and timeout handling for continuation mechanism
 * Hybrid mapping system with Claude fallback for comprehensive data extraction
 * MERGED: English column mappings and interface from TurkishDataExtractor.gs
 */

/**
 * Add-on installation trigger
 * Sets up initial menu and configuration
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Document open trigger
 * Creates the add-on menu in Google Sheets
 */
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🎯 Yolwise Lead Scoring')
    .addItem('🎯 Analyze & Score', 'showScoringDialog')
    .addItem('⚙️ Settings', 'showSettingsDialog')
    .addItem('📖 Help', 'showHelpDialog')
    .addToUi();
}

/**
 * File scope granted trigger - required by manifest
 * Called when user grants file access permissions
 */
function onFileScopeGranted(e) {
  Logger.log('File scope granted:', e);
  onOpen(e);
}

/**
 * Main function to show the scoring dialog
 */
function showScoringDialog() {
  const html = HtmlService.createHtmlOutputFromFile('index')
    .setWidth(600)
    .setHeight(500)
    .setTitle('Yolwise Lead Scoring');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Yolwise Lead Scoring');
}

/**
 * Shows settings dialog for API configuration
 */
function showSettingsDialog() {
  const html = HtmlService.createHtml(`
    <div style="padding: 20px; font-family: Arial, sans-serif;">
      <h3>API Configuration</h3>
      <p><strong>Required API Keys:</strong></p>
      <ul>
        <li>Claude API Key (Anthropic)</li>
        <li>Yolwise API Key (from Replit)</li>
      </ul>
      <p><strong>Setup Instructions:</strong></p>
      <ol>
        <li>Go to <a href="https://console.anthropic.com" target="_blank">Anthropic Console</a></li>
        <li>Create API key for Claude 4 Sonnet</li>
        <li>Configure in Apps Script project settings</li>
        <li>Get Yolwise API key from Replit</li>
      </ol>
      <p><strong>API Configuration:</strong></p>
      <p>Add these keys to Apps Script Project Settings → Script Properties:</p>
      <ul>
        <li><code>CLAUDE_API_KEY</code> - Your Anthropic API key</li>
        <li><code>YOLWISE_API_KEY</code> - Yolwise scoring API key</li>
        <li><code>YOLWISE_API_URL</code> - https://yolwiseleadscoring.replit.app</li>
      </ul>
      <button onclick="google.script.host.close()">Close</button>
    </div>
  `)
    .setWidth(500)
    .setHeight(500)
    .setTitle('Settings');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Settings');
}

/**
 * Shows help dialog with usage instructions
 */
function showHelpDialog() {
  const html = HtmlService.createHtml(`
    <div style="padding: 20px; font-family: Arial, sans-serif;">
      <h3>Yolwise Lead Scoring Help</h3>
      <p><strong>How to use:</strong></p>
      <ol>
        <li>Prepare a table with company data</li>
        <li>Click "Analyze & Score" button</li>
        <li>Select data range and configure options</li>
        <li>Add-on will analyze with Claude API and score with Yolwise API</li>
        <li>Results will be added as new columns</li>
      </ol>
      <p><strong>Required columns:</strong> Company Name (minimum)</p>
      <p><strong>Optional columns:</strong> Industry, Annual Revenue, Number of Employees, City, etc.</p>
      
      <p><strong>Step-by-step setup:</strong></p>
      <ol>
        <li>Configure API keys in Settings</li>
        <li>Test connection with small dataset</li>
        <li>Review results and adjust parameters</li>
        <li>Run analysis on full dataset</li>
      </ol>
      
      <p><strong>Turkish B2B Market:</strong> This system is optimized for B2B service needs in Turkey.</p>
      
      <button onclick="google.script.host.close()">Close</button>
    </div>
  `)
    .setWidth(500)
    .setHeight(450)
    .setTitle('Help');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Help');
}

/**
 * Check if there's an interrupted processing session
 * @returns {Object|null} Processing state or null if none exists
 */
function checkProcessingState() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const stateStr = properties.getProperty('YOLWISE_SCORING_STATE');
    
    Logger.log('Checking for processing state:', stateStr ? 'Found' : 'None');
    
    if (!stateStr) {
      return null;
    }
    
    const state = JSON.parse(stateStr);
    
    // Validate state structure
    if (!state.options || !state.processedRows || typeof state.processedRows !== 'number') {
      Logger.log('Invalid state structure, clearing');
      clearProcessingState();
      return null;
    }
    
    // Check if state is recent (within 24 hours)
    const stateAge = Date.now() - (state.timestamp || 0);
    if (stateAge > 24 * 60 * 60 * 1000) {
      Logger.log('State too old, clearing');
      clearProcessingState();
      return null;
    }
    
    Logger.log('Valid processing state found:', {
      processed: state.processedRows,
      total: state.results ? state.results.total : 'unknown',
      age: Math.round(stateAge / 1000 / 60) + ' minutes'
    });
    
    return state;
    
  } catch (error) {
    Logger.log('Error checking processing state:', error);
    clearProcessingState();
    return null;
  }
}

/**
 * Save current processing state
 * @param {Object} state Processing state to save
 */
function saveProcessingState(state) {
  try {
    if (!state || !state.options || typeof state.processedRows !== 'number') {
      Logger.log('Invalid state provided for saving:', state);
      return;
    }
    
    const properties = PropertiesService.getScriptProperties();
    state.timestamp = Date.now();
    state.version = '2.0-yolwise';
    
    const stateStr = JSON.stringify(state);
    properties.setProperty('YOLWISE_SCORING_STATE', stateStr);
    
    Logger.log('Processing state saved:', {
      processed: state.processedRows,
      total: state.results ? state.results.total : 'unknown',
      timestamp: new Date(state.timestamp).toLocaleString('en-US')
    });
  } catch (error) {
    Logger.log('Error saving processing state:', error);
  }
}

/**
 * Clear processing state
 */
function clearProcessingState() {
  try {
    const properties = PropertiesService.getScriptProperties();
    properties.deleteProperty('YOLWISE_SCORING_STATE');
    Logger.log('Processing state cleared successfully');
  } catch (error) {
    Logger.log('Error clearing processing state:', error);
  }
}

/**
 * Main scoring function with enhanced timeout handling
 * ADAPTED: Turkish business context and Yolwise API integration
 * @param {Object} options Configuration options from UI
 * @returns {Object} Processing results
 */
function processLeadScoring(options) {
  const startTime = Date.now();
  const timeoutMargin = 30000; // 30 seconds before 6-minute limit
  const maxExecutionTime = 6 * 60 * 1000 - timeoutMargin; // 5.5 minutes
  
  try {
    Logger.log('Starting Yolwise lead scoring process with options:', options);
    Logger.log('Execution timeout monitoring: ' + (maxExecutionTime / 1000) + ' seconds');
    
    // Check for existing processing state
    let processingState = checkProcessingState();
    let isResume = false;
    
    if (processingState) {
      Logger.log('Found existing processing state, resuming from row:', processingState.processedRows);
      isResume = true;
    }
    
    // Get active spreadsheet and sheet
    const sheet = SpreadsheetApp.getActiveSheet();
    const range = sheet.getRange(options.dataRange);
    const data = range.getValues();
    
    if (data.length === 0) {
      throw new Error('No data found in specified range');
    }
    
    // Process headers and data
    const headers = data[0];
    const allRows = data.slice(1);
    
    Logger.log(`Processing ${allRows.length} Turkish companies with headers:`, headers.slice(0, 5) + '...');

    // Initialize or restore progress tracking
    let results;
    let startIndex = 0;
    
    if (isResume && processingState) {
      // Resume from where we left off
      results = processingState.results;
      startIndex = processingState.processedRows;
      Logger.log(`Resuming from row ${startIndex + 1} of ${allRows.length}`);
    } else {
      // Start fresh - clear any existing state
      clearProcessingState();
      results = {
        total: allRows.length,
        processed: 0,
        successful: 0,
        failed: 0,
        results: []
      };
    }
    
    // Get rows to process (remaining rows)
    const rowsToProcess = allRows.slice(startIndex);
    
    // Process companies individually with frequent state saves
    for (let i = 0; i < rowsToProcess.length; i++) {
      const currentTime = Date.now();
      const elapsedTime = currentTime - startTime;
      
      // Check if we're approaching timeout
      if (elapsedTime > maxExecutionTime) {
        Logger.log(`Approaching timeout (${Math.round(elapsedTime/1000)}s elapsed), saving state and stopping`);
        
        // Save state for continuation
        saveProcessingState({
          options: options,
          headers: headers,
          processedRows: startIndex + i,
          results: results
        });
        
        // Return partial results with continuation info
        return {
          ...results,
          incomplete: true,
          continuationAvailable: true,
          message: `Processed: ${results.processed} / ${results.total} companies. Use "Continue" button to proceed.`
        };
      }
      
      const row = rowsToProcess[i];
      const actualRowIndex = startIndex + i;
      
      try {
        // Extract company data from row using HYBRID MAPPING SYSTEM (adapted for English)
        const companyData = extractCompanyDataHybrid(row, headers);
        
        if (!companyData.company_name) {
          results.results.push({
            success: false,
            error: 'Company name not found',
            companyData: companyData,
            qualityAnalysis: companyData.quality_analysis || 'Company name not found'
          });
          results.failed++;
        } else {
          Logger.log(`Processing Turkish company ${actualRowIndex + 1}/${results.total}: ${companyData.company_name}`);
          
          // Analyze with Claude API (Turkish context)
          const claudeAnalysis = analyzeWithClaude(companyData, options);
          
          // Score with Yolwise API
          const scoringResult = scoreWithYolwise(claudeAnalysis);
          
          results.results.push({
            success: true,
            companyData: companyData,
            claudeAnalysis: claudeAnalysis,
            scoringResult: scoringResult,
            qualityAnalysis: companyData.quality_analysis || 'All core fields successfully extracted'
          });
          
          results.successful++;
        }
        
        results.processed++;
        
        // Save state every 5 companies
        if ((i + 1) % 5 === 0) {
          saveProcessingState({
            options: options,
            headers: headers,
            processedRows: actualRowIndex + 1,
            results: results
          });
          
          Logger.log(`Progress: ${results.processed}/${results.total} (${Math.round(results.processed/results.total*100)}%)`);
        }
        
        // Small delay to avoid rate limits
        Utilities.sleep(1500);
        
      } catch (companyError) {
        Logger.log('Error processing company:', companyError);
        results.results.push({
          success: false,
          error: companyError.toString(),
          companyData: extractCompanyDataHybrid(row, headers),
          qualityAnalysis: 'Critical processing error'
        });
        results.processed++;
        results.failed++;
      }
    }
    
    // Insert results into spreadsheet
    insertResults(sheet, range, results.results, options);
    
    // Clear processing state on successful completion
    clearProcessingState();
    
    Logger.log('Yolwise lead scoring completed successfully:', results);
    return results;
    
  } catch (error) {
    Logger.log('Error in processLeadScoring:', error);
    
    // Save state on error for continuation
    const elapsedTime = Date.now() - startTime;
    Logger.log(`Error after ${Math.round(elapsedTime/1000)} seconds - state should be preserved for continuation`);
    
    throw error;
  }
}

/**
 * Continue processing from saved state
 * @returns {Object} Processing results
 */
function continueProcessing() {
  try {
    const processingState = checkProcessingState();
    
    if (!processingState) {
      throw new Error('Processing state not found. Session may have expired.');
    }
    
    Logger.log('Continuing processing from saved state');
    return processLeadScoring(processingState.options);
    
  } catch (error) {
    Logger.log('Error in continueProcessing:', error);
    throw error;
  }
}

/**
 * Get available data ranges in the current sheet
 * @returns {Array} Available ranges
 */
function getAvailableRanges() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    if (lastRow === 0 || lastCol === 0) {
      return [];
    }
    
    // Suggest common ranges
    const ranges = [];
    
    // Full data range
    ranges.push({
      label: `All data (A1:${sheet.getRange(lastRow, lastCol).getA1Notation().match(/[A-Z]+/)[0]}${lastRow})`,
      value: `A1:${sheet.getRange(lastRow, lastCol).getA1Notation().match(/[A-Z]+/)[0]}${lastRow}`,
      description: `${lastRow} rows, ${lastCol} columns`
    });
    
    // Data without header
    if (lastRow > 1) {
      ranges.push({
        label: `Data only (A2:${sheet.getRange(lastRow, lastCol).getA1Notation().match(/[A-Z]+/)[0]}${lastRow})`,
        value: `A2:${sheet.getRange(lastRow, lastCol).getA1Notation().match(/[A-Z]+/)[0]}${lastRow}`,
        description: `${lastRow - 1} data rows, ${lastCol} columns`
      });
    }
    
    return ranges;
    
  } catch (error) {
    Logger.log('Error getting available ranges:', error);
    return [];
  }
}

/**
 * HYBRID MAPPING SYSTEM: Extract company data with structured mapping + Claude fallback
 * Primary extraction function for Turkish business data with English column mapping
 * @param {Array} row Spreadsheet row data
 * @param {Array} headers Column headers
 * @returns {Object} Structured company data with quality analysis
 */
function extractCompanyDataHybrid(row, headers) {
  Logger.log('🔄 Starting hybrid mapping system for company data');
  
  // Step 1: Use structured (code-based) mapping for business data
  const structuredData = extractCompanyDataStructured(row, headers);
  
  // Step 2: Identify critical missing fields
  const criticalFields = ['company_name', 'industry', 'revenue_estimate', 'employees_estimate', 'headquarters'];
  const missingCriticalFields = criticalFields.filter(field => !structuredData[field] || structuredData[field] === '');
  
  // Step 3: Track quality analysis
  const qualityAnalysis = {
    structured_mapping_success: criticalFields.length - missingCriticalFields.length,
    total_critical_fields: criticalFields.length,
    missing_fields: missingCriticalFields,
    claude_fallback_used: false,
    claude_mapping_success: 0,
    final_data_completeness: 0
  };
  
  Logger.log(`📊 Structured mapping results: ${qualityAnalysis.structured_mapping_success}/${qualityAnalysis.total_critical_fields} critical fields found`);
  
  // Step 4: Apply Claude-based fallback if critical fields are missing
  let finalData = { ...structuredData };
  
  if (missingCriticalFields.length > 0) {
    Logger.log(`🤖 Applying Claude fallback for missing fields: ${missingCriticalFields.join(', ')}`);
    
    try {
      const claudeMappedData = mapHeadersWithClaude(headers, row, missingCriticalFields);
      
      // Merge Claude results for missing fields only
      missingCriticalFields.forEach(field => {
        if (claudeMappedData[field] && claudeMappedData[field] !== '') {
          finalData[field] = claudeMappedData[field];
          qualityAnalysis.claude_mapping_success++;
          Logger.log(`✅ Claude successfully mapped: ${field} = ${claudeMappedData[field]}`);
        }
      });
      
      qualityAnalysis.claude_fallback_used = true;
      
    } catch (claudeError) {
      Logger.log('❌ Claude fallback mapping failed:', claudeError);
      qualityAnalysis.claude_fallback_error = claudeError.toString();
    }
  }
  
  // Step 5: Calculate final quality metrics
  const finalCriticalFields = criticalFields.filter(field => finalData[field] && finalData[field] !== '');
  qualityAnalysis.final_data_completeness = (finalCriticalFields.length / criticalFields.length) * 100;
  
  // Step 6: Generate quality report
  const qualityReport = generateQualityReport(qualityAnalysis, headers.length);
  finalData.quality_analysis = qualityReport;
  
  Logger.log(`📈 Final mapping quality: ${Math.round(qualityAnalysis.final_data_completeness)}% completeness`);
  
  return finalData;
}

/**
 * Structured (code-based) mapping for business data
 * ADAPTED: English business data based on Target Leads.xlsx structure
 * @param {Array} row Spreadsheet row data
 * @param {Array} headers Column headers
 * @returns {Object} Structured company data
 */
function extractCompanyDataStructured(row, headers) {
  const data = {};
  
  // COMPREHENSIVE field mappings for English business data (based on Target Leads.xlsx)
  const fieldMappings = {
    // Company name variations (English)
    'company_name': [
      'company name', 'company', 'name', 'firm', 'organization', 'entity',
      'business name', 'corp', 'corporation', 'inc', 'ltd', 'llc',
      'enterprise', 'group', 'holdings', 'business', 'establishment'
    ],
    
    // Industry and business activity (English)
    'industry': [
      'industry', 'sector', 'business', 'field', 'market', 'vertical',
      'category', 'classification', 'type', 'domain', 'area',
      'specialization', 'focus', 'activity', 'segment', 'niche'
    ],
    
    // Financial data (Revenue/Sales)
    'revenue_estimate': [
      'annual revenue', 'revenue', 'sales', 'turnover', 'income', 'earnings',
      'gross revenue', 'net revenue', 'total revenue', 'yearly revenue',
      'financial performance', 'sales volume', 'gross income'
    ],
    
    // Employee information (English)
    'employees_estimate': [
      'number of employees', 'employees', 'staff', 'workforce', 'headcount',
      'team size', 'personnel', 'workers', 'employee count', 'staff size',
      'total employees', 'company size', 'organizational size'
    ],
    
    // Business type and legal form (English)
    'business_type': [
      'company size', 'business type', 'legal form', 'entity type',
      'incorporation type', 'structure', 'organization type', 'format'
    ],
    
    // Location and address (English geography)
    'headquarters': [
      'city', 'location', 'headquarters', 'head office', 'address',
      'main office', 'primary location', 'base', 'region', 'state',
      'country', 'geography', 'locale', 'territory'
    ],
    
    // Website and digital presence
    'website': [
      'website', 'domain', 'url', 'web', 'site', 'homepage',
      'web address', 'internet address', 'online presence',
      'company domain name', 'web site', 'portal'
    ],
    
    // Contact information
    'contact_info': [
      'phone', 'telephone', 'contact', 'phone number', 'tel',
      'mobile', 'cell', 'landline', 'contact number', 'business phone'
    ],
    
    // Email information
    'email': [
      'email', 'e-mail', 'electronic mail', 'contact email',
      'business email', 'corporate email', 'company email'
    ],
    
    // Founded/Established information
    'founding_year': [
      'year founded', 'founded', 'established', 'inception',
      'start year', 'creation date', 'establishment date', 'launch year'
    ],
    
    // Social media presence
    'linkedin_url': [
      'linkedin', 'linkedin company page', 'linkedin profile',
      'linkedin url', 'professional network', 'business network'
    ],
    
    'facebook_url': [
      'facebook', 'facebook company page', 'facebook page',
      'facebook url', 'social media', 'fb page'
    ],
    
    // Additional business context
    'description': [
      'description', 'about', 'overview', 'summary', 'profile',
      'company description', 'business description', 'details',
      'information', 'background', 'company overview'
    ],
    
    // Address components
    'street_address': [
      'street address', 'address', 'street', 'location', 'physical address'
    ],
    
    'postal_code': [
      'postal code', 'zip code', 'postcode', 'zip', 'mail code'
    ],
    
    'state_region': [
      'state', 'region', 'province', 'territory', 'state/region'
    ],
    
    // Record and system fields
    'record_id': [
      'record id', 'id', 'identifier', 'unique id', 'system id'
    ],
    
    'timezone': [
      'time zone', 'timezone', 'tz', 'time zone info'
    ]
  };
  
  // Track which headers were successfully mapped
  const mappedHeaders = new Set();
  
  // Extract data based on comprehensive English header mappings
  headers.forEach((header, index) => {
    const normalizedHeader = header.toLowerCase().trim();
    
    for (const [field, variations] of Object.entries(fieldMappings)) {
      if (variations.some(v => normalizedHeader.includes(v))) {
        // Store the raw value
        const cellValue = row[index];
        if (cellValue !== null && cellValue !== undefined && cellValue !== '') {
          data[field] = cellValue;
          mappedHeaders.add(index);
          Logger.log(`Mapped English header "${header}" to field "${field}" with value:`, cellValue);
        }
        break;
      }
    }
  });
  
  // INTELLIGENT CATEGORIZATION of unmapped fields with business context
  const businessCategories = {
    financial_data: [],
    legal_data: [], 
    operational_data: [],
    contact_data: [],
    other_data: []
  };
  
  headers.forEach((header, index) => {
    if (!mappedHeaders.has(index) && row[index] && row[index] !== '') {
      const headerLower = header.toLowerCase();
      const value = row[index];
      const entry = `${header}: ${value}`;
      
      // Categorize based on English business terminology
      if (headerLower.includes('revenue') || headerLower.includes('sales') || 
          headerLower.includes('income') || headerLower.includes('financial') ||
          headerLower.includes('budget') || headerLower.includes('profit')) {
        businessCategories.financial_data.push(entry);
      } else if (headerLower.includes('llc') || headerLower.includes('inc') ||
                 headerLower.includes('corp') || headerLower.includes('ltd') ||
                 headerLower.includes('legal') || headerLower.includes('license') ||
                 headerLower.includes('registration')) {
        businessCategories.legal_data.push(entry);
      } else if (headerLower.includes('branch') || headerLower.includes('office') ||
                 headerLower.includes('department') || headerLower.includes('division') ||
                 headerLower.includes('operation') || headerLower.includes('facility')) {
        businessCategories.operational_data.push(entry);
      } else if (headerLower.includes('phone') || headerLower.includes('email') ||
                 headerLower.includes('contact') || headerLower.includes('address') ||
                 headerLower.includes('fax') || headerLower.includes('mail')) {
        businessCategories.contact_data.push(entry);
      } else {
        businessCategories.other_data.push(entry);
      }
    }
  });
  
  // Combine all categorized data for Claude analysis
  const discoveredFacts = [];
  Object.entries(businessCategories).forEach(([category, items]) => {
    if (items.length > 0) {
      discoveredFacts.push(`${category}: ${items.join('; ')}`);
    }
  });
  
  data.discovered_facts = discoveredFacts;
  
  // Enhanced logging for debugging
  Logger.log(`Extracted company data for ${data.company_name || 'Unknown company'}:`, {
    mappedFields: Object.keys(data).filter(k => k !== 'discovered_facts'),
    categorizedFacts: Object.entries(businessCategories).map(([k,v]) => `${k}: ${v.length} items`),
    totalFactsExtracted: discoveredFacts.length
  });
  
  return data;
}

/**
 * Claude-based header mapping for business data
 * ADAPTED: English business terminology and context
 * @param {Array} headers All column headers
 * @param {Array} row Corresponding row data
 * @param {Array} missingFields List of critical fields that need mapping
 * @returns {Object} Claude-mapped data for missing fields
 */
function mapHeadersWithClaude(headers, row, missingFields) {
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
    if (!apiKey) {
      throw new Error('Claude API key not configured');
    }
    
    // Prepare unmapped headers with their values for Claude analysis
    const unmappedHeaders = [];
    headers.forEach((header, index) => {
      if (row[index] && row[index] !== '') {
        unmappedHeaders.push({
          header: header,
          value: row[index],
          index: index
        });
      }
    });
    
    const prompt = `Analyze table headers containing business data and find mappings for missing fields.

MISSING FIELDS (need to be found): ${missingFields.join(', ')}

AVAILABLE HEADERS AND VALUES:
${unmappedHeaders.map(item => `"${item.header}": "${item.value}"`).join('\n')}

TASK: Find the most appropriate header from the above list for each missing field.

MAPPING RULES:
- company_name: company names, firm names, business names
- industry: sectors, business areas, activity fields, industries
- revenue_estimate: revenue, sales, income, turnover amounts
- employees_estimate: employee count, staff size, workforce
- headquarters: cities, states, addresses, locations

RETURN ONLY JSON format:
{
  "field_name": "extracted_value",
  "field_name2": "extracted_value2"
}

If field cannot be found, do not include it in response.`;

    // Make API request to Claude
    const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      payload: JSON.stringify({
        model: 'claude-sonnet-4-20250514',
        max_tokens: 1000,
        temperature: 0.1,
        messages: [{
          role: 'user',
          content: prompt
        }]
      })
    });
    
    if (response.getResponseCode() !== 200) {
      throw new Error(`Claude API error: ${response.getResponseCode()}`);
    }
    
    const responseData = JSON.parse(response.getContentText());
    let content = responseData.content[0].text.trim();
    
    // Extract JSON from response
    if (content.includes('```json')) {
      const match = content.match(/```json\s*([\s\S]*?)\s*```/);
      if (match) {
        content = match[1];
      }
    }
    
    const claudeMappedData = JSON.parse(content);
    
    Logger.log('🤖 Claude successfully mapped headers:', Object.keys(claudeMappedData));
    return claudeMappedData;
    
  } catch (error) {
    Logger.log('❌ Claude header mapping failed:', error);
    throw error;
  }
}

/**
 * Generate quality analysis report
 * @param {Object} qualityAnalysis Quality metrics object
 * @param {number} totalHeaders Total number of headers in spreadsheet
 * @returns {string} Human-readable quality report
 */
function generateQualityReport(qualityAnalysis, totalHeaders) {
  const reports = [];
  
  // Overall completeness
  const completeness = Math.round(qualityAnalysis.final_data_completeness);
  if (completeness >= 80) {
    reports.push(`✅ Excellent data quality (${completeness}%)`);
  } else if (completeness >= 60) {
    reports.push(`⚠️ Good data quality (${completeness}%)`);
  } else {
    reports.push(`❌ Poor data quality (${completeness}%)`);
  }
  
  // Structured mapping performance
  reports.push(`Structured mapping: ${qualityAnalysis.structured_mapping_success}/${qualityAnalysis.total_critical_fields}`);
  
  // Claude fallback usage
  if (qualityAnalysis.claude_fallback_used) {
    if (qualityAnalysis.claude_mapping_success > 0) {
      reports.push(`Claude mapping: +${qualityAnalysis.claude_mapping_success} fields`);
    } else {
      reports.push(`Claude mapping: failed`);
    }
  }
  
  // Missing fields warning
  if (qualityAnalysis.missing_fields.length > 0) {
    reports.push(`Missing: ${qualityAnalysis.missing_fields.join(', ')}`);
  }
  
  return reports.join(' | ');
}

/**
 * Insert scoring results into spreadsheet with English labels
 * ADAPTED: English column headers and color coding
 * @param {Sheet} sheet Target sheet
 * @param {Range} originalRange Original data range
 * @param {Array} results Scoring results
 * @param {Object} options Configuration options
 */
function insertResults(sheet, originalRange, results, options) {
  try {
    // Define result columns with English labels
    const resultHeaders = [
      'API Score',
      'Industry Multiplier', 
      'LLM Adjustment',
      'Final Score',
      'Priority',
      'Explanation',
      'Data Quality'
    ];
    
    // Find next available column
    const lastCol = originalRange.getLastColumn();
    const resultStartCol = lastCol + 1;
    
    // Insert headers
    const headerRange = sheet.getRange(originalRange.getRow(), resultStartCol, 1, resultHeaders.length);
    headerRange.setValues([resultHeaders]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#E8F0FE');
    
    // Prepare result data with English labels
    const resultData = results.map(result => {
      if (!result.success) {
        return [
          'ERROR',
          '',
          '',
          '',
          'ERROR',
          result.error || 'Unknown error',
          result.qualityAnalysis || 'Critical processing error'
        ];
      }
      
      const scoring = result.scoringResult;
      
      // Calculate scores according to Yolwise specification
      const apiScore = scoring.industry_adjusted_score || scoring.base_score || 0;
      const llmAdjustment = Math.max(-25, Math.min(25, scoring.llm_adjustment || 0));
      const finalScore = Math.max(0, Math.min(100, apiScore + llmAdjustment));
      
      // Priority based on final score (60+ = target)
      const priority = finalScore >= 60 ? 'target' : 'non_target';
      const priorityEn = finalScore >= 60 ? 'TARGET' : 'NON-TARGET';
      
      return [
        Math.round(apiScore * 10) / 10,
        scoring.industry_multiplier || 1.0,
        llmAdjustment,
        Math.round(finalScore * 10) / 10,
        priorityEn,
        scoring.reasoning || 'No explanation available',
        result.qualityAnalysis || 'Quality analysis not available'
      ];
    });
    
    // Insert result data
    if (resultData.length > 0) {
      const dataRange = sheet.getRange(
        originalRange.getRow() + 1,
        resultStartCol,
        resultData.length,
        resultHeaders.length
      );
      dataRange.setValues(resultData);
      
      // Apply conditional formatting for priority
      const priorityCol = resultStartCol + 4; // Priority column
      const priorityRange = sheet.getRange(
        originalRange.getRow() + 1,
        priorityCol,
        resultData.length,
        1
      );
      
      // Color coding for priorities
      const values = priorityRange.getValues();
      const backgrounds = values.map(([priority]) => {
        const priorityStr = String(priority).toLowerCase();
        if (priorityStr.includes('target') && !priorityStr.includes('non')) {
          return ['#D4EDDA']; // Green for target
        } else if (priorityStr.includes('non-target') || priorityStr.includes('non_target')) {
          return ['#F8D7DA']; // Red for non-target
        } else {
          return ['#FFF3CD']; // Yellow for unknown
        }
      });
      priorityRange.setBackgrounds(backgrounds);
      
      // Apply conditional formatting for quality analysis
      const qualityCol = resultStartCol + 6; // Quality analysis column
      const qualityRange = sheet.getRange(
        originalRange.getRow() + 1,
        qualityCol,
        resultData.length,
        1
      );
      
      // Color coding for quality
      const qualityValues = qualityRange.getValues();
      const qualityBackgrounds = qualityValues.map(([quality]) => {
        const qualityStr = String(quality).toLowerCase();
        if (qualityStr.includes('excellent') || qualityStr.includes('✅')) {
          return ['#D4EDDA']; // Green for excellent quality
        } else if (qualityStr.includes('good') || qualityStr.includes('⚠️')) {
          return ['#FFF3CD']; // Yellow for good quality
        } else if (qualityStr.includes('poor') || qualityStr.includes('❌')) {
          return ['#F8D7DA']; // Red for poor quality
        } else {
          return ['#F8F9FA']; // Light gray for neutral
        }
      });
      qualityRange.setBackgrounds(qualityBackgrounds);
    }
    
    Logger.log('Results inserted successfully with English labels and quality analysis');
    
  } catch (error) {
    Logger.log('Error inserting results:', error);
    throw error;
  }
}

/**
 * BACKWARD COMPATIBILITY: Keep original function for existing systems
 * @param {Array} row Spreadsheet row data
 * @param {Array} headers Column headers
 * @returns {Object} Structured company data
 */
function extractCompanyData(row, headers) {
  // Call the new hybrid system
  return extractCompanyDataHybrid(row, headers);
}

/**
 * Analyze company data with Claude API for Yolwise context
 * @param {Object} companyData Extracted company data
 * @param {Object} options Analysis options
 * @returns {Object} Claude analysis results
 */
function analyzeWithClaude(companyData, options) {
  // Placeholder - implement Claude analysis for Turkish B2B context
  Logger.log('Analyzing with Claude API for company:', companyData.company_name);
  
  // This would call Claude API with Yolwise-specific prompts
  // For now, return mock data structure
  return {
    company_analysis: `Turkish B2B analysis for ${companyData.company_name}`,
    b2b_service_propensity: 'medium',
    market_indicators: ['established_business', 'service_oriented'],
    risk_factors: [],
    growth_potential: 'moderate'
  };
}

/**
 * Score company with Yolwise API
 * @param {Object} analysisData Claude analysis results
 * @returns {Object} Yolwise scoring results
 */
function scoreWithYolwise(analysisData) {
  // Placeholder - implement Yolwise API scoring
  Logger.log('Scoring with Yolwise API');
  
  // This would call the Yolwise API at https://yolwiseleadscoring.replit.app
  // For now, return mock scoring structure
  return {
    base_score: 72,
    industry_multiplier: 1.15,
    industry_adjusted_score: 82.8,
    llm_adjustment: 5,
    final_score: 87.8,
    priority_recommendation: 'target',
    reasoning: 'Strong B2B service propensity for Turkish market',
    detected_industry: 'Manufacturing',
    confidence: 'high'
  };
}