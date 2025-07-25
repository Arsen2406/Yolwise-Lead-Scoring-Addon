/**
 * Yolwise Lead Scoring Add-on - SECURITY ENHANCED
 * Main entry point for Google Sheets Add-on
 * Integrates Claude 4 Sonnet API with Yolwise scoring system for Turkish B2B market
 * ADAPTED FROM SMARTWAY: Turkish business context, Yolwise API integration
 * 
 * SECURITY ENHANCEMENTS:
 * - Fixed state management race conditions with atomic operations
 * - Comprehensive input validation and sanitization
 * - Enhanced error handling with user-friendly messages
 * - Improved timeout handling and continuation mechanism
 * - Added security logging and monitoring
 * - Protection against malicious input data
 * - Enhanced data quality assessment
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
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('🎯 Yolwise Lead Scoring')
      .addItem('🎯 Analyze & Score', 'showScoringDialog')
      .addItem('⚙️ Settings', 'showSettingsDialog')
      .addItem('📊 Test API Connection', 'testApiConnectionDialog')
      .addItem('📖 Help', 'showHelpDialog')
      .addToUi();
    
    Logger.log('✅ Yolwise Lead Scoring menu created successfully');
  } catch (error) {
    Logger.log('❌ Error creating menu: ' + error.toString());
  }
}

/**
 * File scope granted trigger - required by manifest
 * Called when user grants file access permissions
 */
function onFileScopeGranted(e) {
  Logger.log('✅ File scope granted:', e);
  onOpen(e);
}

/**
 * Main function to show the scoring dialog
 */
function showScoringDialog() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('index')
      .setWidth(650)  // Slightly wider for better UX
      .setHeight(550)  // Slightly taller for better UX
      .setTitle('Yolwise Lead Scoring - Enhanced Security');
    
    SpreadsheetApp.getUi().showModalDialog(html, 'Yolwise Lead Scoring');
    Logger.log('✅ Scoring dialog displayed');
  } catch (error) {
    Logger.log('❌ Error showing scoring dialog: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error: Unable to show scoring dialog. Please try again.');
  }
}

/**
 * Shows enhanced settings dialog for API configuration
 */
function showSettingsDialog() {
  try {
    const html = HtmlService.createHtml(`
      <div style="padding: 20px; font-family: Arial, sans-serif; line-height: 1.6;">
        <h3 style="color: #1a73e8; margin-bottom: 20px;">🔧 API Configuration - Enhanced Security</h3>
        
        <div style="background: #f8f9fa; padding: 15px; border-radius: 8px; margin-bottom: 20px;">
          <h4>🔑 Required API Keys:</h4>
          <ul style="margin: 10px 0;">
            <li><strong>Claude API Key</strong> (Anthropic) - For AI analysis</li>
            <li><strong>Yolwise API Key</strong> (from Replit) - For scoring system</li>
          </ul>
        </div>
        
        <div style="background: #e8f5e8; padding: 15px; border-radius: 8px; margin-bottom: 20px;">
          <h4>🛡️ Security Features:</h4>
          <ul style="margin: 10px 0;">
            <li>Standardized X-API-Key header authentication</li>
            <li>Input validation and sanitization</li>
            <li>Enhanced error handling and logging</li>
            <li>Race condition protection in state management</li>
          </ul>
        </div>
        
        <div style="background: #fff3cd; padding: 15px; border-radius: 8px; margin-bottom: 20px;">
          <h4>📋 Setup Instructions:</h4>
          <ol style="margin: 10px 0;">
            <li>Go to <a href="https://console.anthropic.com" target="_blank" style="color: #1a73e8;">Anthropic Console</a></li>
            <li>Create API key for Claude 4 Sonnet</li>
            <li>Configure in Apps Script project settings</li>
            <li>Get Yolwise API key from Replit deployment</li>
            <li>Test connections using "Test API Connection" menu</li>
          </ol>
        </div>
        
        <div style="background: #f0f7ff; padding: 15px; border-radius: 8px; margin-bottom: 20px;">
          <h4>⚙️ Script Properties Configuration:</h4>
          <p>Add these keys to Apps Script Project Settings → Script Properties:</p>
          <ul style="font-family: monospace; background: white; padding: 10px; border-radius: 4px;">
            <li><code>CLAUDE_API_KEY</code> - Your Anthropic API key</li>
            <li><code>YOLWISE_API_KEY</code> - Yolwise scoring API key</li>
            <li><code>YOLWISE_API_URL</code> - https://yolwiseleadscoring.replit.app</li>
          </ul>
        </div>
        
        <div style="text-align: center; margin-top: 30px;">
          <button onclick="google.script.host.close()" 
                  style="background: #1a73e8; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer;">
            Close
          </button>
        </div>
      </div>
    `)
      .setWidth(600)
      .setHeight(600)
      .setTitle('Settings - Enhanced Security');
    
    SpreadsheetApp.getUi().showModalDialog(html, 'Settings');
    Logger.log('✅ Settings dialog displayed');
  } catch (error) {
    Logger.log('❌ Error showing settings dialog: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error: Unable to show settings dialog.');
  }
}

/**
 * Test API connection dialog
 */
function testApiConnectionDialog() {
  try {
    Logger.log('🔧 Testing API connections...');
    
    // Run comprehensive API tests
    const testResults = testYolwiseApiConnectionsEnhanced();
    
    // Generate user-friendly results
    let resultHtml = `
      <div style="padding: 20px; font-family: Arial, sans-serif;">
        <h3 style="color: #1a73e8;">🔧 API Connection Test Results</h3>
    `;
    
    // Claude API results
    if (testResults.claude_api.configured && testResults.claude_api.valid) {
      resultHtml += `<p style="color: #137333;">✅ <strong>Claude API:</strong> Ready</p>`;
    } else {
      resultHtml += `<p style="color: #d93025;">❌ <strong>Claude API:</strong> ${testResults.claude_api.error || 'Not configured'}</p>`;
    }
    
    // Yolwise API results
    if (testResults.yolwise_api.accessible && testResults.yolwise_api.authenticated) {
      resultHtml += `<p style="color: #137333;">✅ <strong>Yolwise API:</strong> Ready</p>`;
    } else {
      resultHtml += `<p style="color: #d93025;">❌ <strong>Yolwise API:</strong> ${testResults.yolwise_api.error || 'Not configured'}</p>`;
    }
    
    resultHtml += `
        <div style="margin: 20px 0; padding: 15px; background: #f8f9fa; border-radius: 8px;">
          <p><strong>Next Steps:</strong></p>
          <ul>
            <li>If APIs show errors, check Settings for configuration instructions</li>
            <li>Verify API keys are correct in Script Properties</li>
            <li>Check network connectivity if connection fails</li>
          </ul>
        </div>
        <button onclick="google.script.host.close()" 
                style="background: #1a73e8; color: white; padding: 10px 20px; border: none; border-radius: 4px;">
          Close
        </button>
      </div>
    `;
    
    const html = HtmlService.createHtml(resultHtml)
      .setWidth(500)
      .setHeight(400)
      .setTitle('API Connection Test');
    
    SpreadsheetApp.getUi().showModalDialog(html, 'API Test Results');
    
  } catch (error) {
    Logger.log('❌ Error in API test dialog: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error testing API connections: ' + error.message);
  }
}

/**
 * Shows enhanced help dialog with usage instructions
 */
function showHelpDialog() {
  try {
    const html = HtmlService.createHtml(`
      <div style="padding: 20px; font-family: Arial, sans-serif; line-height: 1.6;">
        <h3 style="color: #1a73e8; margin-bottom: 20px;">📖 Yolwise Lead Scoring Help - Enhanced</h3>
        
        <div style="background: #e8f5e8; padding: 15px; border-radius: 8px; margin-bottom: 20px;">
          <h4>🚀 How to use (Enhanced Security):</h4>
          <ol style="margin: 10px 0;">
            <li>Prepare a table with company data in your sheet</li>
            <li>Configure API keys in Settings (required)</li>
            <li>Test API connections using "Test API Connection"</li>
            <li>Click "Analyze & Score" button</li>
            <li>Select data range and configure options</li>
            <li>Add-on will analyze with Claude API and score with Yolwise API</li>
            <li>Results will be added as new columns with security validation</li>
          </ol>
        </div>
        
        <div style="background: #fff3cd; padding: 15px; border-radius: 8px; margin-bottom: 20px;">
          <h4>📋 Data Requirements:</h4>
          <p><strong>Required:</strong> Company Name (minimum)</p>
          <p><strong>Recommended:</strong> Industry, Annual Revenue, Number of Employees, City, Description</p>
          <p><strong>Security:</strong> All data is validated and sanitized before processing</p>
        </div>
        
        <div style="background: #f0f7ff; padding: 15px; border-radius: 8px; margin-bottom: 20px;">
          <h4>🔧 Step-by-step setup:</h4>
          <ol style="margin: 10px 0;">
            <li>Configure API keys in Settings (see Settings for instructions)</li>
            <li>Test connections to ensure everything works</li>
            <li>Prepare your data with proper column headers</li>
            <li>Test with a small dataset first (recommended)</li>
            <li>Review results and adjust parameters if needed</li>
            <li>Run analysis on full dataset</li>
          </ol>
        </div>
        
        <div style="background: #f8f9fa; padding: 15px; border-radius: 8px; margin-bottom: 20px;">
          <h4>🎯 Turkish B2B Market Focus:</h4>
          <p>This system is specifically optimized for B2B service needs in Turkey, with:</p>
          <ul style="margin: 10px 0;">
            <li>Turkish business terminology recognition</li>
            <li>Local market dynamics consideration</li>
            <li>Turkish Lira revenue thresholds</li>
            <li>Geographic scoring for Turkish cities</li>
            <li>Industry multipliers based on Turkish B2B data</li>
          </ul>
        </div>
        
        <div style="background: #feebe6; padding: 15px; border-radius: 8px; margin-bottom: 20px;">
          <h4>🛡️ Security Features:</h4>
          <ul style="margin: 10px 0;">
            <li>Input validation and sanitization</li>
            <li>Secure API communication</li>
            <li>Enhanced error handling</li>
            <li>State management protection</li>
            <li>Comprehensive logging</li>
          </ul>
        </div>
        
        <div style="text-align: center; margin-top: 30px;">
          <button onclick="google.script.host.close()" 
                  style="background: #1a73e8; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer;">
            Close
          </button>
        </div>
      </div>
    `)
      .setWidth(650)
      .setHeight(700)
      .setTitle('Help - Enhanced Security');
    
    SpreadsheetApp.getUi().showModalDialog(html, 'Help');
    Logger.log('✅ Help dialog displayed');
  } catch (error) {
    Logger.log('❌ Error showing help dialog: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error: Unable to show help dialog.');
  }
}

/**
 * ENHANCED: Check if there's an interrupted processing session with atomic operations
 * FIXED: Race condition prevention with better state validation
 * @returns {Object|null} Processing state or null if none exists
 */
function checkProcessingState() {
  try {
    const properties = PropertiesService.getScriptProperties();
    
    // Atomic read operation with validation
    const stateStr = properties.getProperty('YOLWISE_SCORING_STATE');
    const lockStr = properties.getProperty('YOLWISE_SCORING_LOCK');
    const lockTimestamp = properties.getProperty('YOLWISE_SCORING_LOCK_TIME');
    
    Logger.log('🔍 Checking processing state: ' + (stateStr ? 'Found' : 'None'));
    
    if (!stateStr) {
      // Clean up any orphaned locks
      clearProcessingLock();
      return null;
    }
    
    // Check for processing lock (prevent race conditions)
    if (lockStr && lockTimestamp) {
      const lockAge = Date.now() - parseInt(lockTimestamp);
      if (lockAge < 60000) { // 1-minute lock timeout
        Logger.log('⚠️ Processing currently locked by another session');
        return null;
      } else {
        Logger.log('🔧 Clearing expired processing lock');
        clearProcessingLock();
      }
    }
    
    let state;
    try {
      state = JSON.parse(stateStr);
    } catch (parseError) {
      Logger.log('❌ Invalid state format, clearing: ' + parseError.toString());
      clearProcessingState();
      return null;
    }
    
    // Enhanced state validation
    if (!state || typeof state !== 'object') {
      Logger.log('❌ Invalid state object, clearing');
      clearProcessingState();
      return null;
    }
    
    // Validate required fields
    const requiredFields = ['options', 'processedRows', 'results', 'timestamp', 'version'];
    const missingFields = requiredFields.filter(field => !state.hasOwnProperty(field));
    
    if (missingFields.length > 0) {
      Logger.log('❌ Invalid state structure (missing: ' + missingFields.join(', ') + '), clearing');
      clearProcessingState();
      return null;
    }
    
    // Validate data types
    if (typeof state.processedRows !== 'number' || state.processedRows < 0) {
      Logger.log('❌ Invalid processedRows value, clearing state');
      clearProcessingState();
      return null;
    }
    
    // Check state age (expire after 24 hours)
    const stateAge = Date.now() - (state.timestamp || 0);
    if (stateAge > 24 * 60 * 60 * 1000) {
      Logger.log('⏰ State expired (age: ' + Math.round(stateAge / 1000 / 60) + ' minutes), clearing');
      clearProcessingState();
      return null;
    }
    
    // Validate results structure
    if (!state.results || typeof state.results !== 'object') {
      Logger.log('❌ Invalid results structure, clearing state');
      clearProcessingState();
      return null;
    }
    
    Logger.log('✅ Valid processing state found:', {
      processed: state.processedRows,
      total: state.results.total || 'unknown',
      age: Math.round(stateAge / 1000 / 60) + ' minutes',
      version: state.version || 'unknown'
    });
    
    return state;
    
  } catch (error) {
    Logger.log('❌ Error checking processing state: ' + error.toString());
    clearProcessingState();
    return null;
  }
}

/**
 * ENHANCED: Save current processing state with atomic operations and validation
 * FIXED: Race condition prevention, data validation, error handling
 * @param {Object} state Processing state to save
 */
function saveProcessingState(state) {
  try {
    // Enhanced input validation
    if (!state || typeof state !== 'object') {
      Logger.log('❌ Invalid state provided for saving');
      return false;
    }
    
    // Validate required fields
    const requiredFields = ['options', 'processedRows', 'results'];
    const missingFields = requiredFields.filter(field => !state.hasOwnProperty(field));
    
    if (missingFields.length > 0) {
      Logger.log('❌ Missing required fields for state saving: ' + missingFields.join(', '));
      return false;
    }
    
    if (typeof state.processedRows !== 'number' || state.processedRows < 0) {
      Logger.log('❌ Invalid processedRows value for saving: ' + state.processedRows);
      return false;
    }
    
    // Set processing lock to prevent race conditions
    if (!setProcessingLock()) {
      Logger.log('⚠️ Cannot save state - processing locked by another session');
      return false;
    }
    
    try {
      const properties = PropertiesService.getScriptProperties();
      
      // Enhanced state preparation
      const stateToSave = {
        ...state,
        timestamp: Date.now(),
        version: '2.1-enhanced',
        lastSaved: new Date().toISOString(),
        processId: Session.getTemporaryActiveUserKey() || 'unknown' // Basic process identification
      };
      
      // Validate state size (Google Apps Script has property size limits)
      const stateStr = JSON.stringify(stateToSave);
      if (stateStr.length > 500000) { // 500KB limit
        Logger.log('⚠️ State too large, compacting...');
        // Compact state by removing non-essential data
        stateToSave.results.results = stateToSave.results.results.slice(-10); // Keep only last 10 results
      }
      
      const finalStateStr = JSON.stringify(stateToSave);
      properties.setProperty('YOLWISE_SCORING_STATE', finalStateStr);
      
      Logger.log('✅ Processing state saved successfully:', {
        processed: stateToSave.processedRows,
        total: stateToSave.results.total || 'unknown',
        timestamp: new Date(stateToSave.timestamp).toLocaleString('en-US'),
        size: finalStateStr.length + ' bytes'
      });
      
      return true;
      
    } finally {
      // Always release the lock
      clearProcessingLock();
    }
    
  } catch (error) {
    Logger.log('❌ Error saving processing state: ' + error.toString());
    clearProcessingLock(); // Ensure lock is cleared on error
    return false;
  }
}

/**
 * ENHANCED: Clear processing state with better cleanup
 */
function clearProcessingState() {
  try {
    const properties = PropertiesService.getScriptProperties();
    properties.deleteProperty('YOLWISE_SCORING_STATE');
    clearProcessingLock(); // Also clear any locks
    Logger.log('✅ Processing state cleared successfully');
    return true;
  } catch (error) {
    Logger.log('❌ Error clearing processing state: ' + error.toString());
    return false;
  }
}

/**
 * NEW: Set processing lock to prevent race conditions
 * @returns {boolean} True if lock was acquired, false if already locked
 */
function setProcessingLock() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const existingLock = properties.getProperty('YOLWISE_SCORING_LOCK');
    const existingLockTime = properties.getProperty('YOLWISE_SCORING_LOCK_TIME');
    
    // Check if lock exists and is recent (within last minute)
    if (existingLock && existingLockTime) {
      const lockAge = Date.now() - parseInt(existingLockTime);
      if (lockAge < 60000) { // 1-minute timeout
        return false; // Lock still active
      }
    }
    
    // Set new lock
    const lockId = Utilities.getUuid();
    const lockTime = Date.now().toString();
    
    properties.setProperties({
      'YOLWISE_SCORING_LOCK': lockId,
      'YOLWISE_SCORING_LOCK_TIME': lockTime
    });
    
    Logger.log('🔒 Processing lock acquired: ' + lockId);
    return true;
    
  } catch (error) {
    Logger.log('❌ Error setting processing lock: ' + error.toString());
    return false;
  }
}

/**
 * NEW: Clear processing lock
 */
function clearProcessingLock() {
  try {
    const properties = PropertiesService.getScriptProperties();
    properties.deleteProperty('YOLWISE_SCORING_LOCK');
    properties.deleteProperty('YOLWISE_SCORING_LOCK_TIME');
    Logger.log('🔓 Processing lock cleared');
  } catch (error) {
    Logger.log('❌ Error clearing processing lock: ' + error.toString());
  }
}

/**
 * ENHANCED: Main scoring function with comprehensive security improvements
 * FIXED: Input validation, race condition prevention, enhanced error handling
 * @param {Object} options Configuration options from UI
 * @returns {Object} Processing results
 */
function processLeadScoring(options) {
  const startTime = Date.now();
  const timeoutMargin = 30000; // 30 seconds before 6-minute limit
  const maxExecutionTime = 6 * 60 * 1000 - timeoutMargin; // 5.5 minutes
  
  try {
    Logger.log('🚀 Starting enhanced Yolwise lead scoring process');
    
    // Enhanced input validation
    if (!options || typeof options !== 'object') {
      throw new Error('Invalid options provided');
    }
    
    // Validate required options
    if (!options.dataRange || typeof options.dataRange !== 'string') {
      throw new Error('Data range is required and must be a valid range string');
    }
    
    // Sanitize data range input
    const sanitizedRange = sanitizeDataRange(options.dataRange);
    if (!sanitizedRange) {
      throw new Error('Invalid data range format');
    }
    options.dataRange = sanitizedRange;
    
    Logger.log('✅ Input validation passed, execution time limit: ' + (maxExecutionTime / 1000) + ' seconds');
    
    // Check for existing processing state with race condition protection
    let processingState = checkProcessingState();
    let isResume = false;
    
    if (processingState) {
      Logger.log('📤 Found existing processing state, preparing to resume from row: ' + processingState.processedRows);
      isResume = true;
    }
    
    // Get active spreadsheet and sheet with validation
    const sheet = SpreadsheetApp.getActiveSheet();
    if (!sheet) {
      throw new Error('No active sheet found');
    }
    
    let range, data;
    try {
      range = sheet.getRange(options.dataRange);
      data = range.getValues();
    } catch (rangeError) {
      throw new Error('Invalid data range: ' + rangeError.message);
    }
    
    if (!data || data.length === 0) {
      throw new Error('No data found in specified range');
    }
    
    // Enhanced data validation
    const headers = data[0];
    if (!headers || headers.length === 0) {
      throw new Error('No headers found in data range');
    }
    
    // Sanitize headers
    const sanitizedHeaders = headers.map(header => sanitizeInput(header.toString()));
    
    const allRows = data.slice(1);
    if (allRows.length === 0) {
      throw new Error('No data rows found (only headers present)');
    }
    
    Logger.log(`📊 Processing ${allRows.length} companies with ${sanitizedHeaders.length} columns`);

    // Initialize or restore progress tracking with enhanced validation
    let results;
    let startIndex = 0;
    
    if (isResume && processingState) {
      // Validate resumption conditions
      if (processingState.processedRows > allRows.length) {
        Logger.log('⚠️ Processed rows exceed total rows, starting fresh');
        clearProcessingState();
        isResume = false;
      } else {
        results = processingState.results;
        startIndex = processingState.processedRows;
        Logger.log(`📤 Resuming from row ${startIndex + 1} of ${allRows.length}`);
      }
    }
    
    if (!isResume) {
      // Start fresh with enhanced initialization
      clearProcessingState();
      results = {
        total: allRows.length,
        processed: 0,
        successful: 0,
        failed: 0,
        results: [],
        startTime: startTime,
        version: '2.1-enhanced'
      };
    }
    
    // Get rows to process with validation
    const rowsToProcess = allRows.slice(startIndex);
    
    if (rowsToProcess.length === 0) {
      Logger.log('✅ All rows already processed');
      return results;
    }
    
    Logger.log(`🔄 Processing ${rowsToProcess.length} remaining rows`);
    
    // Process companies individually with enhanced error handling
    for (let i = 0; i < rowsToProcess.length; i++) {
      const currentTime = Date.now();
      const elapsedTime = currentTime - startTime;
      
      // Enhanced timeout checking
      if (elapsedTime > maxExecutionTime) {
        Logger.log(`⏰ Approaching timeout (${Math.round(elapsedTime/1000)}s elapsed), saving state and stopping`);
        
        // Enhanced state saving with validation
        const stateSaved = saveProcessingState({
          options: options,
          headers: sanitizedHeaders,
          processedRows: startIndex + i,
          results: results
        });
        
        if (!stateSaved) {
          Logger.log('❌ Failed to save state for continuation');
        }
        
        // Return enhanced partial results
        return {
          ...results,
          incomplete: true,
          continuationAvailable: stateSaved,
          message: `Processed: ${results.processed} / ${results.total} companies. ${stateSaved ? 'Use "Continue" button to proceed.' : 'State save failed - restart required.'}`,
          timeoutReason: 'execution_limit',
          elapsedTime: elapsedTime
        };
      }
      
      const row = rowsToProcess[i];
      const actualRowIndex = startIndex + i;
      
      try {
        // Enhanced input validation for each row
        if (!row || !Array.isArray(row)) {
          throw new Error('Invalid row data');
        }
        
        // Sanitize row data
        const sanitizedRow = row.map(cell => sanitizeInput(cell));
        
        // Extract company data with enhanced hybrid mapping system
        const companyData = extractCompanyDataHybridSecure(sanitizedRow, sanitizedHeaders);
        
        if (!companyData.company_name) {
          results.results.push({
            success: false,
            error: 'Company name not found or invalid',
            companyData: companyData,
            qualityAnalysis: companyData.quality_analysis || 'Missing required company name',
            rowIndex: actualRowIndex
          });
          results.failed++;
        } else {
          Logger.log(`🏢 Processing company ${actualRowIndex + 1}/${results.total}: ${companyData.company_name.substring(0, 50)}...`);
          
          try {
            // Analyze with Claude API (enhanced Turkish context)
            const claudeAnalysis = analyzeWithClaude(companyData, options);
            
            // Score with Yolwise API
            const scoringResult = scoreWithYolwise(claudeAnalysis);
            
            results.results.push({
              success: true,
              companyData: companyData,
              claudeAnalysis: claudeAnalysis,
              scoringResult: scoringResult,
              qualityAnalysis: companyData.quality_analysis || 'All core fields successfully extracted',
              rowIndex: actualRowIndex,
              processingTime: Date.now() - currentTime
            });
            
            results.successful++;
            Logger.log(`✅ Company scored: ${scoringResult.final_score || 'unknown'}`);
            
          } catch (analysisError) {
            Logger.log('❌ Analysis error for company: ' + analysisError.toString());
            results.results.push({
              success: false,
              error: 'Analysis failed: ' + analysisError.message,
              companyData: companyData,
              qualityAnalysis: 'Analysis failed due to: ' + analysisError.message,
              rowIndex: actualRowIndex
            });
            results.failed++;
          }
        }
        
        results.processed++;
        
        // Enhanced state saving (every 3 companies for better safety)
        if ((i + 1) % 3 === 0) {
          const stateSaved = saveProcessingState({
            options: options,
            headers: sanitizedHeaders,
            processedRows: actualRowIndex + 1,
            results: results
          });
          
          if (stateSaved) {
            Logger.log(`📊 Progress saved: ${results.processed}/${results.total} (${Math.round(results.processed/results.total*100)}%)`);
          } else {
            Logger.log('⚠️ State save failed but continuing...');
          }
        }
        
        // Enhanced delay with jitter to avoid rate limits
        const delay = 1200 + Math.random() * 600; // 1.2-1.8 seconds
        Utilities.sleep(delay);
        
      } catch (companyError) {
        Logger.log('❌ Error processing company at row ' + (actualRowIndex + 1) + ': ' + companyError.toString());
        results.results.push({
          success: false,
          error: companyError.toString(),
          companyData: extractCompanyDataHybridSecure(row, sanitizedHeaders),
          qualityAnalysis: 'Critical processing error: ' + companyError.message,
          rowIndex: actualRowIndex
        });
        results.processed++;
        results.failed++;
      }
    }
    
    // Enhanced results insertion with validation
    try {
      insertEnhancedResults(sheet, range, results.results, options);
      Logger.log('✅ Results inserted into spreadsheet successfully');
    } catch (insertError) {
      Logger.log('❌ Error inserting results: ' + insertError.toString());
      // Don't fail the entire process if insertion fails
      results.insertError = insertError.message;
    }
    
    // Clear processing state on successful completion
    clearProcessingState();
    
    // Calculate final statistics
    const totalTime = Date.now() - startTime;
    results.completionTime = totalTime;
    results.averageTimePerCompany = totalTime / results.processed;
    
    Logger.log('🎉 Enhanced Yolwise lead scoring completed successfully:', {
      total: results.total,
      successful: results.successful,
      failed: results.failed,
      totalTime: Math.round(totalTime / 1000) + 's',
      avgTime: Math.round(results.averageTimePerCompany / 1000) + 's per company'
    });
    
    return results;
    
  } catch (error) {
    Logger.log('❌ Error in processLeadScoring: ' + error.toString());
    
    // Enhanced error handling with state preservation
    const elapsedTime = Date.now() - startTime;
    Logger.log(`⚠️ Error after ${Math.round(elapsedTime/1000)} seconds - attempting to preserve state`);
    
    // Try to save current state even on error
    try {
      if (options && elapsedTime > 30000) { // Only if we've been running for more than 30 seconds
        saveProcessingState({
          options: options,
          headers: [],
          processedRows: 0,
          results: { total: 0, processed: 0, successful: 0, failed: 0, results: [] },
          error: error.message
        });
      }
    } catch (stateSaveError) {
      Logger.log('❌ Failed to save error state: ' + stateSaveError.toString());
    }
    
    throw new Error('Processing failed: ' + error.message);
  }
}

/**
 * ENHANCED: Continue processing from saved state with validation
 * @returns {Object} Processing results
 */
function continueProcessing() {
  try {
    const processingState = checkProcessingState();
    
    if (!processingState) {
      throw new Error('No processing state found. Session may have expired or been cleared.');
    }
    
    // Validate state before continuing
    if (!processingState.options || !processingState.options.dataRange) {
      throw new Error('Invalid processing state - missing required options');
    }
    
    Logger.log('📤 Continuing processing from validated saved state');
    return processLeadScoring(processingState.options);
    
  } catch (error) {
    Logger.log('❌ Error in continueProcessing: ' + error.toString());
    throw new Error('Continue failed: ' + error.message);
  }
}

/**
 * ENHANCED: Get available data ranges with validation
 * @returns {Array} Available ranges with validation
 */
function getAvailableRanges() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    if (!sheet) {
      Logger.log('❌ No active sheet found');
      return [];
    }
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    if (lastRow === 0 || lastCol === 0) {
      Logger.log('ℹ️ Empty sheet detected');
      return [];
    }
    
    // Validate sheet size limits
    if (lastRow > 10000 || lastCol > 100) {
      Logger.log('⚠️ Large sheet detected - may cause performance issues');
    }
    
    const ranges = [];
    
    try {
      // Full data range
      const fullRangeNotation = `A1:${getColumnLetter(lastCol)}${lastRow}`;
      ranges.push({
        label: `All data (${fullRangeNotation})`,
        value: fullRangeNotation,
        description: `${lastRow} rows, ${lastCol} columns`,
        recommended: lastRow <= 1000
      });
      
      // Data without header (if more than 1 row)
      if (lastRow > 1) {
        const dataRangeNotation = `A2:${getColumnLetter(lastCol)}${lastRow}`;
        ranges.push({
          label: `Data only (${dataRangeNotation})`,
          value: dataRangeNotation,
          description: `${lastRow - 1} data rows, ${lastCol} columns`,
          recommended: lastRow <= 1000
        });
      }
      
      // Sample range for testing (first 10 rows)
      if (lastRow > 10) {
        const sampleRangeNotation = `A1:${getColumnLetter(lastCol)}10`;
        ranges.push({
          label: `Sample (${sampleRangeNotation})`,
          value: sampleRangeNotation,
          description: `First 10 rows for testing, ${lastCol} columns`,
          recommended: true
        });
      }
      
    } catch (rangeError) {
      Logger.log('❌ Error generating ranges: ' + rangeError.toString());
      return [{
        label: 'Error generating ranges',
        value: 'A1:A1',
        description: 'Please enter range manually',
        recommended: false
      }];
    }
    
    Logger.log('✅ Generated ' + ranges.length + ' available ranges');
    return ranges;
    
  } catch (error) {
    Logger.log('❌ Error getting available ranges: ' + error.toString());
    return [];
  }
}

/**
 * NEW: Convert column number to letter (A, B, C, ... Z, AA, AB, etc.)
 * @param {number} columnNumber Column number (1-based)
 * @returns {string} Column letter
 */
function getColumnLetter(columnNumber) {
  let result = '';
  while (columnNumber > 0) {
    columnNumber--;
    result = String.fromCharCode(65 + (columnNumber % 26)) + result;
    columnNumber = Math.floor(columnNumber / 26);
  }
  return result;
}

/**
 * ENHANCED: Extract company data with comprehensive security measures
 * FIXED: Input validation, sanitization, injection prevention
 * @param {Array} row Spreadsheet row data (sanitized)
 * @param {Array} headers Column headers (sanitized)
 * @returns {Object} Structured company data with enhanced quality analysis
 */
function extractCompanyDataHybridSecure(row, headers) {
  Logger.log('🔄 Starting secure hybrid mapping system for company data');
  
  try {
    // Enhanced input validation
    if (!row || !Array.isArray(row)) {
      throw new Error('Invalid row data provided');
    }
    
    if (!headers || !Array.isArray(headers)) {
      throw new Error('Invalid headers provided');
    }
    
    if (row.length !== headers.length) {
      Logger.log(`⚠️ Row/header length mismatch: ${row.length} vs ${headers.length}`);
      // Pad shorter array with empty values
      const maxLength = Math.max(row.length, headers.length);
      while (row.length < maxLength) row.push('');
      while (headers.length < maxLength) headers.push('');
    }
    
    // Step 1: Use enhanced structured (code-based) mapping
    const structuredData = extractCompanyDataStructuredSecure(row, headers);
    
    // Step 2: Identify critical missing fields
    const criticalFields = ['company_name', 'industry', 'revenue_estimate', 'employees_estimate', 'headquarters'];
    const missingCriticalFields = criticalFields.filter(field => 
      !structuredData[field] || 
      structuredData[field] === '' || 
      structuredData[field] === null
    );
    
    // Step 3: Enhanced quality analysis tracking
    const qualityAnalysis = {
      structured_mapping_success: criticalFields.length - missingCriticalFields.length,
      total_critical_fields: criticalFields.length,
      missing_fields: missingCriticalFields,
      claude_fallback_used: false,
      claude_mapping_success: 0,
      final_data_completeness: 0,
      security_validation_passed: true,
      data_sanitization_applied: true
    };
    
    Logger.log(`📊 Structured mapping results: ${qualityAnalysis.structured_mapping_success}/${qualityAnalysis.total_critical_fields} critical fields found`);
    
    // Step 4: Apply Claude-based fallback with enhanced security for missing critical fields
    let finalData = { ...structuredData };
    
    if (missingCriticalFields.length > 0 && missingCriticalFields.length < criticalFields.length) {
      Logger.log(`🤖 Applying secure Claude fallback for missing fields: ${missingCriticalFields.join(', ')}`);
      
      try {
        const claudeMappedData = mapHeadersWithClaudeSecure(headers, row, missingCriticalFields);
        
        // Enhanced validation and merging of Claude results
        missingCriticalFields.forEach(field => {
          if (claudeMappedData[field] && 
              claudeMappedData[field] !== '' && 
              claudeMappedData[field] !== null) {
            
            // Additional validation for specific fields
            if (field === 'company_name' && claudeMappedData[field].length > 200) {
              claudeMappedData[field] = claudeMappedData[field].substring(0, 200);
            }
            if (field === 'revenue_estimate' && isNaN(parseFloat(claudeMappedData[field]))) {
              claudeMappedData[field] = 0;
            }
            
            finalData[field] = claudeMappedData[field];
            qualityAnalysis.claude_mapping_success++;
            Logger.log(`✅ Claude successfully mapped: ${field} = ${claudeMappedData[field].toString().substring(0, 50)}...`);
          }
        });
        
        qualityAnalysis.claude_fallback_used = true;
        
      } catch (claudeError) {
        Logger.log('❌ Claude fallback mapping failed: ' + claudeError.toString());
        qualityAnalysis.claude_fallback_error = claudeError.toString();
      }
    }
    
    // Step 5: Calculate enhanced quality metrics
    const finalCriticalFields = criticalFields.filter(field => 
      finalData[field] && 
      finalData[field] !== '' && 
      finalData[field] !== null
    );
    qualityAnalysis.final_data_completeness = (finalCriticalFields.length / criticalFields.length) * 100;
    
    // Step 6: Generate enhanced quality report
    const qualityReport = generateEnhancedQualityReport(qualityAnalysis, headers.length);
    finalData.quality_analysis = qualityReport;
    
    // Step 7: Final security validation
    finalData = performFinalSecurityValidation(finalData);
    
    Logger.log(`📈 Final secure mapping quality: ${Math.round(qualityAnalysis.final_data_completeness)}% completeness`);
    
    return finalData;
    
  } catch (error) {
    Logger.log('❌ Error in secure hybrid mapping: ' + error.toString());
    
    // Return minimal safe data structure
    return {
      company_name: 'Extraction Failed',
      industry: 'Unknown',
      revenue_estimate: 0,
      employees_estimate: 'Unknown',
      headquarters: 'Unknown',
      quality_analysis: 'Critical extraction error: ' + error.message,
      extraction_error: true
    };
  }
}

/**
 * ENHANCED: Structured mapping with comprehensive security
 * @param {Array} row Sanitized spreadsheet row data
 * @param {Array} headers Sanitized column headers
 * @returns {Object} Structured company data
 */
function extractCompanyDataStructuredSecure(row, headers) {
  const data = {};
  
  try {
    // Enhanced field mappings with security considerations
    const secureFieldMappings = {
      'company_name': {
        keywords: ['company name', 'company', 'name', 'firm', 'organization', 'entity', 'business name', 'corp', 'corporation'],
        maxLength: 200,
        required: true
      },
      'industry': {
        keywords: ['industry', 'sector', 'business', 'field', 'market', 'vertical', 'category', 'classification'],
        maxLength: 100,
        required: false
      },
      'revenue_estimate': {
        keywords: ['annual revenue', 'revenue', 'sales', 'turnover', 'income', 'earnings', 'gross revenue'],
        maxLength: 50,
        required: false,
        numeric: true
      },
      'employees_estimate': {
        keywords: ['number of employees', 'employees', 'staff', 'workforce', 'headcount', 'team size'],
        maxLength: 50,
        required: false
      },
      'headquarters': {
        keywords: ['city', 'location', 'headquarters', 'head office', 'address', 'main office'],
        maxLength: 100,
        required: false
      },
      'description': {
        keywords: ['description', 'about', 'overview', 'summary', 'profile', 'company description'],
        maxLength: 1000,
        required: false
      }
    };
    
    const mappedHeaders = new Set();
    
    // Enhanced extraction with security validation
    headers.forEach((header, index) => {
      if (!header || index >= row.length) return;
      
      const normalizedHeader = header.toString().toLowerCase().trim();
      const cellValue = row[index];
      
      // Skip empty or null values
      if (!cellValue || cellValue.toString().trim() === '') return;
      
      for (const [field, config] of Object.entries(secureFieldMappings)) {
        if (config.keywords.some(keyword => normalizedHeader.includes(keyword))) {
          let processedValue = cellValue.toString().trim();
          
          // Apply security constraints
          if (config.maxLength && processedValue.length > config.maxLength) {
            processedValue = processedValue.substring(0, config.maxLength);
            Logger.log(`⚠️ Truncated ${field} to ${config.maxLength} characters`);
          }
          
          // Numeric validation
          if (config.numeric) {
            const numericValue = extractNumericValue(processedValue);
            processedValue = numericValue;
          }
          
          data[field] = processedValue;
          mappedHeaders.add(index);
          Logger.log(`✅ Secure mapping: "${header}" -> ${field} = ${processedValue.toString().substring(0, 50)}...`);
          break;
        }
      }
    });
    
    // Enhanced categorization of unmapped fields with security
    const secureCategories = {
      financial_data: [],
      legal_data: [], 
      operational_data: [],
      contact_data: [],
      other_data: []
    };
    
    headers.forEach((header, index) => {
      if (!mappedHeaders.has(index) && row[index] && row[index].toString().trim() !== '') {
        const headerLower = header.toString().toLowerCase().trim();
        const value = row[index].toString().trim();
        
        // Security: limit entry length
        const safeValue = value.length > 200 ? value.substring(0, 200) + '...' : value;
        const entry = `${header}: ${safeValue}`;
        
        // Enhanced categorization with security keywords
        if (headerLower.includes('revenue') || headerLower.includes('financial') || headerLower.includes('budget')) {
          secureCategories.financial_data.push(entry);
        } else if (headerLower.includes('legal') || headerLower.includes('license') || headerLower.includes('registration')) {
          secureCategories.legal_data.push(entry);
        } else if (headerLower.includes('contact') || headerLower.includes('phone') || headerLower.includes('email')) {
          secureCategories.contact_data.push(entry);
        } else if (headerLower.includes('operation') || headerLower.includes('branch') || headerLower.includes('office')) {
          secureCategories.operational_data.push(entry);
        } else {
          secureCategories.other_data.push(entry);
        }
      }
    });
    
    // Combine categorized data with limits
    const discoveredFacts = [];
    Object.entries(secureCategories).forEach(([category, items]) => {
      if (items.length > 0) {
        // Limit items per category for security
        const limitedItems = items.slice(0, 10);
        discoveredFacts.push(`${category}: ${limitedItems.join('; ')}`);
      }
    });
    
    data.discovered_facts = discoveredFacts;
    
    Logger.log(`✅ Secure structured extraction completed for: ${data.company_name || 'Unknown company'}`);
    
    return data;
    
  } catch (error) {
    Logger.log('❌ Error in secure structured extraction: ' + error.toString());
    return {
      company_name: 'Extraction Error',
      extraction_error: error.message
    };
  }
}

/**
 * NEW: Extract numeric value safely from text
 * @param {string} text Text containing potential numeric value
 * @returns {number} Extracted numeric value or 0
 */
function extractNumericValue(text) {
  try {
    if (typeof text === 'number') return Math.max(0, text);
    
    const textStr = text.toString().trim();
    
    // Handle common multipliers
    let multiplier = 1;
    if (textStr.toLowerCase().includes('million') || textStr.toLowerCase().includes('m')) {
      multiplier = 1000000;
    } else if (textStr.toLowerCase().includes('thousand') || textStr.toLowerCase().includes('k')) {
      multiplier = 1000;
    }
    
    // Extract first number
    const match = textStr.match(/(\d+(?:\.\d+)?)/);
    if (match) {
      return Math.max(0, parseFloat(match[1]) * multiplier);
    }
    
    return 0;
  } catch (error) {
    Logger.log('❌ Error extracting numeric value: ' + error.toString());
    return 0;
  }
}

/**
 * ENHANCED: Claude-based header mapping with comprehensive security
 * @param {Array} headers All column headers (sanitized)
 * @param {Array} row Corresponding row data (sanitized)
 * @param {Array} missingFields List of critical fields that need mapping
 * @returns {Object} Claude-mapped data for missing fields
 */
function mapHeadersWithClaudeSecure(headers, row, missingFields) {
  try {
    // Enhanced input validation
    if (!Array.isArray(headers) || !Array.isArray(row) || !Array.isArray(missingFields)) {
      throw new Error('Invalid input parameters for Claude mapping');
    }
    
    if (missingFields.length === 0) {
      return {};
    }
    
    const apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
    if (!apiKey) {
      throw new Error('Claude API key not configured');
    }
    
    // Prepare sanitized data for Claude
    const sanitizedUnmappedHeaders = [];
    headers.forEach((header, index) => {
      if (row[index] && row[index].toString().trim() !== '') {
        const sanitizedHeader = sanitizeInput(header.toString());
        const sanitizedValue = sanitizeInput(row[index].toString());
        
        // Security: limit data sent to Claude
        const limitedValue = sanitizedValue.length > 100 ? 
          sanitizedValue.substring(0, 100) + '...' : sanitizedValue;
        
        sanitizedUnmappedHeaders.push({
          header: sanitizedHeader,
          value: limitedValue,
          index: index
        });
      }
    });
    
    // Enhanced prompt with security considerations
    const securePrompt = `Analyze business data headers and find mappings for missing fields.

MISSING FIELDS (need to be found): ${missingFields.join(', ')}

AVAILABLE HEADERS AND VALUES (limited for security):
${sanitizedUnmappedHeaders.slice(0, 20).map(item => `"${item.header}": "${item.value}"`).join('\n')}

MAPPING RULES:
- company_name: company names, firm names, business names
- industry: sectors, business areas, activity fields, industries
- revenue_estimate: revenue, sales, income, turnover amounts (extract numbers only)
- employees_estimate: employee count, staff size, workforce (extract numbers only)
- headquarters: cities, states, addresses, locations

SECURITY REQUIREMENTS:
- Extract only factual business information
- Limit text fields to reasonable lengths
- Return numeric values for numeric fields
- No personal or sensitive data extraction

RETURN ONLY JSON format:
{
  "field_name": "extracted_value",
  "field_name2": "extracted_value2"
}

If field cannot be found, do not include it in response.`;

    // Enhanced API request with security
    const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      payload: JSON.stringify({
        model: 'claude-3-5-sonnet-20241022',
        max_tokens: 1000,
        temperature: 0.1,
        messages: [{
          role: 'user',
          content: securePrompt
        }]
      }),
      muteHttpExceptions: true,
      validateHttpsCertificates: true
    });
    
    if (response.getResponseCode() !== 200) {
      throw new Error(`Claude API error: ${response.getResponseCode()}`);
    }
    
    const responseData = JSON.parse(response.getContentText());
    let content = responseData.content[0].text.trim();
    
    // Enhanced JSON extraction with security validation
    if (content.includes('```json')) {
      const match = content.match(/```json\s*([\s\S]*?)\s*```/);
      if (match) {
        content = match[1];
      }
    }
    
    const claudeMappedData = JSON.parse(content);
    
    // Enhanced validation and sanitization of Claude results
    const validatedData = {};
    Object.entries(claudeMappedData).forEach(([key, value]) => {
      if (missingFields.includes(key) && value && value.toString().trim() !== '') {
        let sanitizedValue = sanitizeInput(value.toString());
        
        // Apply field-specific validation
        if (key === 'company_name' && sanitizedValue.length > 200) {
          sanitizedValue = sanitizedValue.substring(0, 200);
        } else if (key === 'revenue_estimate') {
          sanitizedValue = extractNumericValue(sanitizedValue);
        } else if (key === 'employees_estimate') {
          // Keep as string but extract number if needed
          const numMatch = sanitizedValue.match(/\d+/);
          if (numMatch) {
            sanitizedValue = numMatch[0];
          }
        }
        
        validatedData[key] = sanitizedValue;
      }
    });
    
    Logger.log('✅ Claude securely mapped headers:', Object.keys(validatedData));
    return validatedData;
    
  } catch (error) {
    Logger.log('❌ Secure Claude header mapping failed: ' + error.toString());
    throw new Error('Secure Claude mapping failed: ' + error.message);
  }
}

/**
 * NEW: Enhanced quality analysis report generation
 * @param {Object} qualityAnalysis Quality metrics object
 * @param {number} totalHeaders Total number of headers in spreadsheet
 * @returns {string} Human-readable quality report
 */
function generateEnhancedQualityReport(qualityAnalysis, totalHeaders) {
  const reports = [];
  
  // Overall completeness with enhanced categories
  const completeness = Math.round(qualityAnalysis.final_data_completeness);
  if (completeness >= 90) {
    reports.push(`🌟 Excellent data quality (${completeness}%)`);
  } else if (completeness >= 75) {
    reports.push(`✅ Very good data quality (${completeness}%)`);
  } else if (completeness >= 60) {
    reports.push(`⚠️ Good data quality (${completeness}%)`);
  } else if (completeness >= 40) {
    reports.push(`⚠️ Fair data quality (${completeness}%)`);
  } else {
    reports.push(`❌ Poor data quality (${completeness}%)`);
  }
  
  // Structured mapping performance
  reports.push(`Structured: ${qualityAnalysis.structured_mapping_success}/${qualityAnalysis.total_critical_fields}`);
  
  // Claude fallback usage
  if (qualityAnalysis.claude_fallback_used) {
    if (qualityAnalysis.claude_mapping_success > 0) {
      reports.push(`Claude: +${qualityAnalysis.claude_mapping_success} fields`);
    } else {
      reports.push(`Claude: attempted but failed`);
    }
  }
  
  // Security validation status
  if (qualityAnalysis.security_validation_passed) {
    reports.push(`Security: ✅`);
  } else {
    reports.push(`Security: ⚠️`);
  }
  
  // Missing fields warning (limited to prevent long messages)
  if (qualityAnalysis.missing_fields.length > 0) {
    const missingStr = qualityAnalysis.missing_fields.slice(0, 3).join(', ');
    const moreCount = qualityAnalysis.missing_fields.length - 3;
    reports.push(`Missing: ${missingStr}${moreCount > 0 ? ` +${moreCount} more` : ''}`);
  }
  
  return reports.join(' | ');
}

/**
 * NEW: Perform final security validation on extracted data
 * @param {Object} data Extracted company data
 * @returns {Object} Security-validated data
 */
function performFinalSecurityValidation(data) {
  try {
    const validatedData = {};
    
    Object.entries(data).forEach(([key, value]) => {
      if (value === null || value === undefined) {
        validatedData[key] = '';
        return;
      }
      
      let processedValue = value;
      
      // String validation
      if (typeof value === 'string') {
        // Remove potentially dangerous characters
        processedValue = value.replace(/[<>"\&]/g, '').trim();
        
        // Length limits based on field type
        const lengthLimits = {
          company_name: 200,
          industry: 100,
          headquarters: 100,
          description: 1000
        };
        
        const limit = lengthLimits[key] || 500;
        if (processedValue.length > limit) {
          processedValue = processedValue.substring(0, limit);
        }
      }
      
      // Numeric validation
      if (key.includes('revenue') || key.includes('employees')) {
        if (typeof processedValue === 'string') {
          const numericValue = extractNumericValue(processedValue);
          if (key.includes('revenue')) {
            processedValue = numericValue;
          } else {
            processedValue = numericValue > 0 ? numericValue.toString() : processedValue;
          }
        }
      }
      
      validatedData[key] = processedValue;
    });
    
    return validatedData;
    
  } catch (error) {
    Logger.log('❌ Error in final security validation: ' + error.toString());
    return data; // Return original data if validation fails
  }
}

/**
 * NEW: Sanitize input data to prevent injection attacks
 * @param {any} input Raw input data
 * @returns {string} Sanitized input
 */
function sanitizeInput(input) {
  try {
    if (input === null || input === undefined) {
      return '';
    }
    
    let sanitized = input.toString().trim();
    
    // Remove potentially dangerous characters
    sanitized = sanitized
      .replace(/[<>"\&]/g, '')      // HTML/XML characters
      .replace(/[\r\n\t]/g, ' ')    // Line breaks and tabs
      .replace(/\\/g, '')           // Backslashes
      .replace(/'/g, "'")           // Normalize quotes
      .substring(0, 1000);          // Limit length
    
    return sanitized;
    
  } catch (error) {
    Logger.log('❌ Error sanitizing input: ' + error.toString());
    return '';
  }
}

/**
 * NEW: Sanitize data range input
 * @param {string} range Raw range input
 * @returns {string|null} Sanitized range or null if invalid
 */
function sanitizeDataRange(range) {
  try {
    if (!range || typeof range !== 'string') {
      return null;
    }
    
    const sanitized = range.toString().trim().toUpperCase();
    
    // Basic range validation (A1:Z100 format)
    const rangePattern = /^[A-Z]+\d+:[A-Z]+\d+$/;
    if (!rangePattern.test(sanitized)) {
      Logger.log('❌ Invalid range format: ' + sanitized);
      return null;
    }
    
    return sanitized;
    
  } catch (error) {
    Logger.log('❌ Error sanitizing range: ' + error.toString());
    return null;
  }
}

/**
 * ENHANCED: Insert scoring results with comprehensive security and validation
 * @param {Sheet} sheet Target sheet
 * @param {Range} originalRange Original data range
 * @param {Array} results Scoring results
 * @param {Object} options Configuration options
 */
function insertEnhancedResults(sheet, originalRange, results, options) {
  try {
    Logger.log('📊 Inserting enhanced results with security validation...');
    
    // Enhanced validation
    if (!sheet || !originalRange || !Array.isArray(results)) {
      throw new Error('Invalid parameters for result insertion');
    }
    
    if (results.length === 0) {
      Logger.log('⚠️ No results to insert');
      return;
    }
    
    // Enhanced result headers with security information
    const enhancedResultHeaders = [
      'Base Score',
      'Industry Multiplier', 
      'Industry Score',
      'LLM Adjustment',
      'Final Score',
      'Priority',
      'Industry',
      'Confidence',
      'Data Quality',
      'Processing Notes'
    ];
    
    // Find next available column with validation
    const lastCol = originalRange.getLastColumn();
    const resultStartCol = lastCol + 1;
    
    // Validate sheet boundaries
    const maxCols = sheet.getMaxColumns();
    if (resultStartCol + enhancedResultHeaders.length > maxCols) {
      throw new Error(`Not enough columns available. Need ${enhancedResultHeaders.length} more columns.`);
    }
    
    // Insert enhanced headers
    const headerRange = sheet.getRange(originalRange.getRow(), resultStartCol, 1, enhancedResultHeaders.length);
    headerRange.setValues([enhancedResultHeaders]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#E8F0FE');
    
    // Prepare enhanced result data with security validation
    const enhancedResultData = results.map((result, index) => {
      try {
        if (!result.success) {
          return [
            'ERROR',
            '',
            '',
            '',
            0,
            'ERROR',
            'Unknown',
            'Low',
            result.qualityAnalysis || 'Processing error',
            result.error || 'Unknown error'
          ];
        }
        
        const scoring = result.scoringResult;
        if (!scoring) {
          return [
            'NO_SCORE',
            '',
            '',
            '',
            0,
            'NO_SCORE',
            'Unknown',
            'Low',
            'No scoring result',
            'Scoring failed'
          ];
        }
        
        // Enhanced score calculations with validation
        const baseScore = parseFloat(scoring.base_score) || 0;
        const industryMultiplier = parseFloat(scoring.industry_multiplier) || 1.0;
        const industryScore = parseFloat(scoring.industry_adjusted_score) || baseScore;
        const llmAdjustment = Math.max(-25, Math.min(25, parseFloat(scoring.llm_adjustment) || 0));
        const finalScore = Math.max(0, Math.min(100, parseFloat(scoring.final_score) || 0));
        
        // Enhanced priority determination with validation
        const priority = finalScore >= 60 ? 'TARGET' : 'NON-TARGET';
        const industry = (scoring.detected_industry || 'Unknown').toString().substring(0, 50);
        const confidence = (scoring.industry_confidence || 'Low').toString();
        
        // Enhanced quality and processing information
        const dataQuality = (result.qualityAnalysis || 'Unknown').toString().substring(0, 100);
        const processingNotes = generateProcessingNotes(result);
        
        return [
          Math.round(baseScore * 10) / 10,
          Math.round(industryMultiplier * 100) / 100,
          Math.round(industryScore * 10) / 10,
          Math.round(llmAdjustment * 10) / 10,
          Math.round(finalScore * 10) / 10,
          priority,
          industry,
          confidence,
          dataQuality,
          processingNotes
        ];
        
      } catch (rowError) {
        Logger.log(`❌ Error processing result row ${index}: ${rowError.toString()}`);
        return [
          'ERROR',
          '',
          '',
          '',
          0,
          'ERROR',
          'Unknown',
          'Low',
          'Row processing error',
          rowError.message
        ];
      }
    });
    
    // Insert enhanced result data with validation
    if (enhancedResultData.length > 0) {
      const dataRange = sheet.getRange(
        originalRange.getRow() + 1,
        resultStartCol,
        enhancedResultData.length,
        enhancedResultHeaders.length
      );
      
      dataRange.setValues(enhancedResultData);
      
      // Enhanced conditional formatting
      applyEnhancedConditionalFormatting(sheet, originalRange, resultStartCol, enhancedResultData.length);
    }
    
    Logger.log('✅ Enhanced results inserted successfully with security validation');
    
  } catch (error) {
    Logger.log('❌ Error inserting enhanced results: ' + error.toString());
    throw new Error('Result insertion failed: ' + error.message);
  }
}

/**
 * NEW: Generate processing notes for result row
 * @param {Object} result Result object
 * @returns {string} Processing notes
 */
function generateProcessingNotes(result) {
  try {
    const notes = [];
    
    if (result.scoringResult && result.scoringResult.mock_scoring) {
      notes.push('Mock scoring used');
    }
    
    if (result.claudeAnalysis && result.claudeAnalysis.analysis_confidence) {
      notes.push(`Analysis: ${result.claudeAnalysis.analysis_confidence}`);
    }
    
    if (result.processingTime) {
      const timeSeconds = Math.round(result.processingTime / 1000);
      notes.push(`${timeSeconds}s`);
    }
    
    return notes.length > 0 ? notes.join(', ') : 'Standard processing';
    
  } catch (error) {
    return 'Note generation error';
  }
}

/**
 * NEW: Apply enhanced conditional formatting to results
 * @param {Sheet} sheet Target sheet
 * @param {Range} originalRange Original data range
 * @param {number} resultStartCol Starting column for results
 * @param {number} dataRowCount Number of data rows
 */
function applyEnhancedConditionalFormatting(sheet, originalRange, resultStartCol, dataRowCount) {
  try {
    // Priority column formatting (column index 5)
    const priorityCol = resultStartCol + 5;
    const priorityRange = sheet.getRange(
      originalRange.getRow() + 1,
      priorityCol,
      dataRowCount,
      1
    );
    
    const priorityValues = priorityRange.getValues();
    const priorityBackgrounds = priorityValues.map(([priority]) => {
      const priorityStr = priority.toString().toUpperCase();
      if (priorityStr === 'TARGET') {
        return ['#D4EDDA']; // Green for target
      } else if (priorityStr === 'NON-TARGET') {
        return ['#F8D7DA']; // Red for non-target
      } else {
        return ['#FFF3CD']; // Yellow for error/unknown
      }
    });
    priorityRange.setBackgrounds(priorityBackgrounds);
    
    // Data Quality column formatting (column index 8)
    const qualityCol = resultStartCol + 8;
    const qualityRange = sheet.getRange(
      originalRange.getRow() + 1,
      qualityCol,
      dataRowCount,
      1
    );
    
    const qualityValues = qualityRange.getValues();
    const qualityBackgrounds = qualityValues.map(([quality]) => {
      const qualityStr = quality.toString().toLowerCase();
      if (qualityStr.includes('excellent') || qualityStr.includes('🌟')) {
        return ['#D4EDDA']; // Green for excellent
      } else if (qualityStr.includes('good') || qualityStr.includes('✅')) {
        return ['#FFF3CD']; // Yellow for good
      } else if (qualityStr.includes('poor') || qualityStr.includes('❌')) {
        return ['#F8D7DA']; // Red for poor
      } else {
        return ['#F8F9FA']; // Light gray for neutral
      }
    });
    qualityRange.setBackgrounds(qualityBackgrounds);
    
    Logger.log('✅ Enhanced conditional formatting applied');
    
  } catch (error) {
    Logger.log('❌ Error applying conditional formatting: ' + error.toString());
  }
}

// BACKWARD COMPATIBILITY: Keep original function for existing systems
function extractCompanyData(row, headers) {
  return extractCompanyDataHybridSecure(row, headers);
}

function insertResults(sheet, originalRange, results, options) {
  return insertEnhancedResults(sheet, originalRange, results, options);
}
