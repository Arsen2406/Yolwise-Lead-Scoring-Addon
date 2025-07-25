/**
 * Yolwise Lead Scoring Add-on
 * Main entry point for Google Sheets Add-on
 * Integrates Claude 4 Sonnet API with Yolwise scoring system for Turkish B2B market
 * ADAPTED FROM SMARTWAY: Turkish business context, Yolwise API integration
 * Robust state persistence and timeout handling for continuation mechanism
 * Hybrid mapping system with Claude fallback for comprehensive data extraction
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
  ui.createMenu('🎯Yolwise Lead Scoring')
    .addItem('🎯 Analiz ve Skorlama', 'showScoringDialog')
    .addItem('⚙️ Ayarlar', 'showSettingsDialog')
    .addItem('📖 Yardım', 'showHelpDialog')
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
  const html = HtmlService.createHtmlOutput(`
    <div style="padding: 20px; font-family: Arial, sans-serif;">
      <h3>API Yapılandırması</h3>
      <p><strong>Gerekli API Anahtarları:</strong></p>
      <ul>
        <li>Claude API Key (Anthropic)</li>
        <li>Yolwise API Key (Replit'ten alınacak)</li>
      </ul>
      <p><strong>Kurulum Talimatları:</strong></p>
      <ol>
        <li><a href="https://console.anthropic.com" target="_blank">Anthropic Console</a>'a gidin</li>
        <li>Claude 4 Sonnet için API anahtarı oluşturun</li>
        <li>Apps Script proje ayarlarında yapılandırın</li>
        <li>Yolwise API anahtarını Replit'ten alın</li>
      </ol>
      <p><strong>API Yapılandırması:</strong></p>
      <p>Bu anahtarları Apps Script Proje Ayarları → Script Properties'e ekleyin:</p>
      <ul>
        <li><code>CLAUDE_API_KEY</code> - Anthropic API anahtarınız</li>
        <li><code>YOLWISE_API_KEY</code> - Yolwise skorlama API anahtarı</li>
        <li><code>YOLWISE_API_URL</code> - https://yolwiseleadscoring.replit.app</li>
      </ul>
      <button onclick="google.script.host.close()">Kapat</button>
    </div>
  `)
    .setWidth(500)
    .setHeight(500)
    .setTitle('Ayarlar');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Ayarlar');
}

/**
 * Shows help dialog with usage instructions
 */
function showHelpDialog() {
  const html = HtmlService.createHtmlOutput(`
    <div style="padding: 20px; font-family: Arial, sans-serif;">
      <h3>Yolwise Lead Scoring Yardım</h3>
      <p><strong>Nasıl kullanılır:</strong></p>
      <ol>
        <li>Şirket verileri içeren bir tablo hazırlayın</li>
        <li>"Analiz ve Skorlama" butonuna tıklayın</li>
        <li>Veri aralığını seçin ve seçenekleri yapılandırın</li>
        <li>Add-on, Claude API ile analiz yapacak ve Yolwise API ile skorlayacak</li>
        <li>Sonuçlar yeni sütunlar olarak eklenecek</li>
      </ol>
      <p><strong>Gerekli sütunlar:</strong> Şirket Adı (minimum)</p>
      <p><strong>İsteğe bağlı sütunlar:</strong> Sektör, Gelir, Çalışan Sayısı, Şehir, vb.</p>
      
      <p><strong>Adım adım kurulum:</strong></p>
      <ol>
        <li>Ayarlar'da API anahtarlarını yapılandırın</li>
        <li>Küçük bir veri seti ile bağlantıyı test edin</li>
        <li>Sonuçları gözden geçirin ve parametreleri ayarlayın</li>
        <li>Tam veri setinizde analizi çalıştırın</li>
      </ol>
      
      <p><strong>Türk B2B Pazarı:</strong> Bu sistem Türkiye'deki B2B hizmet ihtiyaçlarına göre optimize edilmiştir.</p>
      
      <button onclick="google.script.host.close()">Kapat</button>
    </div>
  `)
    .setWidth(500)
    .setHeight(450)
    .setTitle('Yardım');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Yardım');
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
      timestamp: new Date(state.timestamp).toLocaleString('tr-TR')
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
      throw new Error('Belirtilen aralıkta veri bulunamadı');
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
          message: `İşlenen: ${results.processed} / ${results.total} şirket. Devam etmek için "Devam Et" butonunu kullanın.`
        };
      }
      
      const row = rowsToProcess[i];
      const actualRowIndex = startIndex + i;
      
      try {
        // Extract company data from row using HYBRID MAPPING SYSTEM (adapted for Turkish)
        const companyData = extractTurkishCompanyDataHybrid(row, headers);
        
        if (!companyData.company_name) {
          results.results.push({
            success: false,
            error: 'Şirket adı bulunamadı',
            companyData: companyData,
            qualityAnalysis: companyData.quality_analysis || 'Şirket adı bulunamadı'
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
            qualityAnalysis: companyData.quality_analysis || 'Tüm temel alanlar başarıyla çıkarıldı'
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
          companyData: extractTurkishCompanyDataHybrid(row, headers),
          qualityAnalysis: 'İşleme sırasında kritik hata'
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
      throw new Error('İşleme durumu bulunamadı. Oturum süresi dolmuş olabilir.');
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
      label: `Tüm veriler (A1:${sheet.getRange(lastRow, lastCol).getA1Notation().match(/[A-Z]+/)[0]}${lastRow})`,
      value: `A1:${sheet.getRange(lastRow, lastCol).getA1Notation().match(/[A-Z]+/)[0]}${lastRow}`,
      description: `${lastRow} satır, ${lastCol} sütun`
    });
    
    // Data without header
    if (lastRow > 1) {
      ranges.push({
        label: `Sadece veriler (A2:${sheet.getRange(lastRow, lastCol).getA1Notation().match(/[A-Z]+/)[0]}${lastRow})`,
        value: `A2:${sheet.getRange(lastRow, lastCol).getA1Notation().match(/[A-Z]+/)[0]}${lastRow}`,
        description: `${lastRow - 1} veri satırı, ${lastCol} sütun`
      });
    }
    
    return ranges;
    
  } catch (error) {
    Logger.log('Error getting available ranges:', error);
    return [];
  }
}

/**
 * BACKWARD COMPATIBILITY: Keep original function for existing systems
 * @param {Array} row Spreadsheet row data
 * @param {Array} headers Column headers
 * @returns {Object} Structured company data
 */
function extractCompanyData(row, headers) {
  // Call the new Turkish hybrid system
  return extractTurkishCompanyDataHybrid(row, headers);
}