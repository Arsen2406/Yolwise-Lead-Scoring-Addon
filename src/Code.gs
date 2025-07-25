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
  const html = HtmlService.createHtmlOutput(`
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
      <p>Add these keys to Apps