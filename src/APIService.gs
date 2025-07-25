/**
 * API Service for Claude 4 Sonnet and Yolwise integration - SECURITY ENHANCED
 * Handles external API calls and data processing for Turkish B2B market
 * ADAPTED FROM SMARTWAY: Turkish business context, Yolwise API integration
 * Enhanced Claude prompt for comprehensive Turkish business data analysis
 * Hybrid system integration with quality analysis
 * 
 * SECURITY ENHANCEMENTS:
 * - Standardized authentication with X-API-Key header
 * - Comprehensive error handling with user-friendly messages
 * - Input validation and sanitization
 * - Consistent prompt engineering with language standardization
 * - Enhanced logging and debugging capabilities
 * - Protection against API abuse and rate limiting
 */

/**
 * Analyze Turkish company data with Claude 4 Sonnet API - ENHANCED VERSION
 * SECURITY IMPROVEMENTS: Input validation, error handling, language consistency
 * @param {Object} companyData Raw company data from spreadsheet
 * @param {Object} options Configuration options
 * @returns {Object} Analyzed and structured data for scoring
 */
function analyzeWithClaude(companyData, options) {
  try {
    Logger.log('🔄 Starting Claude analysis for Turkish company: ' + (companyData.company_name || 'Unknown'));
    
    // Enhanced input validation
    if (!companyData || typeof companyData !== 'object') {
      throw new Error('Invalid company data provided');
    }
    
    // Get API key with enhanced validation
    const apiKey = getClaudeApiKey();
    if (!apiKey) {
      throw new Error('Claude API key not configured. Please set CLAUDE_API_KEY in script properties.');
    }
    
    // Validate API key format
    if (!apiKey.startsWith('sk-ant-')) {
      Logger.log('⚠️ Warning: API key format may be invalid');
    }
    
    // Prepare the analysis prompt with enhanced security
    const prompt = buildEnhancedTurkishClaudePrompt(companyData, options);
    
    // Make API call to Claude with enhanced error handling
    const response = callClaudeAPISecure(prompt, apiKey);
    
    // Parse and validate the response
    const analysis = parseClaudeResponseSecure(response);
    
    Logger.log('✅ Claude analysis completed successfully for: ' + (analysis.company_name || 'Unknown'));
    return analysis;
    
  } catch (error) {
    Logger.log('❌ Error in Claude analysis: ' + error.toString());
    
    // Enhanced error categorization for better debugging
    if (error.toString().includes('API key')) {
      throw new Error('Authentication Error: ' + error.message);
    } else if (error.toString().includes('timeout')) {
      throw new Error('Network Error: Request timed out. Please try again.');
    } else if (error.toString().includes('rate limit')) {
      throw new Error('Rate Limit Error: Too many requests. Please wait and try again.');
    } else {
      throw new Error('Analysis Error: ' + error.message);
    }
  }
}

/**
 * Build enhanced analysis prompt for Claude API with consistent language and security
 * FIXED: Language consistency, better structure, input validation
 * @param {Object} companyData Company information from hybrid mapping system
 * @param {Object} options Configuration options
 * @returns {String} Formatted prompt optimized for Turkish B2B data analysis
 */
function buildEnhancedTurkishClaudePrompt(companyData, options) {
  try {
    // Input sanitization - prevent injection attacks
    const sanitizedData = sanitizeCompanyData(companyData);
    
    // Enhanced prompt with consistent language (English for better Claude processing)
    const prompt = `
You are a Turkish B2B market expert specializing in the Yolwise Lead Scoring system. You analyze company data obtained through a hybrid data extraction system (structural mapping + AI analysis) optimized for Turkish business terminology.

**IMPORTANT**: The data was obtained through an advanced hybrid extraction system (structural matching + AI analysis). This system is specifically optimized for Turkish business terminology.

**COMPANY DATA:**
${JSON.stringify(sanitizedData, null, 2)}

**DATA QUALITY:**
${sanitizedData.quality_analysis || 'Quality analysis not available'}

**ANALYSIS OF ALL AVAILABLE FIELDS:**
${Object.entries(sanitizedData).map(([key, value]) => {
  if (key === 'discovered_facts' && Array.isArray(value)) {
    return `${key}: ${value.join(' | ')}`;
  }
  if (key === 'quality_analysis') {
    return `Data quality: ${value}`;
  }
  return `${key}: ${value}`;
}).join('\\n')}

**TURKISH B2B MARKET HYBRID SYSTEM ANALYSIS FEATURES:**

1. **Consider data quality**: If data quality is low (indicated in quality_analysis), be more conservative in your conclusions.

2. **Use ALL available data**: The hybrid system may have extracted data from non-standard headers through Claude analysis - evaluate these thoroughly.

3. **Pay special attention to discovered_facts**: This data contains important information categorized according to Turkish business terminology.

4. **Focus on Turkish B2B service needs**: Consider the characteristics of the B2B service market in Turkey.

**TURKISH B2B MARKET CONTEXT:**
- High B2B service needs especially in finance, construction, manufacturing, logistics sectors
- Large companies (500+ employees) purchase more B2B services
- Business activity concentrated in major cities like Istanbul, Ankara, Izmir
- A.Ş. and large Ltd. Şti. forms indicate higher budget potential

**YOUR TASK:**
1. Analyze ALL available fields and categories considering data quality
2. Determine the sector based on Turkish business data
3. Evaluate company size according to financial indicators and employee count
4. Analyze B2B service need potential and decision-making complexity
5. Identify growth, development, or problem indicators
6. Evaluate the impact of data quality on results

**TURKISH HYBRID DATA SYSTEM PRINCIPLES:**
- If data quality is high (80%+) → provide confident results
- If data quality is medium (60-79%) → provide cautious results with warnings
- If data quality is low (<60%) → provide minimal results, focus on available facts
- Use all categories: financial_data, legal_data, operational_data, contact_data
- Pay special attention to Claude-matched data - these are carefully selected
- Consider Turkish business culture and market dynamics

**RESPONSE FORMAT (must be valid JSON only):**
{
  "company_name": "complete name from data",
  "industry": "sector based on all available data",
  "revenue_estimate": numeric_value_million_TL_or_0,
  "employees_estimate": "number or range based on all data",
  "business_type": "A.Ş./Ltd.Şti./other Turkish legal forms",
  "headquarters": "city from address data",
  "locations": ["all regions and cities found"],
  "key_people": ["managers from data"],
  "discovered_facts": ["all important facts from categorized data"],
  "growth_indicators": "growth/development/problem indicators from all sources",
  "market_position": "market position based on size and activity",
  "business_context": "business characteristics, licenses, activity details",
  "b2b_service_potential": "service potential in Turkish B2B market: sector + size + geography + characteristics",
  "analysis_confidence": "high/medium/low based on data quality and completeness",
  "data_quality_impact": "impact of data quality on analysis",
  "turkish_market_notes": "important notes specific to Turkish B2B market"
}

**HIGH-QUALITY ANALYSIS PRINCIPLES FOR TURKISH B2B MARKET:**
- Integrate all fields from hybrid system, especially evaluate Claude-matched data
- Connect financial, operational, and legal data
- Evaluate B2B service potential according to Turkish market factors: sector + size + geography + characteristics
- Indicate confidence level based on data quality
- Note which aspects of analyzed data are high/low quality

Return ONLY valid JSON, no additional text.`;

    return prompt;
    
  } catch (error) {
    Logger.log('❌ Error building Claude prompt: ' + error.toString());
    throw new Error('Failed to build analysis prompt: ' + error.message);
  }
}

/**
 * Sanitize company data to prevent injection attacks and ensure data integrity
 * @param {Object} data Raw company data
 * @returns {Object} Sanitized company data
 */
function sanitizeCompanyData(data) {
  try {
    const sanitized = {};
    
    for (const [key, value] of Object.entries(data)) {
      if (value === null || value === undefined) {
        sanitized[key] = '';
        continue;
      }
      
      if (typeof value === 'string') {
        // Remove potentially dangerous characters and limit length
        let cleanValue = value.toString()
          .replace(/[<>"\\&]/g, '')  // Remove HTML/XML characters
          .replace(/[\r\n\t]/g, ' ')  // Replace line breaks with spaces
          .substring(0, 1000);  // Limit length
        sanitized[key] = cleanValue;
      } else if (Array.isArray(value)) {
        // Sanitize array elements
        sanitized[key] = value.map(item => 
          typeof item === 'string' ? 
            item.toString().replace(/[<>"\\&]/g, '').substring(0, 500) : 
            item
        );
      } else {
        sanitized[key] = value;
      }
    }
    
    return sanitized;
    
  } catch (error) {
    Logger.log('❌ Error sanitizing company data: ' + error.toString());
    throw new Error('Data sanitization failed: ' + error.message);
  }
}

/**
 * Call Claude 4 Sonnet API with enhanced security and error handling
 * ENHANCED: Better error handling, timeout management, response validation
 * FIXED: Allow redirects for proper API functionality (resolves 307 errors)
 * @param {String} prompt Analysis prompt
 * @param {String} apiKey API key
 * @returns {Object} API response
 */
function callClaudeAPISecure(prompt, apiKey) {
  const url = 'https://api.anthropic.com/v1/messages';
  
  // Enhanced payload with better parameters
  const payload = {
    model: 'claude-3-5-sonnet-20241022',  // Updated to latest model
    max_tokens: 4000,
    temperature: 0.1,  // Low temperature for consistent results
    messages: [{
      role: 'user',
      content: prompt
    }]
  };
  
  // Enhanced request options with security features
  const options = {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
      'User-Agent': 'Yolwise-Lead-Scoring/1.1'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,  // Handle errors manually
    validateHttpsCertificates: true,  // Security best practice
    followRedirects: true  // Allow redirects for proper API functionality (fixes 307 errors)
  };
  
  Logger.log('🔗 Calling Claude API for Turkish business analysis...');
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    // Enhanced error handling with specific error codes
    if (responseCode === 401) {
      throw new Error('Authentication failed: Invalid API key');
    } else if (responseCode === 429) {
      throw new Error('Rate limit exceeded: Please wait before making another request');
    } else if (responseCode === 400) {
      const errorContent = response.getContentText();
      Logger.log('❌ Claude API 400 error: ' + errorContent);
      throw new Error('Invalid request: Please check your data format');
    } else if (responseCode === 500) {
      throw new Error('Claude API server error: Please try again later');
    } else if (responseCode !== 200) {
      Logger.log('❌ Claude API unexpected error: ' + responseCode + ' - ' + response.getContentText());
      throw new Error(`Claude API error: ${responseCode} - Please try again`);
    }
    
    const responseData = JSON.parse(response.getContentText());
    Logger.log('✅ Claude API response received successfully');
    
    // Validate response structure
    if (!responseData.content || !Array.isArray(responseData.content) || responseData.content.length === 0) {
      throw new Error('Invalid response format from Claude API');
    }
    
    return responseData;
    
  } catch (error) {
    Logger.log('❌ Claude API call failed: ' + error.toString());
    
    // Enhanced error categorization
    if (error.toString().includes('timeout')) {
      throw new Error('Network timeout: Please check your connection and try again');
    } else if (error.toString().includes('DNS')) {
      throw new Error('Network error: Cannot reach Claude API');
    } else {
      throw error;  // Re-throw with original message
    }
  }
}

/**
 * Parse Claude API response with enhanced error handling and validation
 * ENHANCED: Better JSON parsing, validation, error recovery
 * @param {Object} response API response
 * @returns {Object} Parsed analysis
 */
function parseClaudeResponseSecure(response) {
  try {
    if (!response.content || !response.content[0] || !response.content[0].text) {
      throw new Error('Invalid Claude API response structure');
    }
    
    const content = response.content[0].text.trim();
    
    // Enhanced JSON extraction with better error handling
    let jsonStr = content;
    
    // Handle potential markdown formatting
    if (content.includes('```json')) {
      const match = content.match(/```json\s*([\s\S]*?)\s*```/);
      if (match) {
        jsonStr = match[1];
      }
    } else if (content.includes('```')) {
      // Handle generic code blocks
      const match = content.match(/```\s*([\s\S]*?)\s*```/);
      if (match) {
        jsonStr = match[1];
      }
    }
    
    // Clean up common JSON formatting issues
    jsonStr = jsonStr
      .replace(/^\s*{/, '{')  // Remove leading whitespace before opening brace
      .replace(/}\s*$/, '}')  // Remove trailing whitespace after closing brace
      .replace(/,\s*}/g, '}')  // Remove trailing commas
      .replace(/,\s*]/g, ']');  // Remove trailing commas in arrays
    
    let analysis;
    try {
      analysis = JSON.parse(jsonStr);
    } catch (parseError) {
      Logger.log('❌ JSON parsing failed, attempting recovery...');
      Logger.log('Raw content: ' + content.substring(0, 500) + '...');
      
      // Attempt to extract key information even if JSON is malformed
      analysis = extractBasicInfoFromText(content);
    }
    
    // Validate and sanitize required fields
    if (!analysis.company_name) {
      throw new Error('Company name not found in Claude analysis');
    }
    
    // Ensure all required fields have default values
    const defaultAnalysis = {
      company_name: 'Unknown',
      industry: 'Unknown',
      revenue_estimate: 0,
      employees_estimate: 'Unknown',
      business_type: 'Unknown',
      headquarters: 'Unknown',
      locations: [],
      key_people: [],
      discovered_facts: [],
      growth_indicators: 'Unknown',
      market_position: 'Unknown',
      business_context: 'Unknown',
      b2b_service_potential: 'Unknown',
      analysis_confidence: 'low',
      data_quality_impact: 'Unknown',
      turkish_market_notes: 'No specific Turkish market analysis available'
    };
    
    // Merge with defaults and validate
    analysis = Object.assign(defaultAnalysis, analysis);
    
    // Validate Turkish-specific fields
    if (analysis.turkish_market_notes && analysis.turkish_market_notes.length > 0) {
      Logger.log('✅ Turkish market analysis completed: ' + analysis.turkish_market_notes.substring(0, 100) + '...');
    }
    
    Logger.log('✅ Claude response parsed successfully for Turkish company: ' + analysis.company_name);
    return analysis;
    
  } catch (error) {
    Logger.log('❌ Error parsing Claude response: ' + error.toString());
    Logger.log('Response structure: ' + JSON.stringify(response, null, 2).substring(0, 500) + '...');
    throw new Error('Failed to parse Claude analysis: ' + error.message);
  }
}

/**
 * Extract basic information from malformed text when JSON parsing fails
 * Fallback function for error recovery
 * @param {String} text Raw response text
 * @returns {Object} Basic extracted information
 */
function extractBasicInfoFromText(text) {
  try {
    Logger.log('🔧 Attempting text-based information extraction...');
    
    const analysis = {
      company_name: 'Unknown',
      industry: 'Unknown', 
      revenue_estimate: 0,
      employees_estimate: 'Unknown',
      analysis_confidence: 'low',
      data_quality_impact: 'JSON parsing failed - extracted from text',
      turkish_market_notes: 'Basic extraction due to parsing error'
    };
    
    // Extract company name
    const nameMatch = text.match(/"company_name"\s*:\s*"([^"]+)"/);
    if (nameMatch) {
      analysis.company_name = nameMatch[1];
    }
    
    // Extract industry
    const industryMatch = text.match(/"industry"\s*:\s*"([^"]+)"/);
    if (industryMatch) {
      analysis.industry = industryMatch[1];
    }
    
    // Extract revenue
    const revenueMatch = text.match(/"revenue_estimate"\s*:\s*(\d+)/);
    if (revenueMatch) {
      analysis.revenue_estimate = parseInt(revenueMatch[1]);
    }
    
    Logger.log('✅ Basic text extraction completed for: ' + analysis.company_name);
    return analysis;
    
  } catch (error) {
    Logger.log('❌ Text extraction also failed: ' + error.toString());
    return {
      company_name: 'Extraction Failed',
      industry: 'Unknown',
      analysis_confidence: 'low',
      data_quality_impact: 'Complete parsing failure'
    };
  }
}

/**
 * Score company with Yolwise API - ENHANCED WITH STANDARDIZED AUTHENTICATION
 * FIXED: Consistent authentication, better error handling, enhanced logging
 * FIXED: Allow redirects for proper API functionality (prevents 307 errors)
 * @param {Object} claudeAnalysis Analyzed company data
 * @returns {Object} Scoring results with proper API format
 */
function scoreWithYolwise(claudeAnalysis) {
  try {
    Logger.log('🎯 Starting Yolwise API scoring for: ' + (claudeAnalysis.company_name || 'Unknown'));
    
    // Enhanced input validation
    if (!claudeAnalysis || typeof claudeAnalysis !== 'object') {
      throw new Error('Invalid Claude analysis data provided');
    }
    
    // Get Yolwise API configuration with validation
    const apiUrl = getYolwiseApiUrl();
    const apiKey = getYolwiseApiKey();
    
    if (!apiUrl) {
      Logger.log('⚠️ Yolwise API URL not configured, using enhanced mock scoring');
      return createEnhancedYolwiseMockScoring(claudeAnalysis);
    }
    
    // Prepare payload for Yolwise API according to enhanced specification
    const payload = {
      company_name: claudeAnalysis.company_name || 'Unknown',
      company_data: {
        // Core business data
        industry: claudeAnalysis.industry || '',
        annual_revenue: claudeAnalysis.revenue_estimate || 0,
        number_of_employees: claudeAnalysis.employees_estimate || '',
        
        // Location data
        city: claudeAnalysis.headquarters || '',
        locations: claudeAnalysis.locations || [],
        
        // Business structure
        business_type: claudeAnalysis.business_type || '',
        
        // Additional context
        key_people: claudeAnalysis.key_people || [],
        discovered_facts: claudeAnalysis.discovered_facts || [],
        growth_indicators: claudeAnalysis.growth_indicators || '',
        market_position: claudeAnalysis.market_position || '',
        description: claudeAnalysis.business_context || '',
        
        // Enhanced Yolwise-specific fields
        data_quality_impact: claudeAnalysis.data_quality_impact || '',
        turkish_market_notes: claudeAnalysis.turkish_market_notes || '',
        analysis_confidence: claudeAnalysis.analysis_confidence || 'medium',
        b2b_service_potential: claudeAnalysis.b2b_service_potential || ''
      }
    };
    
    // Enhanced request options with standardized authentication
    const options = {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'User-Agent': 'Yolwise-GoogleAppsScript/1.1'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,  // Handle errors manually
      validateHttpsCertificates: true,
      followRedirects: true  // Allow redirects for proper API functionality (prevents 307 errors)
    };
    
    // STANDARDIZED AUTHENTICATION - Use X-API-Key header consistently
    if (apiKey) {
      options.headers['X-API-Key'] = apiKey;
      Logger.log('✅ API key authentication configured');
    } else {
      Logger.log('⚠️ No API key configured - requests may fail');
    }
    
    try {
      Logger.log('🔗 Calling Yolwise API: ' + apiUrl + '/score_company');
      const response = UrlFetchApp.fetch(`${apiUrl}/score_company`, options);
      const responseCode = response.getResponseCode();
      
      // Enhanced error handling with specific status codes
      if (responseCode === 401) {
        Logger.log('❌ Yolwise API: Unauthorized - Check API key configuration');
        return createEnhancedYolwiseMockScoring(claudeAnalysis, 'Authentication failed');
      } else if (responseCode === 403) {
        Logger.log('❌ Yolwise API: Forbidden - API key may be invalid or expired');
        return createEnhancedYolwiseMockScoring(claudeAnalysis, 'Access forbidden');
      } else if (responseCode === 429) {
        Logger.log('❌ Yolwise API: Rate limit exceeded');
        return createEnhancedYolwiseMockScoring(claudeAnalysis, 'Rate limit exceeded');
      } else if (responseCode === 400) {
        Logger.log('❌ Yolwise API: Bad request - Invalid payload format');
        return createEnhancedYolwiseMockScoring(claudeAnalysis, 'Invalid request format');
      } else if (responseCode !== 200) {
        Logger.log(`❌ Yolwise API returned ${responseCode}, using enhanced mock scoring`);
        return createEnhancedYolwiseMockScoring(claudeAnalysis, `API error: ${responseCode}`);
      }
      
      const apiResult = JSON.parse(response.getContentText());
      
      if (!apiResult.success || !apiResult.result) {
        Logger.log('❌ Yolwise API returned unsuccessful response, using enhanced mock scoring');
        return createEnhancedYolwiseMockScoring(claudeAnalysis, 'API returned error');
      }
      
      // Process successful API response
      const baseResult = apiResult.result;
      
      // Apply enhanced LLM adjustment based on Claude analysis for Turkish context
      const finalResult = applyEnhancedTurkishLLMAdjustment(baseResult, claudeAnalysis);
      
      Logger.log('✅ Yolwise API scoring completed successfully with score: ' + finalResult.final_score);
      return finalResult;
      
    } catch (apiError) {
      Logger.log('❌ Yolwise API call failed: ' + apiError.toString());
      return createEnhancedYolwiseMockScoring(claudeAnalysis, 'API call failed: ' + apiError.message);
    }
    
  } catch (error) {
    Logger.log('❌ Error in Yolwise scoring process: ' + error.toString());
    return createEnhancedYolwiseMockScoring(claudeAnalysis, 'Scoring process error: ' + error.message);
  }
}

/**
 * Apply enhanced LLM adjustment to API scoring results with Turkish B2B context
 * FIXED: Improved boundary logic, better algorithm, enhanced reasoning
 * @param {Object} apiResult Base scoring result from Yolwise API
 * @param {Object} claudeAnalysis Claude analysis data
 * @returns {Object} Result with proper LLM adjustment applied
 */
function applyEnhancedTurkishLLMAdjustment(apiResult, claudeAnalysis) {
  Logger.log('🧠 Applying enhanced Turkish B2B LLM adjustment for: ' + (apiResult.company_name || 'Unknown'));
  
  try {
    // Enhanced adjustment factors with improved scoring
    const adjustments = [];
    const reasoning = [];
    
    // Base information from API
    reasoning.push(`API Scoring: Base ${apiResult.base_score || 0}, Industry ${apiResult.detected_industry || 'unknown'} (×${apiResult.industry_multiplier || 1.0})`);
    
    // Enhanced Turkish market confidence adjustments
    const analysisConfidence = claudeAnalysis.analysis_confidence || 'medium';
    if (analysisConfidence === 'low') {
      adjustments.push({ value: -8, reason: 'Düşük veri kalitesi Türk pazarı değerlendirmesini olumsuz etkiliyor' });
    } else if (analysisConfidence === 'high') {
      adjustments.push({ value: 6, reason: 'Yüksek veri kalitesi güvenilir Türk pazarı analizi sağlıyor' });
    }
    
    // Enhanced B2B service potential assessment
    if (claudeAnalysis.b2b_service_potential) {
      const potentialText = claudeAnalysis.b2b_service_potential.toLowerCase();
      
      if (potentialText.includes('high') || potentialText.includes('güçlü') || potentialText.includes('yüksek')) {
        adjustments.push({ value: 15, reason: 'Türk B2B pazarında çok yüksek hizmet potansiyeli' });
      } else if (potentialText.includes('strong') || potentialText.includes('significant')) {
        adjustments.push({ value: 12, reason: 'Türk B2B pazarında güçlü hizmet potansiyeli' });
      } else if (potentialText.includes('low') || potentialText.includes('limited') || potentialText.includes('düşük')) {
        adjustments.push({ value: -10, reason: 'Türk B2B pazarında sınırlı hizmet potansiyeli' });
      } else if (potentialText.includes('medium') || potentialText.includes('moderate') || potentialText.includes('orta')) {
        adjustments.push({ value: 5, reason: 'Türk B2B pazarında orta düzey hizmet potansiyeli' });
      }
    }
    
    // Enhanced business context analysis
    if (claudeAnalysis.business_context) {
      const contextText = claudeAnalysis.business_context.toLowerCase();
      
      // Enhanced positive business indicators
      if (contextText.includes('growth') || contextText.includes('expansion') || 
          contextText.includes('büyüme') || contextText.includes('genişleme')) {
        adjustments.push({ value: 12, reason: 'Aktif büyüme ve genişleme stratejisi' });
      }
      
      if (contextText.includes('leader') || contextText.includes('market leader') ||
          contextText.includes('lider') || contextText.includes('pazar lideri')) {
        adjustments.push({ value: 10, reason: 'Pazar liderliği pozisyonu' });
      }
      
      // Enhanced export/international business indicators
      if (contextText.includes('export') || contextText.includes('international') || 
          contextText.includes('ihracat') || contextText.includes('uluslararası')) {
        adjustments.push({ value: 8, reason: 'Uluslararası ticaret faaliyetleri' });
      }
      
      // Technology and innovation indicators
      if (contextText.includes('technology') || contextText.includes('innovation') ||
          contextText.includes('digital') || contextText.includes('teknoloji')) {
        adjustments.push({ value: 7, reason: 'Teknoloji ve inovasyon odaklı yaklaşım' });
      }
      
      // Enhanced negative indicators
      if (contextText.includes('crisis') || contextText.includes('decline') || 
          contextText.includes('kriz') || contextText.includes('düşüş') || contextText.includes('zarar')) {
        adjustments.push({ value: -15, reason: 'Finansal veya operasyonel zorluklar' });
      }
    }
    
    // Enhanced growth indicators assessment
    if (claudeAnalysis.growth_indicators) {
      const growthText = claudeAnalysis.growth_indicators.toLowerCase();
      
      if (growthText.includes('strong growth') || growthText.includes('rapid growth')) {
        adjustments.push({ value: 10, reason: 'Güçlü büyüme performansı' });
      } else if (growthText.includes('growth') || growthText.includes('büyüme')) {
        adjustments.push({ value: 6, reason: 'Pozitif büyüme göstergeleri' });
      } else if (growthText.includes('decline') || growthText.includes('düşüş')) {
        adjustments.push({ value: -8, reason: 'Negatif büyüme göstergeleri' });
      }
    }
    
    // Enhanced market position analysis
    if (claudeAnalysis.market_position) {
      const positionText = claudeAnalysis.market_position.toLowerCase();
      
      if (positionText.includes('dominant') || positionText.includes('leading')) {
        adjustments.push({ value: 12, reason: 'Piyasada baskın/lider pozisyon' });
      } else if (positionText.includes('strong') || positionText.includes('established')) {
        adjustments.push({ value: 8, reason: 'Güçlü pazar pozisyonu' });
      } else if (positionText.includes('small') || positionText.includes('limited')) {
        adjustments.push({ value: -5, reason: 'Sınırlı pazar varlığı' });
      }
    }
    
    // Enhanced Turkish market specific adjustments
    if (claudeAnalysis.turkish_market_notes) {
      const marketNotes = claudeAnalysis.turkish_market_notes.toLowerCase();
      
      if (marketNotes.includes('istanbul') || marketNotes.includes('ankara') || marketNotes.includes('izmir')) {
        adjustments.push({ value: 6, reason: 'Büyük şehir merkezlerinde konum avantajı' });
      }
      
      if (marketNotes.includes('industrial') || marketNotes.includes('manufacturing') ||
          marketNotes.includes('sanayi') || marketNotes.includes('üretim')) {
        adjustments.push({ value: 5, reason: 'Endüstriyel faaliyet alanında konum' });
      }
      
      if (marketNotes.includes('family business') && marketNotes.includes('professional')) {
        adjustments.push({ value: 4, reason: 'Profesyonelleşen aile şirketi yapısı' });
      }
    }
    
    // FIXED: Enhanced boundary logic for optimal adjustment selection
    let totalAdjustment = 0;
    const appliedAdjustments = [];
    
    // Sort by impact priority: positive adjustments first, then negative
    const positiveAdj = adjustments.filter(adj => adj.value > 0).sort((a, b) => b.value - a.value);
    const negativeAdj = adjustments.filter(adj => adj.value < 0).sort((a, b) => a.value - b.value);
    
    // Apply positive adjustments first
    for (const adj of positiveAdj) {
      if (totalAdjustment + adj.value <= 25) {
        totalAdjustment += adj.value;
        appliedAdjustments.push(adj);
        reasoning.push(`LLM: +${adj.value} - ${adj.reason}`);
      }
    }
    
    // Then apply negative adjustments if there's room
    for (const adj of negativeAdj) {
      if (totalAdjustment + adj.value >= -25) {
        totalAdjustment += adj.value;
        appliedAdjustments.push(adj);
        reasoning.push(`LLM: ${adj.value} - ${adj.reason}`);
      }
    }
    
    // Final bounds check
    totalAdjustment = Math.max(-25, Math.min(25, totalAdjustment));
    
    // Calculate final score
    const industryAdjustedScore = apiResult.industry_adjusted_score || apiResult.base_score || 0;
    const finalScore = Math.max(0, Math.min(100, industryAdjustedScore + totalAdjustment));
    
    // Enhanced priority determination
    const finalPriority = finalScore >= 60 ? 'target' : 'non_target';
    const finalPriorityTr = finalScore >= 60 ? 'HEDEFLENEBİLİR' : 'HEDEFLENEMEZ';
    
    Logger.log(`✅ Enhanced Turkish LLM adjustment applied: ${totalAdjustment}, Final score: ${finalScore}`);
    
    return {
      ...apiResult,
      llm_adjustment: totalAdjustment,
      final_score: finalScore,
      priority_recommendation: finalPriority,
      priority_recommendation_tr: finalPriorityTr,
      reasoning: reasoning.join(' | '),
      applied_adjustments: appliedAdjustments,
      claude_business_context: claudeAnalysis.business_context || 'İş bağlamı bilgisi mevcut değil',
      claude_confidence: claudeAnalysis.analysis_confidence || 'orta',
      // Enhanced Turkish-specific fields
      data_quality_impact: claudeAnalysis.data_quality_impact || 'Veri kalitesi analizi mevcut değil',
      turkish_market_notes: claudeAnalysis.turkish_market_notes || 'Türk pazarı özel analizi mevcut değil',
      b2b_service_potential: claudeAnalysis.b2b_service_potential || 'B2B hizmet potansiyeli değerlendirmesi mevcut değil',
      processing_timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    Logger.log('❌ Error in LLM adjustment: ' + error.toString());
    
    // Return basic result with error information
    return {
      ...apiResult,
      llm_adjustment: 0,
      final_score: apiResult.industry_adjusted_score || apiResult.base_score || 0,
      priority_recommendation: (apiResult.industry_adjusted_score || apiResult.base_score || 0) >= 60 ? 'target' : 'non_target',
      reasoning: 'LLM adjustment failed: ' + error.message,
      error_in_adjustment: true
    };
  }
}

/**
 * Create enhanced mock scoring for development/fallback with improved Turkish context
 * ENHANCED: Better scoring logic, more realistic values, improved reasoning
 * @param {Object} claudeAnalysis Claude analysis data
 * @param {String} errorReason Reason for using mock scoring (optional)
 * @returns {Object} Enhanced mock scoring result matching Yolwise API format
 */
function createEnhancedYolwiseMockScoring(claudeAnalysis, errorReason) {
  Logger.log('🔧 Creating enhanced Yolwise mock scoring for: ' + (claudeAnalysis.company_name || 'Unknown'));
  
  try {
    // Enhanced industry scoring with more realistic Turkish market data
    const enhancedTurkishIndustryModifiers = {
      'finance': { multiplier: 1.20, confidence: 'high', reasoning: 'Türk finans sektöründe yüksek B2B hizmet ihtiyacı' },
      'logistics': { multiplier: 1.17, confidence: 'high', reasoning: 'Lojistik sektöründe karmaşık operasyonel B2B ihtiyaçları' },
      'utilities': { multiplier: 1.15, confidence: 'high', reasoning: 'Kamu hizmetlerinde altyapı-ağır B2B gereksinimleri' },
      'food': { multiplier: 1.10, confidence: 'medium', reasoning: 'Gıda sektöründe orta düzey B2B hizmet gereksinimleri' },
      'chemicals': { multiplier: 1.05, confidence: 'medium', reasoning: 'Kimya sektöründe özel teknik hizmetler gerekli' },
      'construction': { multiplier: 1.02, confidence: 'medium', reasoning: 'İnşaat sektöründe proje bazlı B2B hizmetler' },
      'manufacturing': { multiplier: 1.00, confidence: 'medium', reasoning: 'İmalat sektöründe baseline B2B hizmet ihtiyacı' },
      'retail': { multiplier: 0.90, confidence: 'low', reasoning: 'Perakende sektöründe sınırlı B2B ihtiyaç' },
      'automotive': { multiplier: 0.85, confidence: 'low', reasoning: 'Otomotiv sektöründe üretim odaklı operasyonlar' },
      'technology': { multiplier: 0.80, confidence: 'low', reasoning: 'Teknoloji sektöründe self-servis çözümler tercih' },
      'healthcare': { multiplier: 0.75, confidence: 'low', reasoning: 'Sağlık sektöründe özel hizmet gereksinimleri' }
    };
    
    // Enhanced base score calculation
    let baseScore = 50; // Default baseline
    
    // Adjust based on analysis confidence
    const analysisConfidence = claudeAnalysis.analysis_confidence || 'medium';
    if (analysisConfidence === 'low') {
      baseScore = Math.max(25, baseScore - 15);
    } else if (analysisConfidence === 'high') {
      baseScore = Math.min(85, baseScore + 15);
    }
    
    // Enhanced B2B service potential scoring
    if (claudeAnalysis.b2b_service_potential) {
      const potentialText = claudeAnalysis.b2b_service_potential.toLowerCase();
      if (potentialText.includes('high') || potentialText.includes('güçlü') || potentialText.includes('yüksek')) {
        baseScore += 20;
      } else if (potentialText.includes('strong') || potentialText.includes('significant')) {
        baseScore += 15;
      } else if (potentialText.includes('low') || potentialText.includes('düşük') || potentialText.includes('limited')) {
        baseScore -= 12;
      } else if (potentialText.includes('medium') || potentialText.includes('orta')) {
        baseScore += 8;
      }
    }
    
    // Enhanced revenue-based scoring
    if (claudeAnalysis.revenue_estimate && claudeAnalysis.revenue_estimate > 0) {
      const revenue = parseFloat(claudeAnalysis.revenue_estimate);
      if (revenue >= 1000) {  // 1B+ TL
        baseScore += 15;
      } else if (revenue >= 200) {  // 200M+ TL
        baseScore += 10;
      } else if (revenue >= 50) {   // 50M+ TL
        baseScore += 5;
      }
    }
    
    // Enhanced employee-based scoring
    if (claudeAnalysis.employees_estimate) {
      const empText = claudeAnalysis.employees_estimate.toString().toLowerCase();
      const empNumbers = empText.match(/\d+/g);
      if (empNumbers) {
        const empCount = parseInt(empNumbers[0]);
        if (empCount >= 1000) {
          baseScore += 12;
        } else if (empCount >= 200) {
          baseScore += 8;
        } else if (empCount >= 50) {
          baseScore += 4;
        }
      }
    }
    
    // Enhanced industry detection and scoring
    const industryText = (claudeAnalysis.industry || '').toLowerCase();
    let detectedIndustry = 'other';
    let industryMultiplier = 1.0;
    let industryReasoning = 'Standart sektörel değerlendirme';
    
    // Enhanced keyword matching
    const industryKeywords = {
      'finance': ['bank', 'finance', 'insurance', 'financial', 'banka', 'finans', 'sigorta'],
      'logistics': ['logistics', 'transport', 'shipping', 'cargo', 'lojistik', 'nakliye', 'kargo'],
      'utilities': ['electric', 'energy', 'utility', 'gas', 'water', 'elektrik', 'enerji'],
      'food': ['food', 'beverage', 'agriculture', 'farming', 'gıda', 'içecek', 'tarım'],
      'chemicals': ['chemical', 'pharmaceutical', 'drug', 'kimya', 'ilaç'],
      'construction': ['construction', 'building', 'contractor', 'inşaat', 'yapı'],
      'manufacturing': ['manufacturing', 'production', 'factory', 'üretim', 'imalat', 'fabrika'],
      'retail': ['retail', 'store', 'shop', 'perakende', 'mağaza'],
      'automotive': ['automotive', 'automobile', 'vehicle', 'otomotiv', 'araç'],
      'technology': ['technology', 'software', 'IT', 'teknoloji', 'yazılım'],
      'healthcare': ['healthcare', 'hospital', 'medical', 'sağlık', 'hastane', 'tıp']
    };
    
    for (const [industry, keywords] of Object.entries(industryKeywords)) {
      if (keywords.some(keyword => industryText.includes(keyword))) {
        detectedIndustry = industry;
        const industryData = enhancedTurkishIndustryModifiers[industry];
        if (industryData) {
          industryMultiplier = industryData.multiplier;
          industryReasoning = industryData.reasoning;
        }
        break;
      }
    }
    
    // Ensure realistic base score range
    baseScore = Math.max(20, Math.min(90, baseScore));
    
    // Calculate industry-adjusted score
    const industryAdjustedScore = Math.max(0, Math.min(100, Math.round(baseScore * industryMultiplier)));
    
    // Create enhanced mock result
    const mockResult = {
      company_name: claudeAnalysis.company_name || 'Unknown Company',
      base_score: Math.round(baseScore * 10) / 10,
      industry_multiplier: Math.round(industryMultiplier * 100) / 100,
      industry_adjusted_score: industryAdjustedScore,
      detected_industry: detectedIndustry,
      industry_confidence: enhancedTurkishIndustryModifiers[detectedIndustry]?.confidence || 'low',
      industry_explanation: `Türk sektör analizi: ${detectedIndustry}, çarpan: ×${industryMultiplier}, ${industryReasoning}`,
      processing_time_ms: 125,
      priority_recommendation: industryAdjustedScore >= 60 ? 'target' : 'non_target',
      mock_scoring: true,
      mock_reason: errorReason || 'API unavailable - using enhanced mock scoring'
    };
    
    // Apply enhanced Turkish LLM adjustment
    const finalResult = applyEnhancedTurkishLLMAdjustment(mockResult, claudeAnalysis);
    finalResult.mock_scoring = true;
    finalResult.mock_reason = errorReason || 'API unavailable - using enhanced mock scoring';
    
    Logger.log('✅ Enhanced mock scoring completed with final score: ' + finalResult.final_score);
    return finalResult;
    
  } catch (error) {
    Logger.log('❌ Error creating enhanced mock scoring: ' + error.toString());
    
    // Return minimal fallback result
    return {
      company_name: claudeAnalysis.company_name || 'Unknown',
      base_score: 40,
      industry_multiplier: 1.0,
      industry_adjusted_score: 40,
      llm_adjustment: 0,
      final_score: 40,
      priority_recommendation: 'non_target',
      reasoning: 'Mock scoring error: ' + error.message,
      mock_scoring: true,
      error_in_mock_scoring: true
    };
  }
}

/**
 * Get Claude API key from script properties with enhanced validation
 * @returns {String} API key or null
 */
function getClaudeApiKey() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const apiKey = properties.getProperty('CLAUDE_API_KEY');
    
    if (!apiKey) {
      Logger.log('⚠️ Claude API key not found in script properties');
      return null;
    }
    
    // Basic validation
    if (apiKey.length < 10) {
      Logger.log('⚠️ Claude API key appears to be invalid (too short)');
      return null;
    }
    
    return apiKey;
    
  } catch (error) {
    Logger.log('❌ Error retrieving Claude API key: ' + error.toString());
    return null;
  }
}

/**
 * Get Yolwise API URL from script properties with enhanced validation
 * @returns {String} API URL or default
 */
function getYolwiseApiUrl() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const apiUrl = properties.getProperty('YOLWISE_API_URL');
    
    if (!apiUrl) {
      Logger.log('ℹ️ Using default Yolwise API URL');
      return 'https://yolwiseleadscoring.replit.app';
    }
    
    // Basic URL validation
    if (!apiUrl.startsWith('http')) {
      Logger.log('⚠️ Invalid Yolwise API URL format, using default');
      return 'https://yolwiseleadscoring.replit.app';
    }
    
    return apiUrl;
    
  } catch (error) {
    Logger.log('❌ Error retrieving Yolwise API URL: ' + error.toString());
    return 'https://yolwiseleadscoring.replit.app';
  }
}

/**
 * Get Yolwise API key from script properties with enhanced validation
 * @returns {String} API key or null
 */
function getYolwiseApiKey() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const apiKey = properties.getProperty('YOLWISE_API_KEY');
    
    if (!apiKey) {
      Logger.log('⚠️ Yolwise API key not found in script properties');
      return null;
    }
    
    return apiKey;
    
  } catch (error) {
    Logger.log('❌ Error retrieving Yolwise API key: ' + error.toString());
    return null;
  }
}

/**
 * Set API keys for Yolwise with enhanced validation (for setup)
 * @param {String} claudeKey Claude API key
 * @param {String} yolwiseKey Yolwise API key
 * @param {String} yolwiseUrl Yolwise API URL (optional)
 */
function setYolwiseApiKeysSecure(claudeKey, yolwiseKey, yolwiseUrl) {
  try {
    const properties = PropertiesService.getScriptProperties();
    
    if (claudeKey) {
      // Basic validation for Claude API key
      if (claudeKey.length < 10) {
        throw new Error('Claude API key appears to be invalid');
      }
      properties.setProperty('CLAUDE_API_KEY', claudeKey);
      Logger.log('✅ Claude API key updated');
    }
    
    if (yolwiseKey) {
      properties.setProperty('YOLWISE_API_KEY', yolwiseKey);
      Logger.log('✅ Yolwise API key updated');
    }
    
    if (yolwiseUrl) {
      // Basic URL validation
      if (!yolwiseUrl.startsWith('http')) {
        throw new Error('Invalid URL format');
      }
      properties.setProperty('YOLWISE_API_URL', yolwiseUrl);
      Logger.log('✅ Yolwise API URL updated: ' + yolwiseUrl);
    } else {
      properties.setProperty('YOLWISE_API_URL', 'https://yolwiseleadscoring.replit.app');
      Logger.log('✅ Using default Yolwise API URL');
    }
    
    Logger.log('✅ API configuration updated successfully');
    
  } catch (error) {
    Logger.log('❌ Error setting API keys: ' + error.toString());
    throw error;
  }
}

/**
 * Test API connections for Yolwise setup with enhanced diagnostics
 */
function testYolwiseApiConnectionsEnhanced() {
  try {
    Logger.log('🔧 Starting enhanced API connection tests...');
    
    const results = {
      claude_api: { configured: false, valid: false, error: null },
      yolwise_api: { configured: false, accessible: false, authenticated: false, error: null }
    };
    
    // Test Claude API
    const claudeKey = getClaudeApiKey();
    if (claudeKey) {
      results.claude_api.configured = true;
      Logger.log('✅ Claude API key configured');
      
      // Basic format validation
      if (claudeKey.startsWith('sk-ant-')) {
        results.claude_api.valid = true;
        Logger.log('✅ Claude API key format appears valid');
      } else {
        results.claude_api.error = 'API key format may be invalid';
        Logger.log('⚠️ Claude API key format may be invalid');
      }
    } else {
      results.claude_api.error = 'API key not configured';
      Logger.log('❌ Claude API key missing');
    }
    
    // Test Yolwise API
    const yolwiseKey = getYolwiseApiKey();
    const yolwiseUrl = getYolwiseApiUrl();
    
    if (yolwiseKey) {
      results.yolwise_api.configured = true;
      Logger.log('✅ Yolwise API key configured');
    } else {
      results.yolwise_api.error = 'API key not configured';
      Logger.log('❌ Yolwise API key missing');
    }
    
    if (yolwiseUrl) {
      Logger.log('✅ Yolwise API URL configured: ' + yolwiseUrl);
      
      // Test API health endpoint
      try {
        const response = UrlFetchApp.fetch(`${yolwiseUrl}/health`, {
          method: 'GET',
          muteHttpExceptions: true,
          validateHttpsCertificates: true
        });
        
        if (response.getResponseCode() === 200) {
          results.yolwise_api.accessible = true;
          Logger.log('✅ Yolwise API health check passed');
          
          // Test authentication if key is available
          if (yolwiseKey) {
            const authTest = UrlFetchApp.fetch(`${yolwiseUrl}/industries`, {
              method: 'GET',
              headers: { 'X-API-Key': yolwiseKey },
              muteHttpExceptions: true,
              validateHttpsCertificates: true
            });
            
            if (authTest.getResponseCode() === 200) {
              results.yolwise_api.authenticated = true;
              Logger.log('✅ Yolwise API authentication test passed');
            } else {
              results.yolwise_api.error = `Authentication failed: ${authTest.getResponseCode()}`;
              Logger.log('❌ Yolwise API authentication test failed: ' + authTest.getResponseCode());
            }
          }
        } else {
          results.yolwise_api.error = `Health check failed: ${response.getResponseCode()}`;
          Logger.log('❌ Yolwise API health check failed: ' + response.getResponseCode());
        }
      } catch (healthError) {
        results.yolwise_api.error = 'Connection failed: ' + healthError.toString();
        Logger.log('❌ Yolwise API connection failed: ' + healthError.toString());
      }
    } else {
      results.yolwise_api.error = 'API URL not configured';
      Logger.log('❌ Yolwise API URL missing');
    }
    
    // Summary
    Logger.log('🏁 API connection test results:');
    Logger.log('   Claude API: ' + (results.claude_api.configured && results.claude_api.valid ? '✅ Ready' : '❌ Issues'));
    Logger.log('   Yolwise API: ' + (results.yolwise_api.accessible && results.yolwise_api.authenticated ? '✅ Ready' : '❌ Issues'));
    
    return results;
    
  } catch (error) {
    Logger.log('❌ API connection test error: ' + error.toString());
    throw error;
  }
}

// BACKWARD COMPATIBILITY: Keep original function names
function setYolwiseApiKeys(claudeKey, yolwiseKey, yolwiseUrl) {
  return setYolwiseApiKeysSecure(claudeKey, yolwiseKey, yolwiseUrl);
}

function testYolwiseApiConnections() {
  return testYolwiseApiConnectionsEnhanced();
}
