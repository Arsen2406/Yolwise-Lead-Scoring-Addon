/**
 * Turkish Data Extraction Module for Yolwise Lead Scoring
 * Specialized functions for extracting and mapping Turkish business data
 * ADAPTED FROM SMARTWAY: Turkish business terminology and legal structures
 * Hybrid mapping system with intelligent fallback mechanisms
 */

/**
 * HYBRID MAPPING SYSTEM: Extract Turkish company data with structured mapping + Claude fallback
 * Primary extraction function for Turkish business data
 * @param {Array} row Spreadsheet row data
 * @param {Array} headers Column headers
 * @returns {Object} Structured Turkish company data with quality analysis
 */
function extractTurkishCompanyDataHybrid(row, headers) {
  Logger.log('🔄 Starting hybrid mapping system for Turkish company data');
  
  // Step 1: Use structured (code-based) mapping for Turkish business data
  const structuredData = extractTurkishCompanyDataStructured(row, headers);
  
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
      const claudeMappedData = mapTurkishHeadersWithClaude(headers, row, missingCriticalFields);
      
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
  
  // Step 6: Generate quality report for Turkish context
  const qualityReport = generateTurkishQualityReport(qualityAnalysis, headers.length);
  finalData.quality_analysis = qualityReport;
  
  Logger.log(`📈 Final mapping quality: ${Math.round(qualityAnalysis.final_data_completeness)}% completeness`);
  
  return finalData;
}

/**
 * Structured (code-based) mapping for Turkish business data
 * ADAPTED: Turkish business terminology, legal forms, and market context
 * @param {Array} row Spreadsheet row data
 * @param {Array} headers Column headers
 * @returns {Object} Structured Turkish company data
 */
function extractTurkishCompanyDataStructured(row, headers) {
  const data = {};
  
  // COMPREHENSIVE field mappings for Turkish business data
  const fieldMappings = {
    // Company name variations (Turkish context)
    'company_name': [
      'company', 'name', 'company name', 'şirket', 'firma', 'şirket adı', 'şirket adi',
      'firma adı', 'firma adi', 'ticaret unvanı', 'ticaret unvani', 'unvan',
      'kuruluş', 'kuruluş adı', 'organizasyon', 'entity', 'firm', 'organization',
      'işletme', 'işletme adı', 'kurum', 'tüzel kişi', 'anonim şirket', 'limited şirket'
    ],
    
    // Industry and business activity (Turkish)
    'industry': [
      'industry', 'sector', 'sektör', 'sektör', 'endüstri', 'faaliyet', 'iş kolu',
      'ana faaliyet', 'ana faaliyet alanı', 'iş alanı', 'branş', 'uzmanlık',
      'hizmet alanı', 'business', 'activity', 'alan', 'çalışma alanı',
      'faaliyet konusu', 'iş', 'meslek', 'çalışma sahası', 'uzmanlaşma'
    ],
    
    // Financial data (Turkish Lira context)
    'revenue_estimate': [
      'revenue', 'gelir', 'hasılat', 'ciro', 'net satışlar', 'brüt satışlar',
      'yıllık gelir', 'annual revenue', 'sales', 'income', 'earnings',
      'satış hasılatı', 'ticari gelir', 'işletme geliri', 'toplam gelir',
      'gelir tutarı', 'ciro tutarı', 'turnover', 'omzet'
    ],
    
    // Employee information (Turkish)
    'employees_estimate': [
      'employees', 'staff', 'çalışan', 'personel', 'işçi', 'memur',
      'çalışan sayısı', 'personel sayısı', 'kadro', 'workforce', 'headcount',
      'team', 'takım', 'işgücü', 'insan kaynağı', 'toplam çalışan',
      'aktif çalışan', 'kayıtlı çalışan', 'number of employees', 'employee count'
    ],
    
    // Business type and legal form (Turkish legal entities)
    'business_type': [
      'type', 'form', 'tip', 'tür', 'hukuki form', 'yasal form',
      'a.ş.', 'anonim şirket', 'limited şirket', 'ltd.', 'ltd. şti.',
      'şahıs şirketi', 'kolektif şirket', 'komandit şirket', 'kooperatif',
      'dernek', 'vakıf', 'legal form', 'şirket türü', 'kuruluş türü'
    ],
    
    // Location and address (Turkish geography)
    'headquarters': [
      'city', 'location', 'şehir', 'il', 'konum', 'adres', 'merkez',
      'genel merkez', 'ana merkez', 'merkez adresi', 'bulunduğu yer',
      'yerleşim', 'address', 'headquarters', 'head office', 'lokasyon',
      'yer', 'bölge', 'mevki', 'semt', 'mahalle'
    ],
    
    // Website and digital presence
    'website': [
      'site', 'url', 'website', 'web sitesi', 'web site', 'internet sitesi',
      'portal', 'web', 'homepage', 'ana sayfa', 'internet adresi',
      'web adresi', 'domain', 'alan adı', 'online presence'
    ],
    
    // Tax and registration data (Turkish system)
    'tax_info': [
      'vergi no', 'vergi numarası', 'vkn', 'tc kimlik', 'ticaret sicil',
      'ticaret sicil no', 'sicil numarası', 'tax', 'tax id', 'tax number',
      'registration', 'kayıt', 'kayıt numarası', 'kimlik numarası'
    ],
    
    // Contact information
    'contact_info': [
      'telefon', 'phone', 'tel', 'gsm', 'cep telefonu', 'iletişim',
      'e-posta', 'email', 'mail', 'elektronik posta', 'contact',
      'iletişim bilgileri', 'telefon numarası', 'fax', 'faks'
    ],
    
    // Leadership and management (Turkish context)
    'leadership': [
      'yönetici', 'müdür', 'genel müdür', 'ceo', 'başkan', 'ortağı',
      'sahibi', 'kurucusu', 'işletmeci', 'manager', 'head', 'chief',
      'leader', 'sorumlu', 'yetkili', 'temsilci', 'direktör'
    ],
    
    // Operational data
    'operations': [
      'şube', 'şubeler', 'temsilcilik', 'bayilik', 'distribütörlük',
      'franchise', 'franchisor', 'branches', 'subsidiaries', 'operations',
      'operasyon', 'faaliyet gösterdiği yer', 'hizmet verdiği bölge'
    ]
  };
  
  // Track which headers were successfully mapped
  const mappedHeaders = new Set();
  
  // Extract data based on comprehensive Turkish header mappings
  headers.forEach((header, index) => {
    const normalizedHeader = header.toLowerCase().trim();
    
    for (const [field, variations] of Object.entries(fieldMappings)) {
      if (variations.some(v => normalizedHeader.includes(v))) {
        // Store the raw value
        const cellValue = row[index];
        if (cellValue !== null && cellValue !== undefined && cellValue !== '') {
          data[field] = cellValue;
          mappedHeaders.add(index);
          Logger.log(`Mapped Turkish header \"${header}\" to field \"${field}\" with value:`, cellValue);
        }
        break;
      }
    }
  });
  
  // INTELLIGENT CATEGORIZATION of unmapped fields with Turkish business context
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
      
      // Categorize based on Turkish business terminology
      if (headerLower.includes('tl') || headerLower.includes('lira') || 
          headerLower.includes('gelir') || headerLower.includes('ciro') ||
          headerLower.includes('hasılat') || headerLower.includes('kar') ||
          headerLower.includes('vergi') || headerLower.includes('ödeme')) {
        businessCategories.financial_data.push(entry);
      } else if (headerLower.includes('a.ş') || headerLower.includes('ltd') ||
                 headerLower.includes('şirket') || headerLower.includes('kayıt') ||
                 headerLower.includes('sicil') || headerLower.includes('lisans') ||
                 headerLower.includes('ruhsat')) {
        businessCategories.legal_data.push(entry);
      } else if (headerLower.includes('şube') || headerLower.includes('personel') ||
                 headerLower.includes('çalışan') || headerLower.includes('operasyon') ||
                 headerLower.includes('hizmet') || headerLower.includes('üretim')) {
        businessCategories.operational_data.push(entry);
      } else if (headerLower.includes('telefon') || headerLower.includes('email') ||
                 headerLower.includes('adres') || headerLower.includes('iletişim')) {
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
  Logger.log(`Extracted Turkish company data for ${data.company_name || 'Unknown company'}:`, {
    mappedFields: Object.keys(data).filter(k => k !== 'discovered_facts'),
    categorizedFacts: Object.entries(businessCategories).map(([k,v]) => `${k}: ${v.length} items`),
    totalFactsExtracted: discoveredFacts.length
  });
  
  return data;
}

/**
 * Claude-based header mapping for Turkish business data
 * ADAPTED: Turkish business terminology and context
 * @param {Array} headers All column headers
 * @param {Array} row Corresponding row data
 * @param {Array} missingFields List of critical fields that need mapping
 * @returns {Object} Claude-mapped data for missing fields
 */
function mapTurkishHeadersWithClaude(headers, row, missingFields) {
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
    
    const prompt = `Türk iş verilerini içeren tablo başlıklarını analiz et ve eksik alanlar için eşleştirme bul.

EKSİK ALANLAR (bulunması gereken): ${missingFields.join(', ')}

MEVCUT BAŞLIKLAR VE DEĞERLER:
${unmappedHeaders.map(item => `\"${item.header}\": \"${item.value}\"`).join('\\n')}

GÖREV: Her eksik alan için en uygun başlığı yukarıdaki listeden bul.

EŞLEŞTİRME KURALLARI:
- company_name: şirket adları, firma adları, ticaret unvanları
- industry: sektörler, iş kolları, faaliyet alanları, endüstriler
- revenue_estimate: gelir, ciro, hasılat, satış tutarları (TL)
- employees_estimate: çalışan sayısı, personel sayısı, kadro
- headquarters: şehirler, iller, adresler, konumlar

SADECE JSON formatında dön:
{
  \"alan_adı\": \"çıkarılan_değer\",
  \"alan_adı2\": \"çıkarılan_değer2\"
}

Eğer alan bulunamazsa, yanıta dahil etme.`;

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
      const match = content.match(/```json\\s*([\\s\\S]*?)\\s*```/);
      if (match) {
        content = match[1];
      }
    }
    
    const claudeMappedData = JSON.parse(content);
    
    Logger.log('🤖 Claude successfully mapped Turkish headers:', Object.keys(claudeMappedData));
    return claudeMappedData;
    
  } catch (error) {
    Logger.log('❌ Claude Turkish header mapping failed:', error);
    throw error;
  }
}

/**
 * Generate quality analysis report for Turkish context
 * @param {Object} qualityAnalysis Quality metrics object
 * @param {number} totalHeaders Total number of headers in spreadsheet
 * @returns {string} Human-readable quality report in Turkish
 */
function generateTurkishQualityReport(qualityAnalysis, totalHeaders) {
  const reports = [];
  
  // Overall completeness
  const completeness = Math.round(qualityAnalysis.final_data_completeness);
  if (completeness >= 80) {
    reports.push(`✅ Mükemmel veri kalitesi (${completeness}%)`);
  } else if (completeness >= 60) {
    reports.push(`⚠️ İyi veri kalitesi (${completeness}%)`);
  } else {
    reports.push(`❌ Düşük veri kalitesi (${completeness}%)`);
  }
  
  // Structured mapping performance
  reports.push(`Yapısal eşleştirme: ${qualityAnalysis.structured_mapping_success}/${qualityAnalysis.total_critical_fields}`);
  
  // Claude fallback usage
  if (qualityAnalysis.claude_fallback_used) {
    if (qualityAnalysis.claude_mapping_success > 0) {
      reports.push(`Claude-eşleştirme: +${qualityAnalysis.claude_mapping_success} alan`);
    } else {
      reports.push(`Claude-eşleştirme: başarısız`);
    }
  }
  
  // Missing fields warning
  if (qualityAnalysis.missing_fields.length > 0) {
    reports.push(`Eksik: ${qualityAnalysis.missing_fields.join(', ')}`);
  }
  
  return reports.join(' | ');
}

/**
 * Insert scoring results into spreadsheet with Turkish labels
 * ADAPTED: Turkish column headers and color coding
 * @param {Sheet} sheet Target sheet
 * @param {Range} originalRange Original data range
 * @param {Array} results Scoring results
 * @param {Object} options Configuration options
 */
function insertResults(sheet, originalRange, results, options) {
  try {
    // Define result columns with Turkish labels
    const resultHeaders = [
      'API-Skor',
      'Sektör Çarpanı', 
      'LLM-Düzeltme',
      'Final Skor',
      'Öncelik',
      'Açıklama',
      'Veri Kalitesi'
    ];
    
    // Find next available column
    const lastCol = originalRange.getLastColumn();
    const resultStartCol = lastCol + 1;
    
    // Insert headers
    const headerRange = sheet.getRange(originalRange.getRow(), resultStartCol, 1, resultHeaders.length);
    headerRange.setValues([resultHeaders]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#E8F0FE');
    
    // Prepare result data with Turkish labels
    const resultData = results.map(result => {
      if (!result.success) {
        return [
          'HATA',
          '',
          '',
          '',
          'HATA',
          result.error || 'Bilinmeyen hata',
          result.qualityAnalysis || 'Kritik işleme hatası'
        ];
      }
      
      const scoring = result.scoringResult;
      
      // Calculate scores according to Yolwise specification
      const apiScore = scoring.industry_adjusted_score || scoring.base_score || 0;
      const llmAdjustment = Math.max(-25, Math.min(25, scoring.llm_adjustment || 0));
      const finalScore = Math.max(0, Math.min(100, apiScore + llmAdjustment));
      
      // Priority based on final score (60+ = target)
      const priority = finalScore >= 60 ? 'target' : 'non_target';
      const priorityTr = finalScore >= 60 ? 'HEDEFLENEBİLİR' : 'HEDEFLENEMEZ';
      
      return [
        Math.round(apiScore * 10) / 10,
        scoring.industry_multiplier || 1.0,
        llmAdjustment,
        Math.round(finalScore * 10) / 10,
        priorityTr,
        scoring.reasoning || 'Açıklama mevcut değil',
        result.qualityAnalysis || 'Kalite analizi mevcut değil'
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
        if (priorityStr.includes('hedeflenebilir') || priorityStr.includes('target')) {
          return ['#D4EDDA']; // Green for target
        } else if (priorityStr.includes('hedeflenemez') || priorityStr.includes('non_target')) {
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
        if (qualityStr.includes('mükemmel') || qualityStr.includes('✅')) {
          return ['#D4EDDA']; // Green for excellent quality
        } else if (qualityStr.includes('iyi') || qualityStr.includes('⚠️')) {
          return ['#FFF3CD']; // Yellow for good quality
        } else if (qualityStr.includes('düşük') || qualityStr.includes('❌')) {
          return ['#F8D7DA']; // Red for poor quality
        } else {
          return ['#F8F9FA']; // Light gray for neutral
        }
      });
      qualityRange.setBackgrounds(qualityBackgrounds);
    }
    
    Logger.log('Results inserted successfully with Turkish labels and quality analysis');
    
  } catch (error) {
    Logger.log('Error inserting results:', error);
    throw error;
  }
}