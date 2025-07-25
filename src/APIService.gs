/**
 * API Service for Claude 4 Sonnet and Yolwise integration
 * Handles external API calls and data processing for Turkish B2B market
 * ADAPTED FROM SMARTWAY: Turkish business context, Yolwise API integration
 * Enhanced Claude prompt for comprehensive Turkish business data analysis
 * Hybrid system integration with quality analysis
 */

/**
 * Analyze Turkish company data with Claude 4 Sonnet API
 * ADAPTED: Turkish business context and terminology
 * @param {Object} companyData Raw company data from spreadsheet
 * @param {Object} options Configuration options
 * @returns {Object} Analyzed and structured data for scoring
 */
function analyzeWithClaude(companyData, options) {
  try {
    Logger.log('Analyzing Turkish company with Claude:', companyData.company_name);
    
    // Get API key from script properties
    const apiKey = getClaudeApiKey();
    if (!apiKey) {
      throw new Error('Claude API key not configured. Please set CLAUDE_API_KEY in script properties.');
    }
    
    // Prepare the analysis prompt for Turkish B2B context
    const prompt = buildTurkishClaudePrompt(companyData, options);
    
    // Make API call to Claude
    const response = callClaudeAPI(prompt, apiKey);
    
    // Parse and structure the response
    const analysis = parseClaudeResponse(response);
    
    Logger.log('Claude analysis completed for Turkish company:', companyData.company_name);
    return analysis;
    
  } catch (error) {
    Logger.log('Error in Claude analysis:', error);
    throw error;
  }
}

/**
 * Build comprehensive analysis prompt for Claude API with Turkish B2B context
 * ADAPTED: Turkish business terminology and market context
 * @param {Object} companyData Company information from hybrid mapping system
 * @param {Object} options Configuration options
 * @returns {String} Formatted prompt optimized for Turkish B2B data
 */
function buildTurkishClaudePrompt(companyData, options) {
  const prompt = `
Sen Türk B2B pazarı için Yolwise Lead Scoring sistemi uzmanısın. Hibrit veri çıkarma sistemi ile elde edilen şirket verilerini analiz ediyorsun.

**ÖNEMLİ**: Veriler gelişmiş hibrit çıkarma sistemi (yapısal eşleştirme + AI analiz) ile elde edildi. Bu sistem Türk iş terminolojisine özel olarak optimize edilmiş.

**ŞİRKET VERİLERİ:**
${JSON.stringify(companyData, null, 2)}

**VERİ KALİTESİ:**
${companyData.quality_analysis || 'Kalite analizi mevcut değil'}

**TÜM MEVCUT ALANLARIN ANALİZİ:**
${Object.entries(companyData).map(([key, value]) => {
  if (key === 'discovered_facts' && Array.isArray(value)) {
    return `${key}: ${value.join(' | ')}`;
  }
  if (key === 'quality_analysis') {
    return `Veri kalitesi: ${value}`;
  }
  return `${key}: ${value}`;
}).join('\\n')}

**TÜRK B2B PAZARI İÇİN HİBRİT SİSTEM ANALİZ ÖZELLİKLERİ:**

1. **Veri kalitesini dikkate al**: Eğer veri kalitesi düşükse (quality_analysis'te belirtilmiş), sonuçlarda daha temkinli ol.

2. **TÜM mevcut verileri kullan**: Hibrit sistem standart olmayan başlıklardan da Claude-analiz ile veri çıkarmış olabilir - bunları tam olarak değerlendir.

3. **discovered_facts'e özel dikkat**: Bu veriler Türk iş terminolojisine göre kategorize edilmiş önemli bilgiler içerebilir.

4. **Türk B2B hizmet ihtiyaçlarına odaklan**: Türkiye'deki B2B hizmet pazarı özelliklerini göz önünde bulundur.

**TÜRK B2B PAZARI BAĞLAMI:**
- Özellikle finans, inşaat, imalat, lojistik sektörlerinde yüksek B2B hizmet ihtiyacı
- Büyük şirketler (500+ çalışan) daha fazla B2B hizmet satın alıyor
- İstanbul, Ankara, İzmir gibi büyük şehirlerde yoğunlaşan iş aktivitesi
- A.Ş. ve büyük Ltd. Şti. formlarında daha yüksek bütçe potansiyeli

**GÖREVİN:**
1. TÜM mevcut alanları ve kategorileri veri kalitesini göz önünde bulundurarak analiz et
2. Türk iş verilerine dayanarak sektörü belirle
3. Finansal göstergeler ve çalışan sayısına göre şirket büyüklüğünü değerlendir
4. B2B hizmet ihtiyacı potansiyelini ve karar verme karmaşıklığını analiz et
5. Büyüme, gelişme veya sorun belirtilerini tespit et
6. Veri kalitesinin sonuçlara etkisini değerlendir

**TÜRK HİBRİT VERİ SİSTEMİ PRENSİPLERİ:**
- Eğer veri kalitesi yüksekse (80%+) → güvenli sonuçlar çıkar
- Eğer veri kalitesi orta seviyedeyse (60-79%) → temkinli sonuçlar, uyarılarla
- Eğer veri kalitesi düşükse (<60%) → minimum sonuçlar, mevcut gerçeklere odaklan
- Tüm kategorileri kullan: financial_data, legal_data, operational_data, contact_data
- Claude-eşleştirme ile elde edilen verilere özel dikkat - bunlar özenle seçilmiş
- Türk iş kültürü ve pazar dinamiklerini göz önünde bulundur

**YANIT FORMATI (kesinlikle JSON):**
{
  "company_name": "verilerdeki tam ad",
  "industry": "tüm mevcut verilere dayalı sektör",
  "revenue_estimate": sayısal_değer_milyon_TL_veya_0,
  "employees_estimate": "tüm verilere dayalı sayı veya aralık",
  "business_type": "A.Ş./Ltd.Şti./diğer türk hukuki formları",
  "headquarters": "adres verilerinden şehir",
  "locations": ["bulunan tüm bölge ve şehirler"],
  "key_people": ["verilerdeki yöneticiler"],
  "discovered_facts": ["kategorize edilmiş verilerden tüm önemli gerçekler"],
  "growth_indicators": "büyüme/gelişme/sorun belirtileri tüm kaynaklardan",
  "market_position": "boyut ve aktiviteye dayalı pazar pozisyonu",
  "business_context": "iş özellikleri, lisanslar, faaliyet detayları",
  "b2b_service_potential": "türk b2b pazarında hizmet potansiyeli: sektör + boyut + coğrafya + özellikler",
  "analysis_confidence": "yüksek/orta/düşük veri kalitesi ve tamlığına göre",
  "data_quality_impact": "veri kalitesinin analize etkisi",
  "turkish_market_notes": "türk b2b pazarı özelinde önemli notlar"
}

**TÜRK B2B PAZARI İÇİN KALİTELİ ANALİZ PRENSİPLERİ:**
- Hibrit sistemden gelen tüm alanları entegre et, Claude-eşleştirme verilerini özellikle değerlendir
- Finansal, operasyonel ve hukuki veriler arasında bağlantı kur
- B2B hizmet potansiyelini türk pazarındaki faktörlere göre değerlendir: sektör + boyut + coğrafya + özellikler
- Veri kalitesine dayalı güven seviyesi belirt
- Analiz edilen verilerin hangi yönlerinin yüksek/düşük kaliteli olduğunu not et

SADECE geçerli JSON döndür, ek metin yok.`;

  return prompt;
}

/**
 * Call Claude 4 Sonnet API
 * @param {String} prompt Analysis prompt
 * @param {String} apiKey API key
 * @returns {Object} API response
 */
function callClaudeAPI(prompt, apiKey) {
  const url = 'https://api.anthropic.com/v1/messages';
  
  const payload = {
    model: 'claude-sonnet-4-20250514',
    max_tokens: 4000,
    temperature: 0.1,
    messages: [{
      role: 'user',
      content: prompt
    }]
  };
  
  const options = {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(payload)
  };
  
  Logger.log('Calling Claude API for Turkish analysis...');
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      throw new Error(`Claude API error: ${responseCode} - ${response.getContentText()}`);
    }
    
    const responseData = JSON.parse(response.getContentText());
    Logger.log('Claude API response received');
    
    return responseData;
    
  } catch (error) {
    Logger.log('Claude API call failed:', error);
    throw error;
  }
}

/**
 * Parse Claude API response with enhanced error handling
 * @param {Object} response API response
 * @returns {Object} Parsed analysis
 */
function parseClaudeResponse(response) {
  try {
    if (!response.content || !response.content[0] || !response.content[0].text) {
      throw new Error('Invalid Claude API response format');
    }
    
    const content = response.content[0].text.trim();
    
    // Extract JSON from response (handle potential markdown formatting)
    let jsonStr = content;
    if (content.includes('```json')) {
      const match = content.match(/```json\\s*([\\s\\S]*?)\\s*```/);
      if (match) {
        jsonStr = match[1];
      }
    }
    
    const analysis = JSON.parse(jsonStr);
    
    // Validate required fields
    if (!analysis.company_name) {
      throw new Error('Company name not found in Claude analysis');
    }
    
    // Validate Turkish-specific fields
    if (analysis.turkish_market_notes) {
      Logger.log('Turkish market analysis notes:', analysis.turkish_market_notes);
    }
    
    Logger.log('Claude response parsed successfully for Turkish company');
    return analysis;
    
  } catch (error) {
    Logger.log('Error parsing Claude response:', error);
    Logger.log('Response content:', response);
    throw new Error('Failed to parse Claude analysis: ' + error.message);
  }
}

/**
 * Score company with Yolwise API
 * ADAPTED: Yolwise API integration with API key protection
 * @param {Object} claudeAnalysis Analyzed company data
 * @returns {Object} Scoring results with proper API format
 */
function scoreWithYolwise(claudeAnalysis) {
  try {
    Logger.log('Scoring with Yolwise API:', claudeAnalysis.company_name);
    
    // Get Yolwise API configuration
    const apiUrl = getYolwiseApiUrl();
    const apiKey = getYolwiseApiKey();
    
    if (!apiUrl) {
      Logger.log('Yolwise API URL not configured, using mock scoring');
      return createYolwiseMockScoring(claudeAnalysis);
    }
    
    // Prepare payload for Yolwise API according to specification
    const payload = {
      company_name: claudeAnalysis.company_name,
      company_data: {
        industry: claudeAnalysis.industry || '',
        revenue_estimate: claudeAnalysis.revenue_estimate || 0,
        employees_estimate: claudeAnalysis.employees_estimate || '',
        business_type: claudeAnalysis.business_type || '',
        headquarters: claudeAnalysis.headquarters || '',
        locations: claudeAnalysis.locations || [],
        key_people: claudeAnalysis.key_people || [],
        discovered_facts: claudeAnalysis.discovered_facts || [],
        growth_indicators: claudeAnalysis.growth_indicators || '',
        market_position: claudeAnalysis.market_position || '',
        // Yolwise-specific fields
        data_quality_impact: claudeAnalysis.data_quality_impact || '',
        turkish_market_notes: claudeAnalysis.turkish_market_notes || '',
        analysis_confidence: claudeAnalysis.analysis_confidence || 'orta',
        b2b_service_potential: claudeAnalysis.b2b_service_potential || ''
      }
    };
    
    const options = {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload)
    };
    
    // Add API key authentication if available
    if (apiKey) {
      options.headers['Authorization'] = `Bearer ${apiKey}`;
      // Or use X-API-Key header depending on Yolwise API design
      options.headers['X-API-Key'] = apiKey;
    }
    
    try {
      const response = UrlFetchApp.fetch(`${apiUrl}/score_company`, options);
      const responseCode = response.getResponseCode();
      
      if (responseCode === 401) {
        throw new Error('Yolwise API: Unauthorized - Check API key');
      } else if (responseCode === 403) {
        throw new Error('Yolwise API: Forbidden - API key may be invalid');
      } else if (responseCode !== 200) {
        Logger.log(`Yolwise API returned ${responseCode}, falling back to mock scoring`);
        return createYolwiseMockScoring(claudeAnalysis);
      }
      
      const apiResult = JSON.parse(response.getContentText());
      
      if (!apiResult.success || !apiResult.result) {
        Logger.log('Yolwise API returned unsuccessful response, using mock scoring');
        return createYolwiseMockScoring(claudeAnalysis);
      }
      
      // Process successful API response
      const baseResult = apiResult.result;
      
      // Apply LLM adjustment based on Claude analysis for Turkish context
      const finalResult = applyTurkishLLMAdjustment(baseResult, claudeAnalysis);
      
      Logger.log('Yolwise API scoring completed successfully');
      return finalResult;
      
    } catch (apiError) {
      Logger.log('Yolwise API call failed, using mock scoring:', apiError);
      return createYolwiseMockScoring(claudeAnalysis);
    }
    
  } catch (error) {
    Logger.log('Error in Yolwise scoring process:', error);
    return createYolwiseMockScoring(claudeAnalysis);
  }
}

/**
 * Apply LLM adjustment to API scoring results with Turkish B2B context
 * ADAPTED: Turkish market factors and business context
 * @param {Object} apiResult Base scoring result from Yolwise API
 * @param {Object} claudeAnalysis Claude analysis data
 * @returns {Object} Result with proper LLM adjustment applied
 */
function applyTurkishLLMAdjustment(apiResult, claudeAnalysis) {
  Logger.log('Applying Turkish B2B LLM adjustment to API result:', apiResult.company_name);
  
  // Individual adjustment factors with point values
  const adjustments = [];
  const reasoning = [];
  
  // Base information from API
  reasoning.push(`API base: ${apiResult.base_score}, sektör: ${apiResult.detected_industry} (×${apiResult.industry_multiplier})`);
  reasoning.push(apiResult.industry_explanation || 'Standart sektörel değerlendirme');
  
  // Turkish market confidence adjustments
  const analysisConfidence = claudeAnalysis.analysis_confidence || 'orta';
  if (analysisConfidence === 'düşük') {
    adjustments.push({ value: -5, reason: 'Düşük veri güvenilirliği türk pazarı değerlendirmesini etkiliyor' });
  } else if (analysisConfidence === 'yüksek') {
    adjustments.push({ value: 3, reason: 'Yüksek veri güvenilirliği türk pazarı analizini güçlendiriyor' });
  }
  
  // Turkish market data quality bonus
  if (claudeAnalysis.data_quality_impact && claudeAnalysis.data_quality_impact.includes('yüksek kalite')) {
    adjustments.push({ value: 5, reason: 'Yüksek kaliteli türk iş verisi' });
  }
  
  // Turkish B2B service potential assessment
  if (claudeAnalysis.b2b_service_potential) {
    const potentialText = claudeAnalysis.b2b_service_potential.toLowerCase();
    
    if (potentialText.includes('yüksek') || potentialText.includes('güçlü') || potentialText.includes('büyük')) {
      adjustments.push({ value: 12, reason: 'Türk B2B pazarında yüksek hizmet potansiyeli' });
    } else if (potentialText.includes('düşük') || potentialText.includes('sınırlı') || potentialText.includes('zayıf')) {
      adjustments.push({ value: -8, reason: 'Türk B2B pazarında düşük hizmet potansiyeli' });
    } else if (potentialText.includes('orta') || potentialText.includes('ılımlı')) {
      adjustments.push({ value: 4, reason: 'Türk B2B pazarında orta düzey hizmet potansiyeli' });
    }
  }
  
  // Turkish business context analysis
  if (claudeAnalysis.business_context) {
    const contextText = claudeAnalysis.business_context.toLowerCase();
    
    // Positive Turkish business indicators
    if (contextText.includes('büyüme') || contextText.includes('genişleme') || 
        contextText.includes('yeni yatırım') || contextText.includes('ihracat')) {
      adjustments.push({ value: 10, reason: 'Türk şirketinde aktif büyüme ve genişleme' });
    }
    
    // Turkish export/international business
    if (contextText.includes('ihracat') || contextText.includes('uluslararası') || 
        contextText.includes('avrupa') || contextText.includes('export')) {
      adjustments.push({ value: 8, reason: 'Uluslararası iş faaliyetleri' });
    }
    
    // Turkish manufacturing/production indicators
    if (contextText.includes('üretim') || contextText.includes('imalat') || 
        contextText.includes('fabrika') || contextText.includes('tesis')) {
      adjustments.push({ value: 6, reason: 'Türk imalat sektöründe güçlü pozisyon' });
    }
    
    // Negative business indicators
    if (contextText.includes('daralmaya') || contextText.includes('kriz') || 
        contextText.includes('kapanma') || contextText.includes('zarar')) {
      adjustments.push({ value: -12, reason: 'Negatif iş trendleri' });
    }
    
    // Competition concerns specific to Turkish market
    if (contextText.includes('rekabet') && (contextText.includes('zorlanma') || contextText.includes('baskı'))) {
      adjustments.push({ value: -6, reason: 'Türk pazarında rekabet baskısı' });
    }
  }
  
  // Growth indicators assessment
  if (claudeAnalysis.growth_indicators) {
    const growthText = claudeAnalysis.growth_indicators.toLowerCase();
    
    if (growthText.includes('büyüme') || growthText.includes('gelişme') || growthText.includes('artış')) {
      adjustments.push({ value: 8, reason: 'Pozitif büyüme göstergeleri' });
    }
    
    if (growthText.includes('düşüş') || growthText.includes('azalma') || growthText.includes('daralma')) {
      adjustments.push({ value: -6, reason: 'Negatif büyüme göstergeleri' });
    }
  }
  
  // Market position considerations in Turkish context
  if (claudeAnalysis.market_position) {
    const positionText = claudeAnalysis.market_position.toLowerCase();
    
    if (positionText.includes('lider') || positionText.includes('önder') || positionText.includes('büyük')) {
      adjustments.push({ value: 10, reason: 'Türk pazarında lider pozisyon' });
    } else if (positionText.includes('küçük') || positionText.includes('yerel') || positionText.includes('bölgesel')) {
      adjustments.push({ value: -4, reason: 'Sınırlı pazar pozisyonu' });
    }
  }
  
  // Turkish market specific adjustments
  if (claudeAnalysis.turkish_market_notes) {
    const marketNotes = claudeAnalysis.turkish_market_notes.toLowerCase();
    
    if (marketNotes.includes('istanbul') || marketNotes.includes('ankara') || marketNotes.includes('izmir')) {
      adjustments.push({ value: 5, reason: 'Büyük şehir merkezi - yüksek B2B aktivite' });
    }
    
    if (marketNotes.includes('organize sanayi') || marketNotes.includes('teknokent') || marketNotes.includes('endüstri bölgesi')) {
      adjustments.push({ value: 4, reason: 'Endüstriyel bölge konumu - B2B hizmet ihtiyacı' });
    }
    
    if (marketNotes.includes('aile şirketi') && marketNotes.includes('kurumsallaşma')) {
      adjustments.push({ value: 6, reason: 'Kurumsallaşan aile şirketi - hizmet dönüşümü' });
    }
  }
  
  // Apply adjustments with strict ±25 limit enforcement (per specification)
  let totalAdjustment = 0;
  const appliedAdjustments = [];
  
  // Sort adjustments by absolute value (largest impact first)
  adjustments.sort((a, b) => Math.abs(b.value) - Math.abs(a.value));
  
  // Apply adjustments within the ±25 limit
  for (const adj of adjustments) {
    const newTotal = totalAdjustment + adj.value;
    
    // Check if the adjustment would exceed limits
    if (newTotal > 25) {
      const remaining = 25 - totalAdjustment;
      if (remaining > 0) {
        totalAdjustment += remaining;
        appliedAdjustments.push({ value: remaining, reason: adj.reason + ' (sınır: +25)' });
        reasoning.push(`LLM: +${remaining} - ${adj.reason} (kısmi)`);
      }
      break;
    } else if (newTotal < -25) {
      const remaining = -25 - totalAdjustment;
      if (remaining < 0) {
        totalAdjustment += remaining;
        appliedAdjustments.push({ value: remaining, reason: adj.reason + ' (sınır: -25)' });
        reasoning.push(`LLM: ${remaining} - ${adj.reason} (kısmi)`);
      }
      break;
    } else {
      totalAdjustment += adj.value;
      appliedAdjustments.push(adj);
      reasoning.push(`LLM: ${adj.value >= 0 ? '+' : ''}${adj.value} - ${adj.reason}`);
    }
  }
  
  // Calculate final score (API already provided industry_adjusted_score)
  const finalScore = Math.max(0, Math.min(100, apiResult.industry_adjusted_score + totalAdjustment));
  
  // Determine final priority (60+ = target as per Yolwise specification)
  const finalPriority = finalScore >= 60 ? 'target' : 'non_target';
  const finalPriorityTr = finalScore >= 60 ? 'HEDEFLENEBİLİR' : 'HEDEFLENEMEZ';
  
  Logger.log(`Turkish LLM adjustment applied: ${totalAdjustment}, Final score: ${finalScore}`);
  
  return {
    ...apiResult,
    llm_adjustment: totalAdjustment,
    final_score: finalScore,
    priority_recommendation: finalPriority,
    priority_recommendation_tr: finalPriorityTr,
    reasoning: reasoning.join(' | '),
    applied_adjustments: appliedAdjustments,
    claude_business_context: claudeAnalysis.business_context || 'İş bağlamı belirsiz',
    claude_confidence: claudeAnalysis.analysis_confidence || 'orta',
    // Turkish-specific fields
    data_quality_impact: claudeAnalysis.data_quality_impact || 'Veri kalitesi bilgisi yok',
    turkish_market_notes: claudeAnalysis.turkish_market_notes || 'Türk pazarı analizi yok',
    b2b_service_potential: claudeAnalysis.b2b_service_potential || 'B2B hizmet potansiyeli belirsiz'
  };
}

/**
 * Create mock scoring for development/fallback with Turkish context
 * ADAPTED: Yolwise specification and Turkish market factors
 * @param {Object} claudeAnalysis Claude analysis data
 * @returns {Object} Mock scoring result matching Yolwise API format
 */
function createYolwiseMockScoring(claudeAnalysis) {
  Logger.log('Creating Yolwise mock scoring for Turkish company:', claudeAnalysis.company_name);
  
  // Industry modifiers based on Yolwise specification for Turkish B2B market
  const turkishIndustryModifiers = {
    'finans': { multiplier: 1.20, confidence: 'high', reasoning: 'Yüksek B2B hizmet ihtiyacı' },
    'lojistik': { multiplier: 1.17, confidence: 'high', reasoning: 'Karmaşık operasyonel B2B ihtiyaçları' },
    'kamu_hizmetleri': { multiplier: 1.15, confidence: 'high', reasoning: 'Altyapı-ağır B2B gereksinimleri' },
    'gıda_içecek': { multiplier: 1.10, confidence: 'medium', reasoning: 'Orta düzey B2B hizmet gereksinimleri' },
    'kimya': { multiplier: 1.05, confidence: 'medium', reasoning: 'Özel teknik hizmetler gerekli' },
    'inşaat_malzeme': { multiplier: 1.02, confidence: 'medium', reasoning: 'İnşaat ile ilgili B2B hizmetler' },
    'makine_sanayi': { multiplier: 1.00, confidence: 'medium', reasoning: 'Baseline mühendislik hizmet ihtiyacı' },
    'perakende': { multiplier: 0.90, confidence: 'low', reasoning: 'Tüketici odaklı, sınırlı B2B ihtiyaç' },
    'inşaat': { multiplier: 0.88, confidence: 'low', reasoning: 'Proje bazlı hizmet gereksinimleri' },
    'otomotiv': { multiplier: 0.85, confidence: 'low', reasoning: 'İmalat odaklı operasyonlar' },
    'bilgisayar_yazılım': { multiplier: 0.80, confidence: 'low', reasoning: 'Self-servis dijital çözümler tercih' },
    'sağlık': { multiplier: 0.75, confidence: 'low', reasoning: 'Özel hizmet gereksinimleri' }
  };
  
  // Convert Claude analysis to base score with Turkish context
  let baseScore = 50; // Default
  
  // Adjust base score based on analysis confidence and Turkish factors
  const analysisConfidence = claudeAnalysis.analysis_confidence || 'orta';
  if (analysisConfidence === 'düşük') {
    baseScore = Math.max(20, baseScore - 15);
  } else if (analysisConfidence === 'yüksek') {
    baseScore = Math.min(85, baseScore + 10);
  }
  
  // Adjust based on B2B service potential
  if (claudeAnalysis.b2b_service_potential) {
    const potentialText = claudeAnalysis.b2b_service_potential.toLowerCase();
    if (potentialText.includes('yüksek')) {
      baseScore += 15;
    } else if (potentialText.includes('düşük')) {
      baseScore -= 10;
    } else if (potentialText.includes('orta')) {
      baseScore += 5;
    }
  }
  
  // Industry detection for Turkish market
  const industryText = (claudeAnalysis.industry || '').toLowerCase();
  let detectedIndustry = 'other';
  let industryMultiplier = 1.0;
  let industryReasoning = 'Standart değerlendirme';
  
  const turkishIndustryKeywords = {
    'finans': ['banka', 'finans', 'sigorta', 'danışman', 'mali müşavir', 'muhasebe'],
    'lojistik': ['lojistik', 'nakliye', 'taşımacılık', 'kargo', 'depolama', 'dağıtım'],
    'kamu_hizmetleri': ['elektrik', 'su', 'doğalgaz', 'enerji', 'belediye', 'kamu'],
    'gıda_içecek': ['gıda', 'içecek', 'süt', 'tarım', 'beslenme', 'çiftlik'],
    'kimya': ['kimya', 'ilaç', 'biyoteknoloji', 'laboratuvar', 'kimyasal'],
    'inşaat_malzeme': ['inşaat malzemesi', 'çimento', 'çelik', 'beton', 'yapı'],
    'makine_sanayi': ['makine', 'sanayi mühendisliği', 'makina', 'ekipman', 'imalat'],
    'perakende': ['perakende', 'tüketici', 'alışveriş', 'mağaza', 'ticaret', 'satış'],
    'inşaat': ['inşaat', 'yapı', 'mimarlık', 'müteahhit', 'altyapı'],
    'otomotiv': ['otomotiv', 'araç', 'taşıt', 'motor', 'ulaşım ekipmanı'],
    'bilgisayar_yazılım': ['yazılım', 'teknoloji', 'bilgisayar', 'IT', 'dijital', 'programlama'],
    'sağlık': ['hastane', 'sağlık', 'tıp', 'klinik', 'ilaç', 'sağlık hizmetleri']
  };
  
  for (const [industry, keywords] of Object.entries(turkishIndustryKeywords)) {
    if (keywords.some(keyword => industryText.includes(keyword))) {
      detectedIndustry = industry;
      const industryData = turkishIndustryModifiers[industry];
      industryMultiplier = industryData.multiplier;
      industryReasoning = industryData.reasoning;
      break;
    }
  }
  
  // Calculate industry-adjusted score
  const industryAdjustedScore = Math.max(0, Math.min(100, Math.round(baseScore * industryMultiplier)));
  
  // Create mock result for Turkish context
  const mockResult = {
    company_name: claudeAnalysis.company_name,
    base_score: baseScore,
    industry_multiplier: Math.round(industryMultiplier * 100) / 100,
    industry_adjusted_score: industryAdjustedScore,
    detected_industry: detectedIndustry,
    industry_confidence: turkishIndustryModifiers[detectedIndustry]?.confidence || 'low',
    industry_explanation: `Türk sektörü: ${detectedIndustry}, çarpan: ×${industryMultiplier}, ${industryReasoning}`,
    processing_time_ms: 150,
    priority_recommendation: industryAdjustedScore >= 60 ? 'target' : 'non_target'
  };
  
  // Apply Turkish LLM adjustment
  return applyTurkishLLMAdjustment(mockResult, claudeAnalysis);
}

/**
 * Get Claude API key from script properties
 * @returns {String} API key or null
 */
function getClaudeApiKey() {
  const properties = PropertiesService.getScriptProperties();
  return properties.getProperty('CLAUDE_API_KEY');
}

/**
 * Get Yolwise API URL from script properties
 * @returns {String} API URL or default
 */
function getYolwiseApiUrl() {
  const properties = PropertiesService.getScriptProperties();
  return properties.getProperty('YOLWISE_API_URL') || 'https://yolwiseleadscoring.replit.app';
}

/**
 * Get Yolwise API key from script properties
 * @returns {String} API key or null
 */
function getYolwiseApiKey() {
  const properties = PropertiesService.getScriptProperties();
  return properties.getProperty('YOLWISE_API_KEY');
}

/**
 * Set API keys for Yolwise (for setup)
 * @param {String} claudeKey Claude API key
 * @param {String} yolwiseKey Yolwise API key
 * @param {String} yolwiseUrl Yolwise API URL (optional)
 */
function setYolwiseApiKeys(claudeKey, yolwiseKey, yolwiseUrl) {
  const properties = PropertiesService.getScriptProperties();
  
  if (claudeKey) {
    properties.setProperty('CLAUDE_API_KEY', claudeKey);
  }
  
  if (yolwiseKey) {
    properties.setProperty('YOLWISE_API_KEY', yolwiseKey);
  }
  
  if (yolwiseUrl) {
    properties.setProperty('YOLWISE_API_URL', yolwiseUrl);
  } else {
    properties.setProperty('YOLWISE_API_URL', 'https://yolwiseleadscoring.replit.app');
  }
  
  Logger.log('Yolwise API keys updated successfully');
}

/**
 * Test API connections for Yolwise setup
 */
function testYolwiseApiConnections() {
  try {
    // Test Claude API
    const claudeKey = getClaudeApiKey();
    if (claudeKey) {
      Logger.log('Claude API key configured: ✓');
    } else {
      Logger.log('Claude API key missing: ✗');
    }
    
    // Test Yolwise API
    const yolwiseKey = getYolwiseApiKey();
    const yolwiseUrl = getYolwiseApiUrl();
    
    if (yolwiseKey) {
      Logger.log('Yolwise API key configured: ✓');
    } else {
      Logger.log('Yolwise API key missing: ✗');
    }
    
    if (yolwiseUrl) {
      Logger.log('Yolwise API URL configured: ✓ -', yolwiseUrl);
      
      // Test API health endpoint
      try {
        const response = UrlFetchApp.fetch(`${yolwiseUrl}/health`);
        if (response.getResponseCode() === 200) {
          Logger.log('Yolwise API health check: ✓');
        } else {
          Logger.log('Yolwise API health check: ✗ -', response.getResponseCode());
        }
      } catch (healthError) {
        Logger.log('Yolwise API health check failed:', healthError);
      }
    } else {
      Logger.log('Yolwise API URL missing: ✗');
    }
    
  } catch (error) {
    Logger.log('API test error:', error);
  }
}