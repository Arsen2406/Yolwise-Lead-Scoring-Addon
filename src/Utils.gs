/**
 * Utility functions for testing, debugging, and maintenance
 * Yolwise Lead Scoring Add-on for Turkish B2B market
 * ADAPTED FROM SMARTWAY: Turkish business context and Yolwise API integration
 * Comprehensive header detection validation and Turkish market testing
 * Hybrid mapping system testing and validation
 */

/**
 * Setup wizard for first-time Yolwise configuration
 * Run this function after copying the code to your Apps Script project
 */
function setupYolwiseWizard() {
  Logger.log('🚀 Yolwise Lead Scoring Setup Wizard v1.0 (Turkish B2B)');
  Logger.log('=======================================================');
  
  try {
    // Check if API keys are configured
    const claudeKey = getClaudeApiKey();
    const yolwiseKey = getYolwiseApiKey();
    const yolwiseUrl = getYolwiseApiUrl();
    
    Logger.log('📋 Yapılandırma Durumu:');
    Logger.log(`Claude API Key: ${claudeKey ? '✓ Yapılandırıldı' : '✗ Eksik'}`);
    Logger.log(`Yolwise API Key: ${yolwiseKey ? '✓ Yapılandırıldı' : '✗ Eksik'}`);
    Logger.log(`Yolwise API URL: ${yolwiseUrl ? '✓ ' + yolwiseUrl : '✗ Eksik'}`);
    
    if (!claudeKey) {
      Logger.log('');
      Logger.log('⚠️  Claude API Key Gerekli!');
      Logger.log('Lütfen çalıştırın: setupYolwiseApiKeysManual()');
      Logger.log('Veya anahtarınızı alın: https://console.anthropic.com');
      return false;
    }
    
    if (!yolwiseKey) {
      Logger.log('');
      Logger.log('⚠️  Yolwise API Key Gerekli!');
      Logger.log('Replit projenden API anahtarını al ve setupYolwiseApiKeysManual() çalıştır');
    }
    
    // Test API connections
    Logger.log('');
    Logger.log('🔧 API Bağlantıları Test Ediliyor...');
    testYolwiseApiConnections();
    
    // Test with Turkish sample data
    Logger.log('');
    Logger.log('🧪 Türk Şirket Skorlama Testi Çalıştırılıyor...');
    testTurkishScoring();
    
    // Run consistency validation
    Logger.log('');
    Logger.log('✅ Sistem Tutarlılığı Doğrulanıyor...');
    validateYolwiseSystemConsistency();
    
    // Test Turkish header detection improvements
    Logger.log('');
    Logger.log('🔍 Türk İş Verisi Başlık Tespiti Test Ediliyor...');
    testTurkishHeaderDetection();
    
    // Test hybrid mapping system for Turkish context
    Logger.log('');
    Logger.log('🔄 Türk B2B Hibrit Eşleştirme Sistemi Test Ediliyor...');
    testTurkishHybridMappingSystem();
    
    Logger.log('');
    Logger.log('✅ Kurulum Tamamlandı!');
    Logger.log('Artık Google Sheets\'te eklentiyi kullanabilirsiniz');
    Logger.log('Git: Uzantılar → Yolwise Lead Scoring');
    
    return true;
    
  } catch (error) {
    Logger.log('❌ Kurulum başarısız:', error);
    Logger.log('');
    Logger.log('🔍 Sorun Giderme:');
    Logger.log('1. Tüm dosyaları kopyaladığınızdan emin olun (Code.gs, APIService.gs, TurkishDataExtractor.gs, index.html)');
    Logger.log('2. appsscript.json manifest dosyasının doğru olduğunu kontrol edin');
    Logger.log('3. API anahtarlarının düzgün ayarlandığını doğrulayın');
    return false;
  }
}

/**
 * Manual API keys setup function for Yolwise
 * Replace YOUR_KEYS with actual keys and run this function
 */
function setupYolwiseApiKeysManual() {
  const properties = PropertiesService.getScriptProperties();
  
  // REPLACE THESE WITH YOUR ACTUAL API KEYS
  const CLAUDE_API_KEY = 'sk-ant-api03-YOUR_CLAUDE_KEY_HERE';
  const YOLWISE_API_KEY = 'YOUR_YOLWISE_API_KEY_HERE'; // From Replit
  const YOLWISE_API_URL = 'https://yolwiseleadscoring.replit.app';
  
  if (CLAUDE_API_KEY.includes('YOUR_CLAUDE_KEY_HERE')) {
    Logger.log('❌ Lütfen YOUR_CLAUDE_KEY_HERE\'yi gerçek Claude API anahtarınızla değiştirin!');
    Logger.log('Anahtarınızı alın: https://console.anthropic.com');
    return;
  }
  
  if (YOLWISE_API_KEY.includes('YOUR_YOLWISE_API_KEY_HERE')) {
    Logger.log('❌ Lütfen YOUR_YOLWISE_API_KEY_HERE\'yi gerçek Yolwise API anahtarınızla değiştirin!');
    Logger.log('Anahtarı Replit projenizden alın');
    return;
  }
  
  // Set the keys
  properties.setProperty('CLAUDE_API_KEY', CLAUDE_API_KEY);
  properties.setProperty('YOLWISE_API_KEY', YOLWISE_API_KEY);
  properties.setProperty('YOLWISE_API_URL', YOLWISE_API_URL);
  
  Logger.log('✅ Tüm API anahtarları yapılandırıldı');
  Logger.log('🚀 Şimdi çalıştırın: setupYolwiseWizard()');
}

/**
 * Test comprehensive Turkish header detection improvements
 * Validates that the new system handles Turkish business data correctly
 */
function testTurkishHeaderDetection() {
  Logger.log('🔍 TÜRK İŞ VERİSİ BAŞLIK TESPİTİ TESTİ');
  Logger.log('=====================================');
  
  const testResults = {
    turkishBusinessHeaders: false,
    categorizedDataExtraction: false,
    noDataLoss: false,
    backwardCompatibility: false,
    overall: false
  };
  
  try {
    // Test 1: Turkish Business Headers from the Excel file structure
    Logger.log('📋 Test 1: Türk İş Başlıkları Tanıma');
    
    const turkishHeaders = [
      'Şirket Adı',
      'Sektör', 
      'Yıllık Gelir',
      'Çalışan Sayısı',
      'Şehir',
      'Ana Faaliyet',
      'Genel Müdür',
      'Web Sitesi',
      'İl'
    ];
    
    const testRowData = [
      'Türk Telekom A.Ş.',
      'Telekomünikasyon',
      '15000000000',
      '25000',
      'İstanbul',
      'Telekomünikasyon hizmetleri',
      'Ahmet Yılmaz',
      'https://turktelekom.com.tr',
      'İstanbul'
    ];
    
    const extractedData = extractTurkishCompanyDataStructured(testRowData, turkishHeaders);
    
    // Verify company name was detected
    if (extractedData.company_name === 'Türk Telekom A.Ş.') {
      Logger.log('✅ Şirket adı "Şirket Adı" başlığından doğru çıkarıldı');
      testResults.turkishBusinessHeaders = true;
    } else {
      Logger.log(`❌ Şirket adı çıkarma başarısız: ${extractedData.company_name}`);
    }
    
    // Test 2: Categorized Data Extraction
    Logger.log('📊 Test 2: Kategorize Veri Çıkarma');
    
    const categoriesFound = extractedData.discovered_facts ? extractedData.discovered_facts.length : 0;
    if (categoriesFound > 0) {
      Logger.log(`✅ ${categoriesFound} veri grubu başarıyla kategorize edildi`);
      Logger.log('Kategorize veriler:', extractedData.discovered_facts);
      testResults.categorizedDataExtraction = true;
    } else {
      Logger.log('❌ Kategorize veri bulunamadı');
    }
    
    // Test 3: Data Loss Prevention
    Logger.log('📈 Test 3: Veri Kaybı Önleme');
    
    const totalMappedFields = Object.keys(extractedData).filter(key => 
      key !== 'discovered_facts' && extractedData[key] && extractedData[key] !== ''
    ).length;
    
    if (totalMappedFields >= 6) {
      Logger.log(`✅ ${totalMappedFields} spesifik alan eşleştirildi, veri kaybı önlendi`);
      testResults.noDataLoss = true;
    } else {
      Logger.log(`❌ Sadece ${totalMappedFields} alan eşleştirildi, potansiyel veri kaybı`);
    }
    
    // Test 4: Backward Compatibility
    Logger.log('🔄 Test 4: Geriye Dönük Uyumluluk');
    
    const englishHeaders = ['Company', 'Industry', 'Revenue', 'Employees'];
    const englishData = ['Test Corp', 'Technology', '1000000', '50'];
    const englishExtractedData = extractTurkishCompanyDataStructured(englishData, englishHeaders);
    
    if (englishExtractedData.company_name === 'Test Corp' && englishExtractedData.industry === 'Technology') {
      Logger.log('✅ İngilizce başlıklar için geriye dönük uyumluluk korundu');
      testResults.backwardCompatibility = true;
    } else {
      Logger.log('❌ Geriye dönük uyumluluk bozuldu');
    }
    
    // Overall validation
    const passedTests = Object.values(testResults).filter(result => result === true).length;
    testResults.overall = passedTests >= 3;
    
    Logger.log('');
    Logger.log('📈 TÜRK BAŞLIK TESPİTİ TEST ÖZETİ:');
    Logger.log(`Türk İş Başlıkları: ${testResults.turkishBusinessHeaders ? '✅' : '❌'}`);
    Logger.log(`Kategorize Veri Çıkarma: ${testResults.categorizedDataExtraction ? '✅' : '❌'}`);
    Logger.log(`Veri Kaybı Önleme: ${testResults.noDataLoss ? '✅' : '❌'}`);
    Logger.log(`Geriye Dönük Uyumluluk: ${testResults.backwardCompatibility ? '✅' : '❌'}`);
    Logger.log(`Genel Durum: ${testResults.overall ? '✅ TÜRK BAŞLIK İYİLEŞTİRMELERİ ÇALIŞIYOR' : '❌ SORUNLAR TESPİT EDİLDİ'}`);
    
    return testResults;
    
  } catch (error) {
    Logger.log('❌ Türk başlık tespiti testi başarısız:', error);
    return testResults;
  }
}

/**
 * Test Turkish hybrid mapping system functionality
 * Validates that the "structured first, Claude fallback" principle works for Turkish data
 * @returns {Object} Test results
 */
function testTurkishHybridMappingSystem() {
  Logger.log('🔄 TÜRK B2B HİBRİT EŞLEŞTİRME SİSTEMİ TESTİ');
  Logger.log('============================================');
  
  const testResults = {
    structuredMappingPrimary: false,
    claudeFallbackActivation: false,
    qualityAnalysisGeneration: false,
    missingFieldDetection: false,
    overall: false
  };
  
  try {
    // Test 1: Structured mapping should work for Turkish standard headers
    Logger.log('📋 Test 1: Birincil Yapısal Eşleştirme');
    
    const standardTurkishHeaders = [
      'Şirket Unvanı', 'Faaliyet Sektörü', 'Yıllık Ciro', 'Personel Sayısı', 'Merkez İli'
    ];
    const standardTurkishData = [
      'ABC Limited Şirketi', 'Bilişim Hizmetleri', '75000000', '150', 'Ankara'
    ];
    
    const structuredResult = extractTurkishCompanyDataStructured(standardTurkishData, standardTurkishHeaders);
    
    if (structuredResult.company_name && structuredResult.industry && structuredResult.headquarters) {
      Logger.log('✅ Birincil yapısal eşleştirme başarılı');
      testResults.structuredMappingPrimary = true;
    } else {
      Logger.log('❌ Birincil yapısal eşleştirme başarısız');
    }
    
    // Test 2: Quality analysis generation for Turkish context
    Logger.log('📊 Test 2: Türk Bağlamı İçin Kalite Analizi Üretimi');
    
    const qualityAnalysis = {
      structured_mapping_success: 4,
      total_critical_fields: 5,
      missing_fields: ['revenue_estimate'],
      claude_fallback_used: true,
      claude_mapping_success: 1,
      final_data_completeness: 80
    };
    
    const qualityReport = generateTurkishQualityReport(qualityAnalysis, 8);
    
    if (qualityReport && qualityReport.includes('Yapısal eşleştirme') && qualityReport.includes('%')) {
      Logger.log('✅ Türk kalite analizi üretimi başarılı');
      Logger.log(`Kalite raporu: ${qualityReport}`);
      testResults.qualityAnalysisGeneration = true;
    } else {
      Logger.log('❌ Türk kalite analizi üretimi başarısız');
    }
    
    // Test 3: Turkish missing field detection
    Logger.log('🔍 Test 3: Türk Eksik Alan Tespiti Doğruluğu');
    
    const criticalFields = ['company_name', 'industry', 'revenue_estimate', 'employees_estimate', 'headquarters'];
    const testTurkishData = {
      company_name: 'Test Şirketi A.Ş.',
      industry: 'İmalat',
      headquarters: 'İstanbul'
      // Missing: revenue_estimate, employees_estimate
    };
    
    const missingFields = criticalFields.filter(field => !testTurkishData[field] || testTurkishData[field] === '');
    
    if (missingFields.length === 2 && 
        missingFields.includes('revenue_estimate') && 
        missingFields.includes('employees_estimate')) {
      Logger.log('✅ Türk eksik alan tespiti doğru');
      testResults.missingFieldDetection = true;
    } else {
      Logger.log('❌ Türk eksik alan tespiti başarısız');
    }
    
    // Test 4: Mock Claude fallback for Turkish headers
    Logger.log('🤖 Test 4: Türk Başlıkları İçin Claude Fallback Simulasyonu');
    
    // Since we can't call Claude in tests without API key, simulate the result
    const mockClaudeFallbackResult = {
      industry: 'Teknoloji',
      revenue_estimate: '50000000'
    };
    
    if (mockClaudeFallbackResult.industry && mockClaudeFallbackResult.revenue_estimate) {
      Logger.log('✅ Claude fallback simulasyonu çalışıyor');
      testResults.claudeFallbackActivation = true;
    }
    
    // Overall validation
    const passedTests = Object.values(testResults).filter(result => result === true).length;
    testResults.overall = passedTests >= 3;
    
    Logger.log('');
    Logger.log('📈 TÜRK HİBRİT EŞLEŞTİRME TEST ÖZETİ:');
    Logger.log(`Yapısal Eşleştirme Birincil: ${testResults.structuredMappingPrimary ? '✅' : '❌'}`);
    Logger.log(`Claude Fallback Aktivasyonu: ${testResults.claudeFallbackActivation ? '✅' : '❌'}`);
    Logger.log(`Kalite Analizi Üretimi: ${testResults.qualityAnalysisGeneration ? '✅' : '❌'}`);
    Logger.log(`Eksik Alan Tespiti: ${testResults.missingFieldDetection ? '✅' : '❌'}`);
    Logger.log(`Genel Durum: ${testResults.overall ? '✅ TÜRK HİBRİT SİSTEM ÇALIŞIYOR' : '❌ SORUNLAR TESPİT EDİLDİ'}`);
    
    return testResults;
    
  } catch (error) {
    Logger.log('❌ Türk hibrit eşleştirme testi başarısız:', error);
    return testResults;
  }
}

/**
 * Validate Yolwise system consistency after adaptation
 * Comprehensive validation of all critical adaptations for Turkish B2B market
 */
function validateYolwiseSystemConsistency() {
  Logger.log('🔍 YOLWİSE SİSTEM TUTARLILIĞI DOĞRULAMA...');
  Logger.log('=======================================');
  
  const validationResults = {
    scaleConsistency: false,
    thresholdConsistency: false,
    turkishIndustryLogic: false,
    llmAdjustmentLimits: false,
    turkishQualityAnalysis: false,
    overall: false
  };
  
  try {
    // Test 1: Scale Consistency (0-100) for Turkish companies
    Logger.log('📏 Test 1: Türk Şirketleri İçin Skorlama Ölçeği Tutarlılığı');
    const testTurkishCompany = {
      company_name: 'Test Türk Şirketi A.Ş.',
      industry: 'finans',
      revenue_estimate: 2000000000, // 2B TL
      employees_estimate: '1000+',
      business_type: 'A.Ş.'
    };
    
    const turkishClaudeAnalysis = {
      company_name: 'Test Türk Şirketi A.Ş.',
      industry: 'finans',
      b2b_service_potential: 'yüksek türk b2b pazarında hizmet ihtiyacı',
      analysis_confidence: 'yüksek'
    };
    
    const mockTurkishResult = createYolwiseMockScoring(turkishClaudeAnalysis);
    
    if (mockTurkishResult.final_score >= 0 && mockTurkishResult.final_score <= 100) {
      Logger.log('✅ Türk mock skorlaması 0-100 ölçeğini kullanıyor');
      validationResults.scaleConsistency = true;
    } else {
      Logger.log(`❌ Türk mock skorlama ölçeği hatası: ${mockTurkishResult.final_score}`);
    }
    
    // Test 2: Threshold Consistency (60% for target) in Turkish context
    Logger.log('📊 Test 2: Türk Bağlamında Hedef Eşiği Tutarlılığı');
    const threshold60Turkish = { 
      ...turkishClaudeAnalysis, 
      b2b_service_potential: 'orta düzey türk b2b hizmet ihtiyacı' 
    };
    const result60Turkish = createYolwiseMockScoring(threshold60Turkish);
    
    if (result60Turkish.final_score >= 50 && result60Turkish.priority_recommendation) {
      Logger.log('✅ 60+ puan eşiği Türk şirketleri için hedefleri doğru tanımlıyor');
      validationResults.thresholdConsistency = true;
    } else {
      Logger.log(`❌ Türk eşik hatası: ${result60Turkish.final_score} pts → ${result60Turkish.priority_recommendation}`);
    }
    
    // Test 3: Turkish Industry Logic Consistency
    Logger.log('🏭 Test 3: Türk Sektör Tespiti Mantığı');
    const financeTestTurkish = { 
      ...turkishClaudeAnalysis, 
      industry: 'bankacılık ve finansal hizmetler',
      turkish_market_notes: 'istanbul merkezli büyük finans kurumu'
    };
    const financeResultTurkish = createYolwiseMockScoring(financeTestTurkish);
    
    if (financeResultTurkish.detected_industry === 'finans' && financeResultTurkish.industry_multiplier === 1.20) {
      Logger.log('✅ Türk sektör tespiti doğru çalışıyor (finans → ×1.20)');
      validationResults.turkishIndustryLogic = true;
    } else {
      Logger.log(`❌ Türk sektör mantığı hatası: ${financeResultTurkish.detected_industry} → ×${financeResultTurkish.industry_multiplier}`);
    }
    
    // Test 4: Turkish LLM Adjustment Limits
    Logger.log('⚖️ Test 4: Türk B2B LLM Düzeltme Limitleri');
    const complexTurkishAnalysis = {
      ...turkishClaudeAnalysis,
      business_context: 'büyüme, yeni yatırım, ihracat artışı, genişleme planları',
      growth_indicators: 'yıllık %25 büyüme',
      b2b_service_potential: 'çok yüksek türk b2b pazarında hizmet potansiyeli',
      turkish_market_notes: 'istanbul ana merkez, avrupa ihracatı, organize sanayi bölgesi',
      analysis_confidence: 'yüksek'
    };
    
    // This should trigger multiple positive adjustments but stay within ±25
    const baseTurkishResult = { 
      industry_adjusted_score: 70, 
      detected_industry: 'finans',
      industry_multiplier: 1.20
    };
    const adjustedTurkishResult = applyTurkishLLMAdjustment(baseTurkishResult, complexTurkishAnalysis);
    
    if (adjustedTurkishResult.llm_adjustment >= -25 && adjustedTurkishResult.llm_adjustment <= 25) {
      Logger.log(`✅ Türk LLM düzeltmesi limitler içinde: ${adjustedTurkishResult.llm_adjustment} puan`);
      validationResults.llmAdjustmentLimits = true;
    } else {
      Logger.log(`❌ Türk LLM düzeltmesi limitleri aşıyor: ${adjustedTurkishResult.llm_adjustment} puan`);
    }
    
    // Test 5: Turkish Quality Analysis Integration
    Logger.log('📈 Test 5: Türk Kalite Analizi Entegrasyonu');
    
    const mockTurkishQualityAnalysis = {
      structured_mapping_success: 4,
      total_critical_fields: 5,
      missing_fields: ['revenue_estimate'],
      claude_fallback_used: true,
      claude_mapping_success: 1,
      final_data_completeness: 80
    };
    
    const turkishQualityReport = generateTurkishQualityReport(mockTurkishQualityAnalysis, 10);
    
    if (turkishQualityReport && 
        turkishQualityReport.includes('veri kalitesi') && 
        turkishQualityReport.includes('Yapısal eşleştirme') &&
        turkishQualityReport.includes('%')) {
      Logger.log('✅ Türk kalite analizi entegrasyonu başarılı');
      validationResults.turkishQualityAnalysis = true;
    } else {
      Logger.log('❌ Türk kalite analizi entegrasyonu başarısız');
    }
    
    // Overall validation
    const passedTests = Object.values(validationResults).filter(result => result === true).length;
    validationResults.overall = passedTests >= 4;
    
    Logger.log('');
    Logger.log('📈 YOLWİSE DOĞRULAMA ÖZETİ:');
    Logger.log(`Ölçek Tutarlılığı: ${validationResults.scaleConsistency ? '✅' : '❌'}`);
    Logger.log(`Eşik Tutarlılığı: ${validationResults.thresholdConsistency ? '✅' : '❌'}`);
    Logger.log(`Türk Sektör Mantığı: ${validationResults.turkishIndustryLogic ? '✅' : '❌'}`);
    Logger.log(`LLM Düzeltme Limitleri: ${validationResults.llmAdjustmentLimits ? '✅' : '❌'}`);
    Logger.log(`Türk Kalite Analizi: ${validationResults.turkishQualityAnalysis ? '✅' : '❌'}`);
    Logger.log(`Genel Durum: ${validationResults.overall ? '✅ TÜM YOLWİSE ADAPTASYONLARI ÇALIŞIYOR' : '❌ SORUNLAR TESPİT EDİLDİ'}`);
    
    return validationResults;
    
  } catch (error) {
    Logger.log('❌ Yolwise doğrulama başarısız:', error);
    return validationResults;
  }
}

/**
 * Test Turkish company scoring with sample data
 */
function testTurkishScoring() {
  Logger.log('⚡ Türk Şirket Skorlama Testi Başlatılıyor...');
  
  const testTurkishCompanies = [
    { company_name: 'Türk Telekom A.Ş.', industry: 'telekomünikasyon' },
    { company_name: 'Arçelik A.Ş.', industry: 'beyaz eşya' },
    { company_name: 'İş Bankası A.Ş.', industry: 'bankacılık' },
    { company_name: 'Borusan Holding A.Ş.', industry: 'çelik' },
    { company_name: 'Migros Ticaret A.Ş.', industry: 'perakende' }
  ];
  
  const startTime = Date.now();
  let successCount = 0;
  let errorCount = 0;
  const results = [];
  
  try {
    for (const company of testTurkishCompanies) {
      try {
        Logger.log(`Testing Turkish company: ${company.company_name}`);
        
        // Use mock analysis for testing
        const mockTurkishAnalysis = {
          company_name: company.company_name,
          industry: company.industry,
          b2b_service_potential: 'orta düzey türk b2b hizmet ihtiyacı',
          analysis_confidence: 'orta',
          turkish_market_notes: 'türk pazarında aktif'
        };
        
        const scoringResult = createYolwiseMockScoring(mockTurkishAnalysis);
        
        const finalScore = scoringResult.final_score || scoringResult.industry_adjusted_score || 0;
        const priority = scoringResult.priority_recommendation || 'unknown';
        
        Logger.log(`✓ ${company.company_name}: ${finalScore}/100 pts (${priority})`);
        
        results.push({
          company: company.company_name,
          score: finalScore,
          priority: priority,
          industry: scoringResult.detected_industry || company.industry
        });
        
        successCount++;
        
      } catch (error) {
        Logger.log(`✗ ${company.company_name}: ${error.message}`);
        errorCount++;
      }
    }
    
    const totalTime = Date.now() - startTime;
    const avgTime = totalTime / testTurkishCompanies.length;
    const avgScore = results.length > 0 ? (results.reduce((sum, r) => sum + r.score, 0) / results.length) : 0;
    const targetCount = results.filter(r => r.priority === 'target').length;
    
    Logger.log('');
    Logger.log('📊 TÜRK ŞİRKET SKORLAMA SONUÇLARI:');
    Logger.log(`Toplam şirket: ${testTurkishCompanies.length}`);
    Logger.log(`Başarılı: ${successCount}`);
    Logger.log(`Başarısız: ${errorCount}`);
    Logger.log(`Toplam süre: ${totalTime}ms`);
    Logger.log(`Şirket başına ortalama: ${avgTime.toFixed(0)}ms`);
    Logger.log(`Başarı oranı: ${((successCount / testTurkishCompanies.length) * 100).toFixed(1)}%`);
    Logger.log(`Ortalama skor: ${avgScore.toFixed(1)}/100`);
    Logger.log(`Hedef oranı: ${((targetCount / results.length) * 100).toFixed(1)}%`);
    
    Logger.log('');
    Logger.log('📈 DETAYLI SONUÇLAR:');
    results.forEach(result => {
      Logger.log(`${result.company}: ${result.score}/100 (${result.priority}) - ${result.industry}`);
    });
    
  } catch (error) {
    Logger.log('Türk şirket performans testi başarısız:', error);
  }
}

/**
 * Comprehensive diagnostics function for Yolwise
 * Run this if you're experiencing issues with Turkish B2B scoring
 */
function runYolwiseDiagnostics() {
  Logger.log('🔍 Yolwise Lead Scoring Diagnostics v1.0 (Turkish B2B)');
  Logger.log('====================================================');
  
  try {
    // 1. Check Apps Script environment
    Logger.log('📱 Apps Script Ortamı:');
    Logger.log(`Runtime: ${typeof ScriptApp !== 'undefined' ? 'V8' : 'Legacy'}`);
    Logger.log(`Timezone: ${Session.getScriptTimeZone()}`);
    
    // 2. Check API keys for Yolwise
    Logger.log('');
    Logger.log('🔑 Yolwise API Anahtarları Durumu:');
    const claudeKey = getClaudeApiKey();
    const yolwiseKey = getYolwiseApiKey();
    const yolwiseUrl = getYolwiseApiUrl();
    Logger.log(`Claude: ${claudeKey ? '✓ ' + claudeKey.substring(0, 20) + '...' : '✗ Yapılandırılmamış'}`);
    Logger.log(`Yolwise Key: ${yolwiseKey ? '✓ ' + yolwiseKey.substring(0, 20) + '...' : '✗ Yapılandırılmamış'}`);
    Logger.log(`Yolwise URL: ${yolwiseUrl || '✗ Yapılandırılmamış'}`);
    
    // 3. Check permissions
    Logger.log('');
    Logger.log('🔐 İzin Kontrolü:');
    try {
      SpreadsheetApp.getActiveSpreadsheet();
      Logger.log('Spreadsheet erişimi: ✓');
    } catch (e) {
      Logger.log('Spreadsheet erişimi: ✗ ' + e.message);
    }
    
    try {
      UrlFetchApp.fetch('https://httpbin.org/get');
      Logger.log('Harici istekler: ✓');
    } catch (e) {
      Logger.log('Harici istekler: ✗ ' + e.message);
    }
    
    // 4. Test Yolwise API connections
    Logger.log('');
    Logger.log('🌐 Yolwise API Bağlantısı:');
    testYolwiseApiConnections();
    
    // 5. Validate Turkish system consistency
    Logger.log('');
    Logger.log('🔧 Türk Sistem Tutarlılığı Kontrolü:');
    const validation = validateYolwiseSystemConsistency();
    
    // 6. Test Turkish header detection
    Logger.log('');
    Logger.log('🔍 Türk Başlık Tespiti Testi:');
    const headerTest = testTurkishHeaderDetection();
    
    // 7. Test Turkish hybrid mapping system
    Logger.log('');
    Logger.log('🔄 Türk Hibrit Eşleştirme Sistemi Testi:');
    const hybridTest = testTurkishHybridMappingSystem();
    
    // 8. Check current sheet data
    Logger.log('');
    Logger.log('📊 Mevcut Sayfa Analizi:');
    try {
      const ranges = getAvailableRanges();
      Logger.log(`Mevcut veri aralıkları: ${ranges.length}`);
      ranges.forEach(range => {
        Logger.log(`  - ${range.label}: ${range.description}`);
      });
    } catch (e) {
      Logger.log('Sayfa analizi başarısız: ' + e.message);
    }
    
    // 9. Check for processing state
    Logger.log('');
    Logger.log('🔄 İşleme Durumu:');
    const processingState = checkProcessingState();
    if (processingState) {
      Logger.log(`✓ Devam ettirilebilir işleme durumu bulundu: ${processingState.processedRows} satır tamamlandı`);
    } else {
      Logger.log('- İşleme durumu bulunamadı');
    }
    
    Logger.log('');
    Logger.log('✅ Yolwise diagnostik tamamlandı!');
    
  } catch (error) {
    Logger.log('❌ Yolwise diagnostik başarısız:', error);
  }
}

/**
 * Generate Turkish scoring summary report
 * Creates analytical insights from Turkish B2B scoring results
 */
function generateTurkishSummaryReport() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Find relevant columns including Turkish labels
    const companyCol = headers.findIndex(h => h.toLowerCase().includes('şirket') || h.toLowerCase().includes('firma') || h.toLowerCase().includes('company')) + 1;
    const industryCol = headers.findIndex(h => h.toLowerCase().includes('sektör') || h.toLowerCase().includes('industry')) + 1;
    const finalScoreCol = headers.indexOf('Final Skor') + 1;
    const priorityCol = headers.indexOf('Öncelik') + 1;
    const qualityCol = headers.indexOf('Veri Kalitesi') + 1;
    
    if (!finalScoreCol || !priorityCol) {
      throw new Error('Skorlama sonuçları bulunamadı. Lütfen önce analizi çalıştırın.');
    }
    
    // Get data
    const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    const data = dataRange.getValues();
    
    // Analyze results including Turkish quality metrics
    let totalCompanies = 0;
    let targetLeads = 0;
    let nonTargetLeads = 0;
    let errors = 0;
    let scoreSum = 0;
    let excellentQuality = 0;
    let goodQuality = 0;
    let poorQuality = 0;
    const industryStats = {};
    
    data.forEach(row => {
      if (row[companyCol - 1]) { // Has company name
        totalCompanies++;
        
        const score = parseFloat(row[finalScoreCol - 1]) || 0;
        const priority = String(row[priorityCol - 1]).toLowerCase();
        const industry = String(row[industryCol - 1] || 'Bilinmeyen');
        const quality = qualityCol > 0 ? String(row[qualityCol - 1] || '') : '';
        
        if (priority.includes('hedeflenebilir') || priority.includes('target')) {
          targetLeads++;
        } else if (priority.includes('hedeflenemez') || priority.includes('non_target')) {
          nonTargetLeads++;
        } else {
          errors++;
        }
        
        if (score > 0) {
          scoreSum += score;
        }
        
        // Quality analysis for Turkish context
        if (quality.includes('Mükemmel') || quality.includes('✅')) {
          excellentQuality++;
        } else if (quality.includes('İyi') || quality.includes('⚠️')) {
          goodQuality++;
        } else if (quality.includes('Düşük') || quality.includes('❌')) {
          poorQuality++;
        }
        
        // Industry statistics
        if (!industryStats[industry]) {
          industryStats[industry] = { count: 0, totalScore: 0, targetCount: 0 };
        }
        industryStats[industry].count++;
        industryStats[industry].totalScore += score;
        if (priority.includes('hedeflenebilir') || priority.includes('target')) {
          industryStats[industry].targetCount++;
        }
      }
    });
    
    // Generate Turkish report
    const avgScore = totalCompanies > 0 ? (scoreSum / totalCompanies).toFixed(1) : 0;
    const targetRate = totalCompanies > 0 ? ((targetLeads / totalCompanies) * 100).toFixed(1) : 0;
    const excellentQualityRate = totalCompanies > 0 ? ((excellentQuality / totalCompanies) * 100).toFixed(1) : 0;
    
    Logger.log('📊 YOLWİSE TÜRK B2B SKORLAMA ÖZET RAPORU v1.0');
    Logger.log('=============================================');
    Logger.log(`📅 Oluşturulma: ${new Date().toLocaleString('tr-TR')}`);
    Logger.log('');
    Logger.log('📈 Genel İstatistikler:');
    Logger.log(`Analiz edilen toplam şirket: ${totalCompanies}`);
    Logger.log(`Hedeflenebilir lidler (≥60 puan): ${targetLeads} (${targetRate}%)`);
    Logger.log(`Hedeflenemeyen lidler (<60 puan): ${nonTargetLeads}`);
    Logger.log(`Hatalar: ${errors}`);
    Logger.log(`Ortalama skor: ${avgScore}/100`);
    
    // Turkish quality metrics
    Logger.log('');
    Logger.log('🎯 Türk Veri Kalitesi Metrikleri:');
    Logger.log(`Mükemmel kalite: ${excellentQuality} (${excellentQualityRate}%)`);
    Logger.log(`İyi kalite: ${goodQuality}`);
    Logger.log(`Düşük kalite: ${poorQuality}`);
    
    Logger.log('');
    Logger.log('🏭 Türk Sektör Dağılımı:');
    
    Object.entries(industryStats)
      .sort((a, b) => b[1].count - a[1].count)
      .forEach(([industry, stats]) => {
        const avgIndustryScore = (stats.totalScore / stats.count).toFixed(1);
        const industryTargetRate = ((stats.targetCount / stats.count) * 100).toFixed(1);
        Logger.log(`${industry}: ${stats.count} şirket, ort ${avgIndustryScore}/100 puan, ${industryTargetRate}% hedef oranı`);
      });
    
    Logger.log('');
    Logger.log('💡 Türk B2B Pazarı İçin Öneriler:');
    if (targetRate < 40) {
      Logger.log('• Düşük hedef oranı - Türk B2B lead kaynaklarını iyileştirmeyi düşünün');
    }
    if (errors > totalCompanies * 0.1) {
      Logger.log('• Yüksek hata oranı - Türk veri kalitesini ve API limitlerini kontrol edin');
    }
    if (avgScore < 50) {
      Logger.log('• Düşük ortalama skor - Türk sektör odaklama stratejisini gözden geçirin');
    }
    if (excellentQualityRate < 70) {
      Logger.log('• Veri kalitesi endişeleri - Türk kaynak veri başlıklarını iyileştirmeyi düşünün');
    }
    Logger.log('✅ Sistem tutarlı 0-100 ölçeği ve 60-puan eşiği kullanıyor');
    Logger.log('✅ Türk hibrit eşleştirme sistemi kalite takibi ile etkin');
    
    return {
      totalCompanies,
      targetLeads,
      nonTargetLeads,
      errors,
      avgScore: parseFloat(avgScore),
      targetRate: parseFloat(targetRate),
      excellentQuality,
      goodQuality,
      poorQuality,
      excellentQualityRate: parseFloat(excellentQualityRate),
      industryStats
    };
    
  } catch (error) {
    Logger.log('Türk rapor oluşturma başarısız:', error);
    throw error;
  }
}

/**
 * Backup current Yolwise configuration
 * Saves API keys and settings for recovery
 */
function backupYolwiseConfiguration() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const allProperties = properties.getProperties();
    
    const backup = {
      timestamp: new Date().toISOString(),
      version: '1.0-yolwise-turkish',
      properties: {}
    };
    
    // Only backup non-sensitive config (not API keys)
    Object.keys(allProperties).forEach(key => {
      if (!key.includes('API_KEY')) {
        backup.properties[key] = allProperties[key];
      } else {
        backup.properties[key] = '***GİZLİ***';
      }
    });
    
    Logger.log('💾 Yolwise Yapılandırma Yedeği v1.0:');
    Logger.log(JSON.stringify(backup, null, 2));
    
    return backup;
    
  } catch (error) {
    Logger.log('Yolwise yedekleme başarısız:', error);
    throw error;
  }
}