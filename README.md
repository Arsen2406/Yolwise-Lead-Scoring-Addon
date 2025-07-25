# Yolwise Lead Scoring API

**Turkish B2B Market Lead Scoring System with Three-Layer Architecture**

🎯 **Adapted from Smartway Lead Scoring** for the Turkish business market, implementing advanced machine learning-based lead scoring with industry-specific optimizations.

## 🌟 Features

- **Three-Layer Scoring Architecture**: Base Score + Industry Multipliers + LLM Adjustment
- **Turkish Market Optimization**: Specialized for Turkish B2B service propensity patterns
- **API Key Security**: Protected endpoints with environment-based authentication
- **Single & Bulk Scoring**: Support for individual companies and batch processing
- **Industry Intelligence**: 13+ Turkish industry categories with tailored multipliers
- **Real-time Processing**: Fast scoring with sub-second response times
- **Error-Resilient**: Comprehensive error handling and validation

## 📊 Specification Compliance

This implementation follows the complete **`yolwise_scoring_specification.md`** which includes:

- ✅ **Layer 1**: Base Score Components (35% Company Size + 25% Industry + 20% Financial + 10% Geographic + 10% Growth)
- ✅ **Layer 2**: Industry Multipliers (0.70x to 1.20x based on B2B service propensity)
- ✅ **Layer 3**: LLM Adjustment (±25 points for qualitative business factors)
- ✅ **Target Threshold**: 60+ points = Target Lead recommendation
- ✅ **Turkish Context**: TL currency, Turkish cities, local business patterns

## 🚀 Quick Start

### Deploy on Replit

1. **Import this repository** to Replit
2. **Set Environment Variable**:
   - Go to Secrets tab in Replit
   - Add `YOLWISE_API_KEY` with your chosen API key value
3. **Run the application**:
   ```bash
   python main.py
   ```
4. **API will be available** at: `https://your-repl-name.replit.app`

### Local Development

```bash
# Clone the repository
git clone https://github.com/Arsen2406/Yolwise-Lead-Scoring-Addon.git
cd Yolwise-Lead-Scoring-Addon

# Install dependencies
pip install -r requirements.txt

# Set environment variable
export YOLWISE_API_KEY="your-secret-key-here"

# Run the application
python main.py
```

## 📡 API Endpoints

### 🔐 Authentication
All scoring endpoints require an API key via:
- **Header**: `X-API-Key: your-api-key`
- **Query Parameter**: `?api_key=your-api-key`

### 📋 Available Endpoints

| Endpoint | Method | Description | Auth Required |
|----------|--------|-------------|---------------|
| `/` | GET | API information and documentation | ❌ |
| `/health` | GET | Health check and system status | ❌ |
| `/industries` | GET | List supported industries with multipliers | ❌ |
| `/score_company` | POST | Score a single company | ✅ |
| `/score_batch` | POST | Score multiple companies | ✅ |

### 📝 Request Examples

#### Single Company Scoring
```json
POST /score_company
Headers: X-API-Key: your-api-key
Content-Type: application/json

{
  "company_name": "Arçelik A.Ş.",
  "company_data": {
    "Industry": "Mechanical or Industrial Engineering",
    "Annual Revenue": 45000000000,
    "Number of Employees": 15000,
    "City": "Istanbul",
    "Year Founded": 1955,
    "Company Domain Name": "arcelik.com.tr",
    "Description": "Leading home appliances manufacturer in Turkey with international operations and innovative technology solutions."
  }
}
```

#### Batch Company Scoring
```json
POST /score_batch
Headers: X-API-Key: your-api-key
Content-Type: application/json

{
  "companies": [
    {
      "name": "Company A",
      "data": {
        "Industry": "Food & Beverages",
        "Annual Revenue": 25000000,
        "Number of Employees": 150
      }
    },
    {
      "name": "Company B", 
      "data": {
        "Industry": "Computer Software",
        "Annual Revenue": 5000000,
        "Number of Employees": 45
      }
    }
  ]
}
```

### 📊 Response Format

```json
{
  "success": true,
  "result": {
    "company_name": "Arçelik A.Ş.",
    "base_score": 87.5,
    "industry_multiplier": 1.0,
    "industry_adjusted_score": 87.5,
    "llm_adjustment": 12.0,
    "final_score": 99.5,
    "detected_industry": "mechanical or industrial engineering",
    "industry_confidence": "medium",
    "priority_recommendation": "target",
    "reasoning": "High-confidence target • Industry: mechanical or industrial engineering (×1.00) • Large enterprise • Prime Turkish market location",
    "processing_time_ms": 45
  },
  "metadata": {
    "api_version": "1.0-yolwise",
    "scoring_model": "Turkish B2B Market",
    "target_threshold": 60
  }
}
```

## 🏗️ Architecture Details

### Three-Layer Scoring System

1. **Base Score (0-100 points)**
   - **Company Size Indicator** (35%): Employee count + Revenue in Turkish Lira
   - **Industry B2B Propensity** (25%): Sector-specific B2B service likelihood
   - **Financial Capacity** (20%): Revenue stability + Company maturity
   - **Geographic Presence** (10%): Turkish market location advantages
   - **Growth & Digital Indicators** (10%): Digital presence + Growth signals

2. **Industry Multipliers (0.70x - 1.20x)**
   - **High B2B Sectors** (1.15x+): Renewables, Logistics, Utilities
   - **Medium B2B Sectors** (1.00-1.10x): Food, Chemicals, Building Materials
   - **Lower B2B Sectors** (0.85-0.95x): Retail, Construction, Automotive
   - **Low B2B Sectors** (<0.85x): Software, Healthcare, Transportation

3. **LLM Adjustment (±25 points)**
   - Growth indicators and market position
   - Data quality and completeness
   - Business risk factors
   - Service compatibility assessment

### Target Classification
- **≥60 Points**: Target Lead (recommended for pursuit)
- **<60 Points**: Non-Target Lead (lower priority)

## 📈 Industry Support

Currently supports **13 major Turkish industries** with data-driven multipliers:

| Industry | Multiplier | Target Rate | Reasoning |
|----------|------------|-------------|-----------|
| Renewables & Environment | 1.20x | 60.0% | Growing sector, significant B2B investment |
| Logistics and Supply Chain | 1.17x | 58.8% | Complex operational B2B needs |
| Utilities | 1.15x | 54.5% | Infrastructure-heavy requirements |
| Food & Beverages | 1.10x | 50.0% | Moderate B2B service requirements |
| Chemicals | 1.05x | 45.8% | Specialized technical services |
| Building Materials | 1.02x | 44.4% | Construction-related B2B services |
| Mechanical/Industrial Eng. | 1.00x | 42.0% | Baseline engineering services |
| Retail | 0.90x | 37.5% | Consumer-focused, limited B2B |
| Construction | 0.88x | 35.3% | Project-based requirements |
| Automotive | 0.85x | 34.1% | Manufacturing-focused operations |
| Computer Software | 0.80x | 29.6% | Self-service solutions preferred |
| Hospital & Health Care | 0.75x | 30.0% | Specialized service requirements |
| Transportation/Trucking | 0.70x | 23.1% | Operational focus over procurement |

## 🔧 Configuration

### Environment Variables

| Variable | Description | Required | Example |
|----------|-------------|----------|---------|
| `YOLWISE_API_KEY` | API authentication key | ✅ Yes | `yw_prod_abc123xyz789` |
| `PORT` | Server port (default: 5000) | ❌ No | `8080` |

### Replit Secrets Setup

1. Open your Replit project
2. Click on "Secrets" in the left sidebar
3. Add a new secret:
   - **Key**: `YOLWISE_API_KEY`
   - **Value**: Your chosen API key (keep it secure!)

## 🧪 Testing

### Health Check
```bash
curl https://your-repl-name.replit.app/health
```

### Test Scoring (with API key)
```bash
curl -X POST https://your-repl-name.replit.app/score_company \
  -H "X-API-Key: your-api-key" \
  -H "Content-Type: application/json" \
  -d '{
    "company_name": "Test Company",
    "company_data": {
      "Industry": "Food & Beverages",
      "Annual Revenue": 50000000,
      "Number of Employees": 200,
      "City": "Istanbul"
    }
  }'
```

## 📚 Integration

### Google Sheets Add-on Integration
This API is designed to work with the **Yolwise Google Sheets Add-on** located in the `/src` directory:

- **`APIService.gs`**: Handles API communication
- **`Code.gs`**: Main add-on logic and UI
- **`Utils.gs`**: Utility functions for data processing
- **`index.html`**: Add-on interface

### Expected Data Format
The API expects data in the format matching **Target Leads.xlsx** structure:
- Company name
- Industry
- Annual Revenue (in Turkish Lira)
- Number of Employees
- City + State/Region
- Company Domain Name
- Description
- Year Founded
- LinkedIn Company Page
- Other business indicators

## ⚡ Performance

- **Processing Time**: <100ms per company for single scoring
- **Batch Processing**: Efficiently handles 100+ companies
- **Memory Usage**: Optimized for Replit's memory constraints
- **Concurrent Requests**: Thread-safe implementation

## 🔒 Security

- **API Key Authentication**: All scoring endpoints protected
- **Input Validation**: Comprehensive request validation
- **Error Handling**: Secure error responses without data leakage
- **Environment-based Configuration**: Sensitive data in environment variables

## 🐛 Troubleshooting

### Common Issues

1. **403 Unauthorized**: Check your API key in Replit Secrets
2. **500 Internal Server Error**: Verify all required data fields are provided
3. **Connection Timeout**: Ensure Replit instance is awake and running

### Debug Mode
For development, you can enable detailed logging by modifying `main.py`:
```python
app.run(debug=True)  # Only for development!
```

## 📄 License

This project is based on the Smartway Lead Scoring system and adapted for the Turkish B2B market according to the comprehensive Yolwise specification.

## 🤝 Support

For technical support or feature requests, please contact the development team or create an issue in the repository.

---

**⚡ Ready to deploy on Replit and start scoring Turkish B2B leads with precision!**
