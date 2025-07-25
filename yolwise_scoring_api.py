#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Yolwise Lead Scoring API for Turkish B2B Market
ADAPTED FROM SMARTWAY: Turkish business context with English interface
Implements three-layer scoring architecture according to yolwise_scoring_specification.md
Deployed at: https://yolwiseleadscoring.replit.app
Created using Context7 Flask documentation and Yolwise specification
"""

import os
import json
import time
import re
from typing import Dict, List, Any, Optional
from dataclasses import dataclass, asdict
from functools import wraps
from flask import Flask, request, jsonify, abort
from werkzeug.exceptions import HTTPException

app = Flask(__name__)

# API Key validation decorator
def require_api_key(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        api_key = request.headers.get('X-API-Key') or request.args.get('api_key')
        expected_key = os.environ.get('YOLWISE_API_KEY')
        
        if not expected_key:
            return jsonify({
                'success': False,
                'error': 'API key not configured on server',
                'code': 500
            }), 500
            
        if not api_key or api_key != expected_key:
            return jsonify({
                'success': False,
                'error': 'Invalid or missing API key',
                'code': 401
            }), 401
            
        return f(*args, **kwargs)
    return decorated_function

@dataclass
class CompanyScore:
    """Company scoring result according to Yolwise specification"""
    company_name: str
    base_score: float
    industry_multiplier: float
    industry_adjusted_score: float
    llm_adjustment: float
    final_score: float
    detected_industry: str
    industry_confidence: str
    priority_recommendation: str  # "target" or "non_target"
    reasoning: str
    processing_time_ms: int
    score_breakdown: Dict[str, float]

class YolwiseScoring:
    """
    Yolwise Lead Scoring System for Turkish B2B Market
    Three-layer architecture: Base Score + Industry Multipliers + LLM Adjustment
    Based on yolwise_scoring_specification.md
    """

    def __init__(self):
        # Industry multipliers based on Yolwise specification for Turkish B2B market
        self.industry_multipliers = {
            # High B2B Service Sectors (>1.10 multiplier)
            'renewables & environment': {
                'multiplier': 1.20,
                'confidence': 'high',
                'keywords': ['renewable', 'environment', 'solar', 'wind', 'green energy', 'sustainability'],
                'reasoning': '60% target rate, growing sector with significant B2B investment'
            },
            'logistics and supply chain': {
                'multiplier': 1.17,
                'confidence': 'high', 
                'keywords': ['logistics', 'supply chain', 'transportation', 'shipping', 'distribution', 'freight'],
                'reasoning': '58.8% target rate, complex operational B2B needs'
            },
            'utilities': {
                'multiplier': 1.15,
                'confidence': 'high',
                'keywords': ['utilities', 'electric', 'power', 'gas', 'water', 'energy distribution'],
                'reasoning': '54.5% target rate, infrastructure-heavy B2B requirements'
            },

            # Medium B2B Service Sectors (1.00-1.10 multiplier)
            'food & beverages': {
                'multiplier': 1.10,
                'confidence': 'medium',
                'keywords': ['food', 'beverage', 'dairy', 'agriculture', 'nutrition', 'farming'],
                'reasoning': '50% target rate, moderate B2B service requirements'
            },
            'chemicals': {
                'multiplier': 1.05,
                'confidence': 'medium',
                'keywords': ['chemical', 'pharmaceutical', 'biotech', 'laboratory', 'manufacturing chemical'],
                'reasoning': '45.8% target rate, specialized technical services needed'
            },
            'building materials': {
                'multiplier': 1.02,
                'confidence': 'medium',
                'keywords': ['building materials', 'construction materials', 'cement', 'steel', 'concrete'],
                'reasoning': '44.4% target rate, construction-related B2B services'
            },

            # Baseline B2B Service Sectors (0.95-1.00 multiplier)
            'mechanical or industrial engineering': {
                'multiplier': 1.00,
                'confidence': 'medium',
                'keywords': ['mechanical', 'industrial engineering', 'machinery', 'equipment', 'manufacturing'],
                'reasoning': '42% target rate, baseline engineering service needs'
            },
            'mining & metals': {
                'multiplier': 0.98,
                'confidence': 'medium',
                'keywords': ['mining', 'metals', 'metallurgy', 'extraction', 'ore processing'],
                'reasoning': '41.2% target rate, specialized but limited service scope'
            },
            'pharmaceuticals': {
                'multiplier': 0.97,
                'confidence': 'medium',
                'keywords': ['pharmaceuticals', 'pharma', 'medicine', 'drugs', 'healthcare'],
                'reasoning': '40% target rate, highly regulated sector'
            },

            # Lower B2B Service Sectors (0.85-0.95 multiplier)
            'retail': {
                'multiplier': 0.90,
                'confidence': 'low',
                'keywords': ['retail', 'consumer', 'shopping', 'store', 'commerce', 'sales'],
                'reasoning': '37.5% target rate, consumer-focused with limited B2B needs'
            },
            'construction': {
                'multiplier': 0.88,
                'confidence': 'low',
                'keywords': ['construction', 'building', 'architecture', 'contractor', 'infrastructure'],
                'reasoning': '35.3% target rate, project-based service requirements'
            },
            'automotive': {
                'multiplier': 0.85,
                'confidence': 'low',
                'keywords': ['automotive', 'automobile', 'vehicle', 'car', 'transportation equipment'],
                'reasoning': '34.1% target rate, manufacturing-focused operations'
            },

            # Low B2B Service Sectors (<0.85 multiplier)
            'computer software': {
                'multiplier': 0.80,
                'confidence': 'low',
                'keywords': ['computer software', 'technology', 'software', 'it', 'digital', 'programming'],
                'reasoning': '29.6% target rate, self-service digital solutions preferred'
            },
            'hospital & health care': {
                'multiplier': 0.75,
                'confidence': 'low',
                'keywords': ['hospital', 'health care', 'medical', 'healthcare', 'clinic', 'pharmaceutical'],
                'reasoning': '30% target rate, specialized service requirements outside typical B2B'
            },
            'transportation/trucking': {
                'multiplier': 0.70,
                'confidence': 'low',
                'keywords': ['transportation', 'trucking', 'freight', 'delivery', 'shipping'],
                'reasoning': '23.1% target rate, operational focus over service procurement'
            }
        }

        # Turkish city tier classification for geographic scoring
        self.city_tiers = {
            'tier_1_cities': ['istanbul', 'ankara', 'izmir'],
            'tier_2_cities': ['bursa', 'antalya', 'gaziantep', 'konya'],
            'industrial_regions': ['kocaeli', 'tekirdag', 'eskisehir'],
            'other_cities': []  # All other Turkish cities
        }

        # Scoring weights according to Yolwise specification
        self.scoring_weights = {
            'company_size_indicator': 0.35,
            'industry_b2b_propensity': 0.25,
            'financial_capacity': 0.20,
            'geographic_presence': 0.10,
            'growth_digital_indicators': 0.10
        }

    def calculate_score(self, company_name: str, company_data: Dict[str, Any]) -> CompanyScore:
        """
        Main scoring method implementing three-layer Yolwise architecture
        @param company_name: Company name
        @param company_data: Company data matching Target Leads.xlsx structure
        @returns CompanyScore with complete scoring breakdown
        """
        start_time = time.time()

        # Layer 1: Base Score Components (0-100 points)
        base_components = self._calculate_base_components(company_data)
        base_score = sum(score * weight for score, weight in zip(
            base_components.values(), self.scoring_weights.values()))
        base_score = max(0, min(100, base_score))

        # Layer 2: Industry Multipliers
        industry, multiplier, confidence = self._detect_industry(company_name, company_data)
        industry_adjusted_score = base_score * multiplier
        industry_adjusted_score = max(0, min(100, industry_adjusted_score))

        # Layer 3: LLM Adjustment (±25 points) - simplified for this implementation
        llm_adjustment = self._calculate_llm_adjustment(company_data, industry_adjusted_score)
        
        # Final score calculation
        final_score = max(0, min(100, industry_adjusted_score + llm_adjustment))

        # Priority recommendation (60+ = target)
        priority_recommendation = "target" if final_score >= 60 else "non_target"

        processing_time = int((time.time() - start_time) * 1000)

        return CompanyScore(
            company_name=company_name,
            base_score=round(base_score, 1),
            industry_multiplier=multiplier,
            industry_adjusted_score=round(industry_adjusted_score, 1),
            llm_adjustment=llm_adjustment,
            final_score=round(final_score, 1),
            detected_industry=industry,
            industry_confidence=confidence,
            priority_recommendation=priority_recommendation,
            reasoning=self._generate_reasoning(company_data, industry, final_score),
            processing_time_ms=processing_time,
            score_breakdown=base_components
        )

    def _calculate_base_components(self, data: Dict[str, Any]) -> Dict[str, float]:
        """Calculate base score components according to Yolwise specification"""
        return {
            'company_size_score': self._evaluate_company_size(data),
            'industry_propensity_score': self._evaluate_industry_propensity(data),
            'financial_capacity_score': self._evaluate_financial_capacity(data),
            'geographic_score': self._evaluate_geographic_presence(data),
            'growth_digital_score': self._evaluate_growth_digital_indicators(data)
        }

    def _evaluate_company_size(self, data: Dict[str, Any]) -> float:
        """Company Size Indicator (35% weight) - Turkish market context"""
        score = 30  # Base score

        # Employee Count Bonus (based on Yolwise data analysis)
        employees = self._extract_number(str(data.get('Number of Employees', data.get('employees_estimate', 0))))
        if employees >= 5000:
            score += 30  # 71.9% target rate
        elif employees >= 1000:
            score += 25  # 65.3% target rate
        elif employees >= 200:
            score += 15  # 39.4% target rate
        elif employees >= 50:
            score += 10  # 28.2% target rate
        elif employees >= 1:
            score += 5   # 28.1% target rate

        # Revenue Bonus (Turkish Lira - converted thresholds)
        revenue = self._safe_float(data.get('Annual Revenue', data.get('revenue_estimate', 0)))
        if revenue >= 500000000:  # 500M+ TL
            score += 25  # 70.3% target rate
        elif revenue >= 100000000:  # 100-500M TL
            score += 20  # 60.5% target rate
        elif revenue >= 50000000:   # 50-100M TL
            score += 10  # 36.4% target rate
        elif revenue >= 10000000:   # 10-50M TL
            score += 5   # 29.6% target rate
        elif revenue > 0:
            score += 3   # <10M TL - 20% target rate

        return min(score, 100)

    def _evaluate_industry_propensity(self, data: Dict[str, Any]) -> float:
        """Industry B2B Service Propensity (25% weight)"""
        industry = str(data.get('Industry', '')).lower()

        # High B2B Propensity (80-90 points)
        high_propensity = ['renewables', 'environment', 'logistics', 'utilities']
        if any(term in industry for term in high_propensity):
            return 85

        # Medium-High B2B Propensity (60-75 points)
        medium_high = ['food', 'beverage', 'chemical', 'building materials']
        if any(term in industry for term in medium_high):
            return 70

        # Medium B2B Propensity (40-55 points)
        medium = ['mechanical', 'industrial', 'engineering', 'mining', 'pharmaceutical']
        if any(term in industry for term in medium):
            return 50

        # Low-Medium B2B Propensity (25-35 points)
        low_medium = ['retail', 'construction', 'automotive']
        if any(term in industry for term in low_medium):
            return 30

        # Low B2B Propensity (10-20 points)
        low = ['software', 'computer', 'hospital', 'healthcare', 'transportation']
        if any(term in industry for term in low):
            return 15

        return 30  # Unknown industry default

    def _evaluate_financial_capacity(self, data: Dict[str, Any]) -> float:
        """Financial Capacity (20% weight) - Turkish market context"""
        score = 40  # Base score

        # Revenue Stability
        revenue = self._safe_float(data.get('Annual Revenue', 0))
        if revenue > 100000000:  # >100M TL
            score += 30
        elif revenue > 50000000:  # 50-100M TL
            score += 20
        elif revenue > 10000000:  # 10-50M TL
            score += 10
        elif revenue > 0:
            score += 5

        # Company Maturity (based on founding year)
        founded_year = self._extract_number(str(data.get('Year Founded', 0)))
        if founded_year > 0:
            current_year = 2025
            age = current_year - founded_year
            if age >= 20:
                score += 20  # Established, stable
            elif age >= 10:
                score += 15  # Mature business
            elif age >= 5:
                score += 10  # Developing business
            else:
                score += 5   # Startup phase

        # Business Stability Indicators
        if data.get('Company Domain Name'):
            score += 5  # Professional domain
        if data.get('Street Address') and data.get('City'):
            score += 5  # Complete address information

        return min(score, 100)

    def _evaluate_geographic_presence(self, data: Dict[str, Any]) -> float:
        """Geographic Presence (10% weight) - Turkish market focus"""
        score = 50  # Base score (all companies in Turkey)

        city = str(data.get('City', '')).lower()
        
        # Location Tier Bonuses
        if city in self.city_tiers['tier_1_cities']:
            score += 25  # Tier 1 cities
        elif city in self.city_tiers['tier_2_cities']:
            score += 15  # Tier 2 cities
        elif city in self.city_tiers['industrial_regions']:
            score += 10  # Industrial regions
        else:
            score += 5   # Other cities

        # International presence indicators
        description = str(data.get('Description', '')).lower()
        if any(term in description for term in ['international', 'global', 'export', 'worldwide']):
            score += 15

        return min(score, 100)

    def _evaluate_growth_digital_indicators(self, data: Dict[str, Any]) -> float:
        """Growth & Digital Indicators (10% weight)"""
        score = 30  # Base score

        # Digital Presence
        if data.get('Company Domain Name'):
            score += 20  # Professional website
        if data.get('LinkedIn Company Page'):
            score += 15  # LinkedIn presence
        if data.get('Facebook Company Page'):
            score += 10  # Facebook presence

        # Growth Indicators (analyze description)
        description = str(data.get('Description', '')).lower()
        if any(term in description for term in ['leading', 'growing', 'expanding']):
            score += 15
        if any(term in description for term in ['innovation', 'technology', 'solutions']):
            score += 10
        if any(term in description for term in ['development', 'market leader']):
            score += 10

        # Professional Activity
        if len(str(data.get('Description', ''))) > 100:
            score += 10  # Detailed description
        if sum(1 for field in ['Phone Number', 'Company Domain Name'] if data.get(field)) >= 2:
            score += 5   # Multiple contact methods

        return min(score, 100)

    def _detect_industry(self, company_name: str, company_data: Dict[str, Any]) -> tuple:
        """Detect industry using Yolwise classification"""
        text_data = [
            company_name.lower(),
            str(company_data.get('Industry', '')).lower(),
            str(company_data.get('Description', '')).lower()
        ]
        combined_text = ' '.join(text_data)

        # Find best match
        best_match = None
        max_score = 0

        for industry_name, industry_info in self.industry_multipliers.items():
            score = 0
            for keyword in industry_info['keywords']:
                if keyword in combined_text:
                    score += len(keyword)  # Longer keywords = higher weight

            if score > max_score:
                max_score = score
                best_match = industry_name

        if best_match and max_score > 0:
            industry_data = self.industry_multipliers[best_match]
            return best_match, industry_data['multiplier'], industry_data['confidence']
        else:
            return 'other', 1.0, 'low'

    def _calculate_llm_adjustment(self, data: Dict[str, Any], industry_score: float) -> float:
        """
        Simplified LLM adjustment (±25 points)
        In full implementation, this would call Claude API for qualitative analysis
        """
        adjustment = 0

        # Business Growth Indicators
        description = str(data.get('Description', '')).lower()
        if any(term in description for term in ['expansion', 'new facilities', 'growing']):
            adjustment += 10
        if any(term in description for term in ['innovation', 'technology', 'digital']):
            adjustment += 5

        # Market Position Indicators  
        if any(term in description for term in ['leader', 'leading', 'market share']):
            adjustment += 8
        if any(term in description for term in ['partnership', 'strategic']):
            adjustment += 3

        # Data Quality Bonus
        fields_completed = sum(1 for field in ['Company Domain Name', 'Description', 'Year Founded', 
                                             'LinkedIn Company Page'] if data.get(field))
        if fields_completed >= 3:
            adjustment += 5
        elif fields_completed >= 2:
            adjustment += 3

        # Ensure within ±25 range
        return max(-25, min(25, adjustment))

    def _generate_reasoning(self, data: Dict[str, Any], industry: str, final_score: float) -> str:
        """Generate human-readable reasoning for the score"""
        reasons = []
        
        # Score level reasoning
        if final_score >= 80:
            reasons.append("High-confidence target")
        elif final_score >= 60:
            reasons.append("Qualified target candidate")
        else:
            reasons.append("Below target threshold")

        # Industry reasoning
        if industry in self.industry_multipliers:
            industry_data = self.industry_multipliers[industry]
            reasons.append(f"Industry: {industry} (×{industry_data['multiplier']:.2f})")
        
        # Size indicators
        employees = self._extract_number(str(data.get('Number of Employees', 0)))
        if employees >= 1000:
            reasons.append("Large enterprise")
        elif employees >= 200:
            reasons.append("Medium-large company")
        elif employees >= 50:
            reasons.append("Medium company")

        # Geographic advantage
        city = str(data.get('City', '')).lower()
        if city in self.city_tiers['tier_1_cities']:
            reasons.append("Prime Turkish market location")
        
        return " • ".join(reasons)

    def _safe_float(self, value) -> float:
        """Safely convert value to float"""
        try:
            if isinstance(value, (int, float)):
                return float(value)
            elif isinstance(value, str):
                # Extract numbers from string
                numbers = re.findall(r'\d+\.?\d*', str(value))
                return float(numbers[0]) if numbers else 0
            else:
                return 0
        except (ValueError, TypeError, AttributeError):
            return 0

    def _extract_number(self, text: str) -> int:
        """Extract integer from text"""
        try:
            numbers = re.findall(r'\d+', str(text))
            return int(numbers[0]) if numbers else 0
        except (ValueError, TypeError, AttributeError):
            return 0

# Initialize scoring engine
scoring_engine = YolwiseScoring()

# Flask Error Handlers
@app.errorhandler(HTTPException)
def handle_http_exception(e):
    """Return JSON for HTTP errors"""
    return jsonify({
        "success": False,
        "error": e.name,
        "message": e.description,
        "code": e.code
    }), e.code

@app.errorhandler(Exception)
def handle_generic_exception(e):
    """Handle non-HTTP exceptions"""
    if isinstance(e, HTTPException):
        return e

    return jsonify({
        "success": False,
        "error": "Internal Server Error",
        "message": str(e),
        "code": 500
    }), 500

# API Routes
@app.route('/health', methods=['GET'])
def health_check():
    """Health check endpoint"""
    return jsonify({
        'status': 'healthy',
        'timestamp': time.time(),
        'version': '1.0-yolwise',
        'description': 'Yolwise Lead Scoring API for Turkish B2B Market',
        'specification': 'Three-layer architecture with English interface'
    })

@app.route('/score_company', methods=['POST'])
@require_api_key
def score_company():
    """Score a single company"""
    try:
        if not request.json:
            abort(400, description="JSON payload required")

        data = request.json
        company_name = data.get('company_name')

        if not company_name:
            abort(400, description="company_name is required")

        company_data = data.get('company_data', {})

        # Perform Yolwise scoring
        result = scoring_engine.calculate_score(company_name, company_data)

        return jsonify({
            'success': True,
            'result': asdict(result),
            'metadata': {
                'api_version': '1.0-yolwise',
                'scoring_model': 'Turkish B2B Market',
                'target_threshold': 60,
                'processing_time_ms': result.processing_time_ms
            }
        })

    except Exception as e:
        return jsonify({
            'success': False, 
            'error': str(e),
            'code': 500
        }), 500

@app.route('/score_batch', methods=['POST'])
@require_api_key
def score_batch():
    """Bulk scoring for multiple companies"""
    try:
        if not request.json:
            abort(400, description="JSON payload required")

        data = request.json
        companies = data.get('companies', [])

        if not isinstance(companies, list) or not companies:
            abort(400, description="companies list is required and cannot be empty")

        results = []
        start_time = time.time()

        for company_info in companies:
            try:
                # Handle different input formats
                if isinstance(company_info, str):
                    company_name = company_info
                    company_data = {}
                elif isinstance(company_info, dict):
                    company_name = company_info.get('name', company_info.get('company_name', ''))
                    company_data = company_info.get('data', company_info)
                else:
                    continue

                if not company_name:
                    continue

                # Score the company
                result = scoring_engine.calculate_score(company_name, company_data)

                results.append({
                    'company_name': result.company_name,
                    'base_score': result.base_score,
                    'industry_adjusted_score': result.industry_adjusted_score,
                    'final_score': result.final_score,
                    'priority_recommendation': result.priority_recommendation,
                    'detected_industry': result.detected_industry,
                    'industry_multiplier': result.industry_multiplier,
                    'confidence': result.industry_confidence,
                    'reasoning': result.reasoning
                })

            except Exception as e:
                results.append({
                    'company_name': company_name if 'company_name' in locals() else 'Unknown',
                    'base_score': 0,
                    'final_score': 0,
                    'priority_recommendation': 'error',
                    'error': str(e)
                })

        # Sort by final score
        results.sort(key=lambda x: x.get('final_score', 0), reverse=True)

        # Calculate statistics
        target_count = len([r for r in results if r.get('priority_recommendation') == 'target'])
        processing_time = int((time.time() - start_time) * 1000)

        return jsonify({
            'success': True,
            'results': results,
            'summary': {
                'total_companies': len(results),
                'target_recommendations': target_count,
                'non_target_recommendations': len(results) - target_count,
                'target_rate_percent': round((target_count / len(results)) * 100) if results else 0,
                'processing_time_ms': processing_time,
                'top_targets': [r for r in results if r.get('priority_recommendation') == 'target'][:10]
            },
            'metadata': {
                'api_version': '1.0-yolwise',
                'scoring_model': 'Turkish B2B Market',
                'target_threshold': 60,
                'specification': 'Three-layer architecture'
            }
        })

    except Exception as e:
        return jsonify({
            'success': False, 
            'error': str(e),
            'code': 500
        }), 500

@app.route('/industries', methods=['GET'])
def get_industries():
    """Get supported industries with multipliers"""
    industries = {}
    for industry, data in scoring_engine.industry_multipliers.items():
        industries[industry] = {
            'multiplier': data['multiplier'],
            'confidence': data['confidence'],
            'reasoning': data['reasoning'],
            'keywords': data['keywords'][:5]  # First 5 keywords only
        }
    
    return jsonify({
        'success': True,
        'industries': industries,
        'metadata': {
            'total_industries': len(industries),
            'default_multiplier': 1.0,
            'target_threshold': 60
        }
    })

@app.route('/', methods=['GET'])
def api_info():
    """API information and usage"""
    return jsonify({
        'name': 'Yolwise Lead Scoring API',
        'version': '1.0',
        'description': 'Turkish B2B market lead scoring with three-layer architecture',
        'endpoints': {
            '/health': 'GET - Health check',
            '/score_company': 'POST - Score single company (requires API key)',
            '/score_batch': 'POST - Score multiple companies (requires API key)', 
            '/industries': 'GET - List supported industries with multipliers',
            '/': 'GET - This API information'
        },
        'authentication': 'API key required (X-API-Key header or api_key parameter)',
        'specification': 'Based on yolwise_scoring_specification.md',
        'target_threshold': 60,
        'deployment': 'https://yolwiseleadscoring.replit.app'
    })

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)