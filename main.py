#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Yolwise B2B Scoring API for Turkish Market - MATHEMATICAL ONLY
CORRECTED VERSION: Only mathematical scoring, NO LLM-based logic
Adapted from Smartway B2B Scoring with Turkish market adjustments
Created using Context7 documentation Flask and mathematical scoring principles
"""

import os
import json
import time
import re
import functools
from typing import Dict, List, Any, Optional
from dataclasses import dataclass, asdict
from flask import Flask, request, jsonify, abort
from werkzeug.exceptions import HTTPException

app = Flask(__name__)


# API Key Authentication Decorator (Context7 Flask Pattern)
def require_api_key(f):
    """
    API key authentication decorator following Context7 Flask patterns.
    Checks for API key in X-API-Key header or api_key query parameter.
    Validates against YOLWISE_API_KEY environment variable.
    """
    @functools.wraps(f)
    def decorated_function(*args, **kwargs):
        # Get expected API key from environment (Context7 pattern)
        expected_api_key = os.environ.get('YOLWISE_API_KEY')
        
        if not expected_api_key:
            abort(500, description="Server configuration error: API key not configured")
        
        # Check X-API-Key header (Context7 Flask request pattern)
        provided_api_key = request.headers.get('X-API-Key')
        
        # If not in header, check query parameter (Context7 Flask args pattern)
        if not provided_api_key:
            provided_api_key = request.args.get('api_key')
        
        # Validate API key (Context7 authentication pattern)
        if not provided_api_key or provided_api_key != expected_api_key:
            abort(401, description="Unauthorized: Valid API key required")
        
        return f(*args, **kwargs)
    return decorated_function


@dataclass
class CompanyScore:
    """Mathematical scoring result for Turkish B2B market"""
    company_name: str
    base_score: float
    industry_multiplier: float
    industry_adjusted_score: float
    detected_industry: str
    industry_confidence: str
    processing_time_ms: int

    # For classification guidance (not LLM-based)
    priority_recommendation: str  # "target" or "non_target"
    industry_explanation: str
    score_breakdown: Dict[str, float]


class YolwiseScoring:
    """
    Mathematical B2B scoring for Turkish market
    MATHEMATICAL ONLY - No LLM logic, pure calculation-based
    """

    def __init__(self):
        # Turkish B2B industry modifiers (mathematical, data-driven)
        self.industry_modifiers = {
            'renewables_environment': {
                'multiplier': 1.20,
                'confidence': 'high',
                'keywords': ['renewable', 'environment', 'solar', 'wind', 'green energy', 'sustainability', 'çevre', 'yenilenebilir'],
                'reasoning': '60% target rate, growing sector with significant B2B investment'
            },
            'logistics_supply_chain': {
                'multiplier': 1.17,
                'confidence': 'high', 
                'keywords': ['logistics', 'supply chain', 'transportation', 'shipping', 'distribution', 'freight', 'lojistik', 'taşımacılık'],
                'reasoning': '58.8% target rate, complex operational B2B needs'
            },
            'utilities': {
                'multiplier': 1.15,
                'confidence': 'high',
                'keywords': ['utilities', 'electric', 'power', 'gas', 'water', 'energy distribution', 'elektrik', 'enerji'],
                'reasoning': '54.5% target rate, infrastructure-heavy B2B requirements'
            },
            'food_beverages': {
                'multiplier': 1.10,
                'confidence': 'medium',
                'keywords': ['food', 'beverage', 'dairy', 'agriculture', 'nutrition', 'farming', 'gıda', 'içecek', 'tarım'],
                'reasoning': '50% target rate, moderate B2B service requirements'
            },
            'chemicals': {
                'multiplier': 1.05,
                'confidence': 'medium',
                'keywords': ['chemical', 'pharmaceutical', 'biotech', 'laboratory', 'manufacturing chemical', 'kimya', 'ilaç'],
                'reasoning': '45.8% target rate, specialized technical services needed'
            },
            'building_materials': {
                'multiplier': 1.02,
                'confidence': 'medium',
                'keywords': ['building materials', 'construction materials', 'cement', 'steel', 'concrete', 'yapı malzemesi', 'çimento'],
                'reasoning': '44.4% target rate, construction-related B2B services'
            },
            'mechanical_industrial': {
                'multiplier': 1.00,
                'confidence': 'medium',
                'keywords': ['mechanical', 'industrial engineering', 'machinery', 'equipment', 'manufacturing', 'makine', 'mühendislik'],
                'reasoning': '42% target rate, baseline engineering service needs'
            },
            'mining_metals': {
                'multiplier': 0.98,
                'confidence': 'medium',
                'keywords': ['mining', 'metals', 'metallurgy', 'extraction', 'ore processing', 'maden', 'metal'],
                'reasoning': '41.2% target rate, specialized but limited service scope'
            },
            'pharmaceuticals': {
                'multiplier': 0.97,
                'confidence': 'medium',
                'keywords': ['pharmaceuticals', 'pharma', 'medicine', 'drugs', 'healthcare', 'ilaç', 'sağlık'],
                'reasoning': '40% target rate, highly regulated sector'
            },
            'retail': {
                'multiplier': 0.90,
                'confidence': 'low',
                'keywords': ['retail', 'consumer', 'shopping', 'store', 'commerce', 'sales', 'perakende', 'mağaza'],
                'reasoning': '37.5% target rate, consumer-focused with limited B2B needs'
            },
            'construction': {
                'multiplier': 0.88,
                'confidence': 'low',
                'keywords': ['construction', 'building', 'architecture', 'contractor', 'infrastructure', 'inşaat', 'yapı'],
                'reasoning': '35.3% target rate, project-based service requirements'
            },
            'automotive': {
                'multiplier': 0.85,
                'confidence': 'low',
                'keywords': ['automotive', 'automobile', 'vehicle', 'car', 'transportation equipment', 'otomotiv', 'araç'],
                'reasoning': '34.1% target rate, manufacturing-focused operations'
            },
            'computer_software': {
                'multiplier': 0.80,
                'confidence': 'low',
                'keywords': ['computer software', 'technology', 'software', 'it', 'digital', 'programming', 'yazılım', 'teknoloji'],
                'reasoning': '29.6% target rate, self-service digital solutions preferred'
            },
            'hospital_healthcare': {
                'multiplier': 0.75,
                'confidence': 'low',
                'keywords': ['hospital', 'health care', 'medical', 'healthcare', 'clinic', 'pharmaceutical', 'hastane', 'sağlık'],
                'reasoning': '30% target rate, specialized service requirements outside typical B2B'
            },
            'transportation_trucking': {
                'multiplier': 0.70,
                'confidence': 'low',
                'keywords': ['transportation', 'trucking', 'freight', 'delivery', 'shipping', 'nakliye', 'kargo'],
                'reasoning': '23.1% target rate, operational focus over service procurement'
            }
        }

        # Turkish geographic tiers
        self.city_tiers = {
            'tier_1_cities': ['istanbul', 'ankara', 'izmir'],
            'tier_2_cities': ['bursa', 'antalya', 'gaziantep', 'konya', 'adana', 'mersin', 'diyarbakır', 'kayseri'],
            'tier_3_cities': ['eskişehir', 'denizli', 'samsun', 'malatya', 'erzurum', 'van', 'batman', 'şanlıurfa'],
            'industrial_regions': ['kocaeli', 'tekirdağ', 'gebze', 'sakarya', 'çorlu', 'manisa']
        }

        # Mathematical scoring weights
        self.scoring_weights = {
            'company_size_indicator': 0.35,
            'industry_b2b_propensity': 0.25,
            'financial_capacity': 0.20,
            'geographic_presence': 0.10,
            'additional_indicators': 0.10
        }

    def calculate_score(self, company_name: str, company_data: Dict[str, Any]) -> CompanyScore:
        """Main mathematical scoring method - NO LLM logic"""
        start_time = time.time()

        # 1. Base mathematical scoring
        base_components = self._calculate_base_components(company_data)
        base_score = sum(score * weight for score, weight in zip(
            base_components.values(), self.scoring_weights.values()))
        base_score = max(0, min(100, base_score))

        # 2. Industry detection and multiplication (mathematical)
        industry, multiplier, confidence = self._detect_industry(company_name, company_data)

        # 3. Apply industry mathematical modifier
        industry_adjusted_score = base_score * multiplier
        industry_adjusted_score = max(0, min(100, industry_adjusted_score))

        # 4. Mathematical priority recommendation (threshold-based)
        priority_recommendation = "target" if industry_adjusted_score >= 60 else "non_target"

        processing_time = int((time.time() - start_time) * 1000)

        return CompanyScore(
            company_name=company_name,
            base_score=round(base_score, 1),
            industry_multiplier=multiplier,
            industry_adjusted_score=round(industry_adjusted_score, 1),
            detected_industry=industry,
            industry_confidence=confidence,
            processing_time_ms=processing_time,
            priority_recommendation=priority_recommendation,
            industry_explanation=self._get_industry_explanation(industry),
            score_breakdown=base_components
        )

    def _calculate_base_components(self, data: Dict[str, Any]) -> Dict[str, float]:
        """Calculate mathematical base components"""
        return {
            'company_size_score': self._evaluate_company_size(data),
            'industry_propensity_score': self._evaluate_industry_propensity(data),
            'financial_capacity_score': self._evaluate_financial_capacity(data),
            'geographic_score': self._evaluate_geographic_presence(data),
            'additional_score': self._evaluate_additional_indicators(data)
        }

    def _evaluate_company_size(self, data: Dict[str, Any]) -> float:
        """Company Size Mathematical Evaluation (35% weight)"""
        score = 30  # Base score

        # Employee count mathematical evaluation
        employees = self._extract_number(data.get('number_of_employees', data.get('employees_estimate', 0)))
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

        # Revenue mathematical evaluation (Turkish Lira context)
        revenue = self._safe_float(data.get('annual_revenue', data.get('revenue_estimate', 0)))
        if revenue >= 1000000000:  # 1B+ TL
            score += 25  # 70.3% target rate
        elif revenue >= 200000000:  # 200-1000M TL
            score += 20  # 60.5% target rate
        elif revenue >= 100000000:   # 100-200M TL
            score += 15  # 50% target rate
        elif revenue >= 20000000:   # 20-100M TL
            score += 10  # 36.4% target rate
        elif revenue > 0:
            score += 3   # <20M TL - 20% target rate

        return min(score, 100)

    def _evaluate_industry_propensity(self, data: Dict[str, Any]) -> float:
        """Mathematical industry B2B propensity evaluation (25% weight)"""
        industry = str(data.get('industry', '')).lower()

        # Mathematical classification by B2B service potential
        high_b2b = [
            'renewable', 'logistics', 'utilities', 'manufacturing', 'energy',
            'chemical', 'industrial', 'engineering', 'construction materials'
        ]
        medium_b2b = [
            'food', 'pharmaceutical', 'building', 'automotive', 'mining',
            'metals', 'machinery', 'equipment'
        ]
        low_b2b = [
            'retail', 'consumer', 'software', 'it', 'healthcare', 'hospital',
            'transportation', 'trucking'
        ]

        for high_term in high_b2b:
            if high_term in industry:
                return 85

        for medium_term in medium_b2b:
            if medium_term in industry:
                return 60

        for low_term in low_b2b:
            if low_term in industry:
                return 25

        return 40  # Default for unclassified

    def _evaluate_financial_capacity(self, data: Dict[str, Any]) -> float:
        """Mathematical financial capacity evaluation (20% weight)"""
        score = 30  # Base score

        # Company age mathematical indicator
        year_founded = self._safe_float(data.get('year_founded', 0))
        if year_founded > 0:
            age = 2025 - year_founded
            if age >= 20:
                score += 25  # Established company
            elif age >= 10:
                score += 20  # Growing company
            elif age >= 5:
                score += 15  # Young company
            else:
                score += 5   # Startup

        # Company structure mathematical evaluation
        company_name = str(data.get('company_name', '')).lower()
        if any(term in company_name for term in ['a.ş.', 'anonim şirket', 'limited şirket', 'ltd.', 'şti.']):
            score += 20  # Formal company structure

        # Domain presence (mathematical indicator of establishment)
        if data.get('company_domain_name'):
            score += 15

        return min(score, 100)

    def _evaluate_geographic_presence(self, data: Dict[str, Any]) -> float:
        """Mathematical geographic evaluation for Turkish market (10% weight)"""
        score = 40  # Base score for Turkish market

        city = str(data.get('city', data.get('headquarters', ''))).lower()
        
        # Mathematical tier-based scoring
        if any(tier1 in city for tier1 in self.city_tiers['tier_1_cities']):
            score += 30  # Istanbul, Ankara, Izmir
        elif any(tier2 in city for tier2 in self.city_tiers['tier_2_cities']):
            score += 25  # Major cities
        elif any(tier3 in city for tier3 in self.city_tiers['tier_3_cities']):
            score += 20  # Regional centers
        elif any(industrial in city for industrial in self.city_tiers['industrial_regions']):
            score += 22  # Industrial zones
        else:
            score += 10  # Other locations

        return min(score, 100)

    def _evaluate_additional_indicators(self, data: Dict[str, Any]) -> float:
        """Mathematical additional indicators (10% weight)"""
        score = 30  # Base score

        # Description completeness (mathematical data quality indicator)
        description = str(data.get('description', ''))
        if len(description) > 100:
            score += 25
        elif len(description) > 50:
            score += 15
        elif len(description) > 0:
            score += 10

        # Contact information completeness (mathematical indicator)
        contact_score = 0
        if data.get('phone_number'):
            contact_score += 15
        if data.get('company_domain_name'):
            contact_score += 15
        
        score += contact_score

        return min(score, 100)

    def _detect_industry(self, company_name: str, company_data: Dict[str, Any]) -> tuple:
        """Mathematical industry detection using keyword matching"""
        
        # Combine text data for analysis
        text_data = [
            company_name.lower(),
            str(company_data.get('industry', '')).lower(),
            str(company_data.get('description', '')).lower()
        ]
        combined_text = ' '.join(text_data)

        # Mathematical keyword matching with scoring
        best_match = None
        max_score = 0

        for industry_name, industry_info in self.industry_modifiers.items():
            score = 0
            for keyword in industry_info['keywords']:
                if keyword in combined_text:
                    # Mathematical weight: longer keywords = higher score
                    score += len(keyword) * (2 if len(keyword) > 5 else 1)

            if score > max_score:
                max_score = score
                best_match = industry_name

        if best_match and max_score > 0:
            industry_data = self.industry_modifiers[best_match]
            return best_match, industry_data['multiplier'], industry_data['confidence']
        else:
            return 'other', 1.0, 'low'

    def _get_industry_explanation(self, industry: str) -> str:
        """Generate explanation for mathematical industry modifier"""
        if industry in self.industry_modifiers:
            data = self.industry_modifiers[industry]
            return f"Industry: {industry} | Multiplier: ×{data['multiplier']:.2f} | {data['reasoning']}"
        else:
            return f"Industry: {industry} | Multiplier: ×1.0 | Standard mathematical evaluation"

    def _safe_float(self, value) -> float:
        """Safe mathematical float conversion"""
        try:
            if isinstance(value, (int, float)):
                return float(value)
            elif isinstance(value, str):
                numbers = re.findall(r'\d+\.?\d*', str(value))
                return float(numbers[0]) if numbers else 0
            else:
                return 0
        except (ValueError, TypeError, AttributeError):
            return 0

    def _extract_number(self, text: str) -> int:
        """Mathematical number extraction"""
        try:
            if isinstance(text, (int, float)):
                return int(text)
            
            text_str = str(text).strip().lower()
            
            # Handle Turkish number formats
            if 'bin' in text_str or 'k' in text_str:
                number = re.search(r'(\d+(?:\.\d+)?)', text_str)
                if number:
                    return int(float(number.group(1)) * 1000)
            elif 'milyon' in text_str or 'm' in text_str:
                number = re.search(r'(\d+(?:\.\d+)?)', text_str)  
                if number:
                    return int(float(number.group(1)) * 1000000)
            
            # Extract first complete number
            numbers = re.findall(r'\d+', text_str)
            return int(numbers[0]) if numbers else 0
        except (ValueError, TypeError, AttributeError):
            return 0


# Initialize mathematical scoring engine
scoring_engine = YolwiseScoring()


# Flask Error Handlers (from Context7 Flask documentation)
@app.errorhandler(HTTPException)
def handle_exception(e):
    """Return JSON instead of HTML for HTTP errors"""
    return jsonify({
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
        "error": "Internal Server Error",
        "message": str(e),
        "code": 500
    }), 500


@app.route('/health', methods=['GET'])
def health_check():
    """Health check endpoint - Public access"""
    return jsonify({
        'status': 'healthy',
        'timestamp': time.time(),
        'version': '2.0-mathematical',
        'description': 'Yolwise B2B Scoring API - Mathematical Only (No LLM)',
        'specification': 'Turkish B2B market mathematical scoring'
    })


@app.route('/score_company', methods=['POST'])
@require_api_key
def score_company():
    """Mathematical scoring of single company - Requires API key authentication"""
    try:
        # Basic validation (from Context7 Flask patterns)
        if not request.json:
            abort(400, description="JSON payload required")

        data = request.json
        company_name = data.get('company_name')

        if not company_name:
            abort(400, description="company_name is required")

        company_data = data.get('company_data', {})

        # Mathematical scoring only
        result = scoring_engine.calculate_score(company_name, company_data)

        return jsonify({
            'success': True,
            'result': asdict(result),
            'metadata': {
                'api_version': '2.0-mathematical',
                'scoring_type': 'mathematical_only',
                'target_threshold': 60,
                'processing_time_ms': result.processing_time_ms
            }
        })

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/score_batch', methods=['POST'])
@require_api_key
def score_batch():
    """Mathematical batch scoring - Requires API key authentication"""
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

                # Mathematical scoring
                result = scoring_engine.calculate_score(company_name, company_data)

                results.append({
                    'company_name': result.company_name,
                    'base_score': result.base_score,
                    'industry_adjusted_score': result.industry_adjusted_score,
                    'priority_recommendation': result.priority_recommendation,
                    'detected_industry': result.detected_industry,
                    'industry_multiplier': result.industry_multiplier,
                    'confidence': result.industry_confidence
                })

            except Exception as e:
                results.append({
                    'company_name': company_name if 'company_name' in locals() else 'Unknown',
                    'base_score': 0,
                    'industry_adjusted_score': 0,
                    'priority_recommendation': 'error',
                    'error': str(e)
                })

        # Sort by mathematical score
        results.sort(key=lambda x: x.get('industry_adjusted_score', 0), reverse=True)

        # Mathematical statistics
        target_count = len([r for r in results if r.get('priority_recommendation') == 'target'])
        processing_time = int((time.time() - start_time) * 1000)

        return jsonify({
            'success': True,
            'results': results,
            'summary': {
                'total_companies': len(results),
                'target_recommendations': target_count,
                'non_target_recommendations': len(results) - target_count,
                'processing_time_ms': processing_time,
                'top_targets': [r for r in results if r.get('priority_recommendation') == 'target'][:10]
            },
            'metadata': {
                'api_version': '2.0-mathematical',
                'scoring_type': 'mathematical_only',
                'target_threshold': 60
            }
        })

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/industries', methods=['GET'])
def get_industries():
    """Get mathematical industry multipliers - Public access"""
    industries = {}
    for industry, data in scoring_engine.industry_modifiers.items():
        industries[industry] = {
            'multiplier': data['multiplier'],
            'confidence': data['confidence'],
            'reasoning': data['reasoning'],
            'keywords': data['keywords'][:5]  # Limit for response size
        }
    
    return jsonify({
        'success': True,
        'industries': industries,
        'metadata': {
            'total_industries': len(industries),
            'default_multiplier': 1.0,
            'target_threshold': 60,
            'scoring_type': 'mathematical_only'
        }
    })


@app.route('/', methods=['GET'])
def api_info():
    """API information - Public access"""
    return jsonify({
        'name': 'Yolwise Lead Scoring API',
        'version': '2.0-mathematical',
        'description': 'Turkish B2B market mathematical scoring - NO LLM logic',
        'authentication': {
            'required_endpoints': ['/score_company', '/score_batch'],
            'public_endpoints': ['/health', '/industries', '/'],
            'methods': [
                'Header: X-API-Key: your-api-key',
                'Query parameter: ?api_key=your-api-key'
            ],
            'environment_variable': 'YOLWISE_API_KEY'
        },
        'endpoints': {
            '/health': 'GET - Health check (public)',
            '/score_company': 'POST - Score single company (requires API key)',
            '/score_batch': 'POST - Score multiple companies (requires API key)',
            '/industries': 'GET - List supported industries with multipliers (public)',
            '/': 'GET - This API information (public)'
        },
        'scoring_approach': 'Mathematical only - no AI/LLM components',
        'target_threshold': 60,
        'specification': 'Turkish B2B market adapted from Smartway mathematical model'
    })


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
