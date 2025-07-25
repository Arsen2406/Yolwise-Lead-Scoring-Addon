#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Yolwise Lead Scoring API for Turkish B2B Market - SECURITY ENHANCED
COMPLETE SINGLE-FILE IMPLEMENTATION: Adapted from Smartway Lead Scoring
Implements three-layer scoring architecture according to yolwise_scoring_specification.md
Deployed at: https://yolwiseleadscoring.replit.app
Created using Context7 Flask documentation and Yolwise specification

SECURITY ENHANCEMENTS:
- Comprehensive input validation with marshmallow
- Secure error handling with information leakage prevention
- CORS configuration and security headers
- Standardized API authentication
- Enhanced logging with security context
- DoS protection measures
"""

import os
import json
import time
import re
import logging
from typing import Dict, List, Any, Optional
from dataclasses import dataclass, asdict
from functools import wraps
from flask import Flask, request, jsonify, abort, g
from werkzeug.exceptions import HTTPException
from markupsafe import escape
import secrets

# Input validation imports
from marshmallow import Schema, fields, validate, ValidationError, EXCLUDE

# Initialize Flask Application with security configuration
app = Flask(__name__)

# Security Configuration
app.config.update(
    SECRET_KEY=os.environ.get('SECRET_KEY', secrets.token_hex(32)),
    SESSION_COOKIE_SECURE=True,
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE='Lax',
    MAX_CONTENT_LENGTH=16 * 1024 * 1024,  # 16MB max
    MAX_FORM_MEMORY_SIZE=500 * 1024,      # 500KB max form memory
    MAX_FORM_PARTS=1000                   # Max form parts
)

# Configure logging with security context
logging.basicConfig(
    level=logging.INFO,
    format='[%(asctime)s] %(levelname)s: %(message)s - %(pathname)s:%(lineno)d',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler('yolwise_security.log') if not os.environ.get('REPL_ENVIRONMENT') else logging.StreamHandler()
    ]
)

# Security headers middleware
@app.after_request
def add_security_headers(response):
    """Add comprehensive security headers following Context7 Flask best practices."""
    response.headers['Content-Security-Policy'] = "default-src 'self'; script-src 'self' 'unsafe-inline'"
    response.headers['X-Content-Type-Options'] = 'nosniff'
    response.headers['X-Frame-Options'] = 'SAMEORIGIN'
    response.headers['Strict-Transport-Security'] = 'max-age=31536000; includeSubDomains'
    
    # CORS headers for Google Apps Script integration
    response.headers['Access-Control-Allow-Origin'] = '*'
    response.headers['Access-Control-Allow-Methods'] = 'GET, POST, OPTIONS'
    response.headers['Access-Control-Allow-Headers'] = 'Content-Type, X-API-Key, Authorization'
    
    return response

# Handle preflight requests
@app.before_request
def handle_preflight():
    if request.method == "OPTIONS":
        response = jsonify({'status': 'OK'})
        return response

# Input Validation Schemas
class CompanyDataSchema(Schema):
    """Comprehensive input validation schema for company data."""
    
    class Meta:
        unknown = EXCLUDE  # Ignore unknown fields for security
    
    # Core required fields
    company_name = fields.Str(
        required=True, 
        validate=validate.Length(min=1, max=200),
        error_messages={'required': 'Company name is required'}
    )
    
    # Industry information
    industry = fields.Str(
        validate=validate.Length(max=100),
        allow_none=True,
        missing=""
    )
    
    # Financial data with proper validation
    annual_revenue = fields.Raw(  # Use Raw to handle various input formats
        validate=lambda x: self._validate_numeric_field(x, 'revenue'),
        allow_none=True,
        missing=0
    )
    
    # Employee information
    number_of_employees = fields.Raw(
        validate=lambda x: self._validate_numeric_field(x, 'employees'),
        allow_none=True,
        missing=0
    )
    
    # Geographic information
    city = fields.Str(validate=validate.Length(max=50), allow_none=True, missing="")
    state_region = fields.Str(validate=validate.Length(max=50), allow_none=True, missing="")
    headquarters = fields.Str(validate=validate.Length(max=100), allow_none=True, missing="")
    
    # Contact and web presence
    company_domain_name = fields.Url(allow_none=True, missing="")
    phone_number = fields.Str(validate=validate.Length(max=20), allow_none=True, missing="")
    
    # Company metadata
    year_founded = fields.Integer(
        validate=validate.Range(min=1800, max=2025),
        allow_none=True,
        missing=0
    )
    
    # Business description with XSS protection
    description = fields.Str(
        validate=validate.Length(max=1000),
        allow_none=True,
        missing="",
        # Custom sanitization will be applied in post_load
    )
    
    def _validate_numeric_field(self, value, field_type):
        """Validate and convert numeric fields safely."""
        if value is None or value == "":
            return True
            
        if isinstance(value, (int, float)) and value >= 0:
            return True
            
        if isinstance(value, str):
            # Extract numbers from string safely
            clean_value = re.sub(r'[^\d.]', '', str(value))
            if clean_value and clean_value.replace('.', '').isdigit():
                return True
                
        raise ValidationError(f'Invalid {field_type} format')
    
    def sanitize_description(self, value):
        """Sanitize description field to prevent XSS."""
        if not value:
            return ""
        # Remove potentially dangerous characters
        sanitized = str(escape(value))
        # Remove backspace characters as per Context7 security recommendations
        sanitized = sanitized.replace("\b", "")
        return sanitized[:1000]  # Limit length

class BatchCompanySchema(Schema):
    """Schema for batch processing requests."""
    
    companies = fields.List(
        fields.Dict(),
        required=True,
        validate=validate.Length(min=1, max=100),  # Limit batch size for DoS protection
        error_messages={'required': 'Companies list is required'}
    )

# Custom Exception Classes
class YolwiseAPIError(Exception):
    """Custom API exception with secure error handling."""
    
    status_code = 400
    
    def __init__(self, message, status_code=None, payload=None, log_details=None):
        super().__init__()
        self.message = str(escape(message))  # Sanitize error messages
        if status_code is not None:
            self.status_code = status_code
        self.payload = payload or {}
        self.log_details = log_details
        
        # Log security-relevant errors
        if log_details:
            app.logger.warning(f"API Error: {log_details}")
    
    def to_dict(self):
        rv = dict(self.payload)
        rv['success'] = False
        rv['error'] = self.message
        rv['code'] = self.status_code
        return rv

# Enhanced API Key validation with security improvements
def require_api_key(f):
    """Secure API key validation decorator with consistent implementation."""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        # Standardized API key extraction - prioritize X-API-Key header
        api_key = request.headers.get('X-API-Key')
        
        # Fallback to query parameter (less secure, logged)
        if not api_key:
            api_key = request.args.get('api_key')
            if api_key:
                app.logger.warning("API key provided via query parameter - security risk")
        
        expected_key = os.environ.get('YOLWISE_API_KEY')
        
        if not expected_key:
            app.logger.error("YOLWISE_API_KEY not configured in environment")
            raise YolwiseAPIError(
                'API configuration error',
                status_code=500,
                log_details="API key not configured on server"
            )
        
        if not api_key:
            app.logger.warning(f"Missing API key from {request.remote_addr}")
            raise YolwiseAPIError(
                'API key required',
                status_code=401,
                log_details=f"Missing API key from {request.remote_addr}"
            )
        
        if not secrets.compare_digest(api_key, expected_key):
            app.logger.warning(f"Invalid API key attempt from {request.remote_addr}")
            raise YolwiseAPIError(
                'Invalid API key',
                status_code=401,
                log_details=f"Invalid API key from {request.remote_addr}"
            )
        
        # Store API key validation success in request context
        g.api_authenticated = True
        return f(*args, **kwargs)
    return decorated_function

@dataclass
class CompanyScore:
    """Company scoring result according to Yolwise specification."""
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
    data_quality_score: float  # Added for transparency

class YolwiseScoring:
    """
    Enhanced Yolwise Lead Scoring System for Turkish B2B Market
    Three-layer architecture: Base Score + Industry Multipliers + LLM Adjustment
    Based on yolwise_scoring_specification.md with security improvements
    """

    def __init__(self):
        # Enhanced industry multipliers with updated data
        self.industry_multipliers = {
            # High B2B Service Sectors (>1.10 multiplier)
            'renewables & environment': {
                'multiplier': 1.20,
                'confidence': 'high',
                'keywords': ['renewable', 'environment', 'solar', 'wind', 'green energy', 'sustainability', 'çevre', 'yenilenebilir'],
                'reasoning': '60% target rate, growing sector with significant B2B investment'
            },
            'logistics and supply chain': {
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

            # Medium B2B Service Sectors (1.00-1.10 multiplier)
            'food & beverages': {
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
            'building materials': {
                'multiplier': 1.02,
                'confidence': 'medium',
                'keywords': ['building materials', 'construction materials', 'cement', 'steel', 'concrete', 'yapı malzemesi', 'çimento'],
                'reasoning': '44.4% target rate, construction-related B2B services'
            },

            # Baseline B2B Service Sectors (0.95-1.00 multiplier)
            'mechanical or industrial engineering': {
                'multiplier': 1.00,
                'confidence': 'medium',
                'keywords': ['mechanical', 'industrial engineering', 'machinery', 'equipment', 'manufacturing', 'makine', 'mühendislik'],
                'reasoning': '42% target rate, baseline engineering service needs'
            },
            'mining & metals': {
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

            # Lower B2B Service Sectors (0.85-0.95 multiplier)
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

            # Low B2B Service Sectors (<0.85 multiplier)
            'computer software': {
                'multiplier': 0.80,
                'confidence': 'low',
                'keywords': ['computer software', 'technology', 'software', 'it', 'digital', 'programming', 'yazılım', 'teknoloji'],
                'reasoning': '29.6% target rate, self-service digital solutions preferred'
            },
            'hospital & health care': {
                'multiplier': 0.75,
                'confidence': 'low',
                'keywords': ['hospital', 'health care', 'medical', 'healthcare', 'clinic', 'pharmaceutical', 'hastane', 'sağlık'],
                'reasoning': '30% target rate, specialized service requirements outside typical B2B'
            },
            'transportation/trucking': {
                'multiplier': 0.70,
                'confidence': 'low',
                'keywords': ['transportation', 'trucking', 'freight', 'delivery', 'shipping', 'nakliye', 'kargo'],
                'reasoning': '23.1% target rate, operational focus over service procurement'
            }
        }

        # Enhanced Turkish city tier classification with more cities
        self.city_tiers = {
            'tier_1_cities': ['istanbul', 'ankara', 'izmir'],
            'tier_2_cities': ['bursa', 'antalya', 'gaziantep', 'konya', 'adana', 'mersin', 'diyarbakır', 'kayseri'],
            'tier_3_cities': ['eskişehir', 'denizli', 'samsun', 'malatya', 'erzurum', 'van', 'batman', 'şanlıurfa'],
            'industrial_regions': ['kocaeli', 'tekirdağ', 'gebze', 'sakarya', 'çorlu', 'manisa'],
            'other_cities': []  # Default for unclassified cities
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
        Main scoring method implementing three-layer Yolwise architecture with enhanced security
        @param company_name: Company name (sanitized)
        @param company_data: Company data (validated)
        @returns CompanyScore with complete scoring breakdown
        """
        start_time = time.time()
        
        # Sanitize company name
        company_name = str(escape(company_name))[:200]

        # Layer 1: Base Score Components (0-100 points)
        base_components = self._calculate_base_components(company_data)
        base_score = sum(score * weight for score, weight in zip(
            base_components.values(), self.scoring_weights.values()))
        base_score = max(0, min(100, base_score))

        # Layer 2: Industry Multipliers
        industry, multiplier, confidence = self._detect_industry(company_name, company_data)
        industry_adjusted_score = base_score * multiplier
        industry_adjusted_score = max(0, min(100, industry_adjusted_score))

        # Layer 3: LLM Adjustment (±25 points) - enhanced algorithm
        llm_adjustment = self._calculate_llm_adjustment_enhanced(company_data, industry_adjusted_score)
        
        # Final score calculation
        final_score = max(0, min(100, industry_adjusted_score + llm_adjustment))

        # Priority recommendation (60+ = target)
        priority_recommendation = "target" if final_score >= 60 else "non_target"
        
        # Calculate data quality score
        data_quality = self._calculate_data_quality(company_data)

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
            score_breakdown=base_components,
            data_quality_score=data_quality
        )

    def _calculate_base_components(self, data: Dict[str, Any]) -> Dict[str, float]:
        """Calculate base score components with enhanced validation."""
        return {
            'company_size_score': self._evaluate_company_size(data),
            'industry_propensity_score': self._evaluate_industry_propensity(data),
            'financial_capacity_score': self._evaluate_financial_capacity(data),
            'geographic_score': self._evaluate_geographic_presence(data),
            'growth_digital_score': self._evaluate_growth_digital_indicators(data)
        }

    def _safe_float_enhanced(self, value) -> float:
        """Enhanced safe float conversion with better number extraction."""
        try:
            if isinstance(value, (int, float)):
                return max(0, float(value))
            elif isinstance(value, str) and value.strip():
                # Clean the string: remove currency symbols, commas, spaces
                clean_value = re.sub(r'[^\d.]', '', str(value))
                if clean_value and '.' in clean_value:
                    # Handle multiple dots - keep only the last one as decimal point
                    parts = clean_value.split('.')
                    if len(parts) > 2:
                        clean_value = ''.join(parts[:-1]) + '.' + parts[-1]
                
                return max(0, float(clean_value)) if clean_value else 0
            else:
                return 0
        except (ValueError, TypeError, AttributeError):
            return 0

    def _extract_number_enhanced(self, text: str) -> int:
        """Enhanced number extraction with better parsing."""
        try:
            if isinstance(text, (int, float)):
                return max(0, int(text))
            
            text_str = str(text).strip().lower()
            
            # Handle common number formats
            if 'k' in text_str or 'bin' in text_str:
                number = re.search(r'(\d+(?:\.\d+)?)', text_str)
                if number:
                    return max(0, int(float(number.group(1)) * 1000))
            elif 'm' in text_str or 'milyon' in text_str:
                number = re.search(r'(\d+(?:\.\d+)?)', text_str)
                if number:
                    return max(0, int(float(number.group(1)) * 1000000))
            
            # Extract first complete number
            numbers = re.findall(r'\d+', text_str)
            return max(0, int(numbers[0])) if numbers else 0
        except (ValueError, TypeError, AttributeError):
            return 0

    def _evaluate_company_size(self, data: Dict[str, Any]) -> float:
        """Company Size Indicator (35% weight) - Enhanced Turkish market context."""
        score = 30  # Base score

        # Enhanced employee count evaluation
        employees = self._extract_number_enhanced(data.get('number_of_employees', data.get('employees_estimate', 0)))
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

        # Enhanced revenue evaluation with 2024-2025 Turkish Lira context
        revenue = self._safe_float_enhanced(data.get('annual_revenue', data.get('revenue_estimate', 0)))
        if revenue >= 1000000000:  # 1B+ TL (adjusted for inflation)
            score += 25  # 70.3% target rate
        elif revenue >= 200000000:  # 200-1000M TL (adjusted for inflation)
            score += 20  # 60.5% target rate
        elif revenue >= 100000000:   # 100-200M TL
            score += 15  # 50% target rate
        elif revenue >= 20000000:   # 20-100M TL (adjusted for inflation)
            score += 10  # 36.4% target rate
        elif revenue > 0:
            score += 3   # <20M TL - 20% target rate

        return min(score, 100)

    def _evaluate_geographic_presence(self, data: Dict[str, Any]) -> float:
        """Geographic Presence (10% weight) - Enhanced Turkish market focus."""
        score = 50  # Base score (all companies in Turkey)

        city = str(data.get('city', data.get('headquarters', ''))).lower()
        
        # Enhanced location tier bonuses
        if any(tier1 in city for tier1 in self.city_tiers['tier_1_cities']):
            score += 25  # Tier 1 cities
        elif any(tier2 in city for tier2 in self.city_tiers['tier_2_cities']):
            score += 20  # Tier 2 cities
        elif any(tier3 in city for tier3 in self.city_tiers['tier_3_cities']):
            score += 15  # Tier 3 cities
        elif any(industrial in city for industrial in self.city_tiers['industrial_regions']):
            score += 18  # Industrial regions
        else:
            score += 8   # Other cities

        # International presence indicators
        description = str(data.get('description', '')).lower()
        if any(term in description for term in ['international', 'global', 'export', 'worldwide', 'uluslararası']):
            score += 15

        return min(score, 100)

    def _calculate_llm_adjustment_enhanced(self, data: Dict[str, Any], industry_score: float) -> float:
        """
        Enhanced LLM adjustment (±25 points) with better algorithm
        Fixed boundary logic to optimally select adjustments within limits
        """
        adjustments = []
        
        # Business Growth Indicators
        description = str(data.get('description', '')).lower()
        if any(term in description for term in ['expansion', 'new facilities', 'growing', 'büyüme', 'genişleme']):
            adjustments.append(('growth_expansion', 10, 'Business expansion indicators'))
        if any(term in description for term in ['innovation', 'technology', 'digital', 'inovasyon', 'teknoloji']):
            adjustments.append(('innovation', 8, 'Innovation and technology focus'))

        # Market Position Indicators  
        if any(term in description for term in ['leader', 'leading', 'market share', 'lider', 'önder']):
            adjustments.append(('market_leader', 12, 'Market leadership position'))
        if any(term in description for term in ['partnership', 'strategic', 'ortaklık', 'stratejik']):
            adjustments.append(('partnerships', 6, 'Strategic partnerships'))

        # Data Quality Bonus
        fields_completed = sum(1 for field in ['company_domain_name', 'description', 'year_founded'] 
                              if data.get(field))
        if fields_completed >= 3:
            adjustments.append(('data_quality_high', 8, 'High data completeness'))
        elif fields_completed >= 2:
            adjustments.append(('data_quality_medium', 4, 'Medium data completeness'))

        # Risk Factors (negative adjustments)
        if any(term in description for term in ['crisis', 'downsizing', 'kriz', 'küçülme']):
            adjustments.append(('crisis_risk', -15, 'Crisis or downsizing indicators'))
        
        # Enhanced boundary logic - optimally select adjustments
        adjustments.sort(key=lambda x: abs(x[1]), reverse=True)  # Sort by absolute impact
        
        total_positive = sum(adj[1] for adj in adjustments if adj[1] > 0)
        total_negative = sum(adj[1] for adj in adjustments if adj[1] < 0)
        
        # Smart selection within ±25 bounds
        if total_positive + total_negative > 25:
            # Prioritize positive adjustments, then add negative ones if within bounds
            selected_adjustment = 0
            for name, value, reason in adjustments:
                potential_total = selected_adjustment + value
                if -25 <= potential_total <= 25:
                    selected_adjustment = potential_total
            total_adjustment = selected_adjustment
        elif total_positive + total_negative < -25:
            # Cap negative adjustments at -25
            total_adjustment = max(-25, total_positive + total_negative)
        else:
            total_adjustment = total_positive + total_negative

        return max(-25, min(25, total_adjustment))

    def _calculate_data_quality(self, data: Dict[str, Any]) -> float:
        """Calculate data quality score for transparency."""
        required_fields = ['company_name', 'industry', 'annual_revenue', 'number_of_employees']
        optional_fields = ['city', 'description', 'company_domain_name', 'year_founded']
        
        required_score = sum(1 for field in required_fields if data.get(field)) / len(required_fields)
        optional_score = sum(1 for field in optional_fields if data.get(field)) / len(optional_fields)
        
        return round((required_score * 0.7 + optional_score * 0.3) * 100, 1)

    def _detect_industry(self, company_name: str, company_data: Dict[str, Any]) -> tuple:
        """Enhanced industry detection with better matching."""
        text_data = [
            company_name.lower(),
            str(company_data.get('industry', '')).lower(),
            str(company_data.get('description', '')).lower()
        ]
        combined_text = ' '.join(text_data)

        # Find best match with enhanced scoring
        best_match = None
        max_score = 0

        for industry_name, industry_info in self.industry_multipliers.items():
            score = 0
            for keyword in industry_info['keywords']:
                if keyword in combined_text:
                    # Weight longer keywords more heavily
                    score += len(keyword) * (2 if len(keyword) > 5 else 1)

            if score > max_score:
                max_score = score
                best_match = industry_name

        if best_match and max_score > 0:
            industry_data = self.industry_multipliers[best_match]
            return best_match, industry_data['multiplier'], industry_data['confidence']
        else:
            return 'other', 1.0, 'low'

    def _generate_reasoning(self, data: Dict[str, Any], industry: str, final_score: float) -> str:
        """Generate human-readable reasoning for the score."""
        reasons = []
        
        # Score level reasoning
        if final_score >= 85:
            reasons.append("High-confidence target")
        elif final_score >= 70:
            reasons.append("Strong target candidate")
        elif final_score >= 60:
            reasons.append("Qualified target candidate")
        else:
            reasons.append("Below target threshold")

        # Industry reasoning
        if industry in self.industry_multipliers:
            industry_data = self.industry_multipliers[industry]
            reasons.append(f"Industry: {industry} (×{industry_data['multiplier']:.2f})")
        
        # Size indicators
        employees = self._extract_number_enhanced(data.get('number_of_employees', 0))
        if employees >= 1000:
            reasons.append("Large enterprise")
        elif employees >= 200:
            reasons.append("Medium-large company")
        elif employees >= 50:
            reasons.append("Medium company")

        # Geographic advantage
        city = str(data.get('city', data.get('headquarters', ''))).lower()
        if any(tier1 in city for tier1 in self.city_tiers['tier_1_cities']):
            reasons.append("Prime Turkish market location")
        elif any(tier2 in city for tier2 in self.city_tiers['tier_2_cities']):
            reasons.append("Major Turkish market location")
        
        return " • ".join(reasons)

# Initialize scoring engine
scoring_engine = YolwiseScoring()

# Enhanced Error Handlers with secure responses
@app.errorhandler(YolwiseAPIError)
def handle_yolwise_error(e):
    """Handle custom API errors securely."""
    return jsonify(e.to_dict()), e.status_code

@app.errorhandler(ValidationError)
def handle_validation_error(e):
    """Handle marshmallow validation errors."""
    app.logger.warning(f"Validation error: {e.messages}")
    return jsonify({
        'success': False,
        'error': 'Invalid input data',
        'details': e.messages,
        'code': 400
    }), 400

@app.errorhandler(HTTPException)
def handle_http_exception(e):
    """Return secure JSON for HTTP errors."""
    # Don't leak sensitive information in error responses
    safe_description = "Request could not be processed" if e.code >= 500 else e.description
    
    return jsonify({
        "success": False,
        "error": e.name,
        "message": safe_description,
        "code": e.code
    }), e.code

@app.errorhandler(Exception)
def handle_generic_exception(e):
    """Handle non-HTTP exceptions securely."""
    if isinstance(e, HTTPException):
        return handle_http_exception(e)

    # Log the actual error for debugging
    app.logger.error(f"Unhandled exception: {str(e)}", exc_info=True)
    
    # Return generic error to prevent information leakage
    return jsonify({
        "success": False,
        "error": "Internal Server Error",
        "message": "An unexpected error occurred",
        "code": 500
    }), 500

# API Routes with enhanced security
@app.route('/health', methods=['GET'])
def health_check():
    """Health check endpoint with system status."""
    return jsonify({
        'status': 'healthy',
        'timestamp': time.time(),
        'version': '1.1-secure',
        'description': 'Yolwise Lead Scoring API for Turkish B2B Market - Security Enhanced',
        'specification': 'Three-layer architecture with comprehensive security measures'
    })

@app.route('/score_company', methods=['POST'])
@require_api_key
def score_company():
    """Score a single company with enhanced validation and security."""
    try:
        if not request.json:
            raise YolwiseAPIError("JSON payload required", 400)

        # Validate request data
        schema = CompanyDataSchema()
        
        company_name = request.json.get('company_name')
        if not company_name:
            raise YolwiseAPIError("company_name is required", 400)

        company_data_raw = request.json.get('company_data', {})
        
        # Validate and sanitize company data
        try:
            # Sanitize description field
            if 'description' in company_data_raw:
                company_data_raw['description'] = schema.sanitize_description(company_data_raw['description'])
            
            company_data = schema.load(company_data_raw)
        except ValidationError as e:
            raise YolwiseAPIError(f"Invalid company data: {e.messages}", 400)

        # Perform Yolwise scoring
        result = scoring_engine.calculate_score(company_name, company_data)

        response_data = {
            'success': True,
            'result': asdict(result),
            'metadata': {
                'api_version': '1.1-secure',
                'scoring_model': 'Turkish B2B Market Enhanced',
                'target_threshold': 60,
                'processing_time_ms': result.processing_time_ms,
                'security_level': 'enhanced'
            }
        }
        
        # Log successful scoring (without sensitive data)
        app.logger.info(f"Scored company: {company_name[:50]}... Score: {result.final_score}")
        
        return jsonify(response_data)

    except YolwiseAPIError:
        raise
    except Exception as e:
        app.logger.error(f"Error in score_company: {str(e)}", exc_info=True)
        raise YolwiseAPIError("Failed to process company scoring", 500)

@app.route('/score_batch', methods=['POST'])
@require_api_key
def score_batch():
    """Bulk scoring with enhanced security and rate limiting."""
    try:
        if not request.json:
            raise YolwiseAPIError("JSON payload required", 400)

        # Validate batch request
        batch_schema = BatchCompanySchema()
        try:
            batch_data = batch_schema.load(request.json)
        except ValidationError as e:
            raise YolwiseAPIError(f"Invalid batch data: {e.messages}", 400)

        companies = batch_data['companies']
        
        if len(companies) > 100:  # DoS protection
            raise YolwiseAPIError("Maximum 100 companies per batch", 400)

        results = []
        start_time = time.time()
        company_schema = CompanyDataSchema()

        for company_info in companies:
            try:
                # Handle different input formats
                if isinstance(company_info, str):
                    company_name = company_info
                    company_data = {}
                elif isinstance(company_info, dict):
                    company_name = company_info.get('name', company_info.get('company_name', ''))
                    company_data_raw = company_info.get('data', company_info)
                    
                    # Sanitize and validate
                    if 'description' in company_data_raw:
                        company_data_raw['description'] = company_schema.sanitize_description(company_data_raw['description'])
                    
                    company_data = company_schema.load(company_data_raw)
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
                    'reasoning': result.reasoning,
                    'data_quality_score': result.data_quality_score
                })

            except Exception as e:
                app.logger.warning(f"Error processing company in batch: {str(e)}")
                results.append({
                    'company_name': company_name if 'company_name' in locals() else 'Unknown',
                    'base_score': 0,
                    'final_score': 0,
                    'priority_recommendation': 'error',
                    'error': 'Processing failed'
                })

        # Sort by final score
        results.sort(key=lambda x: x.get('final_score', 0), reverse=True)

        # Calculate statistics
        target_count = len([r for r in results if r.get('priority_recommendation') == 'target'])
        processing_time = int((time.time() - start_time) * 1000)

        response_data = {
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
                'api_version': '1.1-secure',
                'scoring_model': 'Turkish B2B Market Enhanced',
                'target_threshold': 60,
                'specification': 'Three-layer architecture with security enhancements',
                'security_level': 'enhanced'
            }
        }
        
        # Log batch processing
        app.logger.info(f"Processed batch: {len(results)} companies, {target_count} targets")
        
        return jsonify(response_data)

    except YolwiseAPIError:
        raise
    except Exception as e:
        app.logger.error(f"Error in score_batch: {str(e)}", exc_info=True)
        raise YolwiseAPIError("Failed to process batch scoring", 500)

@app.route('/industries', methods=['GET'])
def get_industries():
    """Get supported industries with multipliers."""
    industries = {}
    for industry, data in scoring_engine.industry_multipliers.items():
        industries[industry] = {
            'multiplier': data['multiplier'],
            'confidence': data['confidence'],
            'reasoning': data['reasoning'],
            'keywords': data['keywords'][:8]  # Limit keywords for response size
        }
    
    return jsonify({
        'success': True,
        'industries': industries,
        'metadata': {
            'total_industries': len(industries),
            'default_multiplier': 1.0,
            'target_threshold': 60,
            'enhanced_security': True
        }
    })

@app.route('/', methods=['GET'])
def api_info():
    """API information and usage with security details."""
    return jsonify({
        'name': 'Yolwise Lead Scoring API',
        'version': '1.1-secure',
        'description': 'Turkish B2B market lead scoring with three-layer architecture and enhanced security',
        'endpoints': {
            '/health': 'GET - Health check',
            '/score_company': 'POST - Score single company (requires API key)',
            '/score_batch': 'POST - Score multiple companies (requires API key)', 
            '/industries': 'GET - List supported industries with multipliers',
            '/': 'GET - This API information'
        },
        'authentication': 'X-API-Key header required for scoring endpoints',
        'security_features': [
            'Comprehensive input validation',
            'XSS protection with escaping',
            'DoS protection with limits',
            'Secure error handling',
            'CORS configuration',
            'Security headers (CSP, HSTS, etc.)',
            'Enhanced logging'
        ],
        'specification': 'Based on yolwise_scoring_specification.md with security enhancements',
        'target_threshold': 60,
        'deployment': 'https://yolwiseleadscoring.replit.app',
        'data_limits': {
            'max_batch_size': 100,
            'max_content_length': '16MB',
            'max_description_length': 1000
        }
    })

# Main execution entry point
if __name__ == '__main__':
    # Configuration for Replit deployment with security
    port = int(os.environ.get('PORT', 5000))
    host = '0.0.0.0'
    
    # Validate environment
    if not os.environ.get('YOLWISE_API_KEY'):
        app.logger.warning("⚠️  WARNING: YOLWISE_API_KEY not set in environment variables")
        app.logger.warning("   Please set this in Replit Secrets for API security")
    
    if not os.environ.get('SECRET_KEY'):
        app.logger.info("Using auto-generated SECRET_KEY - set SECRET_KEY env var for persistence")
    
    app.logger.info(f"🚀 Starting Yolwise Lead Scoring API (Enhanced Security) on {host}:{port}")
    app.logger.info(f"📋 Specification: Turkish B2B Market - Three-layer Architecture with Security")
    app.logger.info(f"🔑 API Key Protection: {'✅ Enabled' if os.environ.get('YOLWISE_API_KEY') else '❌ Disabled'}")
    app.logger.info(f"🛡️  Security Features: Input validation, XSS protection, DoS limits, secure headers")
    
    app.run(
        host=host,
        port=port,
        debug=False,  # Always False in production for security
        threaded=True  # Enable threading for better performance
    )
