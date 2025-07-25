#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Yolwise Lead Scoring API - Main Entry Point for Replit
Adapted from Smartway Lead Scoring for Turkish B2B Market
This file serves as the entry point for Replit deployment
"""

import os
from yolwise_scoring_api import app

if __name__ == '__main__':
    # Configuration for Replit deployment
    port = int(os.environ.get('PORT', 5000))
    host = '0.0.0.0'
    
    # Set API key from environment for security
    if not os.environ.get('YOLWISE_API_KEY'):
        print("⚠️  WARNING: YOLWISE_API_KEY not set in environment variables")
        print("   Please set this in Replit Secrets for API security")
    
    print(f"🚀 Starting Yolwise Lead Scoring API on {host}:{port}")
    print(f"📋 Specification: Turkish B2B Market - Three-layer Architecture")
    print(f"🔑 API Key Protection: {'✅ Enabled' if os.environ.get('YOLWISE_API_KEY') else '❌ Disabled'}")
    
    app.run(
        host=host,
        port=port,
        debug=False,  # Always False in production
        threaded=True  # Enable threading for better performance
    )
