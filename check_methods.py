#!/usr/bin/env python
"""
Quick script to check what methods are available in your DataMiner class.
"""

import sys
import os

# Add project root to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from src.pipeline.data_miner import DataMiner
    
    # Create instance
    miner = DataMiner()
    
    # List all methods
    print("\nMethods available in DataMiner:")
    print("-" * 40)
    
    methods = [method for method in dir(miner) if not method.startswith('_')]
    for method in sorted(methods):
        print(f"  • {method}")
    
    print("\n" + "-" * 40)
    
    # Specifically check for document mining methods
    doc_methods = [m for m in methods if 'document' in m.lower() or 'mine' in m.lower()]
    if doc_methods:
        print("\nDocument mining methods found:")
        for m in doc_methods:
            print(f"  • {m}")
    else:
        print("\n⚠️  No document mining methods found!")
        
except ImportError as e:
    print(f"Error importing: {e}")
except Exception as e:
    print(f"Error: {e}")