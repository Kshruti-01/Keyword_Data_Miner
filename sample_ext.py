"""
Simple script to mine ONE email file with our keywords and generate a summary.
"""

import sys
import os
import json
from datetime import datetime

# Add project root to path
sys.path.insert(0,os.path.abspath(os.path.dirname(__file__)))
#sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
#sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__),"..")))


#from src.pipeline.data_miner import DataMiner
from src.pipeline.data_miner import DataMiner

def main():

    # 1. Path to email file (stored in data folder)
    EMAIL_FILE = "data/sample.txt"
    
    # 2. Your keywords (what we want to find in the email)
    MY_KEYWORDS = [
        "women",
        "leadership", 
        "healthcare",
        "empower",
        "TalentMatch"
    ]
    
    # 3. Where to save the results
    OUTPUT_FOLDER = "outputs"
    # ============================================================
    
    print("\n" + "="*60)
    print("EMAIL KEYWORD MINER")
    print("="*60)
    
    # Check if file exists
    if not os.path.exists(EMAIL_FILE):
        print(f" Error: File '{EMAIL_FILE}' not found!")
        print(" Please make sure your email is saved in the data folder.")
        return
    
    print(f"\n Reading email from: {EMAIL_FILE}")
    print(f"Looking for keywords: {MY_KEYWORDS}")
    
    # Create the miner
    miner = DataMiner()
    
    # Mine the document with YOUR keywords
    results = miner.mine_document(
        document_path=EMAIL_FILE,
        seed_keywords=MY_KEYWORDS,
        output_dir=OUTPUT_FOLDER
    )
    
    # DISPLAY A NICE SUMMARY
    print("\n" + "="*60)
    print("EXTRACTION SUMMARY")
    print("="*60)
    
    # Get the keywords that were actually found
    found_keywords = results.get('keywords', {})
    
    if found_keywords:
        print(f"\nFound {len(found_keywords)} out of {len(MY_KEYWORDS)} keywords")
        print("\nKEYWORDS FOUND:")
        print("-" * 40)
        
        for keyword, data in found_keywords.items():
            confidence = data.get('confidence', 0)
            occurrences = data.get('occurrences', 0)
            
            # Show confidence with stars
            confidence_percent = int(confidence * 100)
            print(f"\n  • {keyword.upper()}")
            print(f"    Found: {occurrences} time(s) | Confidence: {confidence_percent}%")

            # Show summary if available
            summary = data.get('summary', '')
            if summary:
                print(f"    Summary: {summary[:150]}...")
            
            # Show first context example
            contexts = data.get('contexts', [])
            if contexts:
                ctx = contexts[0]
                print(f"    Example: ...{ctx.get('before', '')[-30:]} {ctx.get('keyword', '')} {ctx.get('after', '')[:30]}...")
    else:
        print("\n No keywords found in the email.")
        print("   Try different keywords or check if the email contains these terms.")
    
    # Show entities found
    entities = results.get('entities', {})
    if entities:
        print("\n" + "-" * 40)
        print("ENTITIES DETECTED:")
        for entity_type, entity_list in entities.items():
            if entity_list:
                print(f"  {entity_type}: {', '.join(entity_list[:5])}")
    
    # Show file locations
    print("\n" + "-" * 40)
    print(" RESULTS SAVED TO:")
    
    # Find the latest files in outputs folder
    output_files = [f for f in os.listdir(OUTPUT_FOLDER) if f.startswith('results_') or f.startswith('report_')]
    json_files = [f for f in output_files if f.startswith('results_')]
    txt_files = [f for f in output_files if f.startswith('report_')]
    
    if json_files:
        print(f" JSON data: {OUTPUT_FOLDER}/{sorted(json_files)[-1]}")
    if txt_files:
        print(f" Text report: {OUTPUT_FOLDER}/{sorted(txt_files)[-1]}")
    
    print("\n" + "="*60)
    print("Done! Check the outputs folder for detailed results.")
    print("="*60)


if __name__ == "__main__":
    main()
