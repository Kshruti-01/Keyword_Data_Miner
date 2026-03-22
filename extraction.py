"""
Simple script to mine ONE email file with YOUR keywords and generate a summary.
"""

import sys
import os
import json
from datetime import datetime

# Add project root to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.pipeline.data_miner import DataMiner


def main():

    # 1. Path to email file (stored in data folder)
    EMAIL_FILE = "data/sample.txt"
    
    # 2. YOUR UPDATED KEYWORDS - MATCHING THE SAMPLE EMAIL
    MY_KEYWORDS = [
        "artificial intelligence",
        "AI",                                    # Added - appears 3 times in email
        "machine learning",
        "Genpact",                               # Added - company name appears multiple times
        "International Women's Day",             # Added - key event
        "Empower",                               # Added - security campaign
        "auto finance",                          # Added - key topic
        "Economic Times",                        # Added - award recognition
        "TalentMatch",                           # Added - hiring tool
        "inclusive leadership",                  # Added - cultural theme
        "employee",                              # Added - people focus
        "awards"                                 # Added - recognition
    ]
    
    # 3. Where to save the results
    OUTPUT_FOLDER = "outputs"
   
    print("\n" + "="*60)
    print("EMAIL KEYWORD MINER")
    print("="*60)
    
    # Check if file exists
    if not os.path.exists(EMAIL_FILE):
        print(f"Error: File '{EMAIL_FILE}' not found!")
        print("  Please make sure your email is saved in the data folder.")
        return
    
    print(f"\nReading email from: {EMAIL_FILE}")
    print(f"Looking for keywords: {MY_KEYWORDS}")
    
    # Create the miner
    miner = DataMiner()
    
    # Mine the document with keywords
    results = miner.mine_document(
        document_path=EMAIL_FILE,
        seed_keywords=MY_KEYWORDS,
        output_dir=OUTPUT_FOLDER
    )
    
    # DISPLAY A SUMMARY
    print("\n" + "="*60)
    print("EXTRACTION SUMMARY")
    print("="*60)
    
    # Get the keywords that were actually found
    found_keywords = results.get('keywords', {})
    
    if found_keywords:
        print(f"\nFound {len(found_keywords)} out of {len(MY_KEYWORDS)} keywords")
        print("\nKEYWORDS FOUND:")
        print("-" * 40)
        
        # Sort by confidence score
        sorted_keywords = sorted(found_keywords.items(), 
                                key=lambda x: x[1].get('confidence', 0), 
                                reverse=True)
        
        for keyword, data in sorted_keywords:
            confidence = data.get('confidence', 0)
            occurrences = data.get('occurrences', 0)
            
            # Show confidence as percentage
            confidence_percent = int(confidence * 100)
            
            # Add confidence indicator
            if confidence_percent >= 70:
                indicator = "HIGH"
            elif confidence_percent >= 40:
                indicator = "MEDIUM"
            else:
                indicator = "LOW"
            
            print(f"\n  • {keyword.upper()}")
            print(f"    Found: {occurrences} time(s) | Confidence: {confidence_percent}% {indicator}")
            
            # Show summary if available
            summary = data.get('summary', '')
            if summary:
                # Limit summary length
                if len(summary) > 150:
                    summary = summary[:150] + "..."
                print(f"    Summary: {summary}")
            
            # Show first context example
            contexts = data.get('contexts', [])
            if contexts:
                ctx = contexts[0]
                before = ctx.get('before', '')[-30:]
                keyword_text = ctx.get('keyword', '')
                after = ctx.get('after', '')[:30]
                print(f"    Example: ...{before}{keyword_text}{after}...")
    else:
        print("\nNo keywords found in the email.")
        print("   Try different keywords or check if the email contains these terms.")
    
    # Show entities found
    entities = results.get('entities', {})
    if entities:
        print("\n" + "-" * 40)
        print("ENTITIES DETECTED:")
        
        # Show entity types in order
        entity_order = ['ORG', 'PERSON', 'DATE', 'EVENT', 'PRODUCT', 'GPE']
        for entity_type in entity_order:
            if entity_type in entities and entities[entity_type]:
                entity_list = entities[entity_type][:5]  # Show top 5
                print(f"  {entity_type}: {', '.join(entity_list)}")
        
        # Show any remaining entity types
        for entity_type, entity_list in entities.items():
            if entity_type not in entity_order and entity_list:
                print(f"  {entity_type}: {', '.join(entity_list[:3])}")
    
    # Show file locations
    print("\n" + "-" * 40)
    print("RESULTS SAVED TO:")
    
    # Find the latest files in outputs folder
    if os.path.exists(OUTPUT_FOLDER):
        output_files = [f for f in os.listdir(OUTPUT_FOLDER) if f.startswith('results_') or f.startswith('report_')]
        json_files = [f for f in output_files if f.startswith('results_')]
        txt_files = [f for f in output_files if f.startswith('report_')]
        
        if json_files:
            print(f"JSON data: {OUTPUT_FOLDER}/{sorted(json_files)[-1]}")
        if txt_files:
            print(f"Text report: {OUTPUT_FOLDER}/{sorted(txt_files)[-1]}")
    else:
        print(f"Output folder: {OUTPUT_FOLDER}/ (created automatically)")
    
    # Show quick statistics
    print("\n" + "-" * 40)
    print("QUICK STATISTICS:")
    print(f"  • Total keywords searched: {len(MY_KEYWORDS)}")
    print(f"  • Keywords found: {len(found_keywords)}")
    print(f"  • Success rate: {int(len(found_keywords)/len(MY_KEYWORDS)*100)}%")
    
    if found_keywords:
        total_occurrences = sum(v.get('occurrences', 0) for v in found_keywords.values())
        print(f"  • Total keyword occurrences: {total_occurrences}")
    
    if entities:
        total_entities = sum(len(v) for v in entities.values())
        print(f"  • Total entities detected: {total_entities}")
    
    print("\n" + "="*60)
    print("Done! Check the outputs folder for detailed results.")
    print("="*60)


if __name__ == "__main__":
    main()