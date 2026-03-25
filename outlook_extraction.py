"""
Extract keywords directly from Outlook emails.
"""

import sys
import os
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.pipeline.data_miner import DataMiner
from src.connectors.outlook_connector import OutlookConnector


def main():
    # Configuration
    MY_KEYWORDS = [
        "artificial intelligence",
        "AI",
        "machine learning",
        "Genpact",
        "International Women's Day",
        "Empower",
        "auto finance",
        "Economic Times",
        "urgent",
        "deadline"
    ]
    
    OUTPUT_FOLDER = "outputs/outlook_results"
    
    print("\n" + "="*60)
    print("📧 OUTLOOK EMAIL KEYWORD MINER")
    print("="*60)
    
    # Step 1: Connect to Outlook
    print("\n1. Connecting to Outlook...")
    connector = OutlookConnector()
    
    if not connector.outlook:
        print("❌ Could not connect to Outlook. Make sure Outlook is running.")
        print("   Alternative: Save emails to files and use run_extraction.py")
        return
    
    # Step 2: Choose extraction mode
    print("\n2. Choose extraction mode:")
    print("   [1] Extract recent emails (last 7 days)")
    print("   [2] Extract unread emails only")
    print("   [3] Search by subject")
    print("   [4] Process all inbox emails")
    
    mode = input("\nEnter choice (1-4): ").strip()
    
    # Step 3: Extract emails based on mode
    print("\n3. Extracting emails...")
    
    if mode == "1":
        days = input("How many days back? (default: 7): ").strip()
        days = int(days) if days else 7
        emails = connector.get_emails(
            folder="Inbox",
            days_back=days,
            limit=100,
            unread_only=False
        )
    
    elif mode == "2":
        emails = connector.get_emails(
            folder="Inbox",
            days_back=30,
            limit=100,
            unread_only=True
        )
    
    elif mode == "3":
        subject_filter = input("Enter subject keyword: ").strip()
        emails = connector.get_emails(
            folder="Inbox",
            days_back=30,
            limit=100,
            unread_only=False,
            subject_filter=subject_filter
        )
    
    else:
        emails = connector.get_emails(
            folder="Inbox",
            days_back=30,
            limit=200,
            unread_only=False
        )
    
    print(f"   Found {len(emails)} emails")
    
    if not emails:
        print("   No emails found matching criteria.")
        return
    
    # Step 4: Save emails to files
    print("\n4. Saving emails to files...")
    saved_files = connector.save_emails_to_files(emails, output_dir="data/outlook_emails")
    print(f"   Saved {len(saved_files)} emails")
    
    # Step 5: Initialize miner
    print("\n5. Initializing keyword miner...")
    miner = DataMiner(relevance_threshold=0.15)
    
    # Step 6: Process each email
    print("\n6. Processing emails for keywords...")
    all_results = []
    
    for i, email_file in enumerate(saved_files, 1):
        print(f"\n   [{i}/{len(saved_files)}] Processing: {os.path.basename(email_file)}")
        
        try:
            results = miner.mine_document(
                document_path=email_file,
                seed_keywords=MY_KEYWORDS,
                output_dir=f"{OUTPUT_FOLDER}/email_{i}"
            )
            
            # Add email metadata to results
            results['email_metadata'] = {
                'subject': results.get('metadata', {}).get('document', ''),
                'processed': True
            }
            
            all_results.append(results)
            
            # Show quick result
            keywords_found = len(results.get('keywords', {}))
            if keywords_found > 0:
                print(f"      ✅ Found {keywords_found} keywords")
                for kw, data in results.get('keywords', {}).items():
                    print(f"         - {kw}: {data.get('confidence', 0)}%")
            else:
                print(f"      ⚠️ No keywords found")
                
        except Exception as e:
            print(f"      ❌ Error: {e}")
            continue
    
    # Step 7: Generate summary report
    print("\n7. Generating summary report...")
    generate_summary_report(all_results, MY_KEYWORDS, OUTPUT_FOLDER)
    
    print("\n" + "="*60)
    print("✅ Outlook extraction complete!")
    print(f"   Processed: {len(all_results)} emails")
    print(f"   Results saved to: {OUTPUT_FOLDER}")
    print("="*60)


def generate_summary_report(results, keywords, output_folder):
    """Generate a summary report of all processed emails."""
    
    import json
    
    # Calculate statistics
    total_emails = len(results)
    emails_with_keywords = len([r for r in results if r.get('keywords')])
    total_keywords_found = sum(len(r.get('keywords', {})) for r in results)
    
    # Keyword frequency across all emails
    keyword_freq = {}
    for keyword in keywords:
        keyword_freq[keyword] = 0
    
    for result in results:
        for keyword in result.get('keywords', {}):
            if keyword in keyword_freq:
                keyword_freq[keyword] += 1
    
    # Prepare summary
    summary = {
        'processed_date': datetime.now().isoformat(),
        'total_emails': total_emails,
        'emails_with_keywords': emails_with_keywords,
        'success_rate': (emails_with_keywords / total_emails * 100) if total_emails > 0 else 0,
        'total_keywords_found': total_keywords_found,
        'keyword_frequency': keyword_freq,
        'detailed_results': results
    }
    
    # Save summary
    summary_file = f"{output_folder}/summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    os.makedirs(output_folder, exist_ok=True)
    
    with open(summary_file, 'w') as f:
        json.dump(summary, f, indent=2, default=str)
    
    # Print summary
    print("\n" + "-" * 40)
    print("📊 SUMMARY REPORT")
    print("-" * 40)
    print(f"Total emails processed: {total_emails}")
    print(f"Emails with keywords: {emails_with_keywords}")
    print(f"Success rate: {summary['success_rate']:.1f}%")
    print(f"Total keywords found: {total_keywords_found}")
    
    print("\nTop keywords found:")
    sorted_keywords = sorted(keyword_freq.items(), key=lambda x: x[1], reverse=True)
    for kw, count in sorted_keywords[:10]:
        if count > 0:
            print(f"  • {kw}: {count} email(s)")


if __name__ == "__main__":
    main()
