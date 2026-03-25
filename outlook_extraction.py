"""
Extract keywords directly from Outlook emails and save as Word document.
"""

import sys
import os
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.pipeline.data_miner import DataMiner
from src.connectors.outlook_connector import OutlookConnector
from src.utils.word_generator import create_quick_report


def main():
    # Configuration - UPDATE THESE WITH YOUR KEYWORDS
    MY_KEYWORDS = [
        "artificial intelligence",
        "AI",
        "machine learning",
        "Genpact",
        "International Women's Day",
        "Empower",
        "auto finance",
        "urgent",
        "deadline",
        "meeting",
        "project"
    ]
    
    OUTPUT_FOLDER = "outputs/outlook_results"
    
    print("\n" + "="*60)
    print("📧 OUTLOOK EMAIL KEYWORD MINER")
    print("="*60)
    
    # Step 1: Connect to Outlook
    print("\n1. Connecting to Outlook...")
    connector = OutlookConnector()
    
    if not connector.outlook:
        print("❌ Could not connect to Outlook.")
        print("\nTroubleshooting:")
        print("  1. Make sure Outlook is open")
        print("  2. Check that pywin32 is installed: pip install pywin32")
        print("  3. Try running as administrator")
        return
    
    # Step 2: Choose extraction mode
    print("\n2. Choose extraction mode:")
    print("   [1] Recent emails (last 7 days)")
    print("   [2] Unread emails only")
    print("   [3] All emails from last 30 days")
    print("   [4] Search by subject")
    
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
        emails = connector.get_emails(
            folder="Inbox",
            days_back=30,
            limit=200,
            unread_only=False
        )
    
    elif mode == "4":
        subject_filter = input("Enter subject keyword: ").strip()
        emails = connector.get_emails(
            folder="Inbox",
            days_back=30,
            limit=100,
            unread_only=False,
            subject_filter=subject_filter
        )
    
    else:
        print("Invalid choice. Using default (last 7 days)")
        emails = connector.get_emails(
            folder="Inbox",
            days_back=7,
            limit=50,
            unread_only=False
        )
    
    print(f"   ✅ Found {len(emails)} emails")
    
    if not emails:
        print("   No emails found matching criteria.")
        return
    
    # Step 4: Save emails to files
    print("\n4. Saving emails to files...")
    os.makedirs("data/outlook_emails", exist_ok=True)
    saved_files = connector.save_emails_to_files(emails, output_dir="data/outlook_emails")
    print(f"   ✅ Saved {len(saved_files)} emails")
    
    # Step 5: Initialize keyword miner
    print("\n5. Initializing keyword miner...")
    miner = DataMiner(relevance_threshold=0.15)
    
    # Step 6: Process each email
    print("\n6. Processing emails for keywords...")
    all_results = []
    emails_with_keywords = 0
    total_keywords_found = 0
    
    for i, email_file in enumerate(saved_files, 1):
        print(f"\n   [{i}/{len(saved_files)}] Processing: {os.path.basename(email_file)}")
        
        try:
            results = miner.mine_document(
                document_path=email_file,
                seed_keywords=MY_KEYWORDS,
                output_dir=f"{OUTPUT_FOLDER}/email_{i}"
            )
            
            # Count keywords found
            keywords_found = len(results.get('keywords', {}))
            if keywords_found > 0:
                emails_with_keywords += 1
                total_keywords_found += keywords_found
                
                print(f"      ✅ Found {keywords_found} keywords:")
                for kw, data in results.get('keywords', {}).items():
                    confidence = int(data.get('confidence', 0) * 100)
                    print(f"         - {kw}: {confidence}% confidence")
            else:
                print(f"      ⚠️ No keywords found")
            
            # Add email metadata
            results['email_metadata'] = {
                'subject': emails[i-1]['subject'],
                'sender': emails[i-1]['sender_name'],
                'date': emails[i-1]['received_time']
            }
            
            all_results.append(results)
                
        except Exception as e:
            print(f"      ❌ Error: {e}")
            continue
    
    # Step 7: Generate Word Document Report
    print("\n7. Generating Word Document Report...")
    
    # Create output directory if it doesn't exist
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    
    # Generate filename with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    word_file = f"{OUTPUT_FOLDER}/keyword_report_{timestamp}.docx"
    
    # Create the Word report
    try:
        create_quick_report(all_results, MY_KEYWORDS, word_file)
    except Exception as e:
        print(f"   ⚠️ Error generating Word document: {e}")
        print(f"   Still saving JSON results...")
    
    # Step 8: Generate JSON summary as backup
    print("\n8. Generating JSON summary...")
    generate_json_summary(all_results, MY_KEYWORDS, OUTPUT_FOLDER, 
                          len(saved_files), emails_with_keywords, total_keywords_found)
    
    # Step 9: Ask if user wants to mark emails as read
    if mode == "2":  # Only for unread emails
        mark_read = input("\nMark processed emails as read? (y/n): ").strip().lower()
        if mark_read == 'y':
            print("\n9. Marking emails as read...")
            for email in emails:
                connector.mark_as_read(email['id'])
            print("   ✅ Emails marked as read")
    
    print("\n" + "="*60)
    print("✅ Outlook extraction complete!")
    print(f"   Processed: {len(saved_files)} emails")
    print(f"   Emails with keywords: {emails_with_keywords}")
    print(f"   Total keywords found: {total_keywords_found}")
    print(f"\n📄 Reports saved to:")
    print(f"   📝 Word Document: {word_file}")
    print(f"   📊 JSON Summary: {OUTPUT_FOLDER}/summary_{timestamp}.json")
    print("="*60)


def generate_json_summary(results, keywords, output_folder, total_emails, 
                         emails_with_keywords, total_keywords_found):
    """Generate a JSON summary as backup."""
    
    import json
    
    # Calculate keyword frequency
    keyword_freq = {kw: 0 for kw in keywords}
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
        'keyword_frequency': {k: v for k, v in keyword_freq.items() if v > 0},
        'detailed_results': results
    }
    
    # Save summary
    summary_file = f"{output_folder}/summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    
    with open(summary_file, 'w') as f:
        json.dump(summary, f, indent=2, default=str)
    
    # Print summary
    print("\n" + "-" * 40)
    print("📊 SUMMARY STATISTICS")
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
