"""
Extract keywords directly from Outlook emails and generate focused summaries.
"""

import sys
import os
import json
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.pipeline.data_miner import DataMiner
from src.connectors.outlook_connector import OutlookConnector


def generate_focused_summary(results, search_keyword, output_folder, total_emails, 
                              emails_with_keyword, total_occurrences):
    """
    Generate a focused summary for the searched keyword.
    """
    print("\n" + "="*80)
    print(f"FOCUSED SUMMARY: '{search_keyword.upper()}'")
    print("="*80)
    
    # Calculate statistics
    success_rate = (emails_with_keyword / total_emails * 100) if total_emails > 0 else 0
    
    print(f"\nSEARCH STATISTICS:")
    print(f" Keyword searched: '{search_keyword}'")
    print(f" Total emails with keyword: {emails_with_keyword}")
    print(f" Total occurrences: {total_occurrences}")
    print(f" Success rate: {success_rate:.1f}%")
    
    # Collect all contexts for the search keyword
    all_contexts = []
    all_summaries = []
    
    for result in results:
        keywords = result.get('keywords', {})
        if search_keyword.lower() in keywords:
            keyword_data = keywords[search_keyword.lower()]
            all_contexts.extend(keyword_data.get('contexts', []))
            if keyword_data.get('summary'):
                all_summaries.append(keyword_data.get('summary'))
    
    # Show contexts
    if all_contexts:
        print(f"\nCONTEXT EXAMPLES:")
        print("-" * 60)
        for i, ctx in enumerate(all_contexts[:5], 1):
            context_text = ctx.get('full_context', '')[:150]
            print(f"\n   {i}. ...{context_text}...")
    
    # Show summaries
    if all_summaries:
        print(f"\n KEY SUMMARIES:")
        print("-" * 60)
        for i, summary in enumerate(all_summaries[:3], 1):
            print(f"\n   {i}. {summary[:200]}...")
    
    # Show emails where keyword was found
    print(f"\n EMAILS WITH '{search_keyword.upper()}':")
    print("-" * 60)
    
    email_count = 0
    for i, result in enumerate(results, 1):
        keywords = result.get('keywords', {})
        if search_keyword.lower() in keywords:
            email_count += 1
            email_meta = result.get('email_metadata', {})
            keyword_data = keywords[search_keyword.lower()]
            confidence = int(keyword_data.get('confidence', 0) * 100)
            occurrences = keyword_data.get('occurrences', 0)
            
            print(f"\n {email_count}. {email_meta.get('subject', 'No Subject')[:70]}")
            print(f" From: {email_meta.get('sender', 'Unknown')}")
            print(f" Date: {email_meta.get('date', 'Unknown')}")
            print(f" '{search_keyword}' found: {occurrences} time(s) | Confidence: {confidence}%")
    
    # Save focused summary to file
    summary_data = {
        'search_keyword': search_keyword,
        'total_emails_scanned': total_emails,
        'emails_with_keyword': emails_with_keyword,
        'total_occurrences': total_occurrences,
        'success_rate': success_rate,
        'contexts': all_contexts[:10],
        'summaries': all_summaries[:5],
        'detailed_results': results
    }
    
    # Save as JSON
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    summary_file = f"{output_folder}/focused_summary_{search_keyword}_{timestamp}.json"
    
    with open(summary_file, 'w') as f:
        json.dump(summary_data, f, indent=2, default=str)
    
    print(f"\nFocused summary saved to: {summary_file}")
    
    # Try to generate Word document for the focused summary
    try:
        from src.utils.word_generator import create_focused_report
        word_file = f"{output_folder}/focused_report_{search_keyword}_{timestamp}.docx"
        create_focused_report(summary_data, search_keyword, word_file)
        print(f"Word report saved to: {word_file}")
    except ImportError:
        print(f" Word generator not available. Install python-docx for Word reports.")
    except Exception as e:
        print(f" Could not create Word report: {e}")


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
        print("could not connect to Outlook.")
        print("\nTroubleshooting:")
        print("  1. Make sure Outlook is open")
        print("  2. Check that pywin32 is installed: pip install pywin32")
        print("  3. Try running as administrator")
        return
    
    # Step 2: Choose extraction mode
    print("\n2. Choose extraction mode:")
    print("   [1] Extract recent emails (last 90 days)")
    print("   [2] Extract unread emails only")
    print("   [3] Search by subject (FOCUSED SUMMARY)")
    print("   [4] Process all inbox emails")
    
    mode = input("\nEnter choice (1-4): ").strip()
    
    # Step 3: Extract emails based on mode
    print("\n3. Extracting emails...")
    
    if mode == "1":
        days = input("How many days back? (default: 90): ").strip()
        days = int(days) if days else 90
        emails = connector.get_emails(
            folder="Inbox",
            days_back=days,
            limit=100,
            unread_only=False
        )
        
        print(f"Found {len(emails)} emails")
        
        if not emails:
            print(" No emails found matching criteria.")
            return
        
        # Save emails to files
        print("\n4. Saving emails to files...")
        os.makedirs("data/outlook_emails", exist_ok=True)
        saved_files = connector.save_emails_to_files(emails, output_dir="data/outlook_emails")
        print(f" Saved {len(saved_files)} emails")
        
        # Initialize keyword miner
        print("\n5. Initializing keyword miner...")
        miner = DataMiner(relevance_threshold=0.15)
        
        # Process each email
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
                
                # Count keywords found
                keywords_found = len(results.get('keywords', {}))
                if keywords_found > 0:
                    print(f"Found {keywords_found} keywords")
                
                # Add email metadata
                results['email_metadata'] = {
                    'subject': emails[i-1]['subject'],
                    'sender': emails[i-1]['sender_name'],
                    'date': emails[i-1]['received_time']
                }
                
                all_results.append(results)
                    
            except Exception as e:
                print(f"Error: {e}")
                continue
        
        # Generate summary
        print("\n7. Generating summary report...")
        generate_json_summary(all_results, MY_KEYWORDS, OUTPUT_FOLDER, 
                              len(saved_files), len(all_results), 
                              sum(len(r.get('keywords', {})) for r in all_results))
    
    elif mode == "2":
        emails = connector.get_emails(
            folder="Inbox",
            days_back=90,
            limit=100,
            unread_only=True
        )
        
        print(f" Found {len(emails)} emails")
        
        if not emails:
            print(" No emails found matching criteria.")
            return
        
        # Save emails to files
        print("\n4. Saving emails to files...")
        os.makedirs("data/outlook_emails", exist_ok=True)
        saved_files = connector.save_emails_to_files(emails, output_dir="data/outlook_emails")
        print(f" Saved {len(saved_files)} emails")
        
        # Initialize keyword miner
        print("\n5. Initializing keyword miner...")
        miner = DataMiner(relevance_threshold=0.15)
        
        # Process each email
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
                
                # Count keywords found
                keywords_found = len(results.get('keywords', {}))
                if keywords_found > 0:
                    print(f" Found {keywords_found} keywords")
                
                # Add email metadata
                results['email_metadata'] = {
                    'subject': emails[i-1]['subject'],
                    'sender': emails[i-1]['sender_name'],
                    'date': emails[i-1]['received_time']
                }
                
                all_results.append(results)
                    
            except Exception as e:
                print(f" Error: {e}")
                continue
        
        # Generate summary
        print("\n7. Generating summary report...")
        generate_json_summary(all_results, MY_KEYWORDS, OUTPUT_FOLDER, 
                              len(saved_files), len(all_results), 
                              sum(len(r.get('keywords', {})) for r in all_results))
        
        # Mark as read
        mark_read = input("\nMark processed emails as read? (y/n): ").strip().lower()
        if mark_read == 'y':
            print("\n8. Marking emails as read...")
            for email in emails:
                connector.mark_as_read(email['id'])
            print(" Emails marked as read")
    
    elif mode == "3":
        # SEARCH BY SUBJECT - FOCUSED SUMMARY
        subject_filter = input("Enter subject keyword: ").strip()
        
        if not subject_filter:
            print(" No keyword entered. Returning to menu.")
            return
        
        print(f"\n Searching for emails with '{subject_filter}' in subject...")
        
        # Get emails with the subject filter
        emails = connector.get_emails(
            folder="Inbox",
            days_back=90,
            limit=100,
            unread_only=False,
            subject_filter=subject_filter
        )
        
        print(f" Found {len(emails)} emails with '{subject_filter}' in subject")
        
        if not emails:
            print(" No emails found matching criteria.")
            return
        
        # Save these emails to files
        print("\n4. Saving emails to files...")
        os.makedirs("data/outlook_emails", exist_ok=True)
        
        # Create a subfolder for this search
        search_folder = f"data/outlook_emails/search_{subject_filter}"
        os.makedirs(search_folder, exist_ok=True)
        
        saved_files = connector.save_emails_to_files(emails, output_dir=search_folder)
        print(f" Saved {len(saved_files)} emails")
        
        # Initialize keyword miner
        print("\n5. Initializing keyword miner...")
        miner = DataMiner(relevance_threshold=0.15)
        
        # Process each email, focusing on the search keyword
        print(f"\n6. Processing emails for keyword: '{subject_filter}'...")
        all_results = []
        emails_with_keyword = 0
        total_occurrences = 0
        
        # Create a results folder for this search
        search_results_folder = f"{OUTPUT_FOLDER}/search_{subject_filter}"
        os.makedirs(search_results_folder, exist_ok=True)
        
        for i, email_file in enumerate(saved_files, 1):
            print(f"\n   [{i}/{len(saved_files)}] Processing: {os.path.basename(email_file)}")
            
            try:
                # Mine the document with ALL keywords
                results = miner.mine_document(
                    document_path=email_file,
                    seed_keywords=MY_KEYWORDS,
                    output_dir=f"{search_results_folder}/email_{i}"
                )
                
                # Check if the search keyword was found in this email
                keyword_found = results.get('keywords', {}).get(subject_filter.lower(), None)
                
                if keyword_found:
                    emails_with_keyword += 1
                    total_occurrences += keyword_found.get('occurrences', 0)
                    print(f" Found '{subject_filter}' {keyword_found.get('occurrences', 0)} time(s)")
                else:
                    print(f" '{subject_filter}' not found in this email")
                
                # Add email metadata
                results['email_metadata'] = {
                    'subject': emails[i-1]['subject'],
                    'sender': emails[i-1]['sender_name'],
                    'date': emails[i-1]['received_time']
                }
                
                all_results.append(results)
                    
            except Exception as e:
                print(f"Error: {e}")
                continue
        
        # Generate a focused summary report for the search keyword
        print("\n7. Generating focused summary report...")
        generate_focused_summary(
            all_results, 
            subject_filter, 
            search_results_folder,
            len(saved_files),
            emails_with_keyword,
            total_occurrences
        )
        
        print("\n" + "="*60)
        print(f" Search for '{subject_filter}' complete!")
        print(f" Total emails with '{subject_filter}': {emails_with_keyword}")
        print(f" Total occurrences of '{subject_filter}': {total_occurrences}")
        print(f" Results saved to: {search_results_folder}")
        print("="*60)
    
    elif mode == "4":
        emails = connector.get_emails(
            folder="Inbox",
            days_back=90,
            limit=200,
            unread_only=False
        )
        
        print(f"Found {len(emails)} emails")
        
        if not emails:
            print(" No emails found matching criteria.")
            return
        
        # Save emails to files
        print("\n4. Saving emails to files...")
        os.makedirs("data/outlook_emails", exist_ok=True)
        saved_files = connector.save_emails_to_files(emails, output_dir="data/outlook_emails")
        print(f"Saved {len(saved_files)} emails")
        
        # Initialize keyword miner
        print("\n5. Initializing keyword miner...")
        miner = DataMiner(relevance_threshold=0.15)
        
        # Process each email
        print("\n6. Processing emails for keywords...")
        all_results = []
        
        for i, email_file in enumerate(saved_files, 1):
            print(f"\n [{i}/{len(saved_files)}] Processing: {os.path.basename(email_file)}")
            
            try:
                results = miner.mine_document(
                    document_path=email_file,
                    seed_keywords=MY_KEYWORDS,
                    output_dir=f"{OUTPUT_FOLDER}/email_{i}"
                )
                
                # Count keywords found
                keywords_found = len(results.get('keywords', {}))
                if keywords_found > 0:
                    print(f"Found {keywords_found} keywords")
                
                # Add email metadata
                results['email_metadata'] = {
                    'subject': emails[i-1]['subject'],
                    'sender': emails[i-1]['sender_name'],
                    'date': emails[i-1]['received_time']
                }
                
                all_results.append(results)
                    
            except Exception as e:
                print(f" Error: {e}")
                continue
        
        # Generate summary
        print("\n7. Generating summary report...")
        generate_json_summary(all_results, MY_KEYWORDS, OUTPUT_FOLDER, 
                              len(saved_files), len(all_results), 
                              sum(len(r.get('keywords', {})) for r in all_results))
    
    else:
        print("Invalid choice. Exiting.")
        return
    
    print("\n" + "="*60)
    print("Outlook extraction complete!")
    print("="*60)


def generate_json_summary(results, keywords, output_folder, total_emails, 
                         emails_with_keywords, total_keywords_found):
    """Generate a JSON summary as backup."""
    
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
    os.makedirs(output_folder, exist_ok=True)
    summary_file = f"{output_folder}/summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    
    with open(summary_file, 'w') as f:
        json.dump(summary, f, indent=2, default=str)
    
    # Print summary
    print("\n" + "-" * 40)
    print("SUMMARY REPORT")
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
    
    print(f"\nFull report saved to: {summary_file}")


if __name__ == "__main__":
    main()
