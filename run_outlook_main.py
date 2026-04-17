"""
Extract keywords from specific emails by subject, then search for keywords.
"""

import sys
import os
import json
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.pipeline.data_miner import DataMiner
from src.connectors.outlook_connector import OutlookConnector


def generate_focused_summary(results, search_keyword, output_folder, total_emails, 
                              emails_with_keyword, total_occurrences, subject_filter):
    """
    Generate a focused summary for the searched keyword within specific emails.
    """
    print("\n" + "="*80)
    print(f"FOCUSED SUMMARY: '{search_keyword.upper()}'")
    print(f"   (Within emails containing: '{subject_filter}')")
    print("="*80)
    
    # Calculate statistics
    success_rate = (emails_with_keyword / total_emails * 100) if total_emails > 0 else 0
    
    print(f"\nSEARCH STATISTICS:")
    print(f"  Email subject filter: '{subject_filter}'")
    print(f"  Keyword searched: '{search_keyword}'")
    print(f"  Total emails with subject filter: {total_emails}")
    print(f"  Emails containing keyword: {emails_with_keyword}")
    print(f"  Total occurrences: {total_occurrences}")
    print(f"  Success rate: {success_rate:.1f}%")
    
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
        print(f"\n CONTEXT EXAMPLES:")
        print("-" * 60)
        for i, ctx in enumerate(all_contexts[:5], 1):
            context_text = ctx.get('full_context', '')[:150]
            print(f"\n   {i}. ...{context_text}...")
    
    # Show summaries
    if all_summaries:
        print(f"\n📄 KEY SUMMARIES:")
        print("-" * 60)
        for i, summary in enumerate(all_summaries[:3], 1):
            print(f"\n   {i}. {summary[:200]}...")
    
    # Show emails where keyword was found
    print(f"\nEMAILS CONTAINING '{search_keyword.upper()}':")
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
            
            print(f"\n   {email_count}. {email_meta.get('subject', 'No Subject')[:70]}")
            print(f"  From: {email_meta.get('sender', 'Unknown')}")
            print(f"  Date: {email_meta.get('date', 'Unknown')}")
            print(f"  '{search_keyword}' found: {occurrences} time(s) | Confidence: {confidence}%")
    
    # Save focused summary to file
    summary_data = {
        'search_criteria': {
            'subject_filter': subject_filter,
            'keyword_searched': search_keyword
        },
        'total_emails_with_subject': total_emails,
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
    
    # Generate Word document
    try:
        from src.utils.word_generator import create_focused_report
        word_file = f"{output_folder}/focused_report_{search_keyword}_{timestamp}.docx"
        create_focused_report(summary_data, search_keyword, subject_filter, word_file)
        print(f"Word report saved to: {word_file}")
    except ImportError:
        print(f"Word generator not available. Install python-docx for Word reports.")
    except Exception as e:
        print(f"Could not create Word report: {e}")
    
    print("\n" + "="*80)
    return summary_file, word_file if 'word_file' in locals() else None


def main():
    # Configuration - UPDATE THESE WITH YOUR KEYWORDS
    MY_KEYWORDS = [
        "artificial intelligence",
        "AI",
        "MRO",
        "Maintanence",
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
    print("OUTLOOK EMAIL KEYWORD MINER")
    print("="*60)
    
    # Step 1: Connect to Outlook
    print("\n1. Connecting to Outlook...")
    connector = OutlookConnector()
    
    if not connector.outlook:
        print("Could not connect to Outlook.")
        print("\nTroubleshooting:")
        print(" 1. Make sure Outlook is open")
        print(" 2. Check that pywin32 is installed: pip install pywin32")
        return
    
    # Step 2: Enter email subject to search for
    print("\n" + "="*60)
    print("STEP 1: Enter email subject to search for")
    print("="*60)
    subject_filter = input("\nEnter email subject (or partial subject): ").strip()
    
    if not subject_filter:
        print(" No subject entered. Exiting.")
        return
    
    print(f"\n Searching for emails with subject containing: '{subject_filter}'...")
    
    # Get emails with the subject filter
    emails = connector.get_emails(
        folder="Inbox",
        days_back=90,
        limit=100,
        unread_only=False,
        subject_filter=subject_filter
    )
    
    print(f" Found {len(emails)} email(s) with subject containing '{subject_filter}'")
    
    if not emails:
        print(" No emails found matching the subject criteria.")
        return
    
    # Display found emails
    print("\n Emails found:")
    for i, email in enumerate(emails, 1):
        print(f" {i}. {email['subject'][:80]}")
        print(f"  From: {email['sender_name']} | Date: {email['received_time']}")
    
    # Step 3: Enter keyword to search within these emails
    print("\n" + "="*60)
    print("STEP 2: Enter keyword to search within these emails")
    print("="*60)
    search_keyword = input("\nEnter keyword to search for: ").strip()
    
    if not search_keyword:
        print(" No keyword entered. Exiting.")
        return
    
    print(f"\n Searching for '{search_keyword}' within the {len(emails)} email(s)...")
    
    # Step 4: Save emails to files
    print("\n Saving emails to files")
    os.makedirs("data/outlook_emails", exist_ok=True)
    
    # Create a subfolder for this search
    safe_subject = "".join(c for c in subject_filter if c.isalnum() or c in ' _-')[:30]
    search_folder = f"data/outlook_emails/search_{safe_subject}"
    os.makedirs(search_folder, exist_ok=True)
    
    saved_files = connector.save_emails_to_files(emails, output_dir=search_folder)
    print(f" Saved {len(saved_files)} email(s)")
    
    # Step 5: Initialize keyword miner
    print("\n Initializing keyword miner...")
    miner = DataMiner(relevance_threshold=0.15)
    
    # Step 6: Process each email, focusing on the search keyword
    print(f"\nProcessing emails for keyword: '{search_keyword}'...")
    all_results = []
    emails_with_keyword = 0
    total_occurrences = 0
    
    # Create a results folder for this search
    results_folder = f"{OUTPUT_FOLDER}/search_{safe_subject}"
    os.makedirs(results_folder, exist_ok=True)
    
    for i, email_file in enumerate(saved_files, 1):
        print(f"\n   [{i}/{len(saved_files)}] Processing: {os.path.basename(email_file)}")
        
        try:
            # Mine the document with ALL keywords
            results = miner.mine_document(
                document_path=email_file,
                seed_keywords=MY_KEYWORDS,
                output_dir=f"{results_folder}/email_{i}"
            )
            
            # Check if the search keyword was found in this email
            keyword_found = results.get('keywords', {}).get(search_keyword.lower(), None)
            
            if keyword_found:
                emails_with_keyword += 1
                total_occurrences += keyword_found.get('occurrences', 0)
                print(f" Found '{search_keyword}' {keyword_found.get('occurrences', 0)} time(s)")
                
                # Show a preview
                summary = keyword_found.get('summary', '')
                if summary:
                    print(f" Preview: {summary[:100]}...")
            else:
                print(f" '{search_keyword}' not found in this email")
            
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
    
    # Step 7: Generate focused summary report
    print("\n" + "="*60)
    print("Generating Summary Report")
    print("="*60)
    
    summary_file, word_file = generate_focused_summary(
        all_results, 
        search_keyword, 
        results_folder,
        len(saved_files),
        emails_with_keyword,
        total_occurrences,
        subject_filter
    )
    
    # Step 8: Final summary
    print("\n" + "="*60)
    print("EXTRACTION COMPLETE!")
    print("="*60)
    print(f"\nSUMMARY:")
    print(f" Email subject searched: '{subject_filter}'")
    print(f" Emails found: {len(saved_files)}")
    print(f" Keyword searched: '{search_keyword}'")
    print(f" Emails containing keyword: {emails_with_keyword}")
    print(f" Total occurrences: {total_occurrences}")
    print(f"\nResults saved to:")
    print(f" JSON: {summary_file}")
    if word_file:
        print(f"Word Document: {word_file}")
    print(f" Folder: {results_folder}")
    print("="*60)


if __name__ == "__main__":
    main()


------------------------------------------------------------------------
updated code

"""
Extract keywords from specific emails by subject, then search for keywords.
"""

import sys
import os
import json
import re
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.pipeline.data_miner import DataMiner
from src.connectors.outlook_connector import OutlookConnector


def generate_focused_summary(results, search_keyword, output_folder, total_emails, 
                              emails_with_keyword, total_occurrences, subject_filter):
    """
    Generate a focused summary for the searched keyword within specific emails.
    """
    print("\n" + "="*80)
    print(f"FOCUSED SUMMARY: '{search_keyword.upper()}'")
    print(f"   (Within emails containing: '{subject_filter}')")
    print("="*80)
    
    # Collect ONLY the searched keyword data
    all_contexts = []
    all_summaries = []
    email_count = 0
    actual_total_occurrences = 0
    
    for result in results:
        keywords = result.get('keywords', {})
        # Look for the exact searched keyword (case-insensitive)
        keyword_data = None
        for kw in keywords.keys():
            if kw.lower() == search_keyword.lower():
                keyword_data = keywords[kw]
                break
        
        if keyword_data:
            email_count += 1
            occurrences = keyword_data.get('occurrences', 0)
            actual_total_occurrences += occurrences
            # Get contexts and clean them
            contexts = keyword_data.get('contexts', [])
            for ctx in contexts[:3]:
                clean_ctx = ctx.copy()
                if 'full_context' in clean_ctx:
                    # Clean the context text
                    clean_ctx['full_context'] = re.sub(r'https?://\S+', '', clean_ctx['full_context'])
                    clean_ctx['full_context'] = re.sub(r'\s+', ' ', clean_ctx['full_context'])
                    clean_ctx['full_context'] = clean_ctx['full_context'][:300]
                all_contexts.append(clean_ctx)
            
            if keyword_data.get('summary'):
                # Clean the summary
                clean_summary = keyword_data['summary']
                clean_summary = re.sub(r'https?://\S+', '', clean_summary)
                clean_summary = re.sub(r'\s+', ' ', clean_summary)
                clean_summary = clean_summary[:300]
                all_summaries.append(clean_summary)
    
    # Calculate statistics
    success_rate = (email_count / total_emails * 100) if total_emails > 0 else 0
    
    print(f"\n SEARCH STATISTICS:")
    print(f" Email subject filter: '{subject_filter}'")
    print(f" Keyword searched: '{search_keyword}'")
    print(f" Total emails with subject filter: {total_emails}")
    print(f" Emails containing '{search_keyword}': {email_count}")
    print(f" Total occurrences: {actual_total_occurrences}")
    print(f" Success rate: {success_rate:.1f}%")
    
    # Show contexts (only for the searched keyword)
    if all_contexts:
        print(f"\n CONTEXT EXAMPLES (for '{search_keyword}'):")
        print("-" * 60)
        for i, ctx in enumerate(all_contexts[:5], 1):
            context_text = ctx.get('full_context', '')[:200]
            print(f"\n   {i}. ...{context_text}...")
    else:
        print(f"\n No context examples found for '{search_keyword}'")
    
    # Show summaries
    if all_summaries:
        print(f"\n KEY SUMMARIES (for '{search_keyword}'):")
        print("-" * 60)
        for i, summary in enumerate(all_summaries[:3], 1):
            print(f"\n {i}. {summary[:200]}...")
    
    # Show emails where keyword was found
    print(f"\n EMAILS CONTAINING '{search_keyword.upper()}':")
    print("-" * 60)
    
    email_index = 0
    for i, result in enumerate(results, 1):
        keywords = result.get('keywords', {})
        keyword_data = None
        for kw in keywords.keys():
            if kw.lower() == search_keyword.lower():
                keyword_data = keywords[kw]
                break
        
        if keyword_data:
            email_index += 1
            email_meta = result.get('email_metadata', {})
            confidence = int(keyword_data.get('confidence', 0) * 100)
            occurrences = keyword_data.get('occurrences', 0)
            
            print(f"\n {email_index}. {email_meta.get('subject', 'No Subject')[:70]}")
            print(f" From: {email_meta.get('sender', 'Unknown')}")
            print(f" Date: {email_meta.get('date', 'Unknown')}")
            print(f" '{search_keyword}' found: {occurrences} time(s) | Confidence: {confidence}%")
            
            # Show summary for this email
            if keyword_data.get('summary'):
                clean_summary = re.sub(r'https?://\S+', '', keyword_data['summary'])
                clean_summary = re.sub(r'\s+', ' ', clean_summary)
                print(f"      Summary: {clean_summary[:150]}...")
    
    if email_count == 0:
        print(f"\n No emails found containing '{search_keyword}'")
        print(f"\n  Available keywords in the processed emails:")
        all_keywords = set()
        for result in results:
            all_keywords.update(result.get('keywords', {}).keys())
        if all_keywords:
            print(f" Found: {', '.join(list(all_keywords)[:10])}")
    
    # Save focused summary (ONLY for the searched keyword)
    summary_data = {
        'search_criteria': {
            'subject_filter': subject_filter,
            'keyword_searched': search_keyword
        },
        'total_emails_with_subject': total_emails,
        'emails_with_keyword': email_count,
        'total_occurrences': actual_total_occurrences,
        'success_rate': success_rate,
        'contexts': all_contexts[:10],
        'summaries': all_summaries[:5],
        'detailed_results': results
    }
    
    # Save as JSON
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_subject = re.sub(r'[^\w\s-]', '', subject_filter)[:30]
    summary_file = f"{output_folder}/focused_summary_{search_keyword}_{timestamp}.json"
    
    with open(summary_file, 'w', encoding='utf-8') as f:
        json.dump(summary_data, f, indent=2, default=str)
    
    print(f"\nFocused summary saved to: {summary_file}")
    
    # Generate Word document
    word_file = None
    try:
        from src.utils.word_generator import create_focused_report
        word_file = f"{output_folder}/focused_report_{search_keyword}_{timestamp}.docx"
        create_focused_report(summary_data, search_keyword, subject_filter, word_file)
        print(f"Word report saved to: {word_file}")
    except ImportError:
        print(f"Word generator not available. Install python-docx: pip install python-docx")
    except Exception as e:
        print(f"Could not create Word report: {e}")
    
    print("\n" + "="*80)
    return summary_file, word_file


def main():
    # Configuration - Keywords to search within emails
    MY_KEYWORDS = [
        "AI",
        "artificial intelligence",
        "machine learning",
        "Genpact"
    ]
    
    OUTPUT_FOLDER = "outputs/outlook_results"
    
    print("\n" + "="*60)
    print(" OUTLOOK EMAIL KEYWORD MINER")
    print("="*60)
    
    # Step 1: Connect to Outlook
    print("\n1. Connecting to Outlook...")
    connector = OutlookConnector()
    
    if not connector.outlook:
        print("Could not connect to Outlook.")
        print("\nTroubleshooting:")
        print("  1. Make sure Outlook is open")
        print("  2. Check that pywin32 is installed: pip install pywin32")
        return
    
    # Step 2: Enter email subject to search for
    print("\n" + "="*60)
    print("STEP 1: Enter email subject to search for")
    print("="*60)
    subject_filter = input("\nEnter email subject (or partial subject): ").strip()
    
    if not subject_filter:
        print("No subject entered. Exiting.")
        return
    
    print(f"\n Searching for emails with subject containing: '{subject_filter}'...")
    
    # Get emails with the subject filter
    emails = connector.get_emails(
        folder="Inbox",
        days_back=90,
        limit=100,
        unread_only=False,
        subject_filter=subject_filter
    )
    
    print(f"Found {len(emails)} email(s) with subject containing '{subject_filter}'")
    
    if not emails:
        print("No emails found matching the subject criteria.")
        return
    
    # Display found emails
    print("\n Emails found:")
    for i, email in enumerate(emails, 1):
        print(f" {i}. {email['subject'][:80]}")
        print(f" From: {email['sender_name']} | Date: {email['received_time']}")
    
    # Step 3: Enter keyword to search within these emails
    print("\n" + "="*60)
    print("STEP 2: Enter keyword to search within these emails")
    print("="*60)
    search_keyword = input("\nEnter keyword to search for: ").strip()
    
    if not search_keyword:
        print(" No keyword entered. Exiting.")
        return
    
    print(f"\n Searching for '{search_keyword}' within the {len(emails)} email(s)...")
    
    # Step 4: Save emails to files
    print("\nSaving emails to files...")
    os.makedirs("data/outlook_emails", exist_ok=True)
    
    # Create a subfolder for this search
    safe_subject = re.sub(r'[^\w\s-]', '', subject_filter)[:30]
    safe_subject = safe_subject.replace(' ', '_')
    search_folder = f"data/outlook_emails/search_{safe_subject}"
    os.makedirs(search_folder, exist_ok=True)
    
    saved_files = connector.save_emails_to_files(emails, output_dir=search_folder)
    print(f"Saved {len(saved_files)} email(s)")
    
    # Step 5: Initialize keyword miner
    print("\n Initializing keyword miner...")
    miner = DataMiner(relevance_threshold=0.15)
    
    # Step 6: Process each email, focusing on the search keyword
    print(f"\nProcessing emails for keyword: '{search_keyword}'...")
    all_results = []
    emails_with_keyword = 0
    total_occurrences = 0
    
    # Create a results folder for this search
    results_folder = f"{OUTPUT_FOLDER}/search_{safe_subject}"
    os.makedirs(results_folder, exist_ok=True)
    
    for i, email_file in enumerate(saved_files, 1):
        print(f"\n [{i}/{len(saved_files)}] Processing: {os.path.basename(email_file)}")
        
        try:
            # Mine the document with ALL keywords
            results = miner.mine_document(
                document_path=email_file,
                seed_keywords=MY_KEYWORDS,
                output_dir=f"{results_folder}/email_{i}"
            )
            
            # Check if the search keyword was found in this email (case-insensitive)
            keyword_found = None
            for kw in results.get('keywords', {}).keys():
                if kw.lower() == search_keyword.lower():
                    keyword_found = results['keywords'][kw]
                    break
            
            if keyword_found:
                emails_with_keyword += 1
                occurrences = keyword_found.get('occurrences', 0)
                total_occurrences += occurrences
                print(f"Found '{search_keyword}' {occurrences} time(s)")
                
                # Show a preview of the summary
                summary = keyword_found.get('summary', '')
                if summary:
                    # Clean the summary
                    clean_summary = re.sub(r'https?://\S+', '', summary)
                    clean_summary = re.sub(r'\s+', ' ', clean_summary)
                    print(f"Preview: {clean_summary[:100]}...")
            else:
                print(f"'{search_keyword}' not found in this email")
                # Debug: Show what keywords were found
                found_keywords = list(results.get('keywords', {}).keys())
                if found_keywords:
                    print(f"Found keywords: {', '.join(found_keywords[:5])}")
            
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
    
    # Step 7: Generate focused summary report
    print("\n" + "="*60)
    print("Generating Summary Report")
    print("="*60)
    
    summary_file, word_file = generate_focused_summary(
        all_results, 
        search_keyword, 
        results_folder,
        len(saved_files),
        emails_with_keyword,
        total_occurrences,
        subject_filter
    )
    
    # Step 8: Final summary
    print("\n" + "="*60)
    print(" EXTRACTION COMPLETE!")
    print("="*60)
    print(f"\n FINAL SUMMARY:")
    print(f" Email subject searched: '{subject_filter}'")
    print(f" Emails found: {len(saved_files)}")
    print(f" Keyword searched: '{search_keyword}'")
    print(f" Emails containing keyword: {emails_with_keyword}")
    print(f" Total occurrences: {total_occurrences}")
    print(f" Success rate: {(emails_with_keyword/len(saved_files)*100) if saved_files else 0:.1f}%")
    
    print(f"\n Results saved to:")
    print(f" JSON: {summary_file}")
    if word_file:
        print(f"  Word Document: {word_file}")
    print(f" Folder: {results_folder}")
    print("="*60)


if __name__ == "__main__":
    main()
