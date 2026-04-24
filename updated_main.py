"""
Extract keywords from specific emails by subject.
Option 1: Enter keywords at runtime
Option 2: Use predefined keywords
"""

import sys
import os
import json
import re
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.pipeline.data_miner import DataMiner
from src.connectors.outlook_connector import OutlookConnector


# ============================================================
# CONFIGURATION - EDIT THESE PREDEFINED KEYWORDS AS NEEDED
# ============================================================
PREDEFINED_KEYWORDS = [
    "mro",
    "aviation", 
    "overhaul",
    "maintenance",
    "service",
    "inventory",
    "supply chain",
    "logistics",
    "procurement",
    "vendor"
]
# ============================================================


def generate_keyword_summary(results, search_keyword, output_folder, total_emails, 
                              emails_with_keyword, total_occurrences, subject_filter):
    """
    Generate a focused summary for a SINGLE searched keyword.
    """
    print("\n" + "="*80)
    print(f" SUMMARY FOR KEYWORD: '{search_keyword.upper()}'")
    print(f" (Within emails containing: '{subject_filter}')")
    print("="*80)
    
    # Collect ONLY the searched keyword data
    all_contexts = []
    all_summaries = []
    email_count = 0
    actual_total_occurrences = 0
    
    for result in results:
        keywords = result.get('keywords', {})
        
        # ONLY look for the searched keyword
        keyword_data = None
        for kw in keywords.keys():
            if kw.lower() == search_keyword.lower():
                keyword_data = keywords[kw]
                break
        
        if keyword_data:
            email_count += 1
            occurrences = keyword_data.get('occurrences', 0)
            actual_total_occurrences += occurrences
            all_contexts.extend(keyword_data.get('contexts', [])[:3])
            if keyword_data.get('summary'):
                clean_summary = keyword_data['summary']
                clean_summary = re.sub(r'https?://\S+', '', clean_summary)
                clean_summary = re.sub(r'\s+', ' ', clean_summary)
                all_summaries.append(clean_summary[:300])
    
    # Calculate statistics
    success_rate = (email_count / total_emails * 100) if total_emails > 0 else 0
    
    print(f"\n SEARCH STATISTICS:")
    print(f" Keyword searched: '{search_keyword}'")
    print(f" Emails containing '{search_keyword}': {email_count}")
    print(f" Total occurrences: {actual_total_occurrences}")
    print(f" Success rate: {success_rate:.1f}%")
    
    # Show contexts
    if all_contexts:
        print(f"\n CONTEXT EXAMPLES (for '{search_keyword}'):")
        for i, ctx in enumerate(all_contexts[:3], 1):
            context_text = ctx.get('full_context', '')[:150]
            context_text = re.sub(r'https?://\S+', '', context_text)
            print(f"\n   {i}. ...{context_text}...")
    else:
        print(f"\n No context examples found for '{search_keyword}'")
    
    # Show emails
    if email_count > 0:
        print(f"\n EMAILS CONTAINING '{search_keyword.upper()}':")
        email_index = 0
        for result in results:
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
                
                print(f"\n {email_index}. {email_meta.get('subject', 'No Subject')[:60]}")
                print(f" From: {email_meta.get('sender', 'Unknown')}")
                print(f" Date: {email_meta.get('date', 'Unknown')}")
                print(f"'{search_keyword}' found: {occurrences} time(s) | Confidence: {confidence}%")
    else:
        print(f"\n No emails found containing '{search_keyword}'")
    
    # Save summary to file
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_subject = re.sub(r'[^\w\s-]', '', subject_filter)[:20].replace(' ', '_')
    
    # Create output directory
    os.makedirs(output_folder, exist_ok=True)
    
    summary_data = {
        'keyword': search_keyword,
        'subject_filter': subject_filter,
        'total_emails_scanned': total_emails,
        'emails_with_keyword': email_count,
        'total_occurrences': actual_total_occurrences,
        'success_rate': success_rate,
        'contexts': all_contexts[:10],
        'summaries': all_summaries[:5],
        'detailed_results': results
    }
    
    # Save JSON
    json_file = f"{output_folder}/{safe_subject}_{search_keyword}_{timestamp}.json"
    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(summary_data, f, indent=2, default=str)
    
    # Save TXT
    txt_file = f"{output_folder}/{safe_subject}_{search_keyword}_{timestamp}.txt"
    with open(txt_file, 'w', encoding='utf-8') as f:
        f.write("="*60 + "\n")
        f.write(f"KEYWORD SUMMARY REPORT\n")
        f.write("="*60 + "\n\n")
        f.write(f"Keyword searched: {search_keyword.upper()}\n")
        f.write(f"Email subject filter: {subject_filter}\n")
        f.write(f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        
        f.write("-"*40 + "\n")
        f.write("STATISTICS\n")
        f.write("-"*40 + "\n")
        f.write(f"Total emails scanned: {total_emails}\n")
        f.write(f"Emails with keyword: {email_count}\n")
        f.write(f"Total occurrences: {actual_total_occurrences}\n")
        f.write(f"Success rate: {success_rate:.1f}%\n\n")
        
        if all_contexts:
            f.write("-"*40 + "\n")
            f.write("CONTEXT EXAMPLES\n")
            f.write("-"*40 + "\n")
            for i, ctx in enumerate(all_contexts[:3], 1):
                ctx_text = ctx.get('full_context', '')[:200]
                ctx_text = re.sub(r'https?://\S+', '', ctx_text)
                f.write(f"\n{i}. ...{ctx_text}...\n")
        
        if all_summaries:
            f.write("\n" + "-"*40 + "\n")
            f.write("KEY SUMMARIES\n")
            f.write("-"*40 + "\n")
            for i, summary in enumerate(all_summaries[:2], 1):
                f.write(f"\n{i}. {summary[:250]}...\n")
    
    print(f"\n JSON saved: {json_file}")
    print(f"Text saved: {txt_file}")
    
    return json_file, txt_file


def main():
    OUTPUT_FOLDER = "outputs/keyword_results"
    
    print("\n" + "="*60)
    print("OUTLOOK EMAIL KEYWORD MINER")
    print("="*60)
    print("\n This tool extracts keywords from your Outlook emails.")
    print(" You can either:")
    print(" 1. Type keywords manually at runtime")
    print(" 2. Use predefined keywords from the configuration")
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
    
    print("Connected to Outlook successfully")
    
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
    
    print(f"\n Found {len(emails)} email(s) with subject containing '{subject_filter}'")
    
    if not emails:
        print("\n   No emails found. Try a different subject or check your inbox.")
        return
    
    # Display found emails
    print("\n Emails found:")
    for i, email in enumerate(emails, 1):
        # Truncate long subjects
        display_subject = email['subject'][:70] + "..." if len(email['subject']) > 70 else email['subject']
        print(f" {i}. {display_subject}")
        print(f" From: {email['sender_name']} | Date: {email['received_time']}")
    
    # Step 3: Choose keyword input method
    print("\n" + "="*60)
    print("STEP 2: Choose keyword input method")
    print("="*60)
    print(" [1] Type keywords manually at runtime")
    print(" [2] Use predefined keywords from configuration")
    print(" [3] Both (predefined + manual)")
    print("-" * 60)
    
    choice = input("\nEnter your choice (1/2/3): ").strip()
    
    keywords_to_search = []
    
    if choice == '1':
        # Manual input only
        print("\n" + "-"*40)
        print("MANUAL KEYWORD ENTRY")
        print("-"*40)
        print(" For single keyword: just type it (e.g., service)")
        print(" For multiple keywords: separate with commas (e.g., service, overhaul, maintenance)")
        
        keywords_input = input("\nEnter keyword(s): ").strip()
        if not keywords_input:
            print("No keyword entered. Exiting.")
            return
        keywords_to_search = [k.strip().lower() for k in keywords_input.split(',')]
        
    elif choice == '2':
        # Predefined keywords only
        print("\n" + "-"*40)
        print("PREDEFINED KEYWORDS")
        print("-"*40)
        print(f"Available keywords: {', '.join(PREDEFINED_KEYWORDS)}")
        
        # Show numbered list for easy selection
        print("\n Select keywords to search (enter numbers separated by commas):")
        for i, kw in enumerate(PREDEFINED_KEYWORDS, 1):
            print(f"{i}. {kw}")
        
        selection = input("\nEnter your choice (e.g., 1,3,5 for first, third and fifth): ").strip()
        
        if selection:
            indices = [int(x.strip()) for x in selection.split(',') if x.strip().isdigit()]
            keywords_to_search = [PREDEFINED_KEYWORDS[i-1] for i in indices if 1 <= i <= len(PREDEFINED_KEYWORDS)]
        
        if not keywords_to_search:
            print(" No valid selection. Using all predefined keywords.")
            keywords_to_search = PREDEFINED_KEYWORDS.copy()
    
    elif choice == '3':
        # Both predefined and manual
        print("\n" + "-"*40)
        print("PREDEFINED KEYWORDS")
        print("-"*40)
        print(f"Available predefined keywords: {', '.join(PREDEFINED_KEYWORDS)}")
        
        use_predefined = input("\nUse predefined keywords? (y/n): ").strip().lower()
        if use_predefined == 'y':
            keywords_to_search.extend(PREDEFINED_KEYWORDS)
        
        print("\n" + "-"*40)
        print("MANUAL KEYWORD ENTRY")
        print("-"*40)
        manual_input = input("Enter additional keywords (comma-separated) or press Enter to skip: ").strip()
        if manual_input:
            manual_keywords = [k.strip().lower() for k in manual_input.split(',')]
            keywords_to_search.extend(manual_keywords)
        
        # Remove duplicates
        keywords_to_search = list(set(keywords_to_search))
    
    else:
        print("Invalid choice. Exiting.")
        return
    
    if not keywords_to_search:
        print("No keywords selected. Exiting.")
        return
    
    print(f"\n Will search for {len(keywords_to_search)} keyword(s):")
    for kw in keywords_to_search:
        print(f"{kw}")
    
    # Step 4: Save emails to files
    print("\n Saving emails to files...")
    safe_subject = re.sub(r'[^\w\s-]', '', subject_filter)[:30].replace(' ', '_')
    search_folder = f"data/outlook_emails/{safe_subject}"
    os.makedirs(search_folder, exist_ok=True)
    
    saved_files = connector.save_emails_to_files(emails, output_dir=search_folder)
    print(f" Saved {len(saved_files)} email(s)")
    
    # Step 5: Initialize keyword miner
    print("\n Initializing keyword miner...")
    miner = DataMiner(relevance_threshold=0.15)
    
    # Create results folder
    results_folder = f"{OUTPUT_FOLDER}/{safe_subject}"
    os.makedirs(results_folder, exist_ok=True)
    
    # Step 6: Process each keyword SEPARATELY
    print("\n" + "="*60)
    print("PROCESSING KEYWORDS")
    print("="*60)
    
    all_summary_files = []
    
    for keyword_idx, keyword in enumerate(keywords_to_search, 1):
        print(f"\n{'='*80}")
        print(f"Processing keyword {keyword_idx}/{len(keywords_to_search)}: '{keyword.upper()}'")
        print(f"{'='*80}")
        
        all_results = []
        emails_with_keyword = 0
        total_occurrences = 0
        
        for i, email_file in enumerate(saved_files, 1):
            print(f"\n   [{i}/{len(saved_files)}] Processing: {os.path.basename(email_file)[:50]}...")
            
            try:
                # Mine ONLY for this specific keyword
                results = miner.mine_document(
                    document_path=email_file,
                    seed_keywords=[keyword],  # ONLY this keyword!
                    output_dir=f"{results_folder}/email_{i}_{keyword}"
                )
                
                # Check if keyword found
                keyword_found = results.get('keywords', {}).get(keyword, None)
                
                if keyword_found:
                    emails_with_keyword += 1
                    occurrences = keyword_found.get('occurrences', 0)
                    total_occurrences += occurrences
                    print(f" Found '{keyword}' {occurrences} time(s)")
                else:
                    print(f"'{keyword}' not found in this email")
                
                # Add metadata
                results['email_metadata'] = {
                    'subject': emails[i-1]['subject'],
                    'sender': emails[i-1]['sender_name'],
                    'date': emails[i-1]['received_time']
                }
                
                all_results.append(results)
                
            except Exception as e:
                print(f" Error: {e}")
                continue
        
        # Generate summary for this keyword
        json_file, txt_file = generate_keyword_summary(
            all_results, 
            keyword, 
            results_folder,
            len(saved_files),
            emails_with_keyword,
            total_occurrences,
            subject_filter
        )
        
        all_summary_files.append((keyword, json_file, txt_file))
    
    # Step 7: Final summary
    print("\n" + "="*60)
    print("ALL KEYWORD EXTRACTIONS COMPLETE!")
    print("="*60)
    
    print("\n FINAL SUMMARY:")
    print(f" Email subject searched: '{subject_filter}'")
    print(f" Total emails found: {len(saved_files)}")
    print(f" Total keywords processed: {len(keywords_to_search)}")
    print(f" Results saved to: {results_folder}")
    
    print("\n Generated files:")
    for keyword, json_file, txt_file in all_summary_files:
        print(f"\n '{keyword.upper()}':")
        print(f" JSON: {os.path.basename(json_file)}")
        print(f" TXT: {os.path.basename(txt_file)}")
    
    print("\n" + "="*60)
    print("EXTRACTION COMPLETE!")
    print(f" All results are in: {results_folder}")
    print("="*60)


if __name__ == "__main__":
    main()
