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


# CONFIGURATION - EDIT THESE PREDEFINED KEYWORDS AS NEEDED
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

def generate_keyword_summary(results, search_keyword, output_folder, total_emails, 
                              emails_with_keyword, total_occurrences, subject_filter):
    """
    Generate a focused summary for a SINGLE searched keyword.
    """
    print("\n" + "="*80)
    print(f"SUMMARY FOR KEYWORD: '{search_keyword.upper()}'")
    print(f" (Within emails containing: '{subject_filter}')")
    print("="*80)
    
    # Collect ONLY the searched keyword data
    all_contexts = []
    all_summaries = []
    all_raw_texts = []  # Store raw email text for better summarization
    email_count = 0
    actual_total_occurrences = 0
    all_entities = {}  # Collect entities from all emails
    
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
            
            # Get contexts
            contexts = keyword_data.get('contexts', [])
            for ctx in contexts[:5]:
                all_contexts.append(ctx)
                # Store raw full context for better summarization
                if ctx.get('full_context'):
                    all_raw_texts.append(ctx.get('full_context'))
            
            # Get summary
            if keyword_data.get('summary'):
                all_summaries.append(keyword_data['summary'])
        
        # Collect entities from this email
        email_entities = result.get('entities', {})
        for entity_type, entity_list in email_entities.items():
            if entity_type not in all_entities:
                all_entities[entity_type] = set()
            for entity in entity_list:
                all_entities[entity_type].add(entity)
    
    # Calculate statistics
    success_rate = (email_count / total_emails * 100) if total_emails > 0 else 0
    
    print(f"\nSEARCH STATISTICS:")
    print(f" Keyword searched: '{search_keyword}'")
    print(f" Emails containing '{search_keyword}': {email_count}")
    print(f" Total occurrences: {actual_total_occurrences}")
    print(f" Success rate: {success_rate:.1f}%")
    
    # Show contexts (raw snippets)
    if all_contexts:
        print(f"\n CONTEXT EXAMPLES (Raw snippets showing the keyword):")
        print("-" * 60)
        for i, ctx in enumerate(all_contexts[:5], 1):
            # Get the full context and clean it
            context_text = ctx.get('full_context', '')
            # Remove extra whitespace and URLs
            context_text = re.sub(r'https?://\S+', '', context_text)
            context_text = re.sub(r'\s+', ' ', context_text)
            # Highlight the keyword
            keyword_lower = search_keyword.lower()
            text_lower = context_text.lower()
            pos = text_lower.find(keyword_lower)
            if pos != -1:
                start = max(0, pos - 40)
                end = min(len(context_text), pos + len(search_keyword) + 60)
                highlighted = context_text[start:end]
                # Replace the keyword with **bold** markers (for console)
                highlighted = highlighted.replace(context_text[pos:pos+len(search_keyword)], 
                                                  f"**{context_text[pos:pos+len(search_keyword)]}**")
                print(f"\n{i}. ...{highlighted}...")
            else:
                print(f"\n {i}. ...{context_text[:200]}...")
    else:
        print(f"\n No context examples found for '{search_keyword}'")
    
    # Generate BETTER SUMMARY from raw text
    print(f"\nDETAILED SUMMARY (from email content):")
    print("-" * 60)
    
    if all_raw_texts:
        # Combine all raw texts
        combined_text = ' '.join(all_raw_texts)
        # Clean the text
        combined_text = re.sub(r'https?://\S+', '', combined_text)
        combined_text = re.sub(r'\s+', ' ', combined_text)
        
        # Extract key sentences (simple approach - look for sentences with the keyword)
        sentences = re.split(r'[.!?]+', combined_text)
        relevant_sentences = []
        for sent in sentences:
            if search_keyword.lower() in sent.lower():
                # Clean and add
                clean_sent = sent.strip()
                if clean_sent and len(clean_sent) > 10:
                    relevant_sentences.append(clean_sent)
        
        if relevant_sentences:
            print(f"\n Here are the complete sentences containing '{search_keyword}':\n")
            for i, sent in enumerate(relevant_sentences[:10], 1):
                # Clean the sentence
                sent = re.sub(r'\s+', ' ', sent)
                print(f" {i}. {sent}")
        else:
            # Fallback to the generated summary
            for summary in all_summaries[:3]:
                clean_summary = re.sub(r'https?://\S+', '', summary)
                clean_summary = re.sub(r'\s+', ' ', clean_summary)
                print(f"\n   {clean_summary[:500]}")
    else:
        print(f"\n No detailed content available for '{search_keyword}'")
    
    # Show emails
    if email_count > 0:
        print(f"\nEMAILS CONTAINING '{search_keyword.upper()}':")
        print("-" * 60)
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
                
                print(f"\n{email_index}. {email_meta.get('subject', 'No Subject')[:70]}")
                print(f" From: {email_meta.get('sender', 'Unknown')}")
                print(f" Date: {email_meta.get('date', 'Unknown')}")
                print(f"'{search_keyword}' found: {occurrences} time(s) | Confidence: {confidence}%")
    
    # Show entities found
    if all_entities:
        print(f"\nENTITIES FOUND IN THE EMAIL:")
        print("-" * 60)
        # Sort entity types for better display
        for entity_type in sorted(all_entities.keys()):
            entity_set = all_entities[entity_type]
            if entity_set:
                print(f"   {entity_type}: {', '.join(list(entity_set)[:10])}")
    else:
        print(f"\n No entities found. To enable entity detection, run:")
        print(f" python -m spacy download en_core_web_sm")
    
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
        'summaries': [s for s in all_summaries[:5] if s],
        'raw_texts': all_raw_texts[:5],
        'entities': {k: list(v) for k, v in all_entities.items()},
        'detailed_results': results
    }
    
    # Save JSON
    json_file = f"{output_folder}/{safe_subject}_{search_keyword}_{timestamp}.json"
    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(summary_data, f, indent=2, default=str)
    
    # Save TXT with better formatting
    txt_file = f"{output_folder}/{safe_subject}_{search_keyword}_{timestamp}.txt"
    with open(txt_file, 'w', encoding='utf-8') as f:
        f.write("="*70 + "\n")
        f.write(f"KEYWORD SUMMARY REPORT: {search_keyword.upper()}\n")
        f.write("="*70 + "\n\n")
        
        f.write(f"Email Subject Filter: {subject_filter}\n")
        f.write(f"Keyword Searched: {search_keyword}\n")
        f.write(f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        
        f.write("-"*50 + "\n")
        f.write("STATISTICS\n")
        f.write("-"*50 + "\n")
        f.write(f"Total emails scanned: {total_emails}\n")
        f.write(f"Emails containing '{search_keyword}': {email_count}\n")
        f.write(f"Total occurrences: {actual_total_occurrences}\n")
        f.write(f"Success rate: {success_rate:.1f}%\n\n")
        
        # Full context examples
        if all_contexts:
            f.write("-"*50 + "\n")
            f.write("COMPLETE CONTEXT EXAMPLES (where the keyword appears)\n")
            f.write("-"*50 + "\n\n")
            for i, ctx in enumerate(all_contexts[:5], 1):
                full_context = ctx.get('full_context', '')
                full_context = re.sub(r'https?://\S+', '', full_context)
                full_context = re.sub(r'\s+', ' ', full_context)
                f.write(f"Context {i}:\n")
                f.write(f"   ...{full_context}...\n\n")
        
        # Complete sentences with the keyword
        if all_raw_texts:
            f.write("-"*50 + "\n")
            f.write("COMPLETE SENTENCES CONTAINING THE KEYWORD\n")
            f.write("-"*50 + "\n\n")
            
            combined_text = ' '.join(all_raw_texts)
            combined_text = re.sub(r'https?://\S+', '', combined_text)
            sentences = re.split(r'[.!?]+', combined_text)
            
            sentence_count = 0
            for sent in sentences:
                if search_keyword.lower() in sent.lower():
                    sentence_count += 1
                    clean_sent = sent.strip()
                    clean_sent = re.sub(r'\s+', ' ', clean_sent)
                    f.write(f"{sentence_count}. {clean_sent}\n\n")
        
        # Entities
        if all_entities:
            f.write("-"*50 + "\n")
            f.write("ENTITIES DETECTED\n")
            f.write("-"*50 + "\n")
            for entity_type in sorted(all_entities.keys()):
                entity_set = all_entities[entity_type]
                if entity_set:
                    f.write(f"{entity_type}: {', '.join(list(entity_set)[:10])}\n")
            f.write("\n")
        else:
            f.write("-"*50 + "\n")
            f.write("ENTITIES DETECTED\n")
            f.write("-"*50 + "\n")
            f.write("No entities found. To enable, run: python -m spacy download en_core_web_sm\n\n")
        
        # Emails list
        f.write("-"*50 + "\n")
        f.write("EMAILS WITH KEYWORD\n")
        f.write("-"*50 + "\n")
        email_idx = 0
        for result in results:
            keywords = result.get('keywords', {})
            keyword_data = None
            for kw in keywords.keys():
                if kw.lower() == search_keyword.lower():
                    keyword_data = keywords[kw]
                    break
            if keyword_data:
                email_idx += 1
                email_meta = result.get('email_metadata', {})
                confidence = int(keyword_data.get('confidence', 0) * 100)
                occurrences = keyword_data.get('occurrences', 0)
                f.write(f"\nEmail {email_idx}:\n")
                f.write(f" Subject: {email_meta.get('subject', 'No Subject')}\n")
                f.write(f" From: {email_meta.get('sender', 'Unknown')}\n")
                f.write(f" Date: {email_meta.get('date', 'Unknown')}\n")
                f.write(f" '{search_keyword}' found: {occurrences} times (Confidence: {confidence}%)\n")
        
        f.write("\n" + "="*70 + "\n")
        f.write("END OF REPORT\n")
        f.write("="*70 + "\n")
    
    print(f"\n JSON saved: {json_file}")
    print(f"Text saved: {txt_file}")
    
    return json_file, txt_file


def main():
    OUTPUT_FOLDER = "outputs/keyword_results"
    
    print("\n" + "="*60)
    print("OUTLOOK EMAIL KEYWORD MINER")
    print("="*60)
    print("\nThis tool extracts keywords from your Outlook emails.")
    print(" You can either:")
    print(" 1. Type keywords manually at runtime")
    print(" 2. Use predefined keywords from the configuration")
    print("="*60)
    
    # Step 1: Connect to Outlook
    print("\n1. Connecting to Outlook...")
    connector = OutlookConnector()
    
    if not connector.outlook:
        print("could not connect to Outlook.")
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
    
    print(f"\n Found {len(emails)} email(s) with subject containing '{subject_filter}'")
    
    if not emails:
        print("\n No emails found. Try a different subject or check your inbox.")
        return
    
    # Display found emails
    print("\n Emails found:")
    for i, email in enumerate(emails, 1):
        display_subject = email['subject'][:70] + "..." if len(email['subject']) > 70 else email['subject']
        print(f"  {i}. {display_subject}")
        print(f" From: {email['sender_name']} | Date: {email['received_time']}")
    
    # Step 3: Choose keyword input method
    print("\n" + "="*60)
    print("STEP 2: Choose keyword input method")
    print("="*60)
    print("[1] Type keywords manually at runtime")
    print("[2] Use predefined keywords from configuration")
    print("[3] Both (predefined + manual)")
    print("-" * 60)
    
    choice = input("\nEnter your choice (1/2/3): ").strip()
    
    keywords_to_search = []
    
    if choice == '1':
        # Manual input only
        print("\n" + "-"*40)
        print("MANUAL KEYWORD ENTRY")
        print("-"*40)
        print(" For single keyword: just type it (e.g., registration)")
        print(" For multiple keywords: separate with commas (e.g., registration, certification, exam)")
        print(" Tip: Try different word forms if not found (e.g., 'register' instead of 'registration')")
        
        keywords_input = input("\nEnter keyword(s): ").strip()
        if not keywords_input:
            print("   No keyword entered. Exiting.")
            return
        keywords_to_search = [k.strip().lower() for k in keywords_input.split(',')]
        
    elif choice == '2':
        # Predefined keywords only
        print("\n" + "-"*40)
        print("PREDEFINED KEYWORDS")
        print("-"*40)
        print(f"Available keywords: {', '.join(PREDEFINED_KEYWORDS)}")
        
        print("\n   Select keywords to search (enter numbers separated by commas):")
        for i, kw in enumerate(PREDEFINED_KEYWORDS, 1):
            print(f"{i}. {kw}")
        
        selection = input("\nEnter your choice (e.g., 1,3,5 for first, third and fifth): ").strip()
        
        if selection:
            indices = [int(x.strip()) for x in selection.split(',') if x.strip().isdigit()]
            keywords_to_search = [PREDEFINED_KEYWORDS[i-1] for i in indices if 1 <= i <= len(PREDEFINED_KEYWORDS)]
        
        if not keywords_to_search:
            print("   No valid selection. Using all predefined keywords.")
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
        print(f" {kw}")
    
    # Step 4: Save emails to files
    print("\n Saving emails to files...")
    safe_subject = re.sub(r'[^\w\s-]', '', subject_filter)[:30].replace(' ', '_')
    search_folder = f"data/outlook_emails/{safe_subject}"
    os.makedirs(search_folder, exist_ok=True)
    
    saved_files = connector.save_emails_to_files(emails, output_dir=search_folder)
    print(f"Saved {len(saved_files)} email(s)")
    
    # Step 5: Initialize keyword miner (will be created per keyword with low threshold)
    print("\n Initializing keyword miner...")
    
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
        print(f" Processing keyword {keyword_idx}/{len(keywords_to_search)}: '{keyword.upper()}'")
        print(f"{'='*80}")
        
        all_results = []
        emails_with_keyword = 0
        total_occurrences = 0
        
        # Create a new miner for each keyword with LOW threshold for better matching
        miner = DataMiner(relevance_threshold=0.05)
        
        for i, email_file in enumerate(saved_files, 1):
            print(f"\n [{i}/{len(saved_files)}] Processing: {os.path.basename(email_file)[:50]}...")
            
            try:
                # Read the email content to show preview
                with open(email_file, 'r', encoding='utf-8') as f:
                    email_content = f.read()
                
                # Quick preview of where the keyword might appear
                if keyword.lower() in email_content.lower():
                    # Find a small preview
                    pos = email_content.lower().find(keyword.lower())
                    if pos != -1:
                        preview_start = max(0, pos - 30)
                        preview_end = min(len(email_content), pos + len(keyword) + 50)
                        preview = email_content[preview_start:preview_end].replace('\n', ' ')
                        print(f" Preview found: ...{preview}...")
                
                # Mine ONLY for this specific keyword
                results = miner.mine_document(
                    document_path=email_file,
                    seed_keywords=[keyword],
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
                    print(f" '{keyword}' not found in this email")
                
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
    print(f"All results are in: {results_folder}")
    print("="*60)


if __name__ == "__main__":
    main()
