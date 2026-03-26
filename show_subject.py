"""
Show all subjects in inbox to see what keywords actually exist.
"""

import sys
import os
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.connectors.outlook_connector import OutlookConnector

print("=" * 80)
print("SHOWING ALL EMAIL SUBJECTS IN INBOX")
print("=" * 80)

# Connect to Outlook
connector = OutlookConnector()

if not connector.outlook:
    print("Could not connect to Outlook")
    sys.exit(1)

print("Connected to Outlook")

# Get Inbox folder
inbox = connector.get_folder("Inbox")
if not inbox:
    print("Could not find Inbox folder")
    sys.exit(1)

print("Found Inbox folder")

# Get items
items = inbox.Items
items.Sort("[ReceivedTime]", True)

print(f"\nTotal emails in inbox: {items.Count}")
print("-" * 80)

# Show first 100 subjects
count = 0
subjects_list = []

for item in items:
    if count >= 100:
        break
    
    try:
        subject = str(item.Subject) if item.Subject else "[No Subject]"
        received = item.ReceivedTime
        
        subjects_list.append(subject)
        count += 1
        
        print(f"{count}. {subject[:70]}")
        
    except Exception as e:
        print(f"Error: {e}")
        continue

print("\n" + "=" * 80)
print(f"Showing first {count} emails")
print("=" * 80)

# Check for keywords
keywords_to_check = ["Genpact", "genpact", "GENPACT", "Genpact", "meeting", "project", "update", "report", "urgent", "deadline"]

print("\nChecking for keywords in subjects:")
print("-" * 60)

for keyword in keywords_to_check:
    found = False
    matching_subjects = []
    for subj in subjects_list:
        if keyword.lower() in subj.lower():
            found = True
            matching_subjects.append(subj[:60])
    
    if found:
        print(f"\n'{keyword}': Found in {len(matching_subjects)} emails")
        for sub in matching_subjects[:3]:
            print(f"   - {sub}")
    else:
        print(f"\n'{keyword}': NOT found in any of the first 100 emails")

print("\n" + "=" * 80)
print("If you're looking for 'Genpact' but don't see it in the subjects,")
print("the emails might be:")
print("  1. Older than 100 emails (scroll further back)")
print("  2. In a different folder (Sent Items, Archive, etc.)")
print("  3. Not containing 'Genpact' in the subject (maybe only in body)")
print("=" * 80)
