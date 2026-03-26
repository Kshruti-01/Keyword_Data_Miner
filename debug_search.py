"""
Debug the actual search process.
"""

import sys
import os
from datetime import datetime, timedelta

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.connectors.outlook_connector import OutlookConnector

print("=" * 80)
print("DEBUG: ACTUAL SEARCH PROCESS")
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

# Calculate date cutoff (90 days)
date_cutoff = datetime.now() - timedelta(days=90)
print(f"Looking back 90 days (since {date_cutoff.strftime('%Y-%m-%d')})")

print("\n" + "=" * 80)
print("CHECKING FIRST 50 EMAILS WITHIN DATE RANGE")
print("=" * 80)

count = 0
genpact_count = 0

for item in items:
    if count >= 50:
        break
    
    try:
        received_time = item.ReceivedTime
        if received_time < date_cutoff:
            continue
        
        count += 1
        subject = str(item.Subject) if item.Subject else ""
        
        print(f"\n{count}. Subject: {subject[:80]}")
        print(f"   Date: {received_time}")
        print(f"   Contains 'Genpact'? ", end="")
        
        if 'Genpact' in subject or 'genpact' in subject:
            print("YES")
            genpact_count += 1
        else:
            print("NO")
            
    except Exception as e:
        print(f"Error: {e}")
        continue

print("\n" + "=" * 80)
print(f"SUMMARY: Found {genpact_count} emails with 'Genpact' in first {count} emails")
print("=" * 80)

# Now test the actual get_emails method with debug
print("\n" + "=" * 80)
print("TESTING get_emails METHOD DIRECTLY")
print("=" * 80)

# Manually test the get_emails method
emails = connector.get_emails(
    folder="Inbox",
    days_back=90,
    limit=50,
    subject_filter="Genpact"
)

print(f"\nget_emails returned: {len(emails)} emails")
