"""
Outlook Email Connector
Extracts emails directly from Microsoft Outlook.
"""

import os
import sys
from datetime import datetime, timedelta
from typing import List, Dict, Optional
import logging

logger = logging.getLogger(__name__)

# Try to import Windows Outlook connector
try:
    import win32com.client
    from win32com.client import constants
    OUTLOOK_AVAILABLE = True
except ImportError:
    OUTLOOK_AVAILABLE = False
    logger.warning("pywin32 not installed. Outlook integration disabled.")
    logger.warning("Install with: pip install pywin32")


class OutlookConnector:
    """
    Connects to Microsoft Outlook and extracts emails.
    Supports both local Outlook and Exchange/Office 365.
    """
    
    def __init__(self, profile_name=None):
        """
        Initialize Outlook connector.
        
        Args:
            profile_name: Optional Outlook profile name
        """
        self.profile_name = profile_name
        self.outlook = None
        self.namespace = None
        
        if OUTLOOK_AVAILABLE:
            self._connect_outlook()
    
    def _connect_outlook(self):
        """Establish connection to Outlook."""
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            
            # Use specific profile if provided
            if self.profile_name:
                self.namespace.Logon(self.profile_name)
            
            logger.info("✅ Connected to Outlook successfully")
        except Exception as e:
            logger.error(f"❌ Failed to connect to Outlook: {e}")
            self.outlook = None
    
    def get_inbox(self):
        """Get Inbox folder."""
        if not self.namespace:
            return None
        return self.namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
    
    def get_folder(self, folder_name: str):
        """
        Get a specific folder by name.
        
        Args:
            folder_name: Name of the folder (e.g., "Inbox", "Sent Items")
        """
        if not self.namespace:
            return None
        
        # Try to find folder in default folders
        folders = {
            "Inbox": 6,
            "Sent Items": 5,
            "Outbox": 4,
            "Drafts": 16,
            "Junk": 23
        }
        
        if folder_name in folders:
            return self.namespace.GetDefaultFolder(folders[folder_name])
        
        # Search in all folders
        return self._find_folder_by_name(self.namespace.Folders, folder_name)
    
    def _find_folder_by_name(self, folders, target_name):
        """Recursively search for folder by name."""
        for folder in folders:
            if folder.Name == target_name:
                return folder
            if folder.Folders.Count > 0:
                result = self._find_folder_by_name(folder.Folders, target_name)
                if result:
                    return result
        return None
    
    def get_emails(self, folder="Inbox", days_back=7, limit=100, 
                   unread_only=False, subject_filter=None):
        """
        Extract emails from Outlook.
        
        Args:
            folder: Folder name to extract from
            days_back: Number of days to look back
            limit: Maximum number of emails to extract
            unread_only: Only extract unread emails
            subject_filter: Filter by subject contains string
            
        Returns:
            List of email dictionaries
        """
        if not self.outlook:
            logger.error("Outlook not connected")
            return []
        
        # Get the folder
        target_folder = self.get_folder(folder)
        if not target_folder:
            logger.error(f"Folder '{folder}' not found")
            return []
        
        # Calculate date range
        date_cutoff = datetime.now() - timedelta(days=days_back)
        
        # Get items
        items = target_folder.Items
        items.Sort("[ReceivedTime]", True)  # Sort by received time descending
        
        # Filter by date
        items = items.Restrict(f"[ReceivedTime] >= '{date_cutoff.strftime('%m/%d/%Y')}'")
        
        # Filter by unread status
        if unread_only:
            items = items.Restrict("[UnRead] = True")
        
        # Filter by subject
        if subject_filter:
            items = items.Restrict(f"[Subject] LIKE '%{subject_filter}%'")
        
        # Limit results
        if limit:
            items = items[:limit]
        
        # Extract email data
        emails = []
        for item in items:
            try:
                email_data = self._extract_email_data(item)
                emails.append(email_data)
            except Exception as e:
                logger.error(f"Error extracting email: {e}")
                continue
        
        logger.info(f"✅ Extracted {len(emails)} emails from {folder}")
        return emails
    
    def _extract_email_data(self, item):
        """
        Extract relevant data from an Outlook mail item.
        
        Returns:
            Dictionary with email content and metadata
        """
        email = {
            'id': str(item.EntryID),
            'subject': str(item.Subject) if item.Subject else "",
            'sender_name': str(item.SenderName) if hasattr(item, 'SenderName') else "",
            'sender_email': str(item.SenderEmailAddress) if hasattr(item, 'SenderEmailAddress') else "",
            'recipients': self._get_recipients(item),
            'received_time': item.ReceivedTime.strftime('%Y-%m-%d %H:%M:%S') if item.ReceivedTime else "",
            'sent_time': item.SentOn.strftime('%Y-%m-%d %H:%M:%S') if hasattr(item, 'SentOn') and item.SentOn else "",
            'body': str(item.Body) if item.Body else "",
            'html_body': str(item.HTMLBody) if hasattr(item, 'HTMLBody') and item.HTMLBody else "",
            'importance': item.Importance,  # 0=Low, 1=Normal, 2=High
            'unread': item.UnRead,
            'categories': list(item.Categories) if item.Categories else [],
            'attachments': self._get_attachments(item)
        }
        
        return email
    
    def _get_recipients(self, item):
        """Extract recipient list."""
        recipients = []
        if hasattr(item, 'Recipients'):
            for recipient in item.Recipients:
                recipients.append({
                    'name': str(recipient.Name) if recipient.Name else "",
                    'email': str(recipient.Address) if hasattr(recipient, 'Address') else "",
                    'type': recipient.Type  # 1=To, 2=Cc, 3=Bcc
                })
        return recipients
    
    def _get_attachments(self, item):
        """Extract attachment information."""
        attachments = []
        if hasattr(item, 'Attachments'):
            for attachment in item.Attachments:
                attachments.append({
                    'name': str(attachment.FileName),
                    'size': attachment.Size
                })
        return attachments
    
    def search_emails(self, keywords: List[str], folder="Inbox", 
                      days_back=30, limit=100):
        """
        Search for emails containing specific keywords.
        
        Args:
            keywords: List of keywords to search for
            folder: Folder to search in
            days_back: Days to look back
            limit: Maximum results
            
        Returns:
            List of matching emails
        """
        if not self.outlook:
            return []
        
        # Get emails first
        emails = self.get_emails(folder, days_back, limit)
        
        # Filter by keywords
        matching_emails = []
        for email in emails:
            text_to_search = f"{email['subject']} {email['body']}".lower()
            matches = []
            
            for keyword in keywords:
                if keyword.lower() in text_to_search:
                    matches.append(keyword)
            
            if matches:
                email['matched_keywords'] = matches
                matching_emails.append(email)
        
        return matching_emails
    
    def mark_as_read(self, email_id):
        """Mark an email as read."""
        if not self.outlook:
            return False
        
        try:
            item = self.namespace.GetItemFromID(email_id)
            item.UnRead = False
            item.Save()
            return True
        except Exception as e:
            logger.error(f"Failed to mark email as read: {e}")
            return False
    
    def save_emails_to_files(self, emails, output_dir="data/emails"):
        """
        Save extracted emails to text files for processing.
        
        Args:
            emails: List of email dictionaries
            output_dir: Directory to save files
        """
        os.makedirs(output_dir, exist_ok=True)
        
        saved_files = []
        for email in emails:
            # Create filename from subject and timestamp
            subject = email['subject'][:50].replace('/', '_').replace('\\', '_')
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"{subject}_{timestamp}.txt"
            filepath = os.path.join(output_dir, filename)
            
            # Write email to file
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(f"Subject: {email['subject']}\n")
                f.write(f"From: {email['sender_name']} <{email['sender_email']}>\n")
                f.write(f"Date: {email['received_time']}\n")
                f.write(f"To: {', '.join([r['email'] for r in email['recipients'] if r['type'] == 1])}\n")
                
                cc_list = [r['email'] for r in email['recipients'] if r['type'] == 2]
                if cc_list:
                    f.write(f"Cc: {', '.join(cc_list)}\n")
                
                f.write(f"Importance: {['Low', 'Normal', 'High'][email['importance']]}\n")
                f.write("-" * 50 + "\n\n")
                f.write(email['body'])
            
            saved_files.append(filepath)
        
        logger.info(f"✅ Saved {len(saved_files)} emails to {output_dir}")
        return saved_files
