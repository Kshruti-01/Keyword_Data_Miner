"""
Outlook Email Connector
Extracts emails directly from Microsoft Outlook.
"""

import os
import re
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
            
            logger.info("Connected to Outlook successfully")
        except Exception as e:
            logger.error(f"Failed to connect to Outlook: {e}")
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
        
        # Extract emails with manual filtering (more reliable than Restrict)
        emails = []
        count = 0
        
        # Show search progress
        if subject_filter:
            print(f"   Searching for emails with subject containing: '{subject_filter}'...")
        elif unread_only:
            print(f"   Searching for unread emails from last {days_back} days...")
        else:
            print(f"   Searching for emails from last {days_back} days...")
        
        for item in items:
            if count >= limit:
                break
            
            try:
                # Date filter
                received_time = item.ReceivedTime
                if received_time < date_cutoff:
                    continue
                
                # Unread filter
                if unread_only:
                    try:
                        if not item.UnRead:
                            continue
                    except:
                        pass
                
                # Subject filter
                if subject_filter:
                    try:
                        subject = str(item.Subject) if item.Subject else ""
                        if subject_filter.lower() not in subject.lower():
                            continue
                    except:
                        continue
                
                # Extract email data
                email_data = self._extract_email_data(item)
                emails.append(email_data)
                count += 1
                
                # Show progress for subject search
                if subject_filter and count % 5 == 0:
                    print(f"      Found {count} emails so far...")
                    
            except Exception as e:
                logger.debug(f"Skipping email due to error: {e}")
                continue
        
        logger.info(f"Extracted {len(emails)} emails from {folder}")
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
            
        Returns:
            List of saved file paths
        """
        # Create output directory if it doesn't exist
        os.makedirs(output_dir, exist_ok=True)
        
        saved_files = []
        
        for i, email in enumerate(emails):
            # Get original subject
            original_subject = email.get('subject', '')
            
            # Create safe filename from subject
            if original_subject:
                # Clean subject for filename
                safe_subject = original_subject[:50]  # Limit length to 50 chars
                
                # Remove all invalid characters for Windows filenames
                # Invalid: \ / : * ? " < > | & % # @ ! $ ^ ( ) { } [ ] ; , . (space)
                invalid_chars = r'[\\/*?:"<>|&%#@!$^(){}[\];,]'
                safe_subject = re.sub(invalid_chars, '_', safe_subject)
                
                # Replace spaces with underscores
                safe_subject = re.sub(r'[\s]+', '_', safe_subject)
                
                # Remove multiple consecutive underscores
                safe_subject = re.sub(r'_+', '_', safe_subject)
                
                # Remove leading and trailing underscores and dots
                safe_subject = safe_subject.strip('_.')
                
                # Ensure not empty after sanitization
                if not safe_subject:
                    safe_subject = f"email_{i+1}"
            else:
                safe_subject = f"email_{i+1}"
            
            # Create timestamp for uniqueness
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            # Create filename
            filename = f"{safe_subject}_{timestamp}.txt"
            
            # Ensure filename isn't too long (Windows max is 255)
            if len(filename) > 200:
                safe_subject = safe_subject[:100]
                filename = f"{safe_subject}_{timestamp}.txt"
            
            filepath = os.path.join(output_dir, filename)
            
            # Handle duplicate filenames
            counter = 1
            original_filepath = filepath
            while os.path.exists(filepath):
                name_part = original_filepath.replace('.txt', '')
                filepath = f"{name_part}_{counter}.txt"
                counter += 1
            
            # Write email to file
            try:
                with open(filepath, 'w', encoding='utf-8') as f:
                    # Write email header
                    f.write("=" * 60 + "\n")
                    f.write("EMAIL EXTRACTED FROM OUTLOOK\n")
                    f.write("=" * 60 + "\n\n")
                    
                    f.write(f"Subject: {email.get('subject', 'No Subject')}\n")
                    f.write(f"From: {email.get('sender_name', 'Unknown')} <{email.get('sender_email', 'unknown')}>\n")
                    f.write(f"Date: {email.get('received_time', 'Unknown')}\n")
                    
                    # To recipients
                    recipients = email.get('recipients', [])
                    to_list = [r.get('email', '') for r in recipients if r.get('type') == 1]
                    if to_list:
                        f.write(f"To: {', '.join(to_list)}\n")
                    
                    # Cc recipients
                    cc_list = [r.get('email', '') for r in recipients if r.get('type') == 2]
                    if cc_list:
                        f.write(f"Cc: {', '.join(cc_list)}\n")
                    
                    # Bcc recipients (if any)
                    bcc_list = [r.get('email', '') for r in recipients if r.get('type') == 3]
                    if bcc_list:
                        f.write(f"Bcc: {', '.join(bcc_list)}\n")
                    
                    # Importance
                    importance_map = {0: 'Low', 1: 'Normal', 2: 'High'}
                    f.write(f"Importance: {importance_map.get(email.get('importance', 1), 'Normal')}\n")
                    
                    # Categories
                    categories = email.get('categories', [])
                    if categories:
                        f.write(f"Categories: {', '.join(categories)}\n")
                    
                    # Attachments
                    attachments = email.get('attachments', [])
                    if attachments:
                        f.write(f"Attachments: {', '.join([a['name'] for a in attachments])}\n")
                    
                    f.write("-" * 60 + "\n\n")
                    
                    # Write email body
                    f.write("EMAIL BODY:\n")
                    f.write("-" * 40 + "\n")
                    f.write(email.get('body', 'No content'))
                    f.write("\n\n")
                    
                    # Write footer
                    f.write("-" * 60 + "\n")
                    f.write(f"Extracted: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                
                saved_files.append(filepath)
                logger.debug(f"Saved: {filename}")
                
            except Exception as e:
                logger.error(f"Failed to save email: {e}")
                
                # Fallback - use index-based filename
                fallback_filename = f"email_{i+1}_{timestamp}.txt"
                fallback_filepath = os.path.join(output_dir, fallback_filename)
                
                try:
                    with open(fallback_filepath, 'w', encoding='utf-8') as f:
                        f.write(f"Subject: {email.get('subject', 'No Subject')}\n")
                        f.write(f"From: {email.get('sender_name', 'Unknown')}\n")
                        f.write(f"Date: {email.get('received_time', 'Unknown')}\n")
                        f.write("-" * 40 + "\n\n")
                        f.write(email.get('body', 'No content'))
                    
                    saved_files.append(fallback_filepath)
                    logger.info(f"Saved with fallback name: {fallback_filename}")
                    
                except Exception as e2:
                    logger.error(f"Completely failed to save email {i+1}: {e2}")
        
        logger.info(f"Successfully saved {len(saved_files)} out of {len(emails)} emails to {output_dir}")
        return saved_files


-------------------------------------------------------------------------------------------------
updated outlook_connectors

"""
Outlook Email Connector
Extracts emails directly from Microsoft Outlook.
"""

import os
import re
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
            
            logger.info("Connected to Outlook successfully")
        except Exception as e:
            logger.error(f"Failed to connect to Outlook: {e}")
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
    
    def _make_naive(self, dt):
        """Convert timezone-aware datetime to naive (remove timezone info)."""
        if dt is None:
            return None
        if dt.tzinfo is not None:
            return dt.replace(tzinfo=None)
        return dt
    
    def get_emails(self, folder="Inbox", days_back=90, limit=100, 
                   unread_only=False, subject_filter=None):
        """
        Extract emails from Outlook.
        
        Args:
            folder: Folder name to extract from
            days_back: Number of days to look back
            limit: Maximum number of emails to extract
            unread_only: Only extract unread emails
            subject_filter: Filter by subject contains string (case insensitive)
            
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
        
        print(f" Connected to folder: {folder}")
        
        # Calculate date range
        now_naive = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        date_cutoff = now_naive - timedelta(days=days_back)
        print(f" Looking back: {days_back} days (since {date_cutoff.strftime('%Y-%m-%d')})")
        
        # Get items
        items = target_folder.Items
        items.Sort("[ReceivedTime]", True)  # Sort by received time descending
        
        print(f" Total emails in folder: {items.Count}")
        
        # Extract emails with manual filtering
        emails = []
        count = 0
        total_checked = 0
        
        # Show search progress
        if subject_filter:
            print(f" Searching for emails with '{subject_filter}' in subject...")
        elif unread_only:
            print(f" Searching for unread emails from last {days_back} days...")
        else:
            print(f" Searching for emails from last {days_back} days...")
        
        for item in items:
            if count >= limit:
                break
            
            total_checked += 1
            
            # Show progress every 50 emails
            if total_checked % 50 == 0:
                print(f" Checked {total_checked} emails, found {count} matches...")
            
            try:
                # Get received time and make it naive for comparison
                received_time = item.ReceivedTime
                if received_time:
                    received_naive = self._make_naive(received_time)
                else:
                    continue
                
                # Date filter
                if received_naive < date_cutoff:
                    continue
                
                # Get subject
                subject = ""
                try:
                    if item.Subject:
                        subject = str(item.Subject)
                    else:
                        subject = ""
                except:
                    subject = ""
                
                # Unread filter
                if unread_only:
                    try:
                        if not item.UnRead:
                            continue
                    except:
                        pass
                
                # Subject filter - case insensitive
                if subject_filter:
                    if subject_filter.lower() not in subject.lower():
                        continue
                    else:
                        print(f" Found match #{count+1}: {subject[:60]}")
                
                # Extract email data (uses the updated _extract_email_data with HTML cleaning)
                email_data = self._extract_email_data(item)
                emails.append(email_data)
                count += 1
                
            except Exception as e:
                logger.debug(f"Error processing email: {e}")
                continue
        
        # Print summary
        print(f"\n Search Summary:")
        print(f" Total emails checked: {total_checked}")
        print(f" MATCHES FOUND: {len(emails)}")
        
        if len(emails) == 0 and subject_filter:
            print(f"\n No emails found with '{subject_filter}' in subject.")
        
        logger.info(f"Extracted {len(emails)} emails from {folder}")
        return emails
    
    def _extract_clean_text(self, html_content):
        """
        Extract clean text from HTML content.
        Removes HTML tags, URLs, and other unwanted elements.
        
        Args:
            html_content: Raw HTML content from email
            
        Returns:
            Clean plain text
        """
        if not html_content:
            return ""
        
        # Remove HTML tags
        clean = re.sub(r'<[^>]+>', ' ', html_content)
        
        # Remove extra whitespace
        clean = re.sub(r'\s+', ' ', clean)
        
        # Remove URLs
        clean = re.sub(r'https?://\S+', '', clean)
        
        # Remove image references
        clean = re.sub(r'\[\w+\]', '', clean)
        
        # Remove email forwarding markers
        clean = re.sub(r'^>+\s*', '', clean, flags=re.MULTILINE)
        
        # Remove common HTML entities
        clean = clean.replace('&nbsp;', ' ')
        clean = clean.replace('&amp;', '&')
        clean = clean.replace('&lt;', '<')
        clean = clean.replace('&gt;', '>')
        clean = clean.replace('&quot;', '"')
        
        # Remove any remaining non-printable characters
        clean = ''.join(char for char in clean if char.isprintable() or char == ' ')
        
        # Collapse multiple spaces
        clean = re.sub(r' +', ' ', clean)
        
        return clean.strip()
    
    def _extract_email_data(self, item):
        """
        Extract relevant data from an Outlook mail item.
        
        Args:
            item: Outlook MailItem object
            
        Returns:
            Dictionary with email content and metadata
        """
        # Get HTML body if available, otherwise plain text
        html_body = ""
        clean_body = ""
        
        if hasattr(item, 'HTMLBody') and item.HTMLBody:
            html_body = str(item.HTMLBody)
            # Extract clean text from HTML
            clean_body = self._extract_clean_text(html_body)
        elif item.Body:
            clean_body = str(item.Body)
        
        email = {
            'id': str(item.EntryID),
            'subject': str(item.Subject) if item.Subject else "",
            'sender_name': str(item.SenderName) if hasattr(item, 'SenderName') else "",
            'sender_email': str(item.SenderEmailAddress) if hasattr(item, 'SenderEmailAddress') else "",
            'recipients': self._get_recipients(item),
            'received_time': item.ReceivedTime.strftime('%Y-%m-%d %H:%M:%S') if item.ReceivedTime else "",
            'sent_time': item.SentOn.strftime('%Y-%m-%d %H:%M:%S') if hasattr(item, 'SentOn') and item.SentOn else "",
            'body': clean_body,  # Use clean text instead of raw HTML
            'html_body': html_body,
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
                      days_back=90, limit=100):
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
            
        Returns:
            List of saved file paths
        """
        # Create output directory if it doesn't exist
        os.makedirs(output_dir, exist_ok=True)
        
        saved_files = []
        
        for i, email in enumerate(emails):
            # Get original subject
            original_subject = email.get('subject', '')
            
            # Create safe filename from subject
            if original_subject:
                # Clean subject for filename
                safe_subject = original_subject[:50]  # Limit length to 50 chars
                
                # Remove all invalid characters for Windows filenames
                invalid_chars = r'[\\/*?:"<>|&%#@!$^(){}[\];,]'
                safe_subject = re.sub(invalid_chars, '_', safe_subject)
                
                # Replace spaces with underscores
                safe_subject = re.sub(r'[\s]+', '_', safe_subject)
                
                # Remove multiple consecutive underscores
                safe_subject = re.sub(r'_+', '_', safe_subject)
                
                # Remove leading and trailing underscores and dots
                safe_subject = safe_subject.strip('_.')
                
                # Ensure not empty after sanitization
                if not safe_subject:
                    safe_subject = f"email_{i+1}"
            else:
                safe_subject = f"email_{i+1}"
            
            # Create timestamp for uniqueness
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            # Create filename
            filename = f"{safe_subject}_{timestamp}.txt"
            
            # Ensure filename isn't too long (Windows max is 255)
            if len(filename) > 200:
                safe_subject = safe_subject[:100]
                filename = f"{safe_subject}_{timestamp}.txt"
            
            filepath = os.path.join(output_dir, filename)
            
            # Handle duplicate filenames
            counter = 1
            original_filepath = filepath
            while os.path.exists(filepath):
                name_part = original_filepath.replace('.txt', '')
                filepath = f"{name_part}_{counter}.txt"
                counter += 1
            
            # Write email to file
            try:
                with open(filepath, 'w', encoding='utf-8') as f:
                    # Write email header
                    f.write("=" * 60 + "\n")
                    f.write("EMAIL EXTRACTED FROM OUTLOOK\n")
                    f.write("=" * 60 + "\n\n")
                    
                    f.write(f"Subject: {email.get('subject', 'No Subject')}\n")
                    f.write(f"From: {email.get('sender_name', 'Unknown')} <{email.get('sender_email', 'unknown')}>\n")
                    f.write(f"Date: {email.get('received_time', 'Unknown')}\n")
                    
                    # To recipients
                    recipients = email.get('recipients', [])
                    to_list = [r.get('email', '') for r in recipients if r.get('type') == 1]
                    if to_list:
                        f.write(f"To: {', '.join(to_list)}\n")
                    
                    # Cc recipients
                    cc_list = [r.get('email', '') for r in recipients if r.get('type') == 2]
                    if cc_list:
                        f.write(f"Cc: {', '.join(cc_list)}\n")
                    
                    # Bcc recipients (if any)
                    bcc_list = [r.get('email', '') for r in recipients if r.get('type') == 3]
                    if bcc_list:
                        f.write(f"Bcc: {', '.join(bcc_list)}\n")
                    
                    # Importance
                    importance_map = {0: 'Low', 1: 'Normal', 2: 'High'}
                    f.write(f"Importance: {importance_map.get(email.get('importance', 1), 'Normal')}\n")
                    
                    # Categories
                    categories = email.get('categories', [])
                    if categories:
                        f.write(f"Categories: {', '.join(categories)}\n")
                    
                    # Attachments
                    attachments = email.get('attachments', [])
                    if attachments:
                        f.write(f"Attachments: {', '.join([a['name'] for a in attachments])}\n")
                    
                    f.write("-" * 60 + "\n\n")
                    
                    # Write clean email body
                    f.write("EMAIL BODY (CLEAN TEXT):\n")
                    f.write("-" * 40 + "\n")
                    
                    # Write the clean body text
                    body = email.get('body', 'No content')
                    if body:
                        f.write(body)
                    else:
                        f.write("No content")
                    
                    f.write("\n\n")
                    
                    # Write footer
                    f.write("-" * 60 + "\n")
                    f.write(f"Extracted: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                
                saved_files.append(filepath)
                logger.debug(f"Saved: {filename}")
                
            except Exception as e:
                logger.error(f"Failed to save email: {e}")
                
                # Fallback - use index-based filename
                fallback_filename = f"email_{i+1}_{timestamp}.txt"
                fallback_filepath = os.path.join(output_dir, fallback_filename)
                
                try:
                    with open(fallback_filepath, 'w', encoding='utf-8') as f:
                        f.write(f"Subject: {email.get('subject', 'No Subject')}\n")
                        f.write(f"From: {email.get('sender_name', 'Unknown')}\n")
                        f.write(f"Date: {email.get('received_time', 'Unknown')}\n")
                        f.write("-" * 40 + "\n\n")
                        f.write(email.get('body', 'No content'))
                    
                    saved_files.append(fallback_filepath)
                    logger.info(f"Saved with fallback name: {fallback_filename}")
                    
                except Exception as e2:
                    logger.error(f"Completely failed to save email {i+1}: {e2}")
        
        logger.info(f" Successfully saved {len(saved_files)} out of {len(emails)} emails to {output_dir}")
        return saved_files


------------------------------------------------------------
outlook connector new
"""
Outlook Email Connector
Extracts emails directly from Microsoft Outlook.
"""

import os
import re
from datetime import datetime, timedelta, timezone
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
            
            logger.info(" Connected to Outlook successfully")
        except Exception as e:
            logger.error(f"Failed to connect to Outlook: {e}")
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
    
    def _extract_clean_text(self, html_content):
        """
        Extract clean text from HTML content.
        Removes HTML tags, URLs, and other unwanted elements.
        
        Args:
            html_content: Raw HTML content from email
            
        Returns:
            Clean plain text
        """
        if not html_content:
            return ""
        
        # Remove HTML tags
        clean = re.sub(r'<[^>]+>', ' ', html_content)
        
        # Remove extra whitespace
        clean = re.sub(r'\s+', ' ', clean)
        
        # Remove URLs
        clean = re.sub(r'https?://\S+', '', clean)
        
        # Remove image references
        clean = re.sub(r'\[\w+\]', '', clean)
        
        # Remove email forwarding markers
        clean = re.sub(r'^>+\s*', '', clean, flags=re.MULTILINE)
        
        # Remove common HTML entities
        clean = clean.replace('&nbsp;', ' ')
        clean = clean.replace('&amp;', '&')
        clean = clean.replace('&lt;', '<')
        clean = clean.replace('&gt;', '>')
        clean = clean.replace('&quot;', '"')
        
        # Remove any remaining non-printable characters
        clean = ''.join(char for char in clean if char.isprintable() or char == ' ')
        
        # Collapse multiple spaces
        clean = re.sub(r' +', ' ', clean)
        
        return clean.strip()
    
    def get_emails(self, folder="Inbox", days_back=90, limit=100, 
                   unread_only=False, subject_filter=None):
        """
        Extract emails from Outlook.
        
        Args:
            folder: Folder name to extract from
            days_back: Number of days to look back
            limit: Maximum number of emails to extract
            unread_only: Only extract unread emails
            subject_filter: Filter by subject contains string (case insensitive)
            
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
        
        print(f" Connected to folder: {folder}")
        
        # Calculate date cutoff - using UTC to avoid timezone issues
        date_cutoff = datetime.now(timezone.utc) - timedelta(days=days_back)
        print(f"Looking back: {days_back} days (since {date_cutoff.strftime('%Y-%m-%d')})")
        
        # Get items
        items = target_folder.Items
        items.Sort("[ReceivedTime]", True)  # Sort by received time descending
        
        print(f" Total emails in folder: {items.Count}")
        
        # Extract emails with manual filtering
        emails = []
        count = 0
        total_checked = 0
        
        # Show search progress
        if subject_filter:
            print(f" Searching for emails with '{subject_filter}' in subject...")
        elif unread_only:
            print(f" Searching for unread emails from last {days_back} days...")
        else:
            print(f" Searching for emails from last {days_back} days...")
        
        for item in items:
            if count >= limit:
                break
            
            total_checked += 1
            
            # Show progress every 50 emails
            if total_checked % 50 == 0:
                print(f"Checked {total_checked} emails, found {count} matches...")
            
            try:
                # Get received time - handle timezone properly
                received_time = item.ReceivedTime
                if received_time:
                    # Convert to UTC if it has timezone
                    if hasattr(received_time, 'tzinfo') and received_time.tzinfo:
                        received_time = received_time.astimezone(timezone.utc)
                        received_time = received_time.replace(tzinfo=None)
                else:
                    continue
                
                # Date filter - compare dates
                if received_time < date_cutoff.replace(tzinfo=None):
                    continue
                
                # Get subject
                subject = ""
                try:
                    if item.Subject:
                        subject = str(item.Subject)
                    else:
                        subject = ""
                except:
                    subject = ""
                
                # Unread filter
                if unread_only:
                    try:
                        if not item.UnRead:
                            continue
                    except:
                        pass
                
                # Subject filter - case insensitive
                if subject_filter:
                    if subject_filter.lower() not in subject.lower():
                        continue
                    else:
                        print(f" Found match #{count+1}: {subject[:60]}")
                
                # Extract email data
                email_data = self._extract_email_data(item)
                emails.append(email_data)
                count += 1
                
            except Exception as e:
                logger.debug(f"Error processing email: {e}")
                continue
        
        # Print summary
        print(f"\n Search Summary:")
        print(f" Total emails checked: {total_checked}")
        print(f" MATCHES FOUND: {len(emails)}")
        
        if len(emails) == 0 and subject_filter:
            print(f"\n No emails found with '{subject_filter}' in subject.")
            print(f" Showing first 10 subjects to help debug:")
            
            # Show recent subjects for debugging
            debug_count = 0
            for item in items:
                if debug_count >= 10:
                    break
                try:
                    # Check if within date range
                    rcv_time = item.ReceivedTime
                    if rcv_time:
                        if hasattr(rcv_time, 'tzinfo') and rcv_time.tzinfo:
                            rcv_time = rcv_time.astimezone(timezone.utc)
                            rcv_time = rcv_time.replace(tzinfo=None)
                        if rcv_time >= date_cutoff.replace(tzinfo=None):
                            subj = str(item.Subject) if item.Subject else ""
                            if subj:
                                print(f"      {debug_count+1}. {subj[:70]}")
                                debug_count += 1
                except:
                    continue
        
        logger.info(f"Extracted {len(emails)} emails from {folder}")
        return emails
    
    def _extract_email_data(self, item):
        """
        Extract relevant data from an Outlook mail item.
        
        Args:
            item: Outlook MailItem object
            
        Returns:
            Dictionary with email content and metadata
        """
        # Get HTML body if available, otherwise plain text
        html_body = ""
        clean_body = ""
        
        if hasattr(item, 'HTMLBody') and item.HTMLBody:
            html_body = str(item.HTMLBody)
            # Extract clean text from HTML
            clean_body = self._extract_clean_text(html_body)
        elif item.Body:
            clean_body = str(item.Body)
        
        email = {
            'id': str(item.EntryID),
            'subject': str(item.Subject) if item.Subject else "",
            'sender_name': str(item.SenderName) if hasattr(item, 'SenderName') else "",
            'sender_email': str(item.SenderEmailAddress) if hasattr(item, 'SenderEmailAddress') else "",
            'recipients': self._get_recipients(item),
            'received_time': item.ReceivedTime.strftime('%Y-%m-%d %H:%M:%S') if item.ReceivedTime else "",
            'sent_time': item.SentOn.strftime('%Y-%m-%d %H:%M:%S') if hasattr(item, 'SentOn') and item.SentOn else "",
            'body': clean_body,  # Use clean text instead of raw HTML
            'html_body': html_body,
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
                      days_back=90, limit=100):
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
            
        Returns:
            List of saved file paths
        """
        # Create output directory if it doesn't exist
        os.makedirs(output_dir, exist_ok=True)
        
        saved_files = []
        
        for i, email in enumerate(emails):
            # Get original subject
            original_subject = email.get('subject', '')
            
            # Create safe filename from subject
            if original_subject:
                # Clean subject for filename
                safe_subject = original_subject[:50]
                
                # Remove invalid characters
                invalid_chars = r'[\\/*?:"<>|&%#@!$^(){}[\];,]'
                safe_subject = re.sub(invalid_chars, '_', safe_subject)
                safe_subject = re.sub(r'[\s]+', '_', safe_subject)
                safe_subject = re.sub(r'_+', '_', safe_subject)
                safe_subject = safe_subject.strip('_.')
                
                if not safe_subject:
                    safe_subject = f"email_{i+1}"
            else:
                safe_subject = f"email_{i+1}"
            
            # Create timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            # Create filename
            filename = f"{safe_subject}_{timestamp}.txt"
            
            if len(filename) > 200:
                safe_subject = safe_subject[:100]
                filename = f"{safe_subject}_{timestamp}.txt"
            
            filepath = os.path.join(output_dir, filename)
            
            # Handle duplicates
            counter = 1
            original_filepath = filepath
            while os.path.exists(filepath):
                name_part = original_filepath.replace('.txt', '')
                filepath = f"{name_part}_{counter}.txt"
                counter += 1
            
            # Write email to file
            try:
                with open(filepath, 'w', encoding='utf-8') as f:
                    f.write("=" * 60 + "\n")
                    f.write("EMAIL EXTRACTED FROM OUTLOOK\n")
                    f.write("=" * 60 + "\n\n")
                    
                    f.write(f"Subject: {email.get('subject', 'No Subject')}\n")
                    f.write(f"From: {email.get('sender_name', 'Unknown')} <{email.get('sender_email', 'unknown')}>\n")
                    f.write(f"Date: {email.get('received_time', 'Unknown')}\n")
                    
                    # To recipients
                    recipients = email.get('recipients', [])
                    to_list = [r.get('email', '') for r in recipients if r.get('type') == 1]
                    if to_list:
                        f.write(f"To: {', '.join(to_list)}\n")
                    
                    # Cc recipients
                    cc_list = [r.get('email', '') for r in recipients if r.get('type') == 2]
                    if cc_list:
                        f.write(f"Cc: {', '.join(cc_list)}\n")
                    
                    # Importance
                    importance_map = {0: 'Low', 1: 'Normal', 2: 'High'}
                    f.write(f"Importance: {importance_map.get(email.get('importance', 1), 'Normal')}\n")
                    
                    f.write("-" * 60 + "\n\n")
                    f.write("EMAIL BODY (CLEAN TEXT):\n")
                    f.write("-" * 40 + "\n")
                    f.write(email.get('body', 'No content'))
                    f.write("\n\n")
                    f.write("-" * 60 + "\n")
                    f.write(f"Extracted: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                
                saved_files.append(filepath)
                logger.debug(f" Saved: {filename}")
                
            except Exception as e:
                logger.error(f"Failed to save email: {e}")
                fallback_filename = f"email_{i+1}_{timestamp}.txt"
                fallback_filepath = os.path.join(output_dir, fallback_filename)
                
                try:
                    with open(fallback_filepath, 'w', encoding='utf-8') as f:
                        f.write(f"Subject: {email.get('subject', 'No Subject')}\n")
                        f.write(f"From: {email.get('sender_name', 'Unknown')}\n")
                        f.write("-" * 40 + "\n\n")
                        f.write(email.get('body', 'No content'))
                    saved_files.append(fallback_filepath)
                    logger.info(f"Saved with fallback name: {fallback_filename}")
                except Exception as e2:
                    logger.error(f"Failed to save email {i+1}: {e2}")
        
        logger.info(f"Successfully saved {len(saved_files)} out of {len(emails)} emails to {output_dir}")
        return saved_files
