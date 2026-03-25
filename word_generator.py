"""
Word Document Generator for Keyword Mining Results
"""

import os
from datetime import datetime
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT


class WordReportGenerator:
    """
    Generates formatted Word documents from keyword mining results.
    """
    
    def __init__(self):
        self.document = None
    
    def create_report(self, results, keywords, total_emails, emails_with_keywords, 
                      total_keywords_found, output_path):
        """
        Create a comprehensive Word report.
        
        Args:
            results: List of email processing results
            keywords: List of keywords searched
            total_emails: Total emails processed
            emails_with_keywords: Emails containing keywords
            total_keywords_found: Total keywords discovered
            output_path: Where to save the Word document
        """
        # Create new document
        self.document = Document()
        
        # Add header
        self._add_header()
        
        # Add summary section
        self._add_summary_section(total_emails, emails_with_keywords, 
                                  total_keywords_found, keywords)
        
        # Add keyword frequency table
        self._add_keyword_frequency_table(results, keywords)
        
        # Add detailed results for each email
        self._add_detailed_results(results)
        
        # Add footer
        self._add_footer()
        
        # Save document
        self.document.save(output_path)
        print(f"✅ Word report saved to: {output_path}")
        
        return output_path
    
    def _add_header(self):
        """Add title and header to the document."""
        # Title
        title = self.document.add_heading('Email Keyword Mining Report', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Subtitle with date
        date_para = self.document.add_paragraph()
        date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        date_para.add_run(f'Generated on: {datetime.now().strftime("%B %d, %Y at %H:%M:%S")}')
        
        # Add horizontal line
        self.document.add_paragraph('_' * 50)
        
    def _add_summary_section(self, total_emails, emails_with_keywords, 
                             total_keywords_found, keywords):
        """Add executive summary section."""
        self.document.add_heading('Executive Summary', level=1)
        
        # Summary statistics table
        table = self.document.add_table(rows=5, cols=2)
        table.style = 'Light Grid Accent 1'
        
        # Fill table
        cells = [
            ("Total Emails Processed", str(total_emails)),
            ("Emails with Keywords Found", str(emails_with_keywords)),
            ("Success Rate", f"{emails_with_keywords/total_emails*100:.1f}%" if total_emails > 0 else "0%"),
            ("Total Keywords Found", str(total_keywords_found)),
            ("Keywords Searched", ", ".join(keywords[:15]) + ("..." if len(keywords) > 15 else ""))
        ]
        
        for i, (label, value) in enumerate(cells):
            row = table.rows[i]
            row.cells[0].text = label
            row.cells[1].text = value
            
            # Bold the labels
            row.cells[0].paragraphs[0].runs[0].bold = True
        
        self.document.add_paragraph()
    
    def _add_keyword_frequency_table(self, results, keywords):
        """Add keyword frequency analysis table."""
        self.document.add_heading('Keyword Frequency Analysis', level=1)
        
        # Calculate frequency
        keyword_freq = {kw: 0 for kw in keywords}
        for result in results:
            for keyword in result.get('keywords', {}):
                if keyword in keyword_freq:
                    keyword_freq[keyword] += 1
        
        # Filter and sort
        keyword_freq = {k: v for k, v in keyword_freq.items() if v > 0}
        sorted_keywords = sorted(keyword_freq.items(), key=lambda x: x[1], reverse=True)
        
        if sorted_keywords:
            # Create table
            table = self.document.add_table(rows=len(sorted_keywords) + 1, cols=3)
            table.style = 'Light Grid Accent 1'
            
            # Add header
            headers = table.rows[0].cells
            headers[0].text = "Rank"
            headers[1].text = "Keyword"
            headers[2].text = "Frequency (Emails)"
            
            # Bold headers
            for cell in headers:
                cell.paragraphs[0].runs[0].bold = True
            
            # Add data
            for i, (keyword, count) in enumerate(sorted_keywords, 1):
                row = table.rows[i]
                row.cells[0].text = str(i)
                row.cells[1].text = keyword
                row.cells[2].text = str(count)
                
                # Highlight high-frequency keywords
                if count >= 5:
                    row.cells[1].paragraphs[0].runs[0].font.color.rgb = RGBColor(0, 102, 204)
                    row.cells[1].paragraphs[0].runs[0].bold = True
            
            self.document.add_paragraph()
        else:
            self.document.add_paragraph("No keywords found in any emails.")
    
    def _add_detailed_results(self, results):
        """Add detailed results for each email."""
        self.document.add_heading('Detailed Email Analysis', level=1)
        
        for i, result in enumerate(results, 1):
            email_meta = result.get('email_metadata', {})
            keywords = result.get('keywords', {})
            
            # Email header
            self.document.add_heading(f'Email {i}: {email_meta.get("subject", "No Subject")[:80]}', level=2)
            
            # Email metadata
            metadata_table = self.document.add_table(rows=3, cols=2)
            metadata_table.style = 'Light Shading'
            
            metadata = [
                ("From", email_meta.get("sender", "Unknown")),
                ("Date", email_meta.get("date", "Unknown")),
                ("Keywords Found", str(len(keywords)))
            ]
            
            for j, (label, value) in enumerate(metadata):
                row = metadata_table.rows[j]
                row.cells[0].text = label
                row.cells[1].text = value
                row.cells[0].paragraphs[0].runs[0].bold = True
            
            # Keyword details
            if keywords:
                self.document.add_paragraph()
                self.document.add_heading('Keywords Found:', level=3)
                
                for keyword, data in keywords.items():
                    confidence = int(data.get('confidence', 0) * 100)
                    occurrences = data.get('occurrences', 0)
                    summary = data.get('summary', 'No summary available')
                    
                    # Create keyword box
                    p = self.document.add_paragraph()
                    run = p.add_run(f"🔑 {keyword.upper()}")
                    run.bold = True
                    run.font.size = Pt(12)
                    
                    p.add_run(f"\n   Confidence: {confidence}% | Occurrences: {occurrences}")
                    p.add_run(f"\n   Summary: {summary[:200]}...")
                    
                    # Add context example
                    contexts = data.get('contexts', [])
                    if contexts:
                        ctx = contexts[0]
                        p.add_run(f"\n   Example: ...{ctx.get('before', '')[-50:]} {ctx.get('keyword', '')} {ctx.get('after', '')[:50]}...")
                    
                    p.paragraph_format.space_after = Pt(12)
            else:
                self.document.add_paragraph("No keywords found in this email.")
            
            # Add separator between emails
            self.document.add_paragraph('_' * 80)
    
    def _add_footer(self):
        """Add footer to the document."""
        self.document.add_page_break()
        footer_para = self.document.add_paragraph()
        footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        footer_para.add_run(f'Report generated by Email Keyword Miner v1.0')
        footer_para.font.size = Pt(10)
        footer_para.font.italic = True


def create_quick_report(extraction_results, keywords, output_path):
    """
    Quick function to create a report from extraction results.
    
    Args:
        extraction_results: Results from the extraction process
        keywords: List of keywords searched
        output_path: Path to save the Word document
    """
    generator = WordReportGenerator()
    
    # Extract stats from results
    total_emails = len(extraction_results)
    emails_with_keywords = sum(1 for r in extraction_results if r.get('keywords'))
    total_keywords_found = sum(len(r.get('keywords', {})) for r in extraction_results)
    
    return generator.create_report(
        results=extraction_results,
        keywords=keywords,
        total_emails=total_emails,
        emails_with_keywords=emails_with_keywords,
        total_keywords_found=total_keywords_found,
        output_path=output_path
                               )
