"""
Batch processor for large email datasets.
"""

import os
import json
from datetime import datetime
from typing import List, Dict
from tqdm import tqdm  # Progress bar

class BatchProcessor:
    """Process multiple emails efficiently."""
    
    def __init__(self, data_miner, batch_size=100):
        self.miner = data_miner
        self.batch_size = batch_size
        self.results = []
    
    def process_folder(self, folder_path: str, keywords: List[str], output_dir: str = "outputs"):
        """
        Process all emails in a folder.
        """
        # Get all email files
        email_files = [f for f in os.listdir(folder_path) 
                      if f.endswith(('.txt', '.eml', '.msg'))]
        
        print(f"Found {len(email_files)} emails to process")
        
        # Process in batches
        for i in range(0, len(email_files), self.batch_size):
            batch = email_files[i:i + self.batch_size]
            print(f"\nProcessing batch {i//self.batch_size + 1}/{(len(email_files)-1)//self.batch_size + 1}")
            
            for email_file in tqdm(batch, desc="Processing emails"):
                file_path = os.path.join(folder_path, email_file)
                
                try:
                    # Process single email
                    result = self.miner.mine_document(
                        document_path=file_path,
                        seed_keywords=keywords,
                        output_dir=f"{output_dir}/individual/{email_file}"
                    )
                    
                    # Store result
                    self.results.append({
                        'file': email_file,
                        'keywords_found': len(result.get('keywords', {})),
                        'entities_found': len(result.get('entities', {})),
                        'timestamp': datetime.now().isoformat()
                    })
                    
                except Exception as e:
                    print(f"Error processing {email_file}: {e}")
                    continue
        
        # Generate summary report
        self._generate_summary_report(output_dir)
        
        return self.results
    
    def _generate_summary_report(self, output_dir: str):
        """
        Generate summary of all processed emails.
        """
        summary = {
            'processed_date': datetime.now().isoformat(),
            'total_emails': len(self.results),
            'successful': len([r for r in self.results if r.get('keywords_found', 0) > 0]),
            'failed': len([r for r in self.results if r.get('keywords_found', 0) == 0]),
            'total_keywords_found': sum(r.get('keywords_found', 0) for r in self.results),
            'total_entities_found': sum(r.get('entities_found', 0) for r in self.results),
            'results': self.results
        }
        
        # Save summary
        with open(f"{output_dir}/batch_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json", 'w') as f:
            json.dump(summary, f, indent=2)
        
        print(f"\n✅ Batch processing complete!")
        print(f"   Processed: {summary['total_emails']} emails")
        print(f"   Keywords found: {summary['total_keywords_found']}")
        print(f"   Entities found: {summary['total_entities_found']}")
