
"""
Simple script to run the keyword miner on documents.
"""

import sys
import os
import traceback

# Add project root to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.pipeline.data_miner import DataMiner


def main():
    # Create the miner
    miner = DataMiner()
    
    # OPTION 1: Mine a file with auto-detected keywords
    print("Example 1: Auto-detect keywords")
    results = miner.mine_document(
        document_path="data/sample.txt",  # Change to your file
        output_dir="outputs"
    )
    
    # OPTION 2: Mine with specific keywords you care about
    print("\n" + "="*60)
    print("Example 2: Use specific keywords")
    results = miner.mine_document(
        document_path="data/sample.txt",
        seed_keywords=["artificial intelligence", "machine learning", "healthcare"],
        output_dir="outputs"
    )
    
    # OPTION 3: Mine raw text directly
    print("\n" + "="*60)
    print("Example 3: Mine raw text")
    sample_text = """
    Artificial intelligence is changing healthcare. Machine learning models
    can now detect diseases from medical images. Deep learning algorithms
    are particularly effective at finding patterns in complex data.
    """
    
    results = miner.mine_document(
        document_path=sample_text,  # Pass text directly
        seed_keywords=["AI", "healthcare"],
        output_dir="outputs"
    )
    
    # Show final results
    print("\n" + "="*60)
    print("FINAL SUMMARY")
    print("="*60)
    print(f"Found {results['summary']['total_keywords_found']} keywords")
    print(f"Extracted {results['summary']['total_contexts']} context snippets")
    
    print("\nTop keywords:")
    for kw, data in results['summary']['top_keywords'][:5]:
        print(f"  • {kw}: {data['score']} confidence")


if __name__ == "__main__":
    main()