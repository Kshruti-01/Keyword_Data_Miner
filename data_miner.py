"""
Main orchestrator that ties everything together.
This is what we actually run to extract keywords.
"""
import re
import os
import json
from datetime import datetime

# Import core modules
from src.core.text_preprocessor import TextPreprocessor
from src.core.keyword_extractor import KeywordExtractor
from src.core.context_extractor import KeywordContextExtractor
from src.core.relevance_scorer import RelevanceScorer
from src.core.entity_extractor import EntityExtractor

# Import models
from src.models.confidence_scorer import ConfidenceScorer
from src.models.text_summarizer import TextSummarizer


class DataMiner:
    """
    Main class that runs the whole extraction pipeline.
    """
    
    def __init__(self):
        # Initialize all components
        self.preprocessor = TextPreprocessor()
        self.keyword_extractor = KeywordExtractor()
        self.context_extractor = KeywordContextExtractor(window=150)
        self.relevance_scorer = RelevanceScorer(threshold=0.3)
        self.entity_extractor = EntityExtractor()
        self.confidence_scorer = ConfidenceScorer()
        self.summarizer = TextSummarizer()
        
        # Store results
        self.results = {}
    
    def mine_document(self, document_path, seed_keywords=None, output_dir="outputs"):
        """
        Main method - give it a document and get back keywords with context.
        
        Args:
            document_path: Path to text file or raw text string
            seed_keywords: Optional list of keywords to look for
            output_dir: Where to save results
        """
        print(f"\n{'='*60}")
        print(f"Starting extraction for: {document_path}")
        print('='*60)
        
        # Step 1: Load and clean the document
        print("\n1. Loading document...")
        if os.path.exists(document_path):
            with open(document_path, 'r', encoding='utf-8') as f:
                raw_text = f.read()
        else:
            # Assume it's raw text
            raw_text = document_path
        
        # Step 2: Preprocess
        print("2. Cleaning and preprocessing...")
        clean_text = self.preprocessor.clean_text(raw_text)
        sentences = self.preprocessor.get_sentences(clean_text)
        stats = self.preprocessor.get_basic_stats(clean_text)
        print(f"   - {stats['sentence_count']} sentences")
        print(f"   - {stats['word_count']} words")
        
        # Step 3: Extract or use provided keywords
        print("3. Finding keywords...")
        if seed_keywords:
            keywords = seed_keywords
            print(f"   - Using {len(keywords)} provided keywords")
        else:
            # Auto-extract keywords
            keywords = self.keyword_extractor.extract_from_text(clean_text, max_keywords=20)
            print(f"   - Found {len(keywords)} keywords automatically")
        
        # Show the keywords we'll be looking for
        print("\n   Keywords to analyze:")
        for i, kw in enumerate(keywords[:10], 1):
            print(f"     {i}. {kw}")
        if len(keywords) > 10:
            print(f"     ... and {len(keywords)-10} more")
        
        # Step 4: Find contexts for each keyword
        print("\n4. Extracting contexts...")
        all_contexts = {}
        for keyword in keywords:
            contexts = self.context_extractor.find_occurrences(clean_text, keyword)
            if contexts:
                all_contexts[keyword] = contexts
                print(f"   - '{keyword}': {len(contexts)} occurrences")
        
        # Step 5: Score relevance
        print("\n5. Scoring relevance...")
        relevant_contexts = {}
        for keyword, contexts in all_contexts.items():
            # Get just the context text for scoring
            context_texts = [ctx['full_context'] for ctx in contexts]
            
            # Score each context
            scored = []
            for ctx, text in zip(contexts, context_texts):
                score = self.relevance_scorer.combined_score(text, keyword)
                ctx['relevance_score'] = score
                if score >= 0.3:  # Keep if relevant enough
                    scored.append(ctx)
            
            if scored:
                relevant_contexts[keyword] = scored
        
        print(f"   - Kept {sum(len(v) for v in relevant_contexts.values())} relevant contexts")
        
        # Step 6: Extract entities
        print("6. Finding named entities...")
        entities = self.entity_extractor.extract_entities(clean_text)
        entity_summary = self.entity_extractor.get_entity_summary(clean_text)
        print(f"   - Found {entity_summary['total_entities']} entities")
        for etype, elist in list(entities.items())[:3]:
            print(f"     {etype}: {', '.join(elist[:3])}")
        
        # Step 7: Calculate confidence scores
        print("7. Calculating confidence scores...")
        doc_stats = {
            'entities': entities,
            'total_sentences': len(sentences),
            'document_length': len(clean_text)
        }
        
        confidence_results = {}
        for keyword, contexts in relevant_contexts.items():
            score = self.confidence_scorer.score_keyword(keyword, contexts, doc_stats)
            confidence_results[keyword] = {
                'score': score,
                'occurrences': len(contexts)
            }
        
        # Show top confidence scores
        top_keywords = sorted(confidence_results.items(), 
                            key=lambda x: x[1]['score'], reverse=True)[:5]
        print("\n   Top confidence keywords:")
        for kw, data in top_keywords:
            print(f"     {kw}: {data['score']} ({data['occurrences']} occurrences)")
        
        # Step 8: Generate summaries
        print("\n8. Creating summaries...")
        summaries = {}
        for keyword, contexts in relevant_contexts.items():
            if len(contexts) > 0:
                # Combine all relevant contexts
                combined = ' '.join([c['full_context'] for c in contexts[:5]])
                summary = self.summarizer.extractive_summary(combined, num_sentences=3)
                summaries[keyword] = summary
        
        # Step 9: Prepare final results
        print("9. Packaging results...")
        
        self.results = {
            'metadata': {
                'document': document_path,
                'processed_date': datetime.now().isoformat(),
                'document_stats': stats
            },
            'keywords': {},
            'entities': entities,
            'summary': {
                'total_keywords_found': len(relevant_contexts),
                'total_contexts': sum(len(v) for v in relevant_contexts.values()),
                'top_keywords': top_keywords[:10]
            }
        }
        
        # Add detailed keyword data
        for keyword, contexts in relevant_contexts.items():
            self.results['keywords'][keyword] = {
                'occurrences': len(contexts),
                'confidence': confidence_results[keyword]['score'],
                'summary': summaries.get(keyword, ''),
                'contexts': contexts[:5]  # Limit to 5 contexts per keyword
            }
        
        # Step 10: Save results
        print("10. Saving results...")
        os.makedirs(output_dir, exist_ok=True)
        
        # Save full results as JSON
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"{output_dir}/results_{timestamp}.json"
        
        with open(output_file, 'w') as f:
            json.dump(self.results, f, indent=2, default=str)
        
        # Save a simple text report too
        report_file = f"{output_dir}/report_{timestamp}.txt"
        self._save_text_report(report_file)
        
        print(f"\nDone! Results saved to:")
        print(f"   - {output_file}")
        print(f"   - {report_file}")
        
        return self.results
    
    def _save_text_report(self, filename):
        """Save a human-readable report."""
        with open(filename, 'w') as f:
            f.write("KEYWORD EXTRACTION REPORT\n")
            f.write("="*50 + "\n\n")
            
            f.write(f"Document: {self.results['metadata']['document']}\n")
            f.write(f"Date: {self.results['metadata']['processed_date']}\n\n")
            
            f.write("DOCUMENT STATS:\n")
            f.write(f"  Words: {self.results['metadata']['document_stats']['word_count']}\n")
            f.write(f"  Sentences: {self.results['metadata']['document_stats']['sentence_count']}\n\n")
            
            f.write("TOP KEYWORDS (by confidence):\n")
            for kw, data in self.results['summary']['top_keywords']:
                f.write(f"  • {kw}: {data['score']} (found {data['occurrences']} times)\n")
            
            f.write("\n\nDETAILED RESULTS:\n")
            for keyword, data in self.results['keywords'].items():
                f.write(f"\n{'-'*40}\n")
                f.write(f"KEYWORD: {keyword}\n")
                f.write(f"Confidence: {data['confidence']}\n")
                f.write(f"Occurrences: {data['occurrences']}\n")
                
                if data['summary']:
                    f.write(f"Summary: {data['summary']}\n")
                
                f.write("\nContext examples:\n")
                for i, ctx in enumerate(data['contexts'][:3], 1):
                    f.write(f"  {i}. ...{ctx['before'][-30:]} {ctx['keyword']} {ctx['after'][:30]}...\n")
            
            f.write("\n\nENTITIES FOUND:\n")
            for etype, elist in self.results['entities'].items():
                f.write(f"  {etype}: {', '.join(elist[:5])}\n")


-----------------------------------------------------------------------------------------
updated code

"""
Main orchestrator that ties everything together.
This is what we actually run to extract keywords.
"""

import re
import os
import json
from datetime import datetime

# Import core modules
from src.core.text_preprocessor import TextPreprocessor
from src.core.keyword_extractor import KeywordExtractor
from src.core.context_extractor import KeywordContextExtractor
from src.core.relevance_scorer import RelevanceScorer
from src.core.entity_extractor import EntityExtractor

# Import models
from src.models.confidence_scorer import ConfidenceScorer
from src.models.text_summarizer import TextSummarizer


class DataMiner:
    """
    Main class that runs the whole extraction pipeline.
    """
    
    def __init__(self, relevance_threshold=0.15):
        """
        Initialize DataMiner with configurable threshold.
        
        Args:
            relevance_threshold: Minimum relevance score to keep contexts (default: 0.15)
        """
        # Initialize all components
        self.preprocessor = TextPreprocessor()
        self.keyword_extractor = KeywordExtractor()
        self.context_extractor = KeywordContextExtractor(window=150)
        self.relevance_scorer = RelevanceScorer(threshold=relevance_threshold)
        self.entity_extractor = EntityExtractor()
        self.confidence_scorer = ConfidenceScorer()
        self.summarizer = TextSummarizer()
        
        # Store results
        self.results = {}
        self.relevance_threshold = relevance_threshold
        print(f"   DataMiner initialized with relevance threshold: {relevance_threshold}")
    
    def mine_document(self, document_path, seed_keywords=None, output_dir="outputs"):
        """
        Main method - give it a document and get back keywords with context.
        
        Args:
            document_path: Path to text file or raw text string
            seed_keywords: Optional list of keywords to look for
            output_dir: Where to save results
        """
        print(f"\n{'='*60}")
        print(f"Starting extraction for: {document_path}")
        print('='*60)
        
        # Step 1: Load and clean the document
        print("\n1. Loading document...")
        if os.path.exists(document_path):
            with open(document_path, 'r', encoding='utf-8') as f:
                raw_text = f.read()
        else:
            # Assume it's raw text
            raw_text = document_path
        
        # Step 2: Preprocess
        print("Cleaning and preprocessing...")
        clean_text = self.preprocessor.clean_text(raw_text)
        sentences = self.preprocessor.get_sentences(clean_text)
        stats = self.preprocessor.get_basic_stats(clean_text)
        print(f"  {stats['sentence_count']} sentences")
        print(f"  {stats['word_count']} words")
        
        # Step 3: Extract or use provided keywords
        print("3. Finding keywords...")
        if seed_keywords:
            keywords = seed_keywords
            print(f"   - Using {len(keywords)} provided keywords")
        else:
            # Auto-extract keywords
            keywords = self.keyword_extractor.extract_from_text(clean_text, max_keywords=20)
            print(f" Found {len(keywords)} keywords automatically")
        
        # Show the keywords we'll be looking for
        print("\n   Keywords to analyze:")
        for i, kw in enumerate(keywords[:10], 1):
            print(f"     {i}. {kw}")
        if len(keywords) > 10:
            print(f"     ... and {len(keywords)-10} more")
        
        # Step 4: Find contexts for each keyword
        print("\n4. Extracting contexts...")
        all_contexts = {}
        for keyword in keywords:
            contexts = self.context_extractor.find_occurrences(clean_text, keyword)
            if contexts:
                all_contexts[keyword] = contexts
                print(f"   - '{keyword}': {len(contexts)} occurrences")
        
        # Step 5: Score relevance
        print("\n5. Scoring relevance...")
        relevant_contexts = {}
        for keyword, contexts in all_contexts.items():
            # Get just the context text for scoring
            context_texts = [ctx['full_context'] for ctx in contexts]
            
            # Score each context
            scored = []
            for ctx, text in zip(contexts, context_texts):
                score = self.relevance_scorer.combined_score(text, keyword)
                ctx['relevance_score'] = score
                if score >= self.relevance_threshold:  # Use the instance threshold
                    scored.append(ctx)
            
            if scored:
                relevant_contexts[keyword] = scored
        
        print(f"   - Kept {sum(len(v) for v in relevant_contexts.values())} relevant contexts")
        
        # Step 6: Extract entities
        print("6. Finding named entities...")
        entities = self.entity_extractor.extract_entities(clean_text)
        entity_summary = self.entity_extractor.get_entity_summary(clean_text)
        print(f"   - Found {entity_summary['total_entities']} entities")
        for etype, elist in list(entities.items())[:3]:
            print(f"     {etype}: {', '.join(elist[:3])}")
        
        # Step 7: Calculate confidence scores
        print("7. Calculating confidence scores...")
        doc_stats = {
            'entities': entities,
            'total_sentences': len(sentences),
            'document_length': len(clean_text)
        }
        
        confidence_results = {}
        for keyword, contexts in relevant_contexts.items():
            score = self.confidence_scorer.score_keyword(keyword, contexts, doc_stats)
            confidence_results[keyword] = {
                'score': score,
                'occurrences': len(contexts)
            }
        
        # Show top confidence scores
        top_keywords = sorted(confidence_results.items(), 
                            key=lambda x: x[1]['score'], reverse=True)[:5]
        print("\n   Top confidence keywords:")
        for kw, data in top_keywords:
            print(f"     {kw}: {data['score']} ({data['occurrences']} occurrences)")
        
        # Step 8: Generate summaries
        print("\n8. Creating summaries...")
        summaries = {}
        for keyword, contexts in relevant_contexts.items():
            if len(contexts) > 0:
                # Combine all relevant contexts
                combined = ' '.join([c['full_context'] for c in contexts[:5]])
                summary = self.summarizer.extractive_summary(combined, num_sentences=3)
                summaries[keyword] = summary
        
        # Step 9: Prepare final results
        print("9. Packaging results...")
        
        self.results = {
            'metadata': {
                'document': document_path,
                'processed_date': datetime.now().isoformat(),
                'document_stats': stats
            },
            'keywords': {},
            'entities': entities,
            'summary': {
                'total_keywords_found': len(relevant_contexts),
                'total_contexts': sum(len(v) for v in relevant_contexts.values()),
                'top_keywords': top_keywords[:10]
            }
        }
        
        # Add detailed keyword data
        for keyword, contexts in relevant_contexts.items():
            self.results['keywords'][keyword] = {
                'occurrences': len(contexts),
                'confidence': confidence_results[keyword]['score'],
                'summary': summaries.get(keyword, ''),
                'contexts': contexts[:5]  # Limit to 5 contexts per keyword
            }
        
        # Step 10: Save results
        print("10. Saving results...")
        os.makedirs(output_dir, exist_ok=True)
        
        # Save full results as JSON
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"{output_dir}/results_{timestamp}.json"
        
        with open(output_file, 'w') as f:
            json.dump(self.results, f, indent=2, default=str)
        
        # Save a simple text report too
        report_file = f"{output_dir}/report_{timestamp}.txt"
        self._save_text_report(report_file)
        
        print(f"\nDone! Results saved to:")
        print(f" {output_file}")
        print(f" {report_file}")
        
        return self.results
    
    def _save_text_report(self, filename):
        """Save a human-readable report."""
        with open(filename, 'w') as f:
            f.write("KEYWORD EXTRACTION REPORT\n")
            f.write("="*50 + "\n\n")
            
            f.write(f"Document: {self.results['metadata']['document']}\n")
            f.write(f"Date: {self.results['metadata']['processed_date']}\n\n")
            
            f.write("DOCUMENT STATS:\n")
            f.write(f"  Words: {self.results['metadata']['document_stats']['word_count']}\n")
            f.write(f"  Sentences: {self.results['metadata']['document_stats']['sentence_count']}\n\n")
            
            f.write("TOP KEYWORDS (by confidence):\n")
            for kw, data in self.results['summary']['top_keywords']:
                f.write(f"  • {kw}: {data['score']} (found {data['occurrences']} times)\n")
            
            f.write("\n\nDETAILED RESULTS:\n")
            for keyword, data in self.results['keywords'].items():
                f.write(f"\n{'-'*40}\n")
                f.write(f"KEYWORD: {keyword}\n")
                f.write(f"Confidence: {data['confidence']}\n")
                f.write(f"Occurrences: {data['occurrences']}\n")
                
                if data['summary']:
                    f.write(f"Summary: {data['summary']}\n")
                
                f.write("\nContext examples:\n")
                for i, ctx in enumerate(data['contexts'][:3], 1):
                    f.write(f"  {i}. ...{ctx['before'][-30:]} {ctx['keyword']} {ctx['after'][:30]}...\n")
            
            f.write("\n\nENTITIES FOUND:\n")
            for etype, elist in self.results['entities'].items():
                f.write(f"  {etype}: {', '.join(elist[:5])}\n")
