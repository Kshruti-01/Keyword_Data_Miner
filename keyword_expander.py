"""
Keyword Expansion Module
Expands seed keywords using various techniques.
"""

from typing import List, Dict, Set, Optional, Tuple
from collections import Counter, defaultdict
import re
import numpy as np
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import logging

logger = logging.getLogger(__name__)


class KeywordExpander:
    """
    Expands seed keywords using various techniques.
    
    This class provides methods for expanding a set of seed keywords
    using word embeddings, co-occurrence analysis, and semantic similarity.
    
    Attributes:
        min_similarity (float): Minimum similarity for expansion
        max_expansions (int): Maximum number of expansions per keyword
    """
    
    def __init__(self, min_similarity: float = 0.6, max_expansions: int = 10):
        """
        Initialize the KeywordExpander.
        
        Args:
            min_similarity: Minimum similarity score for expansion
            max_expansions: Maximum number of expansions per keyword
        """
        self.min_similarity = min_similarity
        self.max_expansions = max_expansions
        self.vectorizer = TfidfVectorizer(
            stop_words='english',
            ngram_range=(1, 2),
            max_features=1000
        )
    
    def expand_by_cooccurrence(self, seed_keywords: List[str], 
                                 text: str, window: int = 50) -> Dict[str, List[str]]:
        """
        Expand keywords based on co-occurrence in text.
        
        Args:
            seed_keywords: List of seed keywords
            text: Context text
            window: Co-occurrence window size
            
        Returns:
            Dictionary mapping seed keywords to expanded terms
        """
        expansions = defaultdict(list)
        
        # Split text into segments
        words = re.findall(r'\b\w+\b', text.lower())
        
        for keyword in seed_keywords:
            keyword_lower = keyword.lower()
            keyword_words = set(keyword_lower.split())
            
            # Find positions of keyword
            positions = []
            for i, word in enumerate(words):
                if word in keyword_words:
                    positions.append(i)
            
            # Extract co-occurring terms
            co_terms = []
            for pos in positions:
                start = max(0, pos - window)
                end = min(len(words), pos + window)
                
                for i in range(start, end):
                    if i != pos:
                        term = words[i]
                        if len(term) > 2 and term not in keyword_words:
                            co_terms.append(term)
            
            # Count frequencies
            term_counts = Counter(co_terms)
            
            # Add top terms as expansions
            expansions[keyword] = [
                term for term, _ in term_counts.most_common(self.max_expansions)
            ]
        
        return dict(expansions)
    
    def expand_by_similarity(self, seed_keywords: List[str], 
                              candidate_terms: List[str]) -> Dict[str, List[str]]:
        """
        Expand keywords based on TF-IDF similarity.
        
        Args:
            seed_keywords: List of seed keywords
            candidate_terms: List of candidate expansion terms
            
        Returns:
            Dictionary mapping seed keywords to similar terms
        """
        if not seed_keywords or not candidate_terms:
            return {}
        
        expansions = defaultdict(list)
        
        try:
            # Create corpus with seeds and candidates
            corpus = seed_keywords + candidate_terms
            tfidf_matrix = self.vectorizer.fit_transform(corpus)
            
            # Get seed vectors
            seed_vectors = tfidf_matrix[:len(seed_keywords)]
            candidate_vectors = tfidf_matrix[len(seed_keywords):]
            
            # Calculate similarities
            similarities = cosine_similarity(seed_vectors, candidate_vectors)
            
            # Find top similar candidates for each seed
            for i, keyword in enumerate(seed_keywords):
                sim_scores = similarities[i]
                top_indices = np.argsort(sim_scores)[-self.max_expansions:][::-1]
                
                for idx in top_indices:
                    if sim_scores[idx] >= self.min_similarity:
                        expansions[keyword].append(candidate_terms[idx])
            
        except Exception as e:
            logger.error(f"Similarity expansion failed: {e}")
        
        return dict(expansions)
    
    def expand_by_wordnet(self, seed_keywords: List[str]) -> Dict[str, List[str]]:
        """
        Expand keywords using WordNet synonyms.
        
        Args:
            seed_keywords: List of seed keywords
            
        Returns:
            Dictionary mapping seed keywords to synonyms
        """
        expansions = defaultdict(list)
        
        try:
            from nltk.corpus import wordnet
            
            for keyword in seed_keywords:
                synonyms = set()
                
                # Get synsets for each word in keyword
                for word in keyword.split():
                    for syn in wordnet.synsets(word):
                        for lemma in syn.lemmas():
                            synonym = lemma.name().replace('_', ' ')
                            if synonym.lower() != word.lower():
                                synonyms.add(synonym)
                
                expansions[keyword] = list(synonyms)[:self.max_expansions]
                
        except ImportError:
            logger.warning("WordNet not available")
        except Exception as e:
            logger.error(f"WordNet expansion failed: {e}")
        
        return dict(expansions)
    
    def expand_by_embeddings(self, seed_keywords: List[str], 
                              vocabulary: List[str]) -> Dict[str, List[str]]:
        """
        Expand keywords using word embeddings (simplified).
        
        Args:
            seed_keywords: List of seed keywords
            vocabulary: List of vocabulary terms
            
        Returns:
            Dictionary mapping seed keywords to similar terms
        """
        # Simplified embedding-based expansion using character n-grams
        expansions = defaultdict(list)
        
        def get_ngrams(text: str, n: int = 3) -> Set[str]:
            """Get character n-grams from text."""
            text = text.lower()
            ngrams = set()
            for i in range(len(text) - n + 1):
                ngrams.add(text[i:i+n])
            return ngrams
        
        # Create n-gram profiles
        vocab_ngrams = {}
        for term in vocabulary:
            vocab_ngrams[term] = get_ngrams(term)
        
        for keyword in seed_keywords:
            keyword_ngrams = get_ngrams(keyword)
            
            # Calculate Jaccard similarity
            similarities = []
            for term in vocabulary:
                if term.lower() != keyword.lower():
                    intersection = len(keyword_ngrams.intersection(vocab_ngrams[term]))
                    union = len(keyword_ngrams.union(vocab_ngrams[term]))
                    
                    if union > 0:
                        similarity = intersection / union
                        similarities.append((term, similarity))
            
            # Sort by similarity and take top
            similarities.sort(key=lambda x: x[1], reverse=True)
            expansions[keyword] = [
                term for term, sim in similarities[:self.max_expansions]
                if sim >= self.min_similarity
            ]
        
        return dict(expansions)
    
    def combined_expansion(self, seed_keywords: List[str], text: str,
                            vocabulary: List[str]) -> Dict[str, List[str]]:
        """
        Combine multiple expansion methods.
        
        Args:
            seed_keywords: List of seed keywords
            text: Context text
            vocabulary: List of vocabulary terms
            
        Returns:
            Dictionary with combined expansions
        """
        # Get expansions from different methods
        cooccurrence = self.expand_by_cooccurrence(seed_keywords, text)
        similarity = self.expand_by_similarity(seed_keywords, vocabulary)
        wordnet = self.expand_by_wordnet(seed_keywords)
        embeddings = self.expand_by_embeddings(seed_keywords, vocabulary)
        
        # Combine and deduplicate
        combined = {}
        for keyword in seed_keywords:
            all_expansions = set()
            all_expansions.update(cooccurrence.get(keyword, []))
            all_expansions.update(similarity.get(keyword, []))
            all_expansions.update(wordnet.get(keyword, []))
            all_expansions.update(embeddings.get(keyword, []))
            
            combined[keyword] = list(all_expansions)[:self.max_expansions]
        
        return combined