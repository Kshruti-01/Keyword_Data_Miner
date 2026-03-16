"""
Figures out how relevant a piece of text is to a keyword.
"""

import re
import math


class RelevanceScorer:
    """
    Scores text segments based on keyword relevance.
    Uses simple metrics - density, position, etc.
    """
    
    def __init__(self, threshold=0.3):
        self.threshold = threshold
    
    def keyword_density(self, text, keyword):
        """
        What percentage of words is the keyword?
        Simple but effective.
        """
        if not text or not keyword:
            return 0.0
        
        words = re.findall(r'\b\w+\b', text.lower())
        if not words:
            return 0.0
        
        # Count keyword occurrences (handle multi-word)
        keyword_lower = keyword.lower()
        keyword_words = keyword_lower.split()
        
        # If multi-word, look for the exact phrase
        if len(keyword_words) > 1:
            text_lower = text.lower()
            count = text_lower.count(keyword_lower)
        else:
            # Single word - count as word matches
            count = sum(1 for w in words if w == keyword_lower)
        
        return count / len(words)
    
    def term_frequency(self, text, keyword):
        """
        Raw frequency of keyword in text.
        Returns count normalized by text length.
        """
        if not text:
            return 0.0
        
        text_lower = text.lower()
        keyword_lower = keyword.lower()
        
        # Simple count
        count = text_lower.count(keyword_lower)
        
        # Normalize by text length (avoid bias toward long texts)
        max_possible = len(text_lower) / max(1, len(keyword_lower))
        normalized = count / max_possible if max_possible > 0 else 0
        
        return min(normalized, 1.0)  # Cap at 1.0
    
    def position_score(self, text, keyword):
        """
        Keywords appearing earlier in the document often matter more.
        Returns 1.0 if at start, 0.0 if at end.
        """
        if not text:
            return 0.0
        
        text_lower = text.lower()
        keyword_lower = keyword.lower()
        
        pos = text_lower.find(keyword_lower)
        if pos == -1:
            return 0.0
        
        # Score based on position - earlier = higher score
        return 1.0 - (pos / len(text_lower))
    
    def combined_score(self, text, keyword):
        """
        Combine multiple factors into one score.
        What you'll normally call.
        """
        density = self.keyword_density(text, keyword)
        frequency = self.term_frequency(text, keyword)
        position = self.position_score(text, keyword)
        
        # Weight them (tuned through trial and error)
        # Density matters most, frequency next, position last
        score = (density * 0.5) + (frequency * 0.3) + (position * 0.2)
        
        return round(min(score, 1.0), 3)  # Cap at 1.0
    
    def score_segments(self, segments, keyword):
        """
        Score multiple text segments against a keyword.
        Returns list of (segment, score) sorted by score.
        """
        results = []
        for seg in segments:
            score = self.combined_score(seg, keyword)
            if score >= self.threshold:
                results.append((seg, score))
        
        # Sort by score descending
        results.sort(key=lambda x: x[1], reverse=True)
        return results
    
    def is_relevant(self, text, keyword, min_score=None):
        """
        Quick check if text is relevant enough.
        """
        if min_score is None:
            min_score = self.threshold
        
        score = self.combined_score(text, keyword)
        return score >= min_score
    
    def explain_score(self, text, keyword):
        """
        Break down the score components - useful for debugging.
        """
        density = self.keyword_density(text, keyword)
        frequency = self.term_frequency(text, keyword)
        position = self.position_score(text, keyword)
        
        return {
            'keyword': keyword,
            'density': round(density, 3),
            'frequency': round(frequency, 3),
            'position': round(position, 3),
            'final_score': round((density * 0.5) + (frequency * 0.3) + (position * 0.2), 3)
        }