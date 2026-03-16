"""
Finds and ranks keywords in text.
Nothing fancy - just frequency and position based.
"""

import re
from collections import Counter


class KeywordExtractor:
    """
    Simple keyword extraction based on what actually appears in the text.
    No ML, no complex models - just counting and basic filtering.
    """
    
    def __init__(self, min_word_length=3):
        self.min_word_length = min_word_length
        # Common words that are rarely useful as keywords
        self.common_words = set([
            'this', 'that', 'these', 'those', 'with', 'from', 'have', 
            'will', 'can', 'all', 'are', 'was', 'were', 'been', 'has',
            'had', 'but', 'for', 'not', 'get', 'got', 'etc'
        ])
    
    def by_frequency(self, text, top_n=20):
        """
        Just count words and return the most frequent ones.
        Surprisingly effective for many documents.
        """
        # Find all words (3+ chars)
        words = re.findall(r'\b\w{3,}\b', text.lower())
        
        # Filter out common words
        words = [w for w in words if w not in self.common_words]
        
        # Count
        word_counts = Counter(words)
        
        # Return top N
        return word_counts.most_common(top_n)
    
    def by_position(self, text, top_n=15):
        """
        Words that appear early in the document often matter more.
        Combines frequency with position weighting.
        """
        words = re.findall(r'\b\w{3,}\b', text.lower())
        
        # Give higher weight to words appearing earlier
        weighted = {}
        total_words = len(words)
        
        for idx, word in enumerate(words):
            if word in self.common_words:
                continue
            
            # Position weight: earlier words get higher score
            position_weight = 1.0 - (idx / total_words) if total_words > 0 else 1.0
            
            if word in weighted:
                weighted[word] += position_weight
            else:
                weighted[word] = position_weight
        
        # Sort by weighted score
        sorted_words = sorted(weighted.items(), key=lambda x: x[1], reverse=True)
        
        return sorted_words[:top_n]
    
    def get_phrases(self, text, max_words=3, top_n=10):
        """
        Find common multi-word phrases.
        Looks for repeated word sequences.
        """
        # Clean and split
        text = re.sub(r'[^\w\s]', ' ', text.lower())
        words = text.split()
        
        phrases = []
        
        # Look for 2 and 3 word phrases
        for n in [2, 3]:
            for i in range(len(words) - n + 1):
                phrase = ' '.join(words[i:i+n])
                # Skip if it starts/ends with common words
                phrase_words = phrase.split()
                if (phrase_words[0] not in self.common_words and 
                    phrase_words[-1] not in self.common_words):
                    phrases.append(phrase)
        
        # Count frequencies
        phrase_counts = Counter(phrases)
        
        # Return most common
        return [p for p, _ in phrase_counts.most_common(top_n)]
    
    def extract_from_text(self, text, max_keywords=25):
        """
        Combined extraction - uses both frequency and phrases.
        This is what you'd normally call.
        """
        # Get frequent words
        freq_words = self.by_frequency(text, top_n=max_keywords)
        freq_list = [w for w, _ in freq_words]
        
        # Get phrases
        phrases = self.get_phrases(text, top_n=10)
        
        # Combine, remove duplicates
        all_keywords = list(set(freq_list + phrases))
        
        # Score them roughly
        scored = []
        for kw in all_keywords:
            # Count occurrences
            count = text.lower().count(kw.lower())
            # Length bonus (longer phrases often more specific)
            length_bonus = min(1.5, 1.0 + (len(kw.split()) * 0.1))
            score = count * length_bonus
            scored.append((kw, score))
        
        # Sort by score
        scored.sort(key=lambda x: x[1], reverse=True)
        
        return [kw for kw, _ in scored[:max_keywords]]