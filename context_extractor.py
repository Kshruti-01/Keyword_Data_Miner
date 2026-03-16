"""
Grabs the text surrounding keywords so you can see them in context.
"""

import re


class KeywordContextExtractor:
    """
    Finds where keywords appear and pulls out the surrounding text.
    Like grep but with more options.
    """
    
    def __init__(self, window=100):
        """
        window: how many characters to grab before and after
        """
        self.window = window
    
    def find_occurrences(self, text, keyword):
        """
        Find all places where a keyword appears.
        Returns list with position and surrounding text.
        """
        if not text or not keyword:
            return []
        
        occurrences = []
        search_text = text.lower()
        search_keyword = keyword.lower()
        
        start = 0
        while True:
            # Find next occurrence
            pos = search_text.find(search_keyword, start)
            if pos == -1:
                break
            
            # Get context windows
            context_start = max(0, pos - self.window)
            context_end = min(len(text), pos + len(keyword) + self.window)
            
            # Get the actual text (preserving original case)
            before = text[context_start:pos].strip()
            keyword_text = text[pos:pos + len(keyword)]
            after = text[pos + len(keyword):context_end].strip()
            
            # Try to get the full sentence too
            sentence = self._get_sentence_at(text, pos)
            
            occurrences.append({
                'keyword': keyword_text,
                'position': pos,
                'before': before[-50:],  # Last 50 chars before
                'after': after[:50],      # First 50 chars after
                'full_context': before + ' ' + keyword_text + ' ' + after,
                'sentence': sentence
            })
            
            # Move past this occurrence
            start = pos + len(keyword)
        
        return occurrences
    
    def _get_sentence_at(self, text, position):
        """
        Find the sentence containing a position.
        """
        # Find sentence boundaries around this position
        start = position
        end = position
        
        # Move back to sentence start
        while start > 0 and text[start] not in '.!?\n':
            start -= 1
        if start > 0:
            start += 1  # Move past the punctuation
        
        # Move forward to sentence end
        while end < len(text) and text[end] not in '.!?\n':
            end += 1
        if end < len(text):
            end += 1  # Include punctuation
        
        return text[start:end].strip()
    
    def get_contexts_for_keywords(self, text, keywords):
        """
        Get contexts for multiple keywords at once.
        Returns dict with keyword as key.
        """
        results = {}
        for kw in keywords:
            contexts = self.find_occurrences(text, kw)
            if contexts:
                results[kw] = contexts
        return results
    
    def kwic(self, text, keyword, width=40):
        """
        Keyword In Context format - useful for display.
        Returns list of (left, keyword, right) tuples.
        """
        occurrences = self.find_occurrences(text, keyword)
        kwic_list = []
        
        for occ in occurrences:
            left = occ['before'][-width:].rjust(width)
            right = occ['after'][:width].ljust(width)
            kwic_list.append((left, occ['keyword'], right))
        
        return kwic_list
    
    def count_occurrences(self, text, keyword):
        """
        Just count how many times a keyword appears.
        Simple wrapper but useful.
        """
        return len(self.find_occurrences(text, keyword))
    
    def find_nearby_words(self, text, keyword, radius=5):
        """
        Find words that commonly appear near the keyword.
        Useful for understanding context.
        """
        occurrences = self.find_occurrences(text, keyword)
        nearby = []
        
        for occ in occurrences:
            # Get words before and after
            before_words = re.findall(r'\b\w+\b', occ['before'].lower())
            after_words = re.findall(r'\b\w+\b', occ['after'].lower())
            
            nearby.extend(before_words[-radius:])  # Last few before
            nearby.extend(after_words[:radius])    # First few after
        
        # Count frequencies
        from collections import Counter
        return Counter(nearby).most_common(10)