"""
Confidence Scoring Module
Tells you how much to trust the extracted data.
"""
import re
import math
from collections import Counter


class ConfidenceScorer:
    """
    Scores how reliable an extraction is.
    
    Looks at things like:
    - How many times the keyword appears
    - How relevant the surrounding text is
    - Whether there are supporting entities
    - Consistency across different parts of the document
    
    The score is always between 0 and 1, where 1 means "very confident".
    """
    
    def __init__(self):
        # These weights come from testing - they seem to work well
        # but you can tweak them for your specific use case
        self.frequency_weight = 0.3
        self.relevance_weight = 0.4
        self.context_weight = 0.2
        self.entity_weight = 0.1
    
    def score_keyword(self, keyword, contexts, document_stats=None):
        """
        Main method - call this to get a confidence score.
        
        Args:
            keyword: The keyword you're scoring
            contexts: List of context dictionaries from the extractor
            document_stats: Optional dict with overall document info
        
        Returns:
            Float between 0 and 1
        """
        if not contexts:
            return 0.0
        
        # Calculate individual scores
        freq_score = self._frequency_score(contexts)
        relevance_score = self._relevance_score(contexts)
        context_score = self._context_quality_score(contexts)
        entity_score = self._entity_support_score(keyword, contexts, document_stats)
        
        # Combine them
        total = (
            freq_score * self.frequency_weight +
            relevance_score * self.relevance_weight +
            context_score * self.context_weight +
            entity_score * self.entity_weight
        )
        
        # Round to 2 decimal places for readability
        return round(min(total, 1.0), 2)
    
    def _frequency_score(self, contexts):
        """
        More mentions = higher confidence, but with diminishing returns.
        """
        count = len(contexts)
        
        if count == 0:
            return 0.0
        elif count == 1:
            return 0.4
        elif count == 2:
            return 0.6
        elif count <= 5:
            # 3-5 mentions: 0.7 to 0.85
            return 0.7 + ((count - 3) * 0.075)
        elif count <= 10:
            # 6-10 mentions: 0.85 to 0.95
            return 0.85 + ((count - 6) * 0.02)
        else:
            # 10+ mentions: cap at 0.98 (nothing's perfect)
            return 0.98
    
    def _relevance_score(self, contexts):
        """
        Look at how relevant each context is to the keyword.
        Uses the scores from RelevanceScorer if available.
        """
        scores = []
        
        for ctx in contexts:
            # Check if we already have a relevance score
            if 'relevance_score' in ctx:
                scores.append(ctx['relevance_score'])
            elif 'score' in ctx:
                scores.append(ctx['score'])
            else:
                # Rough estimate based on position
                if 'position' in ctx:
                    # Earlier positions are better
                    pos_score = 1.0 - (ctx['position'] / 1000)  # rough
                    scores.append(max(0.3, min(0.8, pos_score)))
                else:
                    scores.append(0.5)  # default
        
        if not scores:
            return 0.5
        
        # Average the scores
        avg_score = sum(scores) / len(scores)
        
        # Also check consistency - if scores vary a lot, be less confident
        if len(scores) > 1:
            variance = sum((s - avg_score) ** 2 for s in scores) / len(scores)
            consistency = 1.0 - min(variance, 0.5)  # penalize high variance
            avg_score = avg_score * (0.8 + (consistency * 0.2))
        
        return min(avg_score, 1.0)
    
    def _context_quality_score(self, contexts):
        """
        Check if the surrounding text is actually meaningful.
        Short or empty contexts get lower scores.
        """
        quality_scores = []
        
        for ctx in contexts:
            score = 0.7  # base score
            
            # Check context length (longer usually better)
            if 'full_context' in ctx:
                length = len(ctx['full_context'])
                if length < 20:
                    score -= 0.3
                elif length > 200:
                    score += 0.1
            elif 'sentence' in ctx:
                length = len(ctx['sentence'])
                if length < 15:
                    score -= 0.2
                elif length > 50:
                    score += 0.1
            
            # Check if context has actual words (not just punctuation)
            if 'full_context' in ctx:
                words = len(re.findall(r'\b\w+\b', ctx['full_context']))
                if words < 3:
                    score -= 0.4
                elif words > 5:
                    score += 0.1
            
            quality_scores.append(max(0.0, min(1.0, score)))
        
        if not quality_scores:
            return 0.5
        
        return sum(quality_scores) / len(quality_scores)
    
    def _entity_support_score(self, keyword, contexts, document_stats):
        """
        Check if there are named entities that support this keyword.
        For example, if keyword is "company", seeing actual company names nearby helps.
        """
        if not document_stats or 'entities' not in document_stats:
            return 0.5  # neutral if no entity info
        
        entities = document_stats.get('entities', {})
        
        # Count total entities
        total_entities = sum(len(v) for v in entities.values())
        if total_entities == 0:
            return 0.5
        
        # Look for entities near our keyword in contexts
        nearby_entities = 0
        for ctx in contexts:
            ctx_text = ctx.get('full_context', '') + ctx.get('sentence', '')
            ctx_lower = ctx_text.lower()
            
            # Check each entity type
            for ent_type, ent_list in entities.items():
                for ent in ent_list:
                    if ent.lower() in ctx_lower and ent.lower() != keyword.lower():
                        nearby_entities += 1
        
        # More nearby entities = higher confidence
        if nearby_entities == 0:
            return 0.4
        elif nearby_entities <= 3:
            return 0.6
        elif nearby_entities <= 10:
            return 0.8
        else:
            return 0.95
    
    def score_multiple_keywords(self, keyword_scores):
        """
        Score a whole set of keywords and return summary stats.
        
        Args:
            keyword_scores: Dict with keyword as key and (contexts, stats) as value
        
        Returns:
            Dict with scores and overall reliability
        """
        results = {}
        all_scores = []
        
        for keyword, (contexts, stats) in keyword_scores.items():
            score = self.score_keyword(keyword, contexts, stats)
            results[keyword] = {
                'score': score,
                'occurrences': len(contexts),
                'reliable': score >= 0.6
            }
            all_scores.append(score)
        
        # Overall reliability of this extraction run
        if all_scores:
            results['_summary'] = {
                'average_score': sum(all_scores) / len(all_scores),
                'high_confidence_keywords': sum(1 for s in all_scores if s >= 0.7),
                'low_confidence_keywords': sum(1 for s in all_scores if s < 0.4),
                'total_keywords': len(all_scores)
            }
        
        return results
    
    def explain_score(self, keyword, contexts, document_stats=None):
        """
        Break down how the score was calculated - useful for debugging.
        """
        if not contexts:
            return {"error": "No contexts provided"}
        
        freq = self._frequency_score(contexts)
        rel = self._relevance_score(contexts)
        ctx_q = self._context_quality_score(contexts)
        ent = self._entity_support_score(keyword, contexts, document_stats)
        
        total = (
            freq * self.frequency_weight +
            rel * self.relevance_weight +
            ctx_q * self.context_weight +
            ent * self.entity_weight
        )
        
        return {
            'keyword': keyword,
            'final_score': round(min(total, 1.0), 3),
            'components': {
                'frequency': round(freq, 3),
                'relevance': round(rel, 3),
                'context_quality': round(ctx_q, 3),
                'entity_support': round(ent, 3)
            },
            'weights': {
                'frequency': self.frequency_weight,
                'relevance': self.relevance_weight,
                'context_quality': self.context_weight,
                'entity_support': self.entity_weight
            },
            'occurrence_count': len(contexts)
        }
    
    def adjust_weights(self, frequency=None, relevance=None, context=None, entity=None):
        """
        Let users tweak the weights if the defaults don't work for them.
        """
        if frequency is not None:
            self.frequency_weight = frequency
        if relevance is not None:
            self.relevance_weight = relevance
        if context is not None:
            self.context_weight = context
        if entity is not None:
            self.entity_weight = entity
        
        # Renormalize to ensure they sum to 1.0
        total = (self.frequency_weight + self.relevance_weight + 
                 self.context_weight + self.entity_weight)
        
        if total != 1.0:
            self.frequency_weight /= total
            self.relevance_weight /= total
            self.context_weight /= total
            self.entity_weight /= total


# Simple utility function if you just need a quick confidence check
def quick_confidence_score(occurrence_count):
    """
    Super simple scorer based only on how many times something appears.
    Use when you don't have context data.
    """
    if occurrence_count == 0:
        return 0.0
    elif occurrence_count == 1:
        return 0.3
    elif occurrence_count == 2:
        return 0.5
    elif occurrence_count <= 5:
        return 0.7
    elif occurrence_count <= 10:
        return 0.85
    else:
        return 0.95