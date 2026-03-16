"""
Entity Extraction Module
Extracts named entities and key phrases from text.
"""

import spacy
from typing import List, Dict, Set, Tuple, Optional
from collections import Counter, defaultdict
import re
import logging

logger = logging.getLogger(__name__)


class EntityExtractor:
    """
    Extracts named entities and key phrases from text.
    
    This class uses spaCy for named entity recognition and provides
    methods for extracting and analyzing entities in text.
    
    Attributes:
        nlp (spacy.Language): Loaded spaCy language model
        entity_types (List[str]): Types of entities to extract
    """
    
    def __init__(self, model_name: str = "en_core_web_sm", 
                 entity_types: Optional[List[str]] = None):
        """
        Initialize the EntityExtractor.
        
        Args:
            model_name: Name of spaCy model to load
            entity_types: List of entity types to extract (None for all)
        """
        self.model_name = model_name
        self.entity_types = entity_types or [
            'PERSON', 'ORG', 'GPE', 'DATE', 'PRODUCT', 'EVENT',
            'LAW', 'LANGUAGE', 'MONEY', 'PERCENT', 'TIME', 'WORK_OF_ART'
        ]
        
        # Load spaCy model
        try:
            self.nlp = spacy.load(model_name)
            logger.info(f"Loaded spaCy model: {model_name}")
        except OSError:
            logger.warning(f"spaCy model {model_name} not found. "
                          f"Run: python -m spacy download {model_name}")
            # Load minimal model as fallback
            self.nlp = spacy.blank("en")
    
    def extract_entities(self, text: str) -> Dict[str, List[str]]:
        """
        Extract named entities from text.
        
        Args:
            text: Input text
            
        Returns:
            Dictionary mapping entity types to lists of entities
        """
        if not text or not isinstance(text, str):
            return {}
        
        doc = self.nlp(text)
        entities = defaultdict(list)
        
        for ent in doc.ents:
            # Filter by entity types if specified
            if self.entity_types and ent.label_ in self.entity_types:
                if ent.text not in entities[ent.label_]:
                    entities[ent.label_].append(ent.text)
        
        return dict(entities)
    
    def extract_entities_with_context(self, text: str, window: int = 50) -> List[Dict]:
        """
        Extract entities with surrounding context.
        
        Args:
            text: Input text
            window: Context window size in characters
            
        Returns:
            List of dictionaries with entity, type, context, and position
        """
        doc = self.nlp(text)
        entities_with_context = []
        
        for ent in doc.ents:
            if self.entity_types and ent.label_ not in self.entity_types:
                continue
            
            # Get context window
            start = max(0, ent.start_char - window)
            end = min(len(text), ent.end_char + window)
            
            context = text[start:end]
            
            entities_with_context.append({
                'entity': ent.text,
                'type': ent.label_,
                'context': context,
                'start_pos': ent.start_char,
                'end_pos': ent.end_char,
                'sentence': ent.sent.text.strip()
            })
        
        return entities_with_context
    
    def extract_key_phrases(self, text: str, top_k: int = 10) -> List[str]:
        """
        Extract key phrases using noun chunks and important terms.
        
        Args:
            text: Input text
            top_k: Number of top phrases to return
            
        Returns:
            List of key phrases
        """
        doc = self.nlp(text)
        
        # Extract noun chunks
        noun_chunks = []
        for chunk in doc.noun_chunks:
            # Filter out very short chunks and those with stopwords only
            if len(chunk.text.split()) >= 2:
                noun_chunks.append(chunk.text.lower())
        
        # Extract proper nouns and important terms
        important_terms = []
        for token in doc:
            if token.pos_ in ['PROPN', 'NOUN', 'ADJ'] and not token.is_stop:
                if len(token.text) > 2:  # Filter very short terms
                    important_terms.append(token.text.lower())
        
        # Count frequencies
        chunk_freq = Counter(noun_chunks)
        term_freq = Counter(important_terms)
        
        # Combine and score
        phrase_scores = {}
        
        # Score noun chunks
        for phrase, freq in chunk_freq.most_common(top_k * 2):
            # Boost score for phrases with proper nouns
            boost = 1.0
            if any(token.pos_ == 'PROPN' for token in doc if token.text.lower() in phrase):
                boost = 1.5
            
            phrase_scores[phrase] = freq * boost
        
        # Score individual terms
        for term, freq in term_freq.most_common(top_k):
            if term not in phrase_scores:
                phrase_scores[term] = freq * 0.5  # Lower weight for single terms
        
        # Sort by score and return top_k
        sorted_phrases = sorted(phrase_scores.items(), key=lambda x: x[1], reverse=True)
        
        return [phrase for phrase, _ in sorted_phrases[:top_k]]
    
    def get_entity_frequencies(self, text: str) -> Dict[str, Dict[str, int]]:
        """
        Get frequency counts for entities by type.
        
        Args:
            text: Input text
            
        Returns:
            Nested dictionary: {entity_type: {entity: count}}
        """
        doc = self.nlp(text)
        frequencies = defaultdict(lambda: defaultdict(int))
        
        for ent in doc.ents:
            if self.entity_types and ent.label_ in self.entity_types:
                frequencies[ent.label_][ent.text] += 1
        
        # Convert to regular dict
        return {k: dict(v) for k, v in frequencies.items()}
    
    def find_entity_relationships(self, text: str, max_distance: int = 100) -> List[Dict]:
        """
        Find relationships between entities based on proximity.
        
        Args:
            text: Input text
            max_distance: Maximum character distance for relationship
            
        Returns:
            List of relationship dictionaries
        """
        entities = self.extract_entities_with_context(text)
        relationships = []
        
        # Sort entities by position
        entities.sort(key=lambda x: x['start_pos'])
        
        # Find nearby entities
        for i, ent1 in enumerate(entities):
            for ent2 in entities[i+1:]:
                # Check distance
                distance = ent2['start_pos'] - ent1['end_pos']
                
                if 0 < distance < max_distance:
                    # Get text between entities
                    between_text = text[ent1['end_pos']:ent2['start_pos']].strip()
                    
                    relationships.append({
                        'entity1': ent1['entity'],
                        'type1': ent1['type'],
                        'entity2': ent2['entity'],
                        'type2': ent2['type'],
                        'distance': distance,
                        'connector_text': between_text[:100] if between_text else "",
                        'context': text[
                            max(0, ent1['start_pos'] - 50):
                            min(len(text), ent2['end_pos'] + 50)
                        ]
                    })
        
        return relationships
    
    def extract_custom_entities(self, text: str, patterns: List[Dict]) -> List[Dict]:
        """
        Extract custom entities using regex patterns.
        
        Args:
            text: Input text
            patterns: List of dicts with 'type' and 'pattern' keys
            
        Returns:
            List of custom entities found
        """
        custom_entities = []
        
        for pattern_dict in patterns:
            pattern = pattern_dict.get('pattern', '')
            entity_type = pattern_dict.get('type', 'CUSTOM')
            
            if not pattern:
                continue
            
            matches = re.finditer(pattern, text, re.IGNORECASE)
            
            for match in matches:
                custom_entities.append({
                    'entity': match.group(),
                    'type': entity_type,
                    'start_pos': match.start(),
                    'end_pos': match.end(),
                    'context': text[
                        max(0, match.start() - 50):
                        min(len(text), match.end() + 50)
                    ]
                })
        
        return custom_entities
    
    def get_entity_summary(self, text: str) -> Dict:
        """
        Get summary statistics of entities in text.
        
        Args:
            text: Input text
            
        Returns:
            Dictionary with entity statistics
        """
        entities = self.extract_entities(text)
        entities_with_context = self.extract_entities_with_context(text)
        
        total_entities = sum(len(e) for e in entities.values())
        
        # Calculate entity density
        words = len(text.split())
        entity_density = (total_entities / words) if words > 0 else 0
        
        # Get unique entities
        unique_entities = set()
        for entity_list in entities.values():
            unique_entities.update(entity_list)
        
        return {
            'total_entities': total_entities,
            'unique_entities': len(unique_entities),
            'entity_types': list(entities.keys()),
            'entity_density': round(entity_density, 4),
            'most_common_type': max(entities.items(), key=lambda x: len(x[1]))[0] if entities else None,
            'entities_found': entities,
            'total_occurrences': len(entities_with_context)
        }