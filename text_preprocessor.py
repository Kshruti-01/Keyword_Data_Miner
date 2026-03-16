"""
Basic text cleaning and preparation functions.
"""

import re
import nltk
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Check if NLTK is available and download data
NLTK_AVAILABLE = True
try:
    # Try to import NLTK modules
    from nltk.tokenize import sent_tokenize, word_tokenize
    from nltk.corpus import stopwords
    
    # Download required data if needed
    try:
        nltk.data.find('tokenizers/punkt_tab')
    except LookupError:
        logger.info("Downloading NLTK punkt_tab data...")
        nltk.download('punkt_tab', quiet=True)
    
    try:
        nltk.data.find('tokenizers/punkt')
    except LookupError:
        logger.info("Downloading NLTK punkt data...")
        nltk.download('punkt', quiet=True)
    
    try:
        nltk.data.find('corpora/stopwords')
    except LookupError:
        logger.info("Downloading NLTK stopwords data...")
        nltk.download('stopwords', quiet=True)
        
except ImportError as e:
    NLTK_AVAILABLE = False
    logger.warning(f"NLTK import failed: {e}. Using regex fallbacks.")
except Exception as e:
    NLTK_AVAILABLE = False
    logger.warning(f"NLTK initialization failed: {e}. Using regex fallbacks.")


class TextPreprocessor:
    """
    Takes raw text and turns it into something usable.
    Handles cleaning, sentence splitting, and basic tokenization.
    """
    
    def __init__(self, language='english'):
        self.language = language
        self.stop_words = set()
        
        # Try to load stopwords if NLTK is available
        if NLTK_AVAILABLE:
            try:
                from nltk.corpus import stopwords
                self.stop_words = set(stopwords.words(language))
                logger.info(f"Loaded stopwords for {language}")
            except Exception as e:
                logger.warning(f"Could not load stopwords: {e}")
        
        # Fallback stopwords if NLTK failed or not available
        if not self.stop_words:
            self.stop_words = set([
                'the', 'a', 'an', 'and', 'or', 'but', 'in', 'on', 'at',
                'to', 'for', 'of', 'with', 'by', 'from', 'up', 'about',
                'into', 'through', 'during', 'before', 'after', 'then',
                'than', 'that', 'this', 'these', 'those', 'have', 'has',
                'had', 'will', 'would', 'could', 'should', 'may', 'might'
            ])
    
    def clean_text(self, text):
        """
        Remove weird characters and normalize whitespace.
        Keeps basic punctuation because we need sentences later.
        """
        if not isinstance(text, str):
            text = str(text)
        
        # Replace multiple spaces/tabs/newlines with single space
        text = re.sub(r'\s+', ' ', text)
        
        # Remove special characters but keep periods, question marks, exclamation
        text = re.sub(r'[^\w\s\.\?\!]', '', text)
        
        # Trim
        text = text.strip()
        
        return text
    
    def get_sentences(self, text):
        """
        Split text into sentences.
        Returns empty list if something goes wrong.
        """
        if not text:
            return []
        
        # Try NLTK first
        if NLTK_AVAILABLE:
            try:
                from nltk.tokenize import sent_tokenize
                sentences = sent_tokenize(text)
                return [s.strip() for s in sentences if s.strip()]
            except Exception as e:
                logger.debug(f"NLTK sentence tokenization failed: {e}")
        
        # Fallback - split on punctuation
        raw = re.split(r'[.!?]+', text)
        return [s.strip() for s in raw if s.strip()]
    
    def tokenize(self, text, remove_stops=False, min_length=2):
        """
        Break text into words. Can filter stopwords and short words.
        """
        if not text:
            return []
        
        words = []
        
        # Try NLTK first
        if NLTK_AVAILABLE:
            try:
                from nltk.tokenize import word_tokenize
                words = word_tokenize(text.lower())
            except Exception as e:
                logger.debug(f"NLTK word tokenization failed: {e}")
        
        # Fallback to regex if NLTK failed or not available
        if not words:
            words = re.findall(r'\b\w+\b', text.lower())
        
        filtered = []
        for w in words:
            # Skip short words
            if len(w) < min_length:
                continue
            # Skip stopwords if requested
            if remove_stops and w in self.stop_words:
                continue
            filtered.append(w)
        
        return filtered
    
    def remove_stopwords(self, words):
        """Just filter out stopwords from a word list."""
        return [w for w in words if w not in self.stop_words]
    
    def get_basic_stats(self, text):
        """
        Quick stats about the text - useful for logging.
        """
        sentences = self.get_sentences(text)
        words = self.tokenize(text)
        
        return {
            'char_count': len(text),
            'word_count': len(words),
            'sentence_count': len(sentences),
            'unique_words': len(set(words)),
            'avg_word_len': sum(len(w) for w in words) / len(words) if words else 0
        }