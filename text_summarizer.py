"""
Text Summarization Module
Generates summaries of extracted text content.
"""
import re
import numpy as np
from typing import List, Dict, Optional
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import networkx as nx
import nltk
from nltk.tokenize import sent_tokenize
import logging

logger = logging.getLogger(__name__)


class TextSummarizer:
    """
    Generates summaries of text content.
    
    This class provides extractive and abstractive summarization methods
    for condensing large amounts of text into key insights.
    
    Attributes:
        compression_ratio (float): Target compression ratio
        max_sentences (int): Maximum sentences in summary
    """
    
    def __init__(self, compression_ratio: float = 0.3, max_sentences: int = 5):
        """
        Initialize the TextSummarizer.
        
        Args:
            compression_ratio: Target ratio of summary to original text
            max_sentences: Maximum number of sentences in summary
        """
        self.compression_ratio = compression_ratio
        self.max_sentences = max_sentences
        self.vectorizer = TfidfVectorizer(
            stop_words='english',
            max_features=1000
        )
    
    def extractive_summary(self, text: str, num_sentences: Optional[int] = None) -> str:
        """
        Generate extractive summary using TextRank algorithm.
        
        Args:
            text: Input text
            num_sentences: Number of sentences in summary (uses max_sentences if None)
            
        Returns:
            Summarized text
        """
        if not text:
            return ""
        
        # Tokenize sentences
        sentences = sent_tokenize(text)
        
        if len(sentences) <= 3:
            return text
        
        if num_sentences is None:
            num_sentences = min(self.max_sentences, len(sentences))
        
        try:
            # Create sentence vectors
            sentence_vectors = self.vectorizer.fit_transform(sentences).toarray()
            
            # Build similarity matrix
            similarity_matrix = np.zeros((len(sentences), len(sentences)))
            for i in range(len(sentences)):
                for j in range(len(sentences)):
                    if i != j:
                        similarity_matrix[i][j] = cosine_similarity(
                            sentence_vectors[i].reshape(1, -1),
                            sentence_vectors[j].reshape(1, -1)
                        )[0][0]
            
            # Apply PageRank
            nx_graph = nx.from_numpy_array(similarity_matrix)
            scores = nx.pagerank(nx_graph)
            
            # Rank sentences by score
            ranked_sentences = sorted(
                ((scores[i], sentence) for i, sentence in enumerate(sentences)),
                reverse=True
            )
            
            # Select top sentences and sort by original order
            top_sentences = [sentence for _, sentence in ranked_sentences[:num_sentences]]
            top_sentences.sort(key=lambda x: sentences.index(x))
            
            return ' '.join(top_sentences)
            
        except Exception as e:
            logger.error(f"Summarization failed: {e}")
            # Fallback: return first few sentences
            return ' '.join(sentences[:num_sentences])
    
    def keyword_focused_summary(self, text: str, keywords: List[str], 
                                 num_sentences: int = 5) -> str:
        """
        Generate summary focused on specific keywords.
        
        Args:
            text: Input text
            keywords: List of keywords to focus on
            num_sentences: Number of sentences in summary
            
        Returns:
            Keyword-focused summary
        """
        sentences = sent_tokenize(text)
        
        if len(sentences) <= num_sentences:
            return text
        
        # Score sentences based on keyword presence
        sentence_scores = []
        keywords_lower = [k.lower() for k in keywords]
        
        for sentence in sentences:
            sentence_lower = sentence.lower()
            score = sum(1 for keyword in keywords_lower if keyword in sentence_lower)
            
            # Boost score for sentences with multiple keywords
            if score > 0:
                score *= (1 + 0.1 * score)
            
            sentence_scores.append((score, sentence))
        
        # Sort by score and get top sentences
        top_sentences = sorted(sentence_scores, key=lambda x: x[0], reverse=True)[:num_sentences]
        top_sentences.sort(key=lambda x: sentences.index(x[1]))
        
        return ' '.join([s[1] for s in top_sentences])
    
    def bullet_point_summary(self, text: str, max_points: int = 5) -> List[str]:
        """
        Generate bullet point summary of key points.
        
        Args:
            text: Input text
            max_points: Maximum number of bullet points
            
        Returns:
            List of bullet point strings
        """
        # Get extractive summary first
        summary = self.extractive_summary(text, num_sentences=max_points * 2)
        
        # Split into sentences
        sentences = sent_tokenize(summary)
        
        # Convert to bullet points (simplify each sentence)
        bullet_points = []
        for sent in sentences[:max_points]:
            # Clean and shorten sentence
            sent = sent.strip()
            if len(sent) > 100:
                sent = sent[:97] + "..."
            bullet_points.append(sent)
        
        return bullet_points
    
    def hierarchical_summary(self, text: str, levels: int = 2) -> Dict[str, any]:
        """
        Generate hierarchical summary with different levels of detail.
        
        Args:
            text: Input text
            levels: Number of summary levels
            
        Returns:
            Dictionary with different summary levels
        """
        sentences = sent_tokenize(text)
        total_sentences = len(sentences)
        
        summaries = {}
        
        for level in range(1, levels + 1):
            num_sent = max(1, int(total_sentences * (0.5 ** level)))
            summaries[f'level_{level}'] = self.extractive_summary(text, num_sent)
        
        return summaries
    
    def extract_insights(self, text: str, num_insights: int = 3) -> List[str]:
        """
        Extract key insights from text.
        
        Args:
            text: Input text
            num_insights: Number of insights to extract
            
        Returns:
            List of insight statements
        """
        sentences = sent_tokenize(text)
        
        if len(sentences) <= num_insights:
            return sentences
        
        # Look for sentences with indicators of important information
        insight_indicators = [
            'important', 'significant', 'key', 'crucial', 'essential',
            'notably', 'particularly', 'especially', 'critical',
            'findings show', 'results indicate', 'conclude that',
            'in summary', 'overall', 'therefore', 'thus'
        ]
        
        scored_sentences = []
        for sentence in sentences:
            score = 0
            sentence_lower = sentence.lower()
            
            # Check for insight indicators
            for indicator in insight_indicators:
                if indicator in sentence_lower:
                    score += 2
            
            # Check for numerical findings
            if re.search(r'\d+%|\d+ percent', sentence_lower):
                score += 3
            
            # Check for sentence length (medium length sentences often contain insights)
            word_count = len(sentence.split())
            if 10 <= word_count <= 30:
                score += 1
            
            scored_sentences.append((score, sentence))
        
        # Get top insights
        top_insights = sorted(scored_sentences, key=lambda x: x[0], reverse=True)[:num_insights]
        
        return [insight for score, insight in top_insights if score > 0]