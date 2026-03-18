"""
Semantic Matching Module
Finds semantically similar content using embeddings.
"""

import numpy as np
from typing import List, Dict, Tuple, Optional
from sklearn.metrics.pairwise import cosine_similarity
import logging

try:
    from sentence_transformers import SentenceTransformer
    TRANSFORMERS_AVAILABLE = True
except ImportError:
    TRANSFORMERS_AVAILABLE = False
    logging.warning("sentence-transformers not available. Using TF-IDF fallback.")

logger = logging.getLogger(__name__)


class SemanticMatcher:
    """
    Finds semantically similar content using embeddings.
    
    This class uses sentence transformers for semantic similarity matching
    between keywords and text segments.
    
    Attributes:
        model_name (str): Name of the sentence transformer model
        similarity_threshold (float): Minimum similarity score
    """
    
    def __init__(self, model_name: str = 'all-MiniLM-L6-v2', 
                 similarity_threshold: float = 0.5):
        """
        Initialize the SemanticMatcher.
        
        Args:
            model_name: Name of sentence transformer model
            similarity_threshold: Minimum similarity score for matches
        """
        self.model_name = model_name
        self.similarity_threshold = similarity_threshold
        self.model = None
        
        if TRANSFORMERS_AVAILABLE:
            try:
                self.model = SentenceTransformer(model_name)
                logger.info(f"Loaded sentence transformer model: {model_name}")
            except Exception as e:
                logger.error(f"Failed to load model: {e}")
    
    def encode_texts(self, texts: List[str]) -> Optional[np.ndarray]:
        """
        Encode texts into embeddings.
        
        Args:
            texts: List of text strings
            
        Returns:
            Array of embeddings or None if encoding fails
        """
        if not texts or not self.model:
            return None
        
        try:
            return self.model.encode(texts, convert_to_numpy=True)
        except Exception as e:
            logger.error(f"Encoding failed: {e}")
            return None
    
    def find_similar(self, query: str, candidates: List[str], 
                     top_k: int = 5) -> List[Tuple[str, float]]:
        """
        Find most similar candidates to query.
        
        Args:
            query: Query text
            candidates: List of candidate texts
            top_k: Number of top matches to return
            
        Returns:
            List of (text, score) tuples
        """
        if not candidates or not self.model:
            return []
        
        try:
            # Encode query and candidates
            query_embedding = self.model.encode([query])
            candidate_embeddings = self.model.encode(candidates)
            
            # Calculate similarities
            similarities = cosine_similarity(query_embedding, candidate_embeddings)[0]
            
            # Get top k matches
            top_indices = np.argsort(similarities)[-top_k:][::-1]
            
            results = []
            for idx in top_indices:
                if similarities[idx] >= self.similarity_threshold:
                    results.append((candidates[idx], float(similarities[idx])))
            
            return results
            
        except Exception as e:
            logger.error(f"Similarity search failed: {e}")
            return []
    
    def find_similar_batch(self, queries: List[str], candidates: List[str],
                            top_k_per_query: int = 3) -> Dict[str, List[Tuple[str, float]]]:
        """
        Find similar candidates for multiple queries.
        
        Args:
            queries: List of query texts
            candidates: List of candidate texts
            top_k_per_query: Number of matches per query
            
        Returns:
            Dictionary mapping queries to results
        """
        if not queries or not candidates or not self.model:
            return {}
        
        results = {}
        
        for query in queries:
            matches = self.find_similar(query, candidates, top_k_per_query)
            results[query] = matches
        
        return results
    
    def semantic_search(self, text_segments: List[str], keywords: List[str],
                         threshold: Optional[float] = None) -> Dict[str, List[Tuple[str, float]]]:
        """
        Perform semantic search for keywords in text segments.
        
        Args:
            text_segments: List of text segments to search
            keywords: List of keywords to search for
            threshold: Similarity threshold (uses self.similarity_threshold if None)
            
        Returns:
            Dictionary mapping keywords to matching segments
        """
        if threshold is None:
            threshold = self.similarity_threshold
        
        results = {}
        
        # Encode all text segments once
        if self.model:
            segment_embeddings = self.encode_texts(text_segments)
            
            if segment_embeddings is not None:
                for keyword in keywords:
                    keyword_embedding = self.encode_texts([keyword])
                    
                    if keyword_embedding is not None:
                        similarities = cosine_similarity(keyword_embedding, segment_embeddings)[0]
                        
                        matches = []
                        for i, score in enumerate(similarities):
                            if score >= threshold:
                                matches.append((text_segments[i], float(score)))
                        
                        # Sort by score descending
                        matches.sort(key=lambda x: x[1], reverse=True)
                        results[keyword] = matches
        
        return results
    
    def cluster_by_semantics(self, texts: List[str], n_clusters: int = 3) -> Dict[int, List[str]]:
        """
        Cluster texts by semantic similarity.
        
        Args:
            texts: List of texts to cluster
            n_clusters: Number of clusters
            
        Returns:
            Dictionary mapping cluster IDs to text lists
        """
        if not texts or len(texts) < n_clusters or not self.model:
            return {0: texts}
        
        try:
            # Get embeddings
            embeddings = self.encode_texts(texts)
            
            if embeddings is None:
                return {0: texts}
            
            # Simple k-means clustering
            from sklearn.cluster import KMeans
            kmeans = KMeans(n_clusters=n_clusters, random_state=42)
            labels = kmeans.fit_predict(embeddings)
            
            # Organize results
            clusters = {}
            for i, label in enumerate(labels):
                if label not in clusters:
                    clusters[label] = []
                clusters[label].append(texts[i])
            
            return clusters
            
        except Exception as e:
            logger.error(f"Clustering failed: {e}")
            return {0: texts}