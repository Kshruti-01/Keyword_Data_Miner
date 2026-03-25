"""
Keyword Data Miner Package
A system for extracting keywords and context from emails and documents.
"""

# Version information
__version__ = "1.0.0"
__author__ = "Shruti Kumari"
__description__ = "Email keyword extraction and mining system"

# Expose main classes for easier imports
from src.core.text_preprocessor import TextPreprocessor
from src.core.keyword_extractor import KeywordExtractor
from src.core.context_extractor import KeywordContextExtractor
from src.core.relevance_scorer import RelevanceScorer
from src.core.entity_extractor import EntityExtractor
from src.pipeline.data_miner import DataMiner

# Define what gets imported with "from src import *"
__all__ = [
    'TextPreprocessor',
    'KeywordExtractor', 
    'KeywordContextExtractor',
    'RelevanceScorer',
    'EntityExtractor',
    'DataMiner'
]

"""
Connectors module for email sources (Outlook, Exchange, etc.)
"""

from .outlook_connector import OutlookConnector

__all__ = ['OutlookConnector']
