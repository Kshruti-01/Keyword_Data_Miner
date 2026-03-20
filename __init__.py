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
](venv) C:\Users\850085869\OneDrive - Genpact\Desktop\Project\keyword_data_miner>python -m src.pipeline.data_miner
Traceback (most recent call last):
  File "<frozen runpy>", line 189, in _run_module_as_main
  File "<frozen runpy>", line 112, in _get_module_details
  File "C:\Users\850085869\OneDrive - Genpact\Desktop\Project\keyword_data_miner\src\__init__.py", line 17, in <module>
    from src.pipeline.data_miner import DataMiner
  File "C:\Users\850085869\OneDrive - Genpact\Desktop\Project\keyword_data_miner\src\pipeline\__init__.py", line 5, in <module>
    from .data_miner import DataMiner
ImportError: cannot import name 'DataMiner' from 'src.pipeline.data_miner' (C:\Users\850085869\OneDrive - Genpact\Desktop\Project\keyword_data_miner\src\pipeline\data_miner.py)
