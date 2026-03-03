"""docx2llm — DOCX → token-optimized HTML5 for LLM ingestion."""
from .converter import convert, ConversionError

__all__ = ['convert', 'ConversionError']
__version__ = '1.0.0'
