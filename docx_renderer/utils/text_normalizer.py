"""
Text normalization utilities for DOCX parsing.

Handles namespace cleanup, whitespace normalization, and special character
decoding for extracted text content.
"""

import re
from typing import Optional, Union
from xml.etree.ElementTree import Element


class TextNormalizer:
    """Normalizes extracted text content from WordprocessingML."""
    
    # Common Word special characters that need normalization
    SPECIAL_CHARS = {
        '\u00a0': ' ',      # Non-breaking space → regular space
        '\u2009': ' ',      # Thin space → regular space  
        '\u2007': ' ',      # Figure space → regular space
        '\u2008': ' ',      # Punctuation space → regular space
        '\u200b': '',       # Zero-width space → remove
        '\u200c': '',       # Zero-width non-joiner → remove
        '\u200d': '',       # Zero-width joiner → remove
        '\ufeff': '',       # Byte order mark → remove
        '\u00ad': '',       # Soft hyphen → remove
        '\u2011': '-',      # Non-breaking hyphen → regular hyphen
        '\u2013': '–',      # En dash (keep as is)
        '\u2014': '—',      # Em dash (keep as is)
        '\u2018': "'",      # Left single quotation mark
        '\u2019': "'",      # Right single quotation mark  
        '\u201c': '"',      # Left double quotation mark
        '\u201d': '"',      # Right double quotation mark
        '\u2026': '...',    # Horizontal ellipsis
    }
    
    # Regex for collapsing multiple whitespace characters
    WHITESPACE_PATTERN = re.compile(r'\s+')
    
    # Regex for removing control characters (except tabs, newlines, carriage returns)
    CONTROL_CHARS_PATTERN = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]')
    
    def __init__(self, preserve_whitespace: bool = False):
        """Initialize text normalizer.
        
        Args:
            preserve_whitespace: If True, preserve exact whitespace formatting.
                                If False, normalize whitespace to single spaces.
        """
        self.preserve_whitespace = preserve_whitespace
    
    def normalize_text(self, text: str) -> str:
        """Normalize text content extracted from WordprocessingML."""
        if not text:
            return text
        
        # Replace special characters
        normalized = self._replace_special_chars(text)
        
        # Remove control characters
        normalized = self._remove_control_chars(normalized)
        
        # Normalize whitespace if not preserving it
        if not self.preserve_whitespace:
            normalized = self._normalize_whitespace(normalized)
        
        return normalized
    
    def normalize_element_text(self, element: Optional[Element]) -> str:
        """Extract and normalize text content from XML element."""
        if element is None:
            return ""
        
        # Get text content, handling both element.text and tail text
        text_parts = []
        
        if element.text:
            text_parts.append(element.text)
        
        # Recursively collect text from child elements
        for child in element:
            text_parts.append(self.normalize_element_text(child))
            if child.tail:
                text_parts.append(child.tail)
        
        raw_text = ''.join(text_parts)
        return self.normalize_text(raw_text)
    
    def extract_plain_text(self, element: Optional[Element]) -> str:
        """Extract plain text content, stripping all XML formatting."""
        if element is None:
            return ""
        
        # Use ElementTree's itertext() for efficient text extraction
        text_generator = element.itertext()
        raw_text = ''.join(text_generator)
        return self.normalize_text(raw_text)
    
    def _replace_special_chars(self, text: str) -> str:
        """Replace special Unicode characters with normalized equivalents."""
        for original, replacement in self.SPECIAL_CHARS.items():
            text = text.replace(original, replacement)
        return text
    
    def _remove_control_chars(self, text: str) -> str:
        """Remove control characters that shouldn't appear in document text."""
        return self.CONTROL_CHARS_PATTERN.sub('', text)
    
    def _normalize_whitespace(self, text: str) -> str:
        """Normalize whitespace to single spaces and trim."""
        # Collapse multiple whitespace characters to single space
        normalized = self.WHITESPACE_PATTERN.sub(' ', text)
        # Trim leading/trailing whitespace
        return normalized.strip()


class NamespaceStripper:
    """Removes namespace prefixes from XML text content."""
    
    # Common WordprocessingML namespaces to strip
    NAMESPACE_PREFIXES = [
        'w:',     # Word processing namespace
        'wp:',    # WordprocessingDrawing namespace
        'a:',     # DrawingML namespace
        'pic:',   # Picture namespace
        'r:',     # Relationships namespace
        'o:',     # Office namespace
        'v:',     # VML namespace
        'm:',     # Math namespace
        'mc:',    # Markup Compatibility namespace
        'w14:',   # Word 2010 namespace
        'w15:',   # Word 2013 namespace
        'w16:',   # Word 2016 namespace
    ]
    
    def strip_namespaces(self, text: str) -> str:
        """Remove namespace prefixes from element names in text."""
        if not text:
            return text
        
        stripped = text
        for prefix in self.NAMESPACE_PREFIXES:
            # Only replace at word boundaries to avoid partial matches
            import re
            pattern = r'\b' + re.escape(prefix)
            stripped = re.sub(pattern, '', stripped)
        
        return stripped
    
    def strip_element_namespaces(self, element: Optional[Element]) -> None:
        """Strip namespaces from element and all its children in-place."""
        if element is None:
            return
        
        # Strip namespace from tag
        if '}' in element.tag:
            element.tag = element.tag.split('}')[-1]
        
        # Strip namespaces from attributes
        new_attrib = {}
        for key, value in element.attrib.items():
            clean_key = key.split('}')[-1] if '}' in key else key
            new_attrib[clean_key] = value
        element.attrib.clear()
        element.attrib.update(new_attrib)
        
        # Recursively process children
        for child in element:
            self.strip_element_namespaces(child)


def normalize_docx_text(text: Union[str, Element, None], 
                       preserve_whitespace: bool = False,
                       strip_namespaces: bool = True) -> str:
    """Convenience function to normalize text from DOCX content.
    
    Args:
        text: Text string or XML element to normalize
        preserve_whitespace: Whether to preserve exact whitespace formatting
        strip_namespaces: Whether to strip XML namespace prefixes
        
    Returns:
        Normalized text string
    """
    normalizer = TextNormalizer(preserve_whitespace=preserve_whitespace)
    
    if isinstance(text, Element):
        result = normalizer.normalize_element_text(text)
    elif isinstance(text, str):
        result = normalizer.normalize_text(text)
    elif text is None:
        result = ""
    else:
        result = normalizer.normalize_text(str(text))
    
    if strip_namespaces:
        stripper = NamespaceStripper()
        result = stripper.strip_namespaces(result)
    
    return result