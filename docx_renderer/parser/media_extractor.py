"""
DOCX Media and Binary Asset Extractor

Extracts and catalogs media files (images, fonts, etc.) from DOCX package
with metadata and access paths for rendering pipeline.
"""

import base64
import mimetypes
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Union

from ..model.elements import MediaAsset
from .docx_loader import DocxPackage
from .rels_parser import Relationships


@dataclass
class MediaCatalog:
    """Catalog of all media assets found in DOCX package."""
    
    # Mapping from relationship ID to MediaAsset
    assets: Dict[str, MediaAsset]
    
    # Embedded fonts by family name
    fonts: Dict[str, MediaAsset]
    
    # Quick access by media type
    images: List[MediaAsset]
    audio: List[MediaAsset]
    video: List[MediaAsset]
    documents: List[MediaAsset]  # Embedded documents
    
    def __post_init__(self):
        """Populate convenience lists after initialization."""
        self.images = []
        self.audio = []
        self.video = []
        self.documents = []
        
        for asset in self.assets.values():
            if asset.media_type.startswith('image/'):
                self.images.append(asset)
            elif asset.media_type.startswith('audio/'):
                self.audio.append(asset)
            elif asset.media_type.startswith('video/'):
                self.video.append(asset)
            elif asset.media_type.startswith('application/'):
                self.documents.append(asset)
    
    def get_by_id(self, relationship_id: str) -> Optional[MediaAsset]:
        """Get media asset by relationship ID."""
        return self.assets.get(relationship_id)
    
    def get_by_target(self, target_path: str) -> Optional[MediaAsset]:
        """Get media asset by target path."""
        for asset in self.assets.values():
            if asset.target_path == target_path:
                return asset
        return None


class MediaExtractor:
    """Extracts media assets and fonts from DOCX package."""
    
    def __init__(self, package: DocxPackage, relationships: Relationships):
        self.package = package
        self.relationships = relationships
    
    def extract_media_catalog(self) -> MediaCatalog:
        """Extract complete media catalog from DOCX package."""
        assets = {}
        fonts = {}
        
        # Extract media from document relationships
        assets.update(self._extract_media_assets())
        
        # Extract embedded fonts
        fonts.update(self._extract_embedded_fonts())
        
        return MediaCatalog(
            assets=assets,
            fonts=fonts,
            images=[],  # Will be populated in __post_init__
            audio=[],
            video=[],
            documents=[]
        )
    
    def _extract_media_assets(self) -> Dict[str, MediaAsset]:
        """Extract media assets from relationships."""
        assets = {}
        
        # Get all media relationships (images, audio, video, etc.)
        media_rels = self.relationships.get_targets_by_type([
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio',
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/video',
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package'
        ])
        
        for rel_id, target in media_rels.items():
            try:
                # Get binary data from package
                binary_data = self.package.get_part_data(target)
                if binary_data is None:
                    continue
                
                # Determine media type
                media_type = self._get_media_type(target)
                
                # Extract metadata
                metadata = self._extract_media_metadata(target, binary_data)
                
                asset = MediaAsset(
                    relationship_id=rel_id,
                    target_path=target,
                    media_type=media_type,
                    binary_data=binary_data,
                    base64_data=base64.b64encode(binary_data).decode('utf-8'),
                    size=len(binary_data),
                    metadata=metadata
                )
                
                assets[rel_id] = asset
                
            except Exception as e:
                print(f"Warning: Failed to extract media asset {target}: {e}")
                continue
        
        return assets
    
    def _extract_embedded_fonts(self) -> Dict[str, MediaAsset]:
        """Extract embedded font assets."""
        fonts = {}
        
        # Look for font relationships
        font_rels = self.relationships.get_targets_by_type([
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/font'
        ])
        
        for rel_id, target in font_rels.items():
            try:
                binary_data = self.package.get_part_data(target)
                if binary_data is None:
                    continue
                
                # Try to extract font family name from filename or metadata
                font_family = self._extract_font_family(target, binary_data)
                
                asset = MediaAsset(
                    relationship_id=rel_id,
                    target_path=target,
                    media_type='application/font-woff',  # Default, may vary
                    binary_data=binary_data,
                    base64_data=base64.b64encode(binary_data).decode('utf-8'),
                    size=len(binary_data),
                    metadata={'font_family': font_family}
                )
                
                fonts[font_family] = asset
                
            except Exception as e:
                print(f"Warning: Failed to extract font {target}: {e}")
                continue
        
        return fonts
    
    def _get_media_type(self, target_path: str) -> str:
        """Determine MIME type from file extension."""
        # First check our fallback for common DOCX media types
        ext = Path(target_path).suffix.lower()
        fallback_types = {
            '.png': 'image/png',
            '.jpg': 'image/jpeg',
            '.jpeg': 'image/jpeg',
            '.gif': 'image/gif',
            '.bmp': 'image/bmp',
            '.tiff': 'image/tiff',
            '.svg': 'image/svg+xml',
            '.emf': 'image/x-emf',
            '.wmf': 'image/x-wmf',
            '.mp3': 'audio/mpeg',
            '.wav': 'audio/wav',
            '.mp4': 'video/mp4',
            '.avi': 'video/x-msvideo',
            '.pdf': 'application/pdf',
            '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        }
        
        if ext in fallback_types:
            return fallback_types[ext]
        
        # Use mimetypes module for other extensions
        mime_type, _ = mimetypes.guess_type(target_path)
        return mime_type or 'application/octet-stream'
    
    def _extract_media_metadata(self, target_path: str, binary_data: bytes) -> Dict[str, Union[str, int]]:
        """Extract metadata from media file."""
        metadata = {
            'filename': Path(target_path).name,
            'extension': Path(target_path).suffix.lower()
        }
        
        # Try to extract basic image metadata
        if target_path.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
            try:
                # Basic image size detection (this is simplified)
                # In a full implementation, you'd use PIL or similar
                metadata.update(self._get_image_dimensions(binary_data))
            except Exception:
                pass
        
        return metadata
    
    def _get_image_dimensions(self, binary_data: bytes) -> Dict[str, int]:
        """Extract image dimensions from binary data (simplified)."""
        # This is a simplified implementation
        # For production, use PIL: Image.open(BytesIO(binary_data)).size
        
        # PNG signature check
        if binary_data.startswith(b'\x89PNG\r\n\x1a\n'):
            try:
                # PNG IHDR chunk contains dimensions at bytes 16-24
                if len(binary_data) >= 24:
                    width = int.from_bytes(binary_data[16:20], 'big')
                    height = int.from_bytes(binary_data[20:24], 'big')
                    return {'width': width, 'height': height}
            except Exception:
                pass
        
        # JPEG marker check
        elif binary_data.startswith(b'\xff\xd8'):
            # JPEG parsing is more complex, return empty for now
            pass
        
        return {}
    
    def _extract_font_family(self, target_path: str, binary_data: bytes) -> str:
        """Extract font family name from font file."""
        # Simplified implementation - in practice you'd parse font metadata
        filename = Path(target_path).stem
        
        # Clean up common font filename patterns
        font_family = filename.replace('_', ' ').replace('-', ' ')
        font_family = font_family.replace('Regular', '').replace('Bold', '').replace('Italic', '')
        font_family = font_family.strip()
        
        return font_family or 'Unknown Font'


class MediaResolver:
    """Maps relationship identifiers to actual media payloads."""

    def __init__(self, relationships: Relationships, media_catalog: MediaCatalog) -> None:
        self._relationships = relationships
        self._media_catalog = media_catalog

    def resolve_image(self, part_name: str, r_id: str) -> Optional[bytes]:
        """Return binary data for an image referenced by a relationship id."""
        rel = self._relationships.find(part_name, r_id)
        if rel is None or rel.is_external:
            return None
        
        # Try to get from media catalog first
        asset = self._media_catalog.get_by_id(r_id)
        if asset:
            return asset.binary_data
        
        # Fallback to legacy lookup
        target = rel.resolved_target or rel.target
        asset = self._media_catalog.get_by_target(target)
        return asset.binary_data if asset else None


def extract_media_from_package(package: DocxPackage, relationships: Relationships) -> MediaCatalog:
    """Convenience function to extract media catalog from DOCX package."""
    extractor = MediaExtractor(package, relationships)
    return extractor.extract_media_catalog()
