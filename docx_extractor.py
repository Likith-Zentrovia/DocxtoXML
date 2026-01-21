#!/usr/bin/env python3
"""
DOCX Content Extractor

This module extracts content from DOCX files including:
- Text with formatting (bold, italic, headings)
- Images with metadata
- Tables with structure
- Lists (bulleted and numbered)
- Styles and hierarchy

The extraction preserves document structure for conversion to DocBook XML.
"""

from __future__ import annotations

import io
import os
import re
import hashlib
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any, Union
from zipfile import ZipFile

# python-docx for DOCX parsing
try:
    from docx import Document
    from docx.document import Document as DocumentType
    from docx.text.paragraph import Paragraph
    from docx.text.run import Run
    from docx.table import Table, _Cell
    from docx.oxml.ns import qn
    from docx.shared import Inches, Pt
    from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
    from docx.oxml import parse_xml
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False
    print("WARNING: python-docx not installed. Install with: pip install python-docx")

# PIL for image processing
try:
    from PIL import Image
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

# lxml for XML manipulation
from lxml import etree


# ============================================================================
# DATA CLASSES
# ============================================================================

@dataclass
class ExtractedImage:
    """Represents an extracted image from DOCX."""
    filename: str
    data: bytes
    content_type: str
    width: Optional[int] = None
    height: Optional[int] = None
    alt_text: str = ""
    caption: str = ""
    position_hint: str = ""  # paragraph/table context


@dataclass
class ExtractedTable:
    """Represents an extracted table from DOCX."""
    rows: List[List[str]]
    header_rows: int = 1
    caption: str = ""
    has_merged_cells: bool = False
    cell_spans: Dict[Tuple[int, int], Tuple[int, int]] = field(default_factory=dict)


@dataclass
class TextBlock:
    """Represents a block of text with formatting."""
    text: str
    style: str = "Normal"
    level: int = 0  # Heading level (0 = not a heading)
    is_bold: bool = False
    is_italic: bool = False
    is_underline: bool = False
    list_type: Optional[str] = None  # "bullet" or "number" or None
    list_level: int = 0
    alignment: str = "left"


@dataclass
class DocxContent:
    """Complete extracted content from a DOCX file."""
    title: str = ""
    authors: List[str] = field(default_factory=list)
    text_blocks: List[TextBlock] = field(default_factory=list)
    images: List[ExtractedImage] = field(default_factory=list)
    tables: List[ExtractedTable] = field(default_factory=list)
    metadata: Dict[str, Any] = field(default_factory=dict)
    
    # Structure tracking
    chapters: List[Dict[str, Any]] = field(default_factory=list)
    current_chapter: int = -1


# ============================================================================
# DOCX EXTRACTOR CLASS
# ============================================================================

class DocxExtractor:
    """
    Extracts content from DOCX files.
    
    This class handles:
    - Text extraction with style preservation
    - Image extraction with metadata
    - Table extraction with structure
    - Document metadata extraction
    """
    
    def __init__(
        self,
        extract_images: bool = True,
        extract_tables: bool = True,
        preserve_formatting: bool = True,
        min_image_size: int = 50
    ):
        if not HAS_DOCX:
            raise ImportError("python-docx is required. Install with: pip install python-docx")
        
        self.extract_images = extract_images
        self.extract_tables = extract_tables
        self.preserve_formatting = preserve_formatting
        self.min_image_size = min_image_size
        
        self._image_counter = 0
        self._table_counter = 0

    def extract(self, docx_path: Union[str, Path]) -> DocxContent:
        """
        Extract all content from a DOCX file.
        
        Args:
            docx_path: Path to the DOCX file
            
        Returns:
            DocxContent with all extracted content
        """
        docx_path = Path(docx_path)
        if not docx_path.exists():
            raise FileNotFoundError(f"DOCX file not found: {docx_path}")
        
        # Reset counters
        self._image_counter = 0
        self._table_counter = 0
        
        # Open document
        doc = Document(str(docx_path))
        
        content = DocxContent()
        
        # Extract metadata
        content.metadata = self._extract_metadata(doc)
        content.title = content.metadata.get("title", docx_path.stem)
        content.authors = content.metadata.get("authors", [])
        
        # Extract images from the DOCX package
        if self.extract_images:
            content.images = self._extract_images_from_package(docx_path)
        
        # Process document body
        self._process_document_body(doc, content)
        
        return content

    def _extract_metadata(self, doc: DocumentType) -> Dict[str, Any]:
        """Extract document metadata."""
        metadata = {}
        
        try:
            core_props = doc.core_properties
            
            if core_props.title:
                metadata["title"] = core_props.title
            if core_props.author:
                metadata["authors"] = [core_props.author]
            if core_props.subject:
                metadata["subject"] = core_props.subject
            if core_props.keywords:
                metadata["keywords"] = core_props.keywords
            if core_props.created:
                metadata["created"] = str(core_props.created)
            if core_props.modified:
                metadata["modified"] = str(core_props.modified)
            if core_props.last_modified_by:
                metadata["last_modified_by"] = core_props.last_modified_by
            if core_props.revision:
                metadata["revision"] = core_props.revision
        except Exception as e:
            print(f"Warning: Could not extract all metadata: {e}")
        
        return metadata

    def _extract_images_from_package(self, docx_path: Path) -> List[ExtractedImage]:
        """Extract all images from the DOCX package."""
        images = []
        
        try:
            with ZipFile(docx_path, 'r') as zf:
                # Find all image files in the package
                for name in zf.namelist():
                    if name.startswith('word/media/'):
                        # Get the image data
                        img_data = zf.read(name)
                        
                        # Determine content type
                        ext = Path(name).suffix.lower()
                        content_types = {
                            '.png': 'image/png',
                            '.jpg': 'image/jpeg',
                            '.jpeg': 'image/jpeg',
                            '.gif': 'image/gif',
                            '.bmp': 'image/bmp',
                            '.tiff': 'image/tiff',
                            '.tif': 'image/tiff',
                            '.emf': 'image/x-emf',
                            '.wmf': 'image/x-wmf',
                        }
                        content_type = content_types.get(ext, 'application/octet-stream')
                        
                        # Get image dimensions if possible
                        width, height = None, None
                        if HAS_PIL and ext in ['.png', '.jpg', '.jpeg', '.gif', '.bmp']:
                            try:
                                img = Image.open(io.BytesIO(img_data))
                                width, height = img.size
                                
                                # Skip tiny images
                                if width < self.min_image_size and height < self.min_image_size:
                                    continue
                            except Exception:
                                pass
                        
                        # Create a clean filename
                        self._image_counter += 1
                        clean_name = f"img_{self._image_counter:04d}{ext}"
                        
                        images.append(ExtractedImage(
                            filename=clean_name,
                            data=img_data,
                            content_type=content_type,
                            width=width,
                            height=height
                        ))
        except Exception as e:
            print(f"Warning: Error extracting images: {e}")
        
        return images

    def _process_document_body(self, doc: DocumentType, content: DocxContent):
        """Process the document body extracting text blocks and tables."""

        # Track current position for image placement
        image_positions = self._map_image_positions(doc)

        # Create a mapping from relationship IDs to image indices
        rel_to_image_idx = self._map_rel_ids_to_images(doc, content.images)

        # Process paragraphs directly from doc.paragraphs for reliability
        print(f"  - Processing {len(doc.paragraphs)} paragraphs...")

        for para_index, para in enumerate(doc.paragraphs):
            if para_index > 0 and para_index % 500 == 0:
                print(f"  - Processing paragraph {para_index}...")

            # Check for inline images in this paragraph
            if self.extract_images and para_index in image_positions:
                rel_ids = image_positions[para_index]
                for rel_id in rel_ids:
                    if rel_id in rel_to_image_idx:
                        img_idx = rel_to_image_idx[rel_id]
                        # Add image marker
                        content.text_blocks.append(TextBlock(
                            text=f"[[IMAGE_{img_idx + 1}]]",
                            style="ImageMarker"
                        ))

            block = self._extract_paragraph(para, image_positions)
            if block and (block.text.strip() or block.list_type):
                content.text_blocks.append(block)

                # Check for chapter/section boundaries
                if block.level >= 1:
                    self._track_chapter(content, block)

        # Process tables separately
        if self.extract_tables:
            print(f"  - Processing {len(doc.tables)} tables...")

            # Build a set of table element IDs for position tracking
            table_positions = {}
            body = doc.element.body
            table_idx = 0
            para_count = 0

            for element in body:
                tag = element.tag.split('}')[-1] if '}' in element.tag else element.tag
                if tag == 'p':
                    para_count += 1
                elif tag == 'tbl':
                    table_positions[table_idx] = para_count
                    table_idx += 1

            # Process each table
            for idx, table in enumerate(doc.tables):
                extracted_table = self._extract_table(table)
                if extracted_table and extracted_table.rows:
                    content.tables.append(extracted_table)
                    self._table_counter += 1

                    # Add a marker in text blocks at the appropriate position
                    content.text_blocks.append(TextBlock(
                        text=f"[[TABLE_{self._table_counter}]]",
                        style="TableMarker"
                    ))

    def _map_rel_ids_to_images(self, doc: DocumentType, images: List[ExtractedImage]) -> Dict[str, int]:
        """Map relationship IDs to image indices in the images list."""
        rel_to_idx = {}

        try:
            # Get the document's relationship part
            rels = doc.part.rels

            # Build a mapping from image filename (in media folder) to index
            image_filenames = {}
            for idx, img in enumerate(images):
                # The filename in images is like "img_0001.png", we need to match against media filenames
                image_filenames[img.filename] = idx

            # Map relationship IDs to indices
            for rel_id, rel in rels.items():
                if hasattr(rel, 'target_part') and rel.target_part:
                    target = rel.target_part
                    if hasattr(target, 'partname') and '/media/' in str(target.partname):
                        # Extract the media filename
                        media_name = str(target.partname).split('/')[-1]
                        # Find matching image by extension and order
                        ext = Path(media_name).suffix.lower()
                        for fname, idx in image_filenames.items():
                            if fname.endswith(ext) and idx not in rel_to_idx.values():
                                rel_to_idx[rel_id] = idx
                                break
        except Exception as e:
            print(f"Warning: Could not map relationship IDs to images: {e}")

        return rel_to_idx

    def _get_paragraph_from_element(self, doc: DocumentType, element) -> Optional[Paragraph]:
        """Get Paragraph object from XML element."""
        try:
            for para in doc.paragraphs:
                if para._element is element:
                    return para
        except Exception:
            pass
        
        # Create a new Paragraph wrapper
        try:
            return Paragraph(element, doc)
        except Exception:
            return None

    def _get_table_from_element(self, doc: DocumentType, element) -> Optional[Table]:
        """Get Table object from XML element."""
        try:
            for table in doc.tables:
                if table._element is element:
                    return table
        except Exception:
            pass
        
        # Create a new Table wrapper
        try:
            return Table(element, doc)
        except Exception:
            return None

    def _map_image_positions(self, doc: DocumentType) -> Dict[int, List[str]]:
        """Map paragraph indices to image references."""
        positions = {}
        
        try:
            for i, para in enumerate(doc.paragraphs):
                # Check for inline images
                inline_shapes = para._element.findall('.//a:blip', 
                    namespaces={'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'})
                
                if inline_shapes:
                    positions[i] = [shape.get(qn('r:embed')) for shape in inline_shapes]
        except Exception:
            pass
        
        return positions

    def _extract_paragraph(self, para: Paragraph, image_positions: Dict) -> Optional[TextBlock]:
        """Extract a paragraph with its formatting."""

        # Get paragraph style
        style_name = para.style.name if para.style else "Normal"

        # Determine heading level
        level = 0
        if style_name.startswith("Heading"):
            try:
                level = int(style_name.replace("Heading", "").strip())
            except ValueError:
                if "1" in style_name:
                    level = 1
                elif "2" in style_name:
                    level = 2
                elif "3" in style_name:
                    level = 3
        elif style_name == "Title":
            level = 1
        elif style_name == "Subtitle":
            level = 2

        # Check for list
        list_type = None
        list_level = 0

        numPr = para._element.find('.//w:numPr',
            namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})

        if numPr is not None:
            # This is a list item
            ilvl = numPr.find('w:ilvl',
                namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
            if ilvl is not None:
                list_level = int(ilvl.get(qn('w:val'), '0'))

            # Try to determine if bullet or numbered
            # Default to bullet, but check the numId for numbered lists
            list_type = "bullet"

            numId = numPr.find('w:numId',
                namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
            if numId is not None:
                # Heuristic: even numIds tend to be numbered lists in many docs
                try:
                    num_val = int(numId.get(qn('w:val'), '0'))
                    if num_val % 2 == 0 or "List Number" in style_name or "Numbered" in style_name:
                        list_type = "number"
                except ValueError:
                    pass

        # Check style for list hints
        if "List Bullet" in style_name or "Bullet" in style_name:
            list_type = "bullet"
        elif "List Number" in style_name or "Numbered" in style_name:
            list_type = "number"

        # Extract text - use paragraph.text as primary source (more reliable)
        text = para.text or ""

        # Also try to get formatting from runs
        is_bold = False
        is_italic = False
        is_underline = False

        # If we want to preserve formatting, extract from runs
        if self.preserve_formatting and para.runs:
            text_parts = []
            for run in para.runs:
                run_text = run.text
                if not run_text:
                    continue

                # Track formatting
                if run.bold:
                    is_bold = True
                    run_text = f"**{run_text}**"

                if run.italic:
                    is_italic = True
                    run_text = f"*{run_text}*"

                if run.underline:
                    is_underline = True

                text_parts.append(run_text)

            # Use runs text if available, otherwise fallback to para.text
            if text_parts:
                text = "".join(text_parts)

                # Clean up markdown artifacts
                text = re.sub(r'\*\*\*\*+', '', text)
                text = re.sub(r'\*\*\s*\*\*', '', text)
                text = re.sub(r'\*\s*\*', '', text)
        else:
            # Check if runs have any formatting
            for run in para.runs:
                if run.bold:
                    is_bold = True
                if run.italic:
                    is_italic = True
                if run.underline:
                    is_underline = True
        
        # Get alignment
        alignment = "left"
        if para.alignment:
            if para.alignment == WD_PARAGRAPH_ALIGNMENT.CENTER:
                alignment = "center"
            elif para.alignment == WD_PARAGRAPH_ALIGNMENT.RIGHT:
                alignment = "right"
            elif para.alignment == WD_PARAGRAPH_ALIGNMENT.JUSTIFY:
                alignment = "justify"
        
        return TextBlock(
            text=text,
            style=style_name,
            level=level,
            is_bold=is_bold,
            is_italic=is_italic,
            is_underline=is_underline,
            list_type=list_type,
            list_level=list_level,
            alignment=alignment
        )

    def _extract_table(self, table: Table) -> Optional[ExtractedTable]:
        """Extract a table with its structure."""
        rows = []
        cell_spans = {}
        has_merged = False
        
        try:
            for row_idx, row in enumerate(table.rows):
                row_cells = []
                for col_idx, cell in enumerate(row.cells):
                    # Get cell text
                    cell_text = cell.text.strip()
                    row_cells.append(cell_text)
                    
                    # Check for merged cells
                    tc = cell._tc
                    gridSpan = tc.find('.//w:gridSpan',
                        namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                    
                    if gridSpan is not None:
                        span_val = int(gridSpan.get(qn('w:val'), '1'))
                        if span_val > 1:
                            has_merged = True
                            cell_spans[(row_idx, col_idx)] = (1, span_val)
                
                if row_cells:
                    rows.append(row_cells)
        except Exception as e:
            print(f"Warning: Error extracting table: {e}")
            return None
        
        if not rows:
            return None
        
        return ExtractedTable(
            rows=rows,
            header_rows=1,  # Assume first row is header
            has_merged_cells=has_merged,
            cell_spans=cell_spans
        )

    def _track_chapter(self, content: DocxContent, block: TextBlock):
        """Track chapter/section structure."""
        if block.level == 1:
            # New chapter
            content.chapters.append({
                "title": block.text,
                "level": 1,
                "sections": []
            })
            content.current_chapter = len(content.chapters) - 1
        elif block.level >= 2 and content.current_chapter >= 0:
            # Section within chapter
            content.chapters[content.current_chapter]["sections"].append({
                "title": block.text,
                "level": block.level
            })


# ============================================================================
# CONVENIENCE FUNCTIONS
# ============================================================================

def extract_docx(docx_path: Union[str, Path], **kwargs) -> DocxContent:
    """
    Convenience function to extract content from a DOCX file.
    
    Args:
        docx_path: Path to the DOCX file
        **kwargs: Additional arguments passed to DocxExtractor
        
    Returns:
        DocxContent with all extracted content
    """
    extractor = DocxExtractor(**kwargs)
    return extractor.extract(docx_path)


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python docx_extractor.py <docx_file>")
        sys.exit(1)
    
    docx_path = sys.argv[1]
    content = extract_docx(docx_path)
    
    print(f"Title: {content.title}")
    print(f"Authors: {content.authors}")
    print(f"Text blocks: {len(content.text_blocks)}")
    print(f"Images: {len(content.images)}")
    print(f"Tables: {len(content.tables)}")
    print(f"Chapters: {len(content.chapters)}")
    
    # Print first few text blocks
    print("\nFirst 5 text blocks:")
    for i, block in enumerate(content.text_blocks[:5]):
        prefix = "#" * block.level if block.level > 0 else ""
        list_prefix = "- " if block.list_type == "bullet" else ("1. " if block.list_type == "number" else "")
        print(f"  {i+1}. [{block.style}] {prefix} {list_prefix}{block.text[:100]}...")
