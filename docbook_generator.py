#!/usr/bin/env python3
"""
DocBook XML 4.2 Generator

This module converts extracted DOCX content to DocBook XML 4.2 format
compliant with the RittDoc DTD specification.

The output format matches the PDF to XML pipeline for consistency.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Union, Any
from datetime import datetime

from lxml import etree

from docx_extractor import DocxContent, TextBlock, ExtractedTable, ExtractedImage


# ============================================================================
# CONSTANTS
# ============================================================================

DOCTYPE_PUBLIC = "-//OASIS//DTD DocBook XML V4.2//EN"
DOCTYPE_SYSTEM = "http://www.oasis-open.org/docbook/xml/4.2/docbookx.dtd"

RITTDOC_DOCTYPE_PUBLIC = "-//RIS Dev//DTD DocBook V4.3 -Based Variant V1.1//EN"
RITTDOC_DOCTYPE_SYSTEM = "http://LOCALHOST/dtd/V1.1/RittDocBook.dtd"


# ============================================================================
# DOCBOOK GENERATOR CLASS
# ============================================================================

class DocBookGenerator:
    """
    Generates DocBook XML 4.2 from extracted DOCX content.
    
    The output follows the RittDoc DTD specification for compatibility
    with the PDF to XML pipeline.
    """
    
    def __init__(
        self,
        use_rittdoc_dtd: bool = True,
        include_bookinfo: bool = True,
        multimedia_prefix: str = "multimedia/"
    ):
        """
        Initialize the DocBook generator.
        
        Args:
            use_rittdoc_dtd: Use RittDoc DTD declaration instead of standard DocBook
            include_bookinfo: Include bookinfo section with metadata
            multimedia_prefix: Prefix for image file references
        """
        self.use_rittdoc_dtd = use_rittdoc_dtd
        self.include_bookinfo = include_bookinfo
        self.multimedia_prefix = multimedia_prefix
        
        self._table_counter = 0
        self._figure_counter = 0
        self._chapter_counter = 0
        self._section_counters = {}

    def generate(
        self,
        content: DocxContent,
        output_path: Optional[Union[str, Path]] = None
    ) -> str:
        """
        Generate DocBook XML from extracted content.
        
        Args:
            content: Extracted DOCX content
            output_path: Optional path to write the XML file
            
        Returns:
            XML string
        """
        # Reset counters
        self._table_counter = 0
        self._figure_counter = 0
        self._chapter_counter = 0
        self._section_counters = {}
        
        # Create root element
        root = etree.Element("book")
        
        # Add bookinfo if enabled
        if self.include_bookinfo:
            bookinfo = self._create_bookinfo(content)
            root.append(bookinfo)
        
        # Process content into chapters
        self._process_content(root, content)
        
        # Generate XML string with proper declaration and DOCTYPE
        xml_str = self._serialize_xml(root)
        
        # Write to file if path provided
        if output_path:
            output_path = Path(output_path)
            output_path.parent.mkdir(parents=True, exist_ok=True)
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(xml_str)
        
        return xml_str

    def _create_bookinfo(self, content: DocxContent) -> etree._Element:
        """Create the bookinfo section with metadata."""
        bookinfo = etree.Element("bookinfo")
        
        # ISBN (use placeholder if not available)
        isbn = etree.SubElement(bookinfo, "isbn")
        isbn.text = content.metadata.get("isbn", "0000000000000")
        
        # Title
        title = etree.SubElement(bookinfo, "title")
        title.text = content.title or "Untitled Document"
        
        # Subtitle if available
        if content.metadata.get("subtitle"):
            subtitle = etree.SubElement(bookinfo, "subtitle")
            subtitle.text = content.metadata["subtitle"]
        
        # Authors
        authorgroup = etree.SubElement(bookinfo, "authorgroup")
        authors = content.authors or ["Unknown Author"]
        
        for author_name in authors:
            author = etree.SubElement(authorgroup, "author")
            personname = etree.SubElement(author, "personname")
            
            # Split name into firstname/surname
            parts = author_name.split()
            if len(parts) >= 2:
                firstname = etree.SubElement(personname, "firstname")
                firstname.text = " ".join(parts[:-1])
                surname = etree.SubElement(personname, "surname")
                surname.text = parts[-1]
            else:
                surname = etree.SubElement(personname, "surname")
                surname.text = author_name
        
        # Publisher
        publisher = etree.SubElement(bookinfo, "publisher")
        publishername = etree.SubElement(publisher, "publishername")
        publishername.text = content.metadata.get("publisher", "Unknown Publisher")
        
        # Publication date
        pubdate = etree.SubElement(bookinfo, "pubdate")
        pubdate.text = content.metadata.get("pubdate", str(datetime.now().year))
        
        # Edition
        edition = etree.SubElement(bookinfo, "edition")
        edition.text = content.metadata.get("edition", "1st Edition")
        
        # Copyright
        copyright_elem = etree.SubElement(bookinfo, "copyright")
        year = etree.SubElement(copyright_elem, "year")
        year.text = content.metadata.get("copyright_year", str(datetime.now().year))
        holder = etree.SubElement(copyright_elem, "holder")
        holder.text = content.metadata.get("copyright_holder", content.metadata.get("publisher", "Copyright Holder"))
        
        return bookinfo

    def _process_content(self, root: etree._Element, content: DocxContent):
        """Process text blocks into DocBook structure."""

        # Track current containers
        current_chapter = None
        current_sect1 = None
        current_sect2 = None
        current_sect3 = None

        # Track list state
        current_list = None
        current_list_type = None

        # Table map for placeholder replacement
        table_map = {i+1: table for i, table in enumerate(content.tables)}

        # Image map for placeholder replacement
        image_map = {i+1: img for i, img in enumerate(content.images)}

        # Track which images have been placed inline
        placed_images = set()

        for block in content.text_blocks:
            # Check for image placeholder
            if block.style == "ImageMarker" and block.text.startswith("[[IMAGE_"):
                # Extract image number and insert figure
                match = re.match(r'\[\[IMAGE_(\d+)\]\]', block.text)
                if match:
                    img_num = int(match.group(1))
                    if img_num in image_map:
                        # Close any open list
                        current_list = None
                        current_list_type = None

                        # Determine parent element
                        parent = current_sect3 or current_sect2 or current_sect1 or current_chapter or root

                        # Create a default chapter if none exists
                        if parent is None or parent is root:
                            self._chapter_counter += 1
                            current_chapter = etree.SubElement(root, "chapter")
                            current_chapter.set("id", f"ch{self._chapter_counter:04d}")
                            title = etree.SubElement(current_chapter, "title")
                            title.text = content.title or "Content"
                            parent = current_chapter

                        figure_elem = self._create_figure_element(image_map[img_num], img_num)
                        parent.append(figure_elem)
                        placed_images.add(img_num)
                continue

            # Check for table placeholder
            if block.style == "TableMarker" and block.text.startswith("[[TABLE_"):
                # Extract table number and insert table
                match = re.match(r'\[\[TABLE_(\d+)\]\]', block.text)
                if match:
                    table_num = int(match.group(1))
                    if table_num in table_map:
                        # Close any open list
                        current_list = None
                        current_list_type = None

                        # Determine parent element
                        parent = current_sect3 or current_sect2 or current_sect1 or current_chapter or root
                        table_elem = self._create_table_element(table_map[table_num])
                        parent.append(table_elem)
                continue
            
            # Handle headings (create structure)
            if block.level == 1:
                # New chapter
                current_list = None
                current_list_type = None
                current_sect1 = None
                current_sect2 = None
                current_sect3 = None
                
                self._chapter_counter += 1
                current_chapter = etree.SubElement(root, "chapter")
                current_chapter.set("id", f"ch{self._chapter_counter:04d}")
                
                title = etree.SubElement(current_chapter, "title")
                title.text = self._clean_text(block.text)
                continue
            
            elif block.level == 2:
                # Section 1
                current_list = None
                current_list_type = None
                current_sect2 = None
                current_sect3 = None
                
                # Ensure we have a chapter
                if current_chapter is None:
                    self._chapter_counter += 1
                    current_chapter = etree.SubElement(root, "chapter")
                    current_chapter.set("id", f"ch{self._chapter_counter:04d}")
                    title = etree.SubElement(current_chapter, "title")
                    title.text = "Content"
                
                sect_id = f"ch{self._chapter_counter:04d}_s1_{self._get_section_counter(1)}"
                current_sect1 = etree.SubElement(current_chapter, "sect1")
                current_sect1.set("id", sect_id)
                
                title = etree.SubElement(current_sect1, "title")
                title.text = self._clean_text(block.text)
                continue
            
            elif block.level == 3:
                # Section 2
                current_list = None
                current_list_type = None
                current_sect3 = None
                
                # Ensure we have a sect1
                if current_sect1 is None:
                    if current_chapter is None:
                        self._chapter_counter += 1
                        current_chapter = etree.SubElement(root, "chapter")
                        current_chapter.set("id", f"ch{self._chapter_counter:04d}")
                        title = etree.SubElement(current_chapter, "title")
                        title.text = "Content"
                    
                    sect_id = f"ch{self._chapter_counter:04d}_s1_{self._get_section_counter(1)}"
                    current_sect1 = etree.SubElement(current_chapter, "sect1")
                    current_sect1.set("id", sect_id)
                    title = etree.SubElement(current_sect1, "title")
                    title.text = "Section"
                
                sect_id = f"ch{self._chapter_counter:04d}_s2_{self._get_section_counter(2)}"
                current_sect2 = etree.SubElement(current_sect1, "sect2")
                current_sect2.set("id", sect_id)
                
                title = etree.SubElement(current_sect2, "title")
                title.text = self._clean_text(block.text)
                continue
            
            elif block.level >= 4:
                # Section 3
                current_list = None
                current_list_type = None
                
                # Ensure we have a sect2
                if current_sect2 is None:
                    if current_sect1 is None:
                        if current_chapter is None:
                            self._chapter_counter += 1
                            current_chapter = etree.SubElement(root, "chapter")
                            current_chapter.set("id", f"ch{self._chapter_counter:04d}")
                            title = etree.SubElement(current_chapter, "title")
                            title.text = "Content"
                        
                        sect_id = f"ch{self._chapter_counter:04d}_s1_{self._get_section_counter(1)}"
                        current_sect1 = etree.SubElement(current_chapter, "sect1")
                        current_sect1.set("id", sect_id)
                        title = etree.SubElement(current_sect1, "title")
                        title.text = "Section"
                    
                    sect_id = f"ch{self._chapter_counter:04d}_s2_{self._get_section_counter(2)}"
                    current_sect2 = etree.SubElement(current_sect1, "sect2")
                    current_sect2.set("id", sect_id)
                    title = etree.SubElement(current_sect2, "title")
                    title.text = "Subsection"
                
                sect_id = f"ch{self._chapter_counter:04d}_s3_{self._get_section_counter(3)}"
                current_sect3 = etree.SubElement(current_sect2, "sect3")
                current_sect3.set("id", sect_id)
                
                title = etree.SubElement(current_sect3, "title")
                title.text = self._clean_text(block.text)
                continue
            
            # Determine parent element for content
            if current_sect3 is not None:
                parent = current_sect3
            elif current_sect2 is not None:
                parent = current_sect2
            elif current_sect1 is not None:
                parent = current_sect1
            else:
                parent = current_chapter
            
            # Create a default chapter if none exists
            if parent is None:
                self._chapter_counter += 1
                current_chapter = etree.SubElement(root, "chapter")
                current_chapter.set("id", f"ch{self._chapter_counter:04d}")
                title = etree.SubElement(current_chapter, "title")
                title.text = content.title or "Content"
                parent = current_chapter
            
            # Handle list items
            if block.list_type:
                list_tag = "itemizedlist" if block.list_type == "bullet" else "orderedlist"
                
                # Start new list if needed
                if current_list is None or current_list_type != block.list_type:
                    current_list = etree.SubElement(parent, list_tag)
                    current_list_type = block.list_type
                
                # Add list item
                listitem = etree.SubElement(current_list, "listitem")
                para = etree.SubElement(listitem, "para")
                self._set_para_content(para, block.text)
            
            else:
                # Close any open list
                current_list = None
                current_list_type = None
                
                # Regular paragraph
                if block.text.strip():
                    para = etree.SubElement(parent, "para")
                    self._set_para_content(para, block.text)
        
        # Add remaining images that weren't placed inline
        if content.images:
            self._add_images_section(root, content.images, current_chapter, placed_images)

    def _create_table_element(self, table: ExtractedTable) -> etree._Element:
        """Create a DocBook table element."""
        self._table_counter += 1
        
        # Create table wrapper
        informal_table = etree.Element("informaltable")
        
        # Create tgroup
        num_cols = max(len(row) for row in table.rows) if table.rows else 1
        tgroup = etree.SubElement(informal_table, "tgroup")
        tgroup.set("cols", str(num_cols))
        
        # Add colspecs
        for i in range(num_cols):
            colspec = etree.SubElement(tgroup, "colspec")
            colspec.set("colname", f"c{i+1}")
        
        # Split into header and body
        header_rows = table.rows[:table.header_rows] if table.header_rows > 0 else []
        body_rows = table.rows[table.header_rows:] if table.header_rows > 0 else table.rows
        
        # Add header if present
        if header_rows:
            thead = etree.SubElement(tgroup, "thead")
            for row in header_rows:
                tr = etree.SubElement(thead, "row")
                for cell_text in row:
                    entry = etree.SubElement(tr, "entry")
                    entry.text = self._clean_text(cell_text)
        
        # Add body
        if body_rows:
            tbody = etree.SubElement(tgroup, "tbody")
            for row in body_rows:
                tr = etree.SubElement(tbody, "row")
                for cell_text in row:
                    entry = etree.SubElement(tr, "entry")
                    entry.text = self._clean_text(cell_text)
        
        return informal_table

    def _create_figure_element(self, img: ExtractedImage, img_num: int) -> etree._Element:
        """Create a DocBook figure element for an image."""
        self._figure_counter += 1

        # Create figure wrapper
        figure = etree.Element("figure")
        figure.set("id", f"fig_{self._figure_counter:04d}")

        # Figure title (required by DTD)
        fig_title = etree.SubElement(figure, "title")
        fig_title.text = img.caption or f"Figure {self._figure_counter}"

        # Mediaobject containing the image
        mediaobject = etree.SubElement(figure, "mediaobject")
        imageobject = etree.SubElement(mediaobject, "imageobject")
        imagedata = etree.SubElement(imageobject, "imagedata")
        imagedata.set("fileref", f"{self.multimedia_prefix}{img.filename}")

        if img.width:
            imagedata.set("width", f"{img.width}px")
        if img.height:
            imagedata.set("depth", f"{img.height}px")

        # Text object for accessibility
        if img.alt_text:
            textobject = etree.SubElement(mediaobject, "textobject")
            phrase = etree.SubElement(textobject, "phrase")
            phrase.text = img.alt_text

        return figure

    def _add_images_section(
        self,
        root: etree._Element,
        images: List[ExtractedImage],
        current_chapter: Optional[etree._Element],
        placed_images: Optional[set] = None
    ):
        """Add a section with remaining images that weren't placed inline."""
        # Filter out already placed images
        remaining_images = []
        for i, img in enumerate(images):
            if placed_images is None or (i + 1) not in placed_images:
                remaining_images.append(img)

        if not remaining_images:
            return

        # Use current chapter or create new one
        parent = current_chapter
        if parent is None:
            self._chapter_counter += 1
            parent = etree.SubElement(root, "chapter")
            parent.set("id", f"ch{self._chapter_counter:04d}")
            title = etree.SubElement(parent, "title")
            title.text = "Figures"

        # Create section for figures
        sect1 = etree.SubElement(parent, "sect1")
        sect1.set("id", "figures")
        title = etree.SubElement(sect1, "title")
        title.text = "Figures"

        for img in remaining_images:
            self._figure_counter += 1
            
            # Create figure
            figure = etree.SubElement(sect1, "figure")
            figure.set("id", f"fig_{self._figure_counter:04d}")
            
            # Figure title
            fig_title = etree.SubElement(figure, "title")
            fig_title.text = img.caption or f"Figure {self._figure_counter}"
            
            # Mediaobject
            mediaobject = etree.SubElement(figure, "mediaobject")
            imageobject = etree.SubElement(mediaobject, "imageobject")
            imagedata = etree.SubElement(imageobject, "imagedata")
            imagedata.set("fileref", f"{self.multimedia_prefix}{img.filename}")
            
            if img.width:
                imagedata.set("width", f"{img.width}px")
            
            # Text object for accessibility
            if img.alt_text:
                textobject = etree.SubElement(mediaobject, "textobject")
                phrase = etree.SubElement(textobject, "phrase")
                phrase.text = img.alt_text

    def _set_para_content(self, para: etree._Element, text: str):
        """Set paragraph content with inline formatting."""
        text = self._clean_text(text)
        
        # Convert markdown-style formatting to DocBook emphasis
        # Bold: **text** -> <emphasis role="bold">text</emphasis>
        # Italic: *text* -> <emphasis>text</emphasis>
        
        # Simple approach: just set text and handle emphasis if present
        if "**" in text or "*" in text:
            # Parse and convert inline formatting
            parts = self._parse_inline_formatting(text)
            
            if len(parts) == 1 and parts[0][1] is None:
                # No formatting, just text
                para.text = parts[0][0]
            else:
                # Has formatting
                para.text = ""
                last_elem = None
                
                for part_text, fmt in parts:
                    if fmt is None:
                        if last_elem is not None:
                            if last_elem.tail:
                                last_elem.tail += part_text
                            else:
                                last_elem.tail = part_text
                        else:
                            if para.text:
                                para.text += part_text
                            else:
                                para.text = part_text
                    else:
                        emphasis = etree.SubElement(para, "emphasis")
                        if fmt == "bold":
                            emphasis.set("role", "bold")
                        emphasis.text = part_text
                        last_elem = emphasis
        else:
            para.text = text

    def _parse_inline_formatting(self, text: str) -> List[tuple]:
        """Parse text with inline formatting markers."""
        parts = []
        current_pos = 0
        
        # Find all formatting markers
        pattern = r'(\*\*\*(.+?)\*\*\*|\*\*(.+?)\*\*|\*(.+?)\*)'
        
        for match in re.finditer(pattern, text):
            # Add text before match
            if match.start() > current_pos:
                parts.append((text[current_pos:match.start()], None))
            
            # Determine formatting type
            if match.group(2):  # ***bold-italic***
                parts.append((match.group(2), "bold"))
            elif match.group(3):  # **bold**
                parts.append((match.group(3), "bold"))
            elif match.group(4):  # *italic*
                parts.append((match.group(4), "italic"))
            
            current_pos = match.end()
        
        # Add remaining text
        if current_pos < len(text):
            parts.append((text[current_pos:], None))
        
        # If no formatting found, return the whole text
        if not parts:
            parts.append((text, None))
        
        return parts

    def _clean_text(self, text: str) -> str:
        """Clean text for XML output."""
        if not text:
            return ""
        
        # Remove control characters
        text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', text)
        
        # Normalize whitespace
        text = re.sub(r'\s+', ' ', text)
        
        return text.strip()

    def _get_section_counter(self, level: int) -> int:
        """Get and increment section counter for a level."""
        if level not in self._section_counters:
            self._section_counters[level] = 0
        self._section_counters[level] += 1
        return self._section_counters[level]

    def _serialize_xml(self, root: etree._Element) -> str:
        """Serialize XML with proper declaration and DOCTYPE."""
        # Choose DOCTYPE
        if self.use_rittdoc_dtd:
            doctype_public = RITTDOC_DOCTYPE_PUBLIC
            doctype_system = RITTDOC_DOCTYPE_SYSTEM
        else:
            doctype_public = DOCTYPE_PUBLIC
            doctype_system = DOCTYPE_SYSTEM
        
        # Create XML declaration
        xml_declaration = '<?xml version="1.0" encoding="UTF-8"?>\n'
        
        # Create DOCTYPE declaration
        doctype = f'<!DOCTYPE book PUBLIC "{doctype_public}"\n  "{doctype_system}">\n'
        
        # Serialize the tree
        xml_content = etree.tostring(
            root,
            encoding="unicode",
            pretty_print=True
        )
        
        return xml_declaration + doctype + xml_content


# ============================================================================
# CONVENIENCE FUNCTIONS
# ============================================================================

def generate_docbook(
    content: DocxContent,
    output_path: Optional[Union[str, Path]] = None,
    **kwargs
) -> str:
    """
    Convenience function to generate DocBook XML from DOCX content.
    
    Args:
        content: Extracted DOCX content
        output_path: Optional path to write the XML file
        **kwargs: Additional arguments passed to DocBookGenerator
        
    Returns:
        XML string
    """
    generator = DocBookGenerator(**kwargs)
    return generator.generate(content, output_path)


if __name__ == "__main__":
    import sys
    from docx_extractor import extract_docx
    
    if len(sys.argv) < 2:
        print("Usage: python docbook_generator.py <docx_file> [output_xml]")
        sys.exit(1)
    
    docx_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else None
    
    # Extract content
    print(f"Extracting content from: {docx_path}")
    content = extract_docx(docx_path)
    
    # Generate DocBook XML
    print("Generating DocBook XML...")
    xml_str = generate_docbook(content, output_path)
    
    if output_path:
        print(f"Written to: {output_path}")
    else:
        print("\n" + "=" * 60)
        print(xml_str[:2000])
        print("..." if len(xml_str) > 2000 else "")
