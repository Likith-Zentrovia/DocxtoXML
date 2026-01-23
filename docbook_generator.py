#!/usr/bin/env python3
"""
DocBook XML Generator (RittDoc DTD Compliant)

Converts extracted DOCX content to DocBook XML following the RittDoc DTD
specification. Matches the output format of the PDFtoXML pipeline.

Key rules followed:
- Chapter IDs: ch0001, ch0002, etc.
- Section IDs: hierarchical (s0101, s0201, etc.)
- Figure naming: Ch{chapter}f{figure_num}.{ext}
- Table format: CALS with <table><title><tgroup cols="N">
- Image attributes: fileref, width="100%", scalefit="1"
- DOCTYPE: RittDoc public/system IDs
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List, Optional, Union, Any
from datetime import datetime

from lxml import etree

from docx_extractor import (
    DocxContent, DocumentElement, TextBlock, ExtractedTable, ExtractedImage
)


# ============================================================================
# CONSTANTS
# ============================================================================

RITTDOC_DOCTYPE_PUBLIC = "-//RIS Dev//DTD DocBook V4.3 -Based Variant V1.1//EN"
RITTDOC_DOCTYPE_SYSTEM = "http://LOCALHOST/dtd/V1.1/RittDocBook.dtd"


# ============================================================================
# DOCBOOK GENERATOR CLASS
# ============================================================================

class DocBookGenerator:
    """
    Generates RittDoc DTD-compliant DocBook XML from extracted DOCX content.

    Processes document elements in order, placing images and tables
    at their exact positions as they appear in the source document.
    """

    def __init__(self, multimedia_prefix: str = ""):
        """
        Args:
            multimedia_prefix: Prefix for image file references (empty = just filename)
        """
        self.multimedia_prefix = multimedia_prefix

        # Counters
        self._chapter_counter = 0
        self._figure_counter = 0
        self._table_counter = 0
        self._sect1_counter = 0
        self._sect2_counter = 0
        self._sect3_counter = 0

        # Current chapter code for image naming
        self._current_chapter_code = "Ch0001"

    def _get_section_code(self) -> str:
        """
        Build the 4-digit section code: s{sect1:02d}{sect2:02d}
        e.g., s0100 (sect1=01), s0101 (sect1=01, sect2=01)
        """
        if self._sect1_counter > 0:
            return f"s{self._sect1_counter:02d}{self._sect2_counter:02d}"
        return "s0000"

    def _get_chapter_id(self) -> str:
        """Get current chapter ID: ch0000 format."""
        return f"ch{self._chapter_counter:04d}"

    def _get_section_id(self) -> str:
        """Get full section ID: ch0000s0000 format."""
        return f"{self._get_chapter_id()}{self._get_section_code()}"

    def generate(self, content: DocxContent, output_path: Optional[Union[str, Path]] = None) -> str:
        """
        Generate DocBook XML from extracted content.

        Args:
            content: Extracted DOCX content with elements in document order
            output_path: Optional path to write the XML file

        Returns:
            XML string
        """
        # Reset counters
        self._chapter_counter = 0
        self._figure_counter = 0
        self._table_counter = 0
        self._sect1_counter = 0
        self._sect2_counter = 0
        self._sect3_counter = 0

        # Create root element
        root = etree.Element("book")

        # Add bookinfo
        bookinfo = self._create_bookinfo(content)
        root.append(bookinfo)

        # Process all elements in document order
        self._process_elements(root, content)

        # Rename images to follow convention
        self._rename_images(content)

        # Serialize
        xml_str = self._serialize_xml(root)

        # Write to file if path provided
        if output_path:
            output_path = Path(output_path)
            output_path.parent.mkdir(parents=True, exist_ok=True)
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(xml_str)

        return xml_str

    def _create_bookinfo(self, content: DocxContent) -> etree._Element:
        """Create bookinfo section with metadata."""
        bookinfo = etree.Element("bookinfo")

        # ISBN
        isbn = etree.SubElement(bookinfo, "isbn")
        isbn.text = content.metadata.get("isbn", "0000000000000")

        # Title
        title = etree.SubElement(bookinfo, "title")
        title.text = content.title or "Untitled Document"

        # Authors
        authorgroup = etree.SubElement(bookinfo, "authorgroup")
        authors = content.authors or ["Unknown Author"]

        for author_name in authors:
            author = etree.SubElement(authorgroup, "author")
            personname = etree.SubElement(author, "personname")
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
        holder.text = content.metadata.get("copyright_holder", "Copyright Holder")

        return bookinfo

    def _process_elements(self, root: etree._Element, content: DocxContent):
        """
        Process all document elements in order.
        Creates chapters, sections, paragraphs, figures, and tables
        at their exact positions.
        """
        # Current containers
        current_chapter = None
        current_sect1 = None
        current_sect2 = None
        current_sect3 = None

        # List tracking
        current_list = None
        current_list_type = None

        for elem in content.elements:
            if elem.element_type == "paragraph" and elem.paragraph:
                block = elem.paragraph

                # Handle headings
                if block.level == 1:
                    # New chapter
                    current_list = None
                    current_list_type = None
                    current_sect1 = None
                    current_sect2 = None
                    current_sect3 = None
                    self._sect1_counter = 0
                    self._sect2_counter = 0
                    self._sect3_counter = 0

                    self._chapter_counter += 1
                    self._figure_counter = 0  # Reset per chapter
                    self._table_counter = 0

                    current_chapter = etree.SubElement(root, "chapter")
                    current_chapter.set("id", f"ch{self._chapter_counter:04d}")
                    self._current_chapter_code = f"Ch{self._chapter_counter:04d}"

                    title = etree.SubElement(current_chapter, "title")
                    title.text = self._clean_text(block.text)
                    continue

                elif block.level == 2:
                    # sect1
                    current_list = None
                    current_list_type = None
                    current_sect2 = None
                    current_sect3 = None
                    self._sect2_counter = 0
                    self._sect3_counter = 0

                    if current_chapter is None:
                        current_chapter = self._ensure_chapter(root, content.title)

                    self._sect1_counter += 1
                    current_sect1 = etree.SubElement(current_chapter, "sect1")
                    current_sect1.set("id", self._get_section_id())

                    title = etree.SubElement(current_sect1, "title")
                    title.text = self._clean_text(block.text)
                    continue

                elif block.level == 3:
                    # sect2
                    current_list = None
                    current_list_type = None
                    current_sect3 = None
                    self._sect3_counter = 0

                    if current_chapter is None:
                        current_chapter = self._ensure_chapter(root, content.title)
                    if current_sect1 is None:
                        self._sect1_counter += 1
                        current_sect1 = etree.SubElement(current_chapter, "sect1")
                        current_sect1.set("id", self._get_section_id())
                        t = etree.SubElement(current_sect1, "title")
                        t.text = "Section"

                    self._sect2_counter += 1
                    current_sect2 = etree.SubElement(current_sect1, "sect2")
                    current_sect2.set("id", self._get_section_id())

                    title = etree.SubElement(current_sect2, "title")
                    title.text = self._clean_text(block.text)
                    continue

                elif block.level >= 4:
                    # sect3
                    current_list = None
                    current_list_type = None

                    if current_chapter is None:
                        current_chapter = self._ensure_chapter(root, content.title)
                    if current_sect1 is None:
                        self._sect1_counter += 1
                        current_sect1 = etree.SubElement(current_chapter, "sect1")
                        current_sect1.set("id", self._get_section_id())
                        t = etree.SubElement(current_sect1, "title")
                        t.text = "Section"
                    if current_sect2 is None:
                        self._sect2_counter += 1
                        current_sect2 = etree.SubElement(current_sect1, "sect2")
                        current_sect2.set("id", self._get_section_id())
                        t = etree.SubElement(current_sect2, "title")
                        t.text = "Subsection"

                    self._sect3_counter += 1
                    current_sect3 = etree.SubElement(current_sect2, "sect3")
                    current_sect3.set("id", self._get_section_id())

                    title = etree.SubElement(current_sect3, "title")
                    title.text = self._clean_text(block.text)
                    continue

                # Regular content - determine parent
                parent = current_sect3 or current_sect2 or current_sect1 or current_chapter
                if parent is None:
                    current_chapter = self._ensure_chapter(root, content.title)
                    parent = current_chapter

                # Handle lists
                if block.list_type:
                    list_tag = "itemizedlist" if block.list_type == "bullet" else "orderedlist"

                    if current_list is None or current_list_type != block.list_type:
                        current_list = etree.SubElement(parent, list_tag)
                        current_list_type = block.list_type

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

            elif elem.element_type == "image" and elem.image:
                # Close any open list
                current_list = None
                current_list_type = None

                parent = current_sect3 or current_sect2 or current_sect1 or current_chapter
                if parent is None:
                    current_chapter = self._ensure_chapter(root, content.title)
                    parent = current_chapter

                self._figure_counter += 1
                figure = self._create_figure(elem.image)
                parent.append(figure)

            elif elem.element_type == "table" and elem.table:
                # Close any open list
                current_list = None
                current_list_type = None

                parent = current_sect3 or current_sect2 or current_sect1 or current_chapter
                if parent is None:
                    current_chapter = self._ensure_chapter(root, content.title)
                    parent = current_chapter

                self._table_counter += 1
                table_elem = self._create_table(elem.table)
                parent.append(table_elem)

    def _ensure_chapter(self, root: etree._Element, title: str) -> etree._Element:
        """Create a default chapter if none exists."""
        self._chapter_counter += 1
        self._figure_counter = 0
        self._table_counter = 0
        self._sect1_counter = 0
        self._sect2_counter = 0
        self._sect3_counter = 0
        self._current_chapter_code = f"Ch{self._chapter_counter:04d}"

        chapter = etree.SubElement(root, "chapter")
        chapter.set("id", self._get_chapter_id())
        t = etree.SubElement(chapter, "title")
        t.text = title or "Content"
        return chapter

    def _create_figure(self, img: ExtractedImage) -> etree._Element:
        """
        Create a DTD-compliant figure element.

        Structure: <figure><title/><mediaobject><imageobject><imagedata/></imageobject></mediaobject></figure>

        ID format: ch0000s0000fg00
        Filename format: Ch0000s0000fg00.ext
        """
        ext = Path(img.filename).suffix
        section_code = self._get_section_code()

        # Figure ID: ch0000s0000fg00
        fig_id = f"{self._get_chapter_id()}{section_code}fg{self._figure_counter:02d}"

        # Filename: Ch0000s0000fg00.ext (uppercase Ch for filename)
        figure_filename = f"{self._current_chapter_code}{section_code}fg{self._figure_counter:02d}{ext}"

        # Update image filename for package
        img.filename = figure_filename

        # Create figure element
        figure = etree.Element("figure")
        figure.set("id", fig_id)

        # Title (required by DTD)
        fig_title = etree.SubElement(figure, "title")
        fig_title.text = img.caption or f"Figure {self._figure_counter}"

        # Mediaobject
        mediaobject = etree.SubElement(figure, "mediaobject")
        imageobject = etree.SubElement(mediaobject, "imageobject")
        imagedata = etree.SubElement(imageobject, "imagedata")
        imagedata.set("fileref", f"{self.multimedia_prefix}{figure_filename}")
        imagedata.set("width", "100%")
        imagedata.set("scalefit", "1")

        return figure

    def _create_table(self, table: ExtractedTable) -> etree._Element:
        """
        Create a DTD-compliant table element.

        Structure: <table><title/><tgroup cols="N"><colspec/><thead/><tbody/></tgroup></table>
        Uses CALS table format (not HTML).

        ID format: ch0000s0000tb00
        """
        # Create table element (not informaltable - DTD requires <table> with <title>)
        table_elem = etree.Element("table")
        section_code = self._get_section_code()
        table_id = f"{self._get_chapter_id()}{section_code}tb{self._table_counter:02d}"
        table_elem.set("id", table_id)

        # Title (required by DTD)
        table_title = etree.SubElement(table_elem, "title")
        table_title.text = table.caption or f"Table {self._table_counter}"

        # Tgroup with cols attribute (required)
        num_cols = table.num_cols or (max(len(row) for row in table.rows) if table.rows else 1)
        tgroup = etree.SubElement(table_elem, "tgroup")
        tgroup.set("cols", str(num_cols))

        # Colspec for each column
        for i in range(num_cols):
            colspec = etree.SubElement(tgroup, "colspec")
            colspec.set("colname", f"c{i+1}")

        # Header rows
        header_rows = table.rows[:table.header_rows] if table.header_rows > 0 else []
        body_rows = table.rows[table.header_rows:] if table.header_rows > 0 else table.rows

        if header_rows:
            thead = etree.SubElement(tgroup, "thead")
            for row in header_rows:
                tr = etree.SubElement(thead, "row")
                for cell_text in row:
                    entry = etree.SubElement(tr, "entry")
                    entry.text = self._clean_text(cell_text)

        # Body rows
        if body_rows:
            tbody = etree.SubElement(tgroup, "tbody")
            for row in body_rows:
                tr = etree.SubElement(tbody, "row")
                for cell_text in row:
                    entry = etree.SubElement(tr, "entry")
                    entry.text = self._clean_text(cell_text)

        return table_elem

    def _rename_images(self, content: DocxContent):
        """Rename images in the content to follow the convention."""
        # Images have already been renamed during figure creation
        # This method is a hook for any additional renaming logic
        pass

    def _set_para_content(self, para: etree._Element, text: str):
        """Set paragraph content with inline formatting."""
        text = self._clean_text(text)

        if "**" in text or "*" in text:
            parts = self._parse_inline_formatting(text)
            if len(parts) == 1 and parts[0][1] is None:
                para.text = parts[0][0]
            else:
                para.text = ""
                last_elem = None
                for part_text, fmt in parts:
                    if fmt is None:
                        if last_elem is not None:
                            last_elem.tail = (last_elem.tail or "") + part_text
                        else:
                            para.text = (para.text or "") + part_text
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
        pattern = r'(\*\*\*(.+?)\*\*\*|\*\*(.+?)\*\*|\*(.+?)\*)'

        for match in re.finditer(pattern, text):
            if match.start() > current_pos:
                parts.append((text[current_pos:match.start()], None))
            if match.group(2):
                parts.append((match.group(2), "bold"))
            elif match.group(3):
                parts.append((match.group(3), "bold"))
            elif match.group(4):
                parts.append((match.group(4), "italic"))
            current_pos = match.end()

        if current_pos < len(text):
            parts.append((text[current_pos:], None))

        if not parts:
            parts.append((text, None))

        return parts

    def _clean_text(self, text: str) -> str:
        """Clean text for XML output."""
        if not text:
            return ""
        text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', text)
        text = re.sub(r'\s+', ' ', text)
        return text.strip()

    def _serialize_xml(self, root: etree._Element) -> str:
        """Serialize XML with proper declaration and DOCTYPE."""
        xml_declaration = '<?xml version="1.0" encoding="UTF-8"?>\n'
        doctype = f'<!DOCTYPE book PUBLIC "{RITTDOC_DOCTYPE_PUBLIC}"\n  "{RITTDOC_DOCTYPE_SYSTEM}">\n'

        xml_content = etree.tostring(
            root,
            encoding="unicode",
            pretty_print=True
        )

        return xml_declaration + doctype + xml_content


# ============================================================================
# CONVENIENCE FUNCTIONS
# ============================================================================

def generate_docbook(content: DocxContent, output_path: Optional[Union[str, Path]] = None, **kwargs) -> str:
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

    print(f"Extracting content from: {docx_path}")
    content = extract_docx(docx_path)

    print("Generating DocBook XML...")
    xml_str = generate_docbook(content, output_path)

    if output_path:
        print(f"Written to: {output_path}")
    else:
        print("\n" + "=" * 60)
        print(xml_str[:2000])
        if len(xml_str) > 2000:
            print("...")
