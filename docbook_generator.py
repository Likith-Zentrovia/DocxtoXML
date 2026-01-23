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

        # Counters (per-chapter, reset on new chapter)
        self._chapter_counter = 0
        self._figure_counter = 0
        self._table_counter = 0
        self._sect1_counter = 0
        self._sect2_counter = 0
        self._sect3_counter = 0

        # Global sequential counters (document-wide, for cross-references)
        self._global_figure_counter = 0
        self._global_table_counter = 0

        # Maps: sequential number -> element ID (for cross-reference linking)
        self._figure_id_map: Dict[int, str] = {}
        self._table_id_map: Dict[int, str] = {}

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
        self._global_figure_counter = 0
        self._global_table_counter = 0
        self._figure_id_map = {}
        self._table_id_map = {}

        # Create root element
        root = etree.Element("book")

        # Add bookinfo
        bookinfo = self._create_bookinfo(content)
        root.append(bookinfo)

        # Process all elements in document order
        self._process_elements(root, content)

        # Post-process: convert "Figure N" / "Table N" references to <link> elements
        self._post_process_references(root)

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
                parent = (current_sect3 if current_sect3 is not None
                          else current_sect2 if current_sect2 is not None
                          else current_sect1 if current_sect1 is not None
                          else current_chapter)
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

                parent = (current_sect3 if current_sect3 is not None
                          else current_sect2 if current_sect2 is not None
                          else current_sect1 if current_sect1 is not None
                          else current_chapter)
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

                parent = (current_sect3 if current_sect3 is not None
                          else current_sect2 if current_sect2 is not None
                          else current_sect1 if current_sect1 is not None
                          else current_chapter)
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

        # Track global figure number for cross-references
        self._global_figure_counter += 1
        self._figure_id_map[self._global_figure_counter] = fig_id

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

        # Track global table number for cross-references
        self._global_table_counter += 1
        self._table_id_map[self._global_table_counter] = table_id

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

    def _post_process_references(self, root: etree._Element):
        """
        Post-process the XML tree to convert figure/table references to links.

        Converts patterns like:
        - "Figure 29" / "Fig. 29" -> <link linkend="fig_id">Figure 29</link>
        - "Table 5" / "Tab. 5" -> <link linkend="table_id">Table 5</link>

        Handles references in:
        - <emphasis role="bold">Figure N</emphasis> (replaces emphasis with link)
        - Plain text within <para> elements
        """
        # Pattern to match figure/table references
        ref_pattern = re.compile(
            r'\b(Figures?|Figs?\.?|Tables?|Tabs?\.?)\s+(\d+(?:\s*[-–&,]\s*\d+)*)',
            re.IGNORECASE
        )

        # Process all emphasis elements first (bold references like "Figure 29")
        for emphasis in root.iter('emphasis'):
            if emphasis.text:
                match = ref_pattern.match(emphasis.text.strip())
                if match:
                    ref_type = match.group(1).lower()
                    ref_nums = match.group(2)
                    linkend = self._resolve_reference(ref_type, ref_nums)
                    if linkend:
                        self._replace_with_link(emphasis, linkend)

        # Process plain text in paragraphs
        for para in root.iter('para'):
            self._linkify_text_references(para, ref_pattern)

    def _resolve_reference(self, ref_type: str, ref_nums: str) -> Optional[str]:
        """
        Resolve a reference type and number to a linkend ID.

        Args:
            ref_type: "figure", "fig.", "table", etc.
            ref_nums: "29" or "29-30" or "29, 30"

        Returns:
            linkend ID string or None if not found
        """
        # Extract first number from the reference
        num_match = re.search(r'(\d+)', ref_nums)
        if not num_match:
            return None
        num = int(num_match.group(1))

        is_figure = ref_type.startswith('fig') or ref_type.startswith('Fig')
        is_table = ref_type.startswith('tab') or ref_type.startswith('Tab')

        if is_figure and num in self._figure_id_map:
            return self._figure_id_map[num]
        elif is_table and num in self._table_id_map:
            return self._table_id_map[num]

        return None

    def _replace_with_link(self, elem: etree._Element, linkend: str):
        """
        Replace an element (like <emphasis>) with a <link> element in-place.
        Preserves text content and tail.
        """
        parent = elem.getparent()
        if parent is None:
            return

        # Create link element
        link = etree.Element("link")
        link.set("linkend", linkend)
        link.text = elem.text
        link.tail = elem.tail

        # Replace in parent
        idx = list(parent).index(elem)
        parent.remove(elem)
        parent.insert(idx, link)

    def _linkify_text_references(self, para: etree._Element, pattern: re.Pattern):
        """
        Find and convert figure/table references in plain text within a paragraph.
        Handles text in para.text and in tail text of child elements.
        """
        # Process para.text
        if para.text:
            new_text, links = self._extract_links_from_text(para.text, pattern)
            if links:
                para.text = new_text
                # Insert link elements at the beginning (after text)
                for i, (link_text, linkend, tail) in enumerate(links):
                    link = etree.Element("link")
                    link.set("linkend", linkend)
                    link.text = link_text
                    link.tail = tail
                    para.insert(i, link)

        # Process tail text of child elements
        for child in list(para):
            if child.tag == 'link':
                continue  # Don't process already-linked text
            if child.tail:
                new_tail, links = self._extract_links_from_text(child.tail, pattern)
                if links:
                    child.tail = new_tail
                    parent = child.getparent()
                    idx = list(parent).index(child)
                    for i, (link_text, linkend, tail) in enumerate(links):
                        link = etree.Element("link")
                        link.set("linkend", linkend)
                        link.text = link_text
                        link.tail = tail
                        parent.insert(idx + 1 + i, link)

    def _extract_links_from_text(self, text: str, pattern: re.Pattern) -> tuple:
        """
        Extract link references from text.

        Returns:
            (remaining_text, [(link_text, linkend, tail_text), ...])
        """
        links = []
        last_end = 0
        remaining = ""

        for match in pattern.finditer(text):
            ref_type = match.group(1).lower()
            ref_nums = match.group(2)
            linkend = self._resolve_reference(ref_type, ref_nums)

            if linkend:
                # Text before this match goes as prefix
                remaining += text[last_end:match.start()]
                # The match itself becomes a link; tail is text after until next match
                links.append((match.group(0), linkend, ""))
                last_end = match.end()

        if not links:
            return text, []

        # Set the tail of the last link to remaining text
        if last_end < len(text):
            if links:
                links[-1] = (links[-1][0], links[-1][1], text[last_end:])
        # For intermediate links, set tails between matches
        # (already handled by the loop - remaining text after each match)

        return remaining, links

    def _set_para_content(self, para: etree._Element, text: str):
        """Set paragraph content with inline formatting (bold, italic, subscript, superscript)."""
        text = self._clean_text(text)

        has_formatting = ("**" in text or "*" in text or
                         "{sub:" in text or "{sup:" in text)

        if has_formatting:
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
                    elif fmt == "bold":
                        elem = etree.SubElement(para, "emphasis")
                        elem.set("role", "bold")
                        elem.text = part_text
                        last_elem = elem
                    elif fmt == "italic":
                        elem = etree.SubElement(para, "emphasis")
                        elem.text = part_text
                        last_elem = elem
                    elif fmt == "subscript":
                        elem = etree.SubElement(para, "subscript")
                        elem.text = part_text
                        last_elem = elem
                    elif fmt == "superscript":
                        elem = etree.SubElement(para, "superscript")
                        elem.text = part_text
                        last_elem = elem
        else:
            para.text = text

    def _parse_inline_formatting(self, text: str) -> List[tuple]:
        """Parse text with inline formatting markers: **bold**, *italic*, {sub:text}, {sup:text}."""
        parts = []
        current_pos = 0
        # Match: {sub:...}, {sup:...}, ***bold-italic***, **bold**, *italic*
        pattern = r'(\{sub:(.+?)\}|\{sup:(.+?)\}|\*\*\*(.+?)\*\*\*|\*\*(.+?)\*\*|\*(.+?)\*)'

        for match in re.finditer(pattern, text):
            if match.start() > current_pos:
                parts.append((text[current_pos:match.start()], None))
            if match.group(2):
                parts.append((match.group(2), "subscript"))
            elif match.group(3):
                parts.append((match.group(3), "superscript"))
            elif match.group(4):
                parts.append((match.group(4), "bold"))
            elif match.group(5):
                parts.append((match.group(5), "bold"))
            elif match.group(6):
                parts.append((match.group(6), "italic"))
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
