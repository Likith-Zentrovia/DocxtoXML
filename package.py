#!/usr/bin/env python3
"""
RittDoc Package Generator

This module creates RittDoc-compliant ZIP packages from the converted
DocBook XML and extracted media files.

The package structure matches the PDF to XML pipeline output:
- Book.xml (main DocBook file)
- multimedia/ (extracted images)
- metadata.csv (optional metadata)
"""

from __future__ import annotations

import csv
import hashlib
import os
import re
import shutil
import tempfile
import zipfile
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Union, Any

from lxml import etree

from docx_extractor import DocxContent, ExtractedImage


# ============================================================================
# CONSTANTS
# ============================================================================

BOOK_DOCTYPE_PUBLIC = "-//RIS Dev//DTD DocBook V4.3 -Based Variant V1.1//EN"
BOOK_DOCTYPE_SYSTEM = "http://LOCALHOST/dtd/V1.1/RittDocBook.dtd"


# ============================================================================
# DATA CLASSES
# ============================================================================

@dataclass
class PackageResult:
    """Result of package creation."""
    success: bool
    package_path: Optional[str] = None
    files_included: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)
    
    # Statistics
    xml_size: int = 0
    media_count: int = 0
    total_size: int = 0


@dataclass
class ImageMetadata:
    """Metadata for a packaged image."""
    filename: str
    original_filename: str
    chapter: str
    figure_number: str
    caption: str
    alt_text: str
    width: int
    height: int
    file_size: str
    format: str


# ============================================================================
# PACKAGE GENERATOR CLASS
# ============================================================================

class PackageGenerator:
    """
    Generates RittDoc-compliant ZIP packages.
    
    The package structure:
    - Book.xml (or custom root name)
    - multimedia/
      - img_0001.png
      - img_0002.jpg
      - ...
    - metadata.csv (optional)
    """
    
    def __init__(
        self,
        book_filename: str = "Book.xml",
        multimedia_folder: str = "multimedia",
        include_metadata_csv: bool = True
    ):
        """
        Initialize the package generator.
        
        Args:
            book_filename: Name for the main XML file
            multimedia_folder: Name for the media folder
            include_metadata_csv: Whether to include metadata.csv
        """
        self.book_filename = book_filename
        self.multimedia_folder = multimedia_folder
        self.include_metadata_csv = include_metadata_csv

    def create_package(
        self,
        xml_content: str,
        images: List[ExtractedImage],
        output_path: Union[str, Path],
        content: Optional[DocxContent] = None
    ) -> PackageResult:
        """
        Create a RittDoc ZIP package.
        
        Args:
            xml_content: The DocBook XML content
            images: List of extracted images
            output_path: Path for the output ZIP file
            content: Optional DocxContent for additional metadata
            
        Returns:
            PackageResult with details of the created package
        """
        output_path = Path(output_path)
        result = PackageResult(success=False)
        
        try:
            # Ensure output directory exists
            output_path.parent.mkdir(parents=True, exist_ok=True)
            
            # Create the ZIP package
            with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                # Add the main XML file
                xml_bytes = xml_content.encode('utf-8')
                zf.writestr(self.book_filename, xml_bytes)
                result.files_included.append(self.book_filename)
                result.xml_size = len(xml_bytes)
                
                # Add images to multimedia folder
                image_metadata = []
                for i, img in enumerate(images, 1):
                    img_path = f"{self.multimedia_folder}/{img.filename}"
                    zf.writestr(img_path, img.data)
                    result.files_included.append(img_path)
                    result.media_count += 1
                    
                    # Collect metadata
                    image_metadata.append(ImageMetadata(
                        filename=img.filename,
                        original_filename=img.filename,
                        chapter="",
                        figure_number=str(i),
                        caption=img.caption,
                        alt_text=img.alt_text,
                        width=img.width or 0,
                        height=img.height or 0,
                        file_size=self._format_size(len(img.data)),
                        format=img.content_type.split('/')[-1].upper()
                    ))
                
                # Add metadata CSV if enabled
                if self.include_metadata_csv and image_metadata:
                    csv_content = self._create_metadata_csv(image_metadata)
                    zf.writestr("metadata.csv", csv_content)
                    result.files_included.append("metadata.csv")
                
                # Add book metadata CSV if content provided
                if content and self.include_metadata_csv:
                    book_csv = self._create_book_metadata_csv(content)
                    zf.writestr("book_metadata.csv", book_csv)
                    result.files_included.append("book_metadata.csv")
            
            # Calculate total size
            result.total_size = output_path.stat().st_size
            result.package_path = str(output_path)
            result.success = True
            
        except Exception as e:
            result.errors.append(str(e))
            result.success = False
        
        return result

    def _create_metadata_csv(self, images: List[ImageMetadata]) -> str:
        """Create CSV content for image metadata."""
        output = BytesIO()
        writer = csv.writer(output)
        
        # Header
        writer.writerow([
            "Filename",
            "Original Filename",
            "Chapter",
            "Figure Number",
            "Caption",
            "Alt Text",
            "Width",
            "Height",
            "File Size",
            "Format"
        ])
        
        # Data rows
        for img in images:
            writer.writerow([
                img.filename,
                img.original_filename,
                img.chapter,
                img.figure_number,
                img.caption,
                img.alt_text,
                img.width,
                img.height,
                img.file_size,
                img.format
            ])
        
        return output.getvalue().decode('utf-8')

    def _create_book_metadata_csv(self, content: DocxContent) -> str:
        """Create CSV content for book metadata."""
        output = BytesIO()
        writer = csv.writer(output)
        
        # Header
        writer.writerow(["Field", "Value"])
        
        # Data rows
        writer.writerow(["Title", content.title])
        writer.writerow(["Authors", "; ".join(content.authors)])
        writer.writerow(["ISBN", content.metadata.get("isbn", "")])
        writer.writerow(["Publisher", content.metadata.get("publisher", "")])
        writer.writerow(["Publication Date", content.metadata.get("pubdate", "")])
        writer.writerow(["Created", content.metadata.get("created", "")])
        writer.writerow(["Modified", content.metadata.get("modified", "")])
        writer.writerow(["Chapters", str(len(content.chapters))])
        writer.writerow(["Tables", str(len(content.tables))])
        writer.writerow(["Images", str(len(content.images))])
        
        return output.getvalue().decode('utf-8')

    def _format_size(self, size_bytes: int) -> str:
        """Format file size in human-readable form."""
        if size_bytes < 1024:
            return f"{size_bytes} B"
        elif size_bytes < 1024 * 1024:
            return f"{size_bytes / 1024:.1f} KB"
        else:
            return f"{size_bytes / (1024 * 1024):.1f} MB"


# ============================================================================
# CONVENIENCE FUNCTIONS
# ============================================================================

def create_rittdoc_package(
    xml_content: str,
    images: List[ExtractedImage],
    output_path: Union[str, Path],
    content: Optional[DocxContent] = None,
    **kwargs
) -> PackageResult:
    """
    Convenience function to create a RittDoc package.
    
    Args:
        xml_content: The DocBook XML content
        images: List of extracted images
        output_path: Path for the output ZIP file
        content: Optional DocxContent for additional metadata
        **kwargs: Additional arguments passed to PackageGenerator
        
    Returns:
        PackageResult with details of the created package
    """
    generator = PackageGenerator(**kwargs)
    return generator.create_package(xml_content, images, output_path, content)


def save_images_to_folder(
    images: List[ExtractedImage],
    output_dir: Union[str, Path]
) -> List[str]:
    """
    Save extracted images to a folder.
    
    Args:
        images: List of extracted images
        output_dir: Directory to save images
        
    Returns:
        List of saved file paths
    """
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    saved_paths = []
    for img in images:
        img_path = output_dir / img.filename
        with open(img_path, 'wb') as f:
            f.write(img.data)
        saved_paths.append(str(img_path))
    
    return saved_paths


if __name__ == "__main__":
    import sys
    from docx_extractor import extract_docx
    from docbook_generator import generate_docbook
    
    if len(sys.argv) < 2:
        print("Usage: python package.py <docx_file> [output_zip]")
        sys.exit(1)
    
    docx_path = sys.argv[1]
    stem = Path(docx_path).stem
    output_zip = sys.argv[2] if len(sys.argv) > 2 else f"{stem}_rittdoc.zip"
    
    # Extract content
    print(f"Extracting content from: {docx_path}")
    content = extract_docx(docx_path)
    
    # Generate DocBook XML
    print("Generating DocBook XML...")
    xml_content = generate_docbook(content)
    
    # Create package
    print(f"Creating package: {output_zip}")
    result = create_rittdoc_package(
        xml_content=xml_content,
        images=content.images,
        output_path=output_zip,
        content=content
    )
    
    if result.success:
        print(f"\nPackage created successfully!")
        print(f"  Path: {result.package_path}")
        print(f"  Files: {len(result.files_included)}")
        print(f"  Media: {result.media_count}")
        print(f"  Total size: {result.total_size / 1024:.1f} KB")
    else:
        print(f"\nPackage creation failed!")
        for error in result.errors:
            print(f"  Error: {error}")
