#!/usr/bin/env python3
"""
DOCX to XML Conversion Orchestrator

This is the main entry point for the DOCX to XML conversion pipeline.
It coordinates extraction, conversion, and packaging of DOCX files
to RittDoc-compliant DocBook XML.

Usage:
    python docx_orchestrator.py input.docx --out ./output
    python docx_orchestrator.py input.docx --out ./output --create-zip
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import sys
import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Any

from config import PipelineConfig, get_config
from docx_extractor import DocxExtractor, DocxContent, extract_docx
from docbook_generator import DocBookGenerator, generate_docbook
from package import PackageGenerator, create_rittdoc_package, save_images_to_folder
from validation_report import validate_xml, ValidationReportGenerator


# ============================================================================
# RESULT DATA CLASS
# ============================================================================

@dataclass
class ConversionResult:
    """Result of DOCX to XML conversion."""
    success: bool
    input_file: str
    output_dir: str
    
    # Output files
    xml_path: Optional[str] = None
    package_path: Optional[str] = None
    multimedia_dir: Optional[str] = None
    validation_report_path: Optional[str] = None
    
    # Statistics
    text_blocks: int = 0
    images: int = 0
    tables: int = 0
    chapters: int = 0
    
    # Timing
    start_time: Optional[datetime] = None
    end_time: Optional[datetime] = None
    duration_seconds: float = 0.0
    
    # Errors
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for JSON serialization."""
        return {
            "success": self.success,
            "input_file": self.input_file,
            "output_dir": self.output_dir,
            "xml_path": self.xml_path,
            "package_path": self.package_path,
            "multimedia_dir": self.multimedia_dir,
            "statistics": {
                "text_blocks": self.text_blocks,
                "images": self.images,
                "tables": self.tables,
                "chapters": self.chapters,
            },
            "timing": {
                "start_time": self.start_time.isoformat() if self.start_time else None,
                "end_time": self.end_time.isoformat() if self.end_time else None,
                "duration_seconds": self.duration_seconds,
            },
            "errors": self.errors,
            "warnings": self.warnings,
        }


# ============================================================================
# ORCHESTRATOR CLASS
# ============================================================================

class DocxOrchestrator:
    """
    Orchestrates the DOCX to XML conversion pipeline.
    
    This class coordinates:
    1. Content extraction from DOCX
    2. DocBook XML generation
    3. Image extraction and organization
    4. Optional RittDoc package creation
    """
    
    def __init__(self, config: Optional[PipelineConfig] = None, verbose: bool = True):
        """
        Initialize the orchestrator.
        
        Args:
            config: Pipeline configuration (uses defaults if not provided)
            verbose: Whether to print progress messages
        """
        self.config = config or get_config()
        self.verbose = verbose
        
        # Initialize components
        self.extractor = DocxExtractor(
            extract_images=self.config.extraction.extract_images,
            extract_tables=self.config.extraction.extract_tables,
            preserve_formatting=self.config.extraction.preserve_formatting,
            min_image_size=self.config.extraction.min_image_size
        )
        
        self.generator = DocBookGenerator(
            multimedia_prefix="multimedia/"
        )
        
        self.packager = PackageGenerator(
            book_filename="Book.xml",
            multimedia_folder="multimedia",
            include_metadata_csv=True
        )

    def convert(
        self,
        docx_path: str | Path,
        output_dir: Optional[str | Path] = None,
        create_package: bool = True
    ) -> ConversionResult:
        """
        Convert a DOCX file to DocBook XML.
        
        Args:
            docx_path: Path to the input DOCX file
            output_dir: Output directory (defaults to config setting)
            create_package: Whether to create a RittDoc ZIP package
            
        Returns:
            ConversionResult with details of the conversion
        """
        docx_path = Path(docx_path).resolve()
        output_dir = Path(output_dir or self.config.output_dir).resolve()
        
        result = ConversionResult(
            success=False,
            input_file=str(docx_path),
            output_dir=str(output_dir),
            start_time=datetime.now()
        )
        
        stem = docx_path.stem
        
        try:
            # Ensure output directory exists
            output_dir.mkdir(parents=True, exist_ok=True)
            
            # Step 1: Extract content from DOCX
            if self.verbose:
                print(f"\n{'=' * 60}")
                print(f"DOCX to XML Conversion: {docx_path.name}")
                print(f"{'=' * 60}")
                print(f"\nStep 1: Extracting content from DOCX...")
            
            content = self.extractor.extract(docx_path)
            
            result.text_blocks = len(content.text_blocks)
            result.images = len(content.images)
            result.tables = len(content.tables)
            result.chapters = len(content.chapters)
            
            if self.verbose:
                print(f"  - Title: {content.title}")
                print(f"  - Text blocks: {result.text_blocks}")
                print(f"  - Images: {result.images}")
                print(f"  - Tables: {result.tables}")
                print(f"  - Chapters: {result.chapters}")
            
            # Step 2: Generate DocBook XML
            if self.verbose:
                print(f"\nStep 2: Generating DocBook XML...")
            
            xml_path = output_dir / f"{stem}_docbook42.xml"
            xml_content = self.generator.generate(content, xml_path)
            result.xml_path = str(xml_path)
            
            if self.verbose:
                print(f"  - XML written to: {xml_path}")
            
            # Step 3: Save images to multimedia folder
            if content.images:
                if self.verbose:
                    print(f"\nStep 3: Saving images to multimedia folder...")
                
                multimedia_dir = output_dir / f"{stem}_multimedia"
                saved_paths = save_images_to_folder(content.images, multimedia_dir)
                result.multimedia_dir = str(multimedia_dir)
                
                if self.verbose:
                    print(f"  - Saved {len(saved_paths)} images to: {multimedia_dir}")
            
            # Step 4: Create RittDoc package (optional)
            if create_package:
                if self.verbose:
                    print(f"\nStep 4: Creating RittDoc package...")
                
                package_path = output_dir / f"{stem}_rittdoc.zip"
                package_result = self.packager.create_package(
                    xml_content=xml_content,
                    images=content.images,
                    output_path=package_path,
                    content=content
                )
                
                if package_result.success:
                    result.package_path = str(package_path)
                    if self.verbose:
                        print(f"  - Package created: {package_path}")
                        print(f"  - Size: {package_result.total_size / 1024:.1f} KB")
                else:
                    for error in package_result.errors:
                        result.warnings.append(f"Package error: {error}")
                    if self.verbose:
                        print(f"  - Package creation had issues: {package_result.errors}")
            
            # Step 5: Generate validation report
            if self.verbose:
                step_num = 5 if create_package else 4
                print(f"\nStep {step_num}: Generating validation report...")

            try:
                validation_result = validate_xml(xml_content, f"{stem}_docbook42.xml")
                report_path = output_dir / f"{stem}_validation_report.xlsx"
                report_gen = ValidationReportGenerator()
                report_gen.generate_report(
                    validation_result, report_path,
                    f"Validation Report: {stem}"
                )
                result.validation_report_path = str(report_path)

                if self.verbose:
                    print(f"  - Errors: {validation_result.total_errors}")
                    print(f"  - Warnings: {validation_result.total_warnings}")
                    print(f"  - Report: {report_path}")
            except Exception as val_err:
                result.warnings.append(f"Validation report error: {val_err}")
                if self.verbose:
                    print(f"  - Warning: Could not generate validation report: {val_err}")

            # Success!
            result.success = True

        except Exception as e:
            result.errors.append(str(e))
            if self.verbose:
                print(f"\nError: {e}")
        
        # Finalize timing
        result.end_time = datetime.now()
        result.duration_seconds = (result.end_time - result.start_time).total_seconds()
        
        if self.verbose:
            print(f"\n{'=' * 60}")
            print(f"Conversion {'completed successfully' if result.success else 'failed'}!")
            print(f"Duration: {result.duration_seconds:.2f} seconds")
            if result.errors:
                print(f"Errors: {len(result.errors)}")
            print(f"{'=' * 60}\n")
        
        return result


# ============================================================================
# CLI INTERFACE
# ============================================================================

def main():
    """Command-line interface for DOCX to XML conversion."""
    parser = argparse.ArgumentParser(
        description="Convert DOCX files to RittDoc-compliant DocBook XML",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s document.docx
  %(prog)s document.docx --out ./output
  %(prog)s document.docx --out ./output --no-package
  %(prog)s document.docx --out ./output --json-result result.json
        """
    )
    
    parser.add_argument(
        "docx_file",
        help="Path to the input DOCX file"
    )
    
    parser.add_argument(
        "--out", "-o",
        default="./output",
        help="Output directory (default: ./output)"
    )
    
    parser.add_argument(
        "--no-package",
        action="store_true",
        help="Skip creating the RittDoc ZIP package"
    )
    
    parser.add_argument(
        "--no-images",
        action="store_true",
        help="Skip extracting images"
    )
    
    parser.add_argument(
        "--no-tables",
        action="store_true",
        help="Skip extracting tables"
    )
    
    parser.add_argument(
        "--json-result",
        help="Write conversion result to JSON file"
    )
    
    parser.add_argument(
        "--quiet", "-q",
        action="store_true",
        help="Suppress progress output"
    )
    
    parser.add_argument(
        "--api-mode",
        action="store_true",
        help="Run in API mode (returns JSON result)"
    )
    
    args = parser.parse_args()
    
    # Validate input file
    docx_path = Path(args.docx_file)
    if not docx_path.exists():
        print(f"Error: File not found: {docx_path}", file=sys.stderr)
        return 1
    
    if not docx_path.suffix.lower() in ['.docx', '.doc']:
        print(f"Warning: File may not be a DOCX file: {docx_path}", file=sys.stderr)
    
    # Configure pipeline
    config = get_config()
    config.extraction.extract_images = not args.no_images
    config.extraction.extract_tables = not args.no_tables
    
    # Create orchestrator
    orchestrator = DocxOrchestrator(
        config=config,
        verbose=not args.quiet and not args.api_mode
    )
    
    # Run conversion
    result = orchestrator.convert(
        docx_path=docx_path,
        output_dir=args.out,
        create_package=not args.no_package
    )
    
    # Write JSON result if requested
    if args.json_result:
        with open(args.json_result, 'w') as f:
            json.dump(result.to_dict(), f, indent=2)
    
    # API mode: print JSON result
    if args.api_mode:
        print(json.dumps(result.to_dict(), indent=2))
    
    return 0 if result.success else 1


if __name__ == "__main__":
    sys.exit(main())
