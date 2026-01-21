#!/usr/bin/env python3
"""
RittDoc Editor Server for DOCX to XML

A Flask-based web editor for viewing and editing RittDoc-compliant DocBook XML
generated from DOCX files.

Usage:
    python editor_server.py output/document_rittdoc.zip
    python editor_server.py output/document_docbook42.xml --multimedia output/document_multimedia
"""

from __future__ import annotations

import argparse
import os
import re
import sys
import tempfile
import webbrowser
import zipfile
from pathlib import Path
from typing import Dict, List, Optional, Any
from io import BytesIO

# Auto-install dependencies
def ensure_dependencies():
    """Ensure required packages are installed."""
    required = ['flask', 'flask_cors', 'lxml']
    missing = []
    for pkg in required:
        try:
            __import__(pkg.replace('-', '_'))
        except ImportError:
            missing.append(pkg)

    if missing:
        print(f"Installing missing dependencies: {', '.join(missing)}")
        import subprocess
        subprocess.check_call([sys.executable, '-m', 'pip', 'install'] + missing)

ensure_dependencies()

from flask import Flask, jsonify, request, send_file, send_from_directory
from flask_cors import CORS
from lxml import etree

# ============================================================================
# FLASK APPLICATION
# ============================================================================

app = Flask(__name__, static_folder='editor_ui', static_url_path='')
CORS(app)

# Global state
editor_state = {
    'xml_path': None,
    'xml_content': None,
    'multimedia_dir': None,
    'package_path': None,
    'temp_dir': None,
    'title': 'Untitled'
}


# ============================================================================
# XML TO HTML RENDERER
# ============================================================================

class XMLToHTMLRenderer:
    """Converts DocBook XML to browser-viewable HTML."""

    def __init__(self, multimedia_prefix: str = 'multimedia/'):
        self.multimedia_prefix = multimedia_prefix
        self.figure_counter = 0
        self.table_counter = 0

    def render(self, xml_content: str) -> str:
        """Convert DocBook XML to HTML."""
        try:
            # Parse XML with recovery mode for entity issues
            parser = etree.XMLParser(recover=True, resolve_entities=False)
            root = etree.fromstring(xml_content.encode('utf-8'), parser=parser)
        except Exception as e:
            return f"<div class='error'>Error parsing XML: {e}</div>"

        html_parts = ['<div class="docbook-content">']
        self._render_element(root, html_parts)
        html_parts.append('</div>')

        return '\n'.join(html_parts)

    def _render_element(self, elem, html_parts: List[str], depth: int = 0):
        """Recursively render an element to HTML."""
        tag = self._local_name(elem.tag)

        if tag == 'book':
            for child in elem:
                self._render_element(child, html_parts, depth)

        elif tag == 'bookinfo':
            html_parts.append('<div class="bookinfo">')
            for child in elem:
                self._render_element(child, html_parts, depth + 1)
            html_parts.append('</div>')

        elif tag == 'title':
            parent_tag = self._local_name(elem.getparent().tag) if elem.getparent() is not None else ''
            if parent_tag == 'book' or parent_tag == 'bookinfo':
                html_parts.append(f'<h1 class="book-title">{self._get_text(elem)}</h1>')
            elif parent_tag == 'chapter':
                html_parts.append(f'<h2 class="chapter-title">{self._get_text(elem)}</h2>')
            elif parent_tag in ('sect1', 'section'):
                html_parts.append(f'<h3 class="section-title">{self._get_text(elem)}</h3>')
            elif parent_tag == 'sect2':
                html_parts.append(f'<h4 class="section-title">{self._get_text(elem)}</h4>')
            elif parent_tag == 'sect3':
                html_parts.append(f'<h5 class="section-title">{self._get_text(elem)}</h5>')
            elif parent_tag == 'figure':
                html_parts.append(f'<figcaption>{self._get_text(elem)}</figcaption>')
            else:
                html_parts.append(f'<h4>{self._get_text(elem)}</h4>')

        elif tag == 'chapter':
            ch_id = elem.get('id', f'ch{depth}')
            html_parts.append(f'<section class="chapter" id="{ch_id}">')
            for child in elem:
                self._render_element(child, html_parts, depth + 1)
            html_parts.append('</section>')

        elif tag in ('sect1', 'sect2', 'sect3', 'section'):
            sec_id = elem.get('id', f'sec{depth}')
            html_parts.append(f'<section class="{tag}" id="{sec_id}">')
            for child in elem:
                self._render_element(child, html_parts, depth + 1)
            html_parts.append('</section>')

        elif tag == 'para':
            html_parts.append(f'<p>{self._render_inline(elem)}</p>')

        elif tag == 'figure':
            self.figure_counter += 1
            fig_id = elem.get('id', f'fig{self.figure_counter}')
            html_parts.append(f'<figure class="figure" id="{fig_id}">')
            for child in elem:
                self._render_element(child, html_parts, depth + 1)
            html_parts.append('</figure>')

        elif tag == 'mediaobject':
            for child in elem:
                self._render_element(child, html_parts, depth + 1)

        elif tag == 'imageobject':
            for child in elem:
                self._render_element(child, html_parts, depth + 1)

        elif tag == 'imagedata':
            fileref = elem.get('fileref', '')
            # Convert multimedia path to API path
            if fileref.startswith('multimedia/'):
                filename = fileref.replace('multimedia/', '')
                src = f'/api/media/{filename}'
            else:
                src = f'/api/media/{fileref}'

            width = elem.get('width', 'auto')
            html_parts.append(f'<img src="{src}" alt="Figure" style="max-width: 100%; width: {width};">')

        elif tag == 'informaltable' or tag == 'table':
            self.table_counter += 1
            html_parts.append(f'<table class="docbook-table" id="tbl{self.table_counter}">')
            for child in elem:
                self._render_element(child, html_parts, depth + 1)
            html_parts.append('</table>')

        elif tag == 'tgroup':
            for child in elem:
                self._render_element(child, html_parts, depth + 1)

        elif tag == 'thead':
            html_parts.append('<thead>')
            for child in elem:
                self._render_element(child, html_parts, depth + 1)
            html_parts.append('</thead>')

        elif tag == 'tbody':
            html_parts.append('<tbody>')
            for child in elem:
                self._render_element(child, html_parts, depth + 1)
            html_parts.append('</tbody>')

        elif tag == 'row':
            html_parts.append('<tr>')
            for child in elem:
                self._render_element(child, html_parts, depth + 1)
            html_parts.append('</tr>')

        elif tag == 'entry':
            parent_tag = self._local_name(elem.getparent().getparent().tag) if elem.getparent() is not None else ''
            cell_tag = 'th' if parent_tag == 'thead' else 'td'
            html_parts.append(f'<{cell_tag}>{self._render_inline(elem)}</{cell_tag}>')

        elif tag == 'itemizedlist':
            html_parts.append('<ul class="itemizedlist">')
            for child in elem:
                self._render_element(child, html_parts, depth + 1)
            html_parts.append('</ul>')

        elif tag == 'orderedlist':
            html_parts.append('<ol class="orderedlist">')
            for child in elem:
                self._render_element(child, html_parts, depth + 1)
            html_parts.append('</ol>')

        elif tag == 'listitem':
            html_parts.append('<li>')
            for child in elem:
                self._render_element(child, html_parts, depth + 1)
            html_parts.append('</li>')

        elif tag == 'author':
            html_parts.append('<div class="author">')
            for child in elem:
                self._render_element(child, html_parts, depth + 1)
            html_parts.append('</div>')

        elif tag in ('firstname', 'surname', 'othername'):
            html_parts.append(f'<span class="{tag}">{self._get_text(elem)} </span>')

        elif tag in ('isbn', 'publisher', 'pubdate', 'copyright'):
            html_parts.append(f'<div class="{tag}"><strong>{tag.upper()}:</strong> {self._get_text(elem)}</div>')

        else:
            # Default: render children
            for child in elem:
                self._render_element(child, html_parts, depth + 1)

    def _render_inline(self, elem) -> str:
        """Render element with inline formatting."""
        parts = []
        if elem.text:
            parts.append(elem.text)

        for child in elem:
            tag = self._local_name(child.tag)
            if tag == 'emphasis':
                role = child.get('role', 'italic')
                if role == 'bold':
                    parts.append(f'<strong>{self._get_text(child)}</strong>')
                else:
                    parts.append(f'<em>{self._get_text(child)}</em>')
            elif tag == 'link':
                href = child.get('href', '#')
                parts.append(f'<a href="{href}">{self._get_text(child)}</a>')
            else:
                parts.append(self._get_text(child))

            if child.tail:
                parts.append(child.tail)

        return ''.join(parts)

    def _get_text(self, elem) -> str:
        """Get all text content from an element."""
        return ''.join(elem.itertext())

    def _local_name(self, tag: str) -> str:
        """Extract local name from potentially namespaced tag."""
        if '}' in tag:
            return tag.split('}')[1]
        return tag


# ============================================================================
# API ENDPOINTS
# ============================================================================

@app.route('/')
def index():
    """Serve the editor UI."""
    return send_from_directory('editor_ui', 'index.html')


@app.route('/api/init')
def api_init():
    """Initialize editor with current document."""
    if not editor_state['xml_content']:
        return jsonify({'error': 'No document loaded'}), 400

    renderer = XMLToHTMLRenderer()
    html_content = renderer.render(editor_state['xml_content'])

    return jsonify({
        'title': editor_state['title'],
        'xml': editor_state['xml_content'],
        'html': html_content,
        'hasMultimedia': editor_state['multimedia_dir'] is not None
    })


@app.route('/api/render-html', methods=['POST'])
def api_render_html():
    """Convert XML to HTML."""
    data = request.get_json()
    xml_content = data.get('xml', '')

    renderer = XMLToHTMLRenderer()
    html_content = renderer.render(xml_content)

    return jsonify({'html': html_content})


@app.route('/api/save', methods=['POST'])
def api_save():
    """Save edited XML."""
    data = request.get_json()
    xml_content = data.get('xml', '')

    if not xml_content:
        return jsonify({'error': 'No XML content provided'}), 400

    # Update state
    editor_state['xml_content'] = xml_content

    # Save to file if path is set
    if editor_state['xml_path']:
        try:
            with open(editor_state['xml_path'], 'w', encoding='utf-8') as f:
                f.write(xml_content)

            # Also update the ZIP package if exists
            if editor_state['package_path']:
                update_package_xml(editor_state['package_path'], xml_content)

            return jsonify({'success': True, 'message': 'Saved successfully'})
        except Exception as e:
            return jsonify({'error': str(e)}), 500

    return jsonify({'success': True, 'message': 'Updated in memory'})


@app.route('/api/media/<path:filename>')
def api_media(filename):
    """Serve multimedia files."""
    if not editor_state['multimedia_dir']:
        return jsonify({'error': 'No multimedia directory'}), 404

    # Try multiple paths
    search_paths = [
        Path(editor_state['multimedia_dir']) / filename,
        Path(editor_state['multimedia_dir']) / 'multimedia' / filename,
    ]

    # Also search in temp dir if extracted from ZIP
    if editor_state['temp_dir']:
        search_paths.extend([
            Path(editor_state['temp_dir']) / 'multimedia' / filename,
            Path(editor_state['temp_dir']) / filename,
        ])

    for path in search_paths:
        if path.exists():
            return send_file(path)

    # Return placeholder
    return generate_placeholder_image(filename)


@app.route('/api/download-xml')
def api_download_xml():
    """Download the current XML file."""
    if not editor_state['xml_content']:
        return jsonify({'error': 'No document loaded'}), 400

    return send_file(
        BytesIO(editor_state['xml_content'].encode('utf-8')),
        mimetype='application/xml',
        as_attachment=True,
        download_name=f"{editor_state['title']}.xml"
    )


@app.route('/api/download-package')
def api_download_package():
    """Download the RittDoc package."""
    if not editor_state['package_path'] or not Path(editor_state['package_path']).exists():
        return jsonify({'error': 'No package available'}), 404

    return send_file(
        editor_state['package_path'],
        mimetype='application/zip',
        as_attachment=True
    )


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def generate_placeholder_image(filename: str):
    """Generate a placeholder image for missing media."""
    # Simple SVG placeholder
    svg = f'''<svg xmlns="http://www.w3.org/2000/svg" width="400" height="300" viewBox="0 0 400 300">
        <rect width="400" height="300" fill="#f0f0f0"/>
        <text x="200" y="140" font-family="Arial" font-size="14" fill="#666" text-anchor="middle">Image not found:</text>
        <text x="200" y="165" font-family="Arial" font-size="12" fill="#999" text-anchor="middle">{filename}</text>
    </svg>'''

    return send_file(
        BytesIO(svg.encode('utf-8')),
        mimetype='image/svg+xml'
    )


def update_package_xml(package_path: str, xml_content: str):
    """Update the XML content in a ZIP package."""
    package_path = Path(package_path)
    if not package_path.exists():
        return

    # Read existing ZIP
    temp_path = package_path.with_suffix('.tmp')

    with zipfile.ZipFile(package_path, 'r') as zf_in:
        with zipfile.ZipFile(temp_path, 'w', zipfile.ZIP_DEFLATED) as zf_out:
            for item in zf_in.namelist():
                if item == 'Book.xml':
                    # Replace with new XML
                    zf_out.writestr(item, xml_content.encode('utf-8'))
                else:
                    # Copy as-is
                    zf_out.writestr(item, zf_in.read(item))

    # Replace original
    temp_path.replace(package_path)


def load_from_zip(zip_path: Path) -> bool:
    """Load document from a RittDoc ZIP package."""
    try:
        temp_dir = tempfile.mkdtemp(prefix='rittdoc_editor_')
        editor_state['temp_dir'] = temp_dir

        with zipfile.ZipFile(zip_path, 'r') as zf:
            zf.extractall(temp_dir)

        # Find Book.xml
        xml_path = Path(temp_dir) / 'Book.xml'
        if not xml_path.exists():
            # Try to find any XML file
            xml_files = list(Path(temp_dir).glob('*.xml'))
            if xml_files:
                xml_path = xml_files[0]
            else:
                print(f"Error: No XML file found in package")
                return False

        # Load XML content
        with open(xml_path, 'r', encoding='utf-8') as f:
            editor_state['xml_content'] = f.read()

        editor_state['xml_path'] = str(xml_path)
        editor_state['package_path'] = str(zip_path)
        editor_state['multimedia_dir'] = temp_dir
        editor_state['title'] = zip_path.stem.replace('_rittdoc', '')

        return True
    except Exception as e:
        print(f"Error loading ZIP: {e}")
        return False


def load_from_xml(xml_path: Path, multimedia_dir: Optional[Path] = None) -> bool:
    """Load document from XML file."""
    try:
        with open(xml_path, 'r', encoding='utf-8') as f:
            editor_state['xml_content'] = f.read()

        editor_state['xml_path'] = str(xml_path)
        editor_state['title'] = xml_path.stem.replace('_docbook42', '')

        # Find multimedia directory
        if multimedia_dir and multimedia_dir.exists():
            editor_state['multimedia_dir'] = str(multimedia_dir)
        else:
            # Try to auto-detect
            parent = xml_path.parent
            stem = xml_path.stem.replace('_docbook42', '')

            candidates = [
                parent / f"{stem}_multimedia",
                parent / 'multimedia',
                parent / f"{stem}_MultiMedia",
                parent / 'MultiMedia',
            ]

            for candidate in candidates:
                if candidate.exists():
                    editor_state['multimedia_dir'] = str(candidate)
                    break

        return True
    except Exception as e:
        print(f"Error loading XML: {e}")
        return False


# ============================================================================
# CLI
# ============================================================================

def main():
    """Command-line interface."""
    parser = argparse.ArgumentParser(
        description="RittDoc Editor for DOCX to XML conversions",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s output/document_rittdoc.zip
  %(prog)s output/document_docbook42.xml
  %(prog)s output/document_docbook42.xml --multimedia output/document_multimedia
        """
    )

    parser.add_argument(
        'input',
        help='Path to ZIP package or XML file'
    )

    parser.add_argument(
        '--multimedia', '-m',
        help='Path to multimedia directory (auto-detected if not specified)'
    )

    parser.add_argument(
        '--port', '-p',
        type=int,
        default=8080,
        help='Port to run server on (default: 8080)'
    )

    parser.add_argument(
        '--no-browser',
        action='store_true',
        help="Don't open browser automatically"
    )

    args = parser.parse_args()

    input_path = Path(args.input)

    if not input_path.exists():
        print(f"Error: File not found: {input_path}")
        return 1

    # Load document
    if input_path.suffix.lower() == '.zip':
        if not load_from_zip(input_path):
            return 1
        print(f"Loaded package: {input_path}")
    else:
        multimedia_dir = Path(args.multimedia) if args.multimedia else None
        if not load_from_xml(input_path, multimedia_dir):
            return 1
        print(f"Loaded XML: {input_path}")

    if editor_state['multimedia_dir']:
        print(f"Multimedia: {editor_state['multimedia_dir']}")

    # Open browser
    url = f"http://localhost:{args.port}"
    if not args.no_browser:
        print(f"\nOpening editor in browser: {url}")
        webbrowser.open(url)
    else:
        print(f"\nEditor available at: {url}")

    # Run server
    print("Press Ctrl+C to stop the server\n")
    app.run(host='0.0.0.0', port=args.port, debug=False)

    return 0


if __name__ == '__main__':
    sys.exit(main())
