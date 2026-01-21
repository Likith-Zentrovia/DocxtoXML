# DOCX to XML Conversion Pipeline (RittDoc)

A fast DOCX-to-DocBook XML conversion pipeline that produces RittDoc DTD-compliant output. This pipeline is designed to complement the PDF to XML pipeline, providing faster and more accurate conversion for DOCX documents since text is directly readable without OCR.

## Features

- **Fast Text Extraction**: Direct text extraction from DOCX without AI/OCR
- **Image Extraction**: Extracts embedded images with metadata
- **Table Extraction**: Preserves table structure including merged cells
- **Style Preservation**: Maintains headings, bold, italic, lists, etc.
- **RittDoc Compliance**: Outputs DocBook XML 4.2 compliant with RittDoc DTD
- **ZIP Packaging**: Creates RittDoc-compatible ZIP packages
- **REST API**: FastAPI-based API for integration
- **Compatible Output**: Same output format as the PDF to XML pipeline

## Quick Start

### Installation

```bash
# Clone the repository
git clone <repository-url>
cd docx-to-xml

# Install Python dependencies
pip install -r requirements.txt
```

### Basic Usage (CLI)

```bash
# Basic conversion
python docx_orchestrator.py document.docx --out ./output

# Skip ZIP package creation
python docx_orchestrator.py document.docx --out ./output --no-package

# Skip image extraction
python docx_orchestrator.py document.docx --out ./output --no-images

# Quiet mode
python docx_orchestrator.py document.docx --out ./output --quiet
```

### REST API Usage

```bash
# Start the API server
uvicorn api:app --host 0.0.0.0 --port 8000

# Or use Python directly
python api.py
```

API Endpoints:
- `POST /api/v1/convert` - Upload and convert a DOCX file
- `GET /api/v1/jobs/{job_id}` - Get job status
- `GET /api/v1/jobs` - List all jobs
- `GET /api/v1/dashboard` - Get conversion statistics
- `GET /api/v1/jobs/{job_id}/files` - List output files
- `GET /api/v1/jobs/{job_id}/files/{filename}` - Download a file

## Project Structure

```
docx-to-xml/
├── docx_orchestrator.py      # Main CLI entry point
├── docx_extractor.py         # DOCX content extraction
├── docbook_generator.py      # DocBook XML generation
├── package.py                # RittDoc ZIP packaging
├── api.py                    # FastAPI REST API
├── config.py                 # Configuration management
│
├── RITTDOCdtd/              # DTD schema files
│   └── v1.1/RittDocBook.dtd
│
├── requirements.txt         # Python dependencies
└── README.md               # This file
```

## Output Files

For `document.docx`, the pipeline produces:
- `document_docbook42.xml` - DocBook XML 4.2
- `document_rittdoc.zip` - RittDoc package containing:
  - `Book.xml` - Main XML file
  - `MultiMedia/` - Extracted images
  - `metadata.csv` - Image metadata
- `document_MultiMedia/` - Extracted images folder

## Configuration

### Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `DOCXTOXML_OUTPUT_DIR` | `./output` | Output directory |
| `DOCXTOXML_AI_ENABLED` | `false` | Enable optional AI enhancement |
| `DOCXTOXML_MODEL` | `claude-sonnet-4-20250514` | Claude model (if AI enabled) |
| `DOCXTOXML_DTD_PATH` | `RITTDOCdtd/v1.1/RittDocBook.dtd` | DTD file path |

### Configuration File

Create a `config.json` file:

```json
{
    "extraction": {
        "extract_images": true,
        "extract_tables": true,
        "preserve_formatting": true,
        "min_image_size": 50
    },
    "output": {
        "output_dir": "output",
        "create_rittdoc_zip": true,
        "include_toc": true
    },
    "ai": {
        "enabled": false,
        "model": "claude-sonnet-4-20250514"
    }
}
```

## Python Module Usage

```python
from docx_extractor import extract_docx
from docbook_generator import generate_docbook
from package import create_rittdoc_package

# Extract content from DOCX
content = extract_docx("document.docx")
print(f"Title: {content.title}")
print(f"Images: {len(content.images)}")
print(f"Tables: {len(content.tables)}")

# Generate DocBook XML
xml_content = generate_docbook(content, "output/document_docbook42.xml")

# Create RittDoc package
result = create_rittdoc_package(
    xml_content=xml_content,
    images=content.images,
    output_path="output/document_rittdoc.zip",
    content=content
)
print(f"Package created: {result.package_path}")
```

## API Integration

```python
import requests

# Upload and convert a DOCX file
with open("document.docx", "rb") as f:
    response = requests.post(
        "http://localhost:8000/api/v1/convert",
        files={"file": f},
        data={
            "extract_images": True,
            "extract_tables": True,
            "create_package": True
        }
    )
job = response.json()
job_id = job["job_id"]

# Poll for completion
import time
while True:
    status = requests.get(f"http://localhost:8000/api/v1/jobs/{job_id}").json()
    if status["status"] in ("completed", "failed"):
        break
    time.sleep(1)

# Download results
if status["status"] == "completed":
    files = requests.get(f"http://localhost:8000/api/v1/jobs/{job_id}/files").json()
    for file_info in files["files"]:
        content = requests.get(f"http://localhost:8000{file_info['download_url']}").content
        with open(file_info["name"], "wb") as f:
            f.write(content)
```

## Comparison with PDF Pipeline

| Feature | PDF Pipeline | DOCX Pipeline |
|---------|-------------|---------------|
| AI Required | Yes (for text extraction) | No (optional enhancement) |
| Speed | Slower (image processing) | Fast (direct text access) |
| Accuracy | Good with AI | Excellent (native format) |
| Image Quality | Re-extracted from PDF | Original quality |
| Table Detection | AI-based detection | Native structure |
| Output Format | DocBook XML 4.2 | DocBook XML 4.2 |
| RittDoc Package | Yes | Yes |

## Troubleshooting

### Common Issues

1. **"python-docx not installed"**
   ```bash
   pip install python-docx
   ```

2. **"lxml not found"**
   ```bash
   pip install lxml
   ```

3. **Images not extracting**
   - Check that images are embedded, not linked
   - Increase `min_image_size` if images are being filtered

4. **Tables not parsing correctly**
   - Complex merged cells may need manual adjustment
   - Nested tables are flattened

## License

[Add your license here]

## Contributing

[Add contribution guidelines here]
