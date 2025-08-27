# Markdown to PPTX Converter

A basic Python script to convert simple Markdown files into PowerPoint presentations.

**⚠️ This is a Proof of Concept (POC) implementation with significant limitations.**

## Features

- Convert Markdown headers to slide titles
- Support for bullet points and regular text
- Basic text formatting (bold and italic)
- Simple geometric shapes as visual elements
- Automatic slide generation based on headers
- Customizable presentation title and subtitle
- Command-line interface

## Supported Markdown Elements

- Headers (`# ## ###`) - Create new slides
- Bullet points (`- *`) - Convert to PowerPoint bullets
- Bold text (`**text**` or `__text__`) - Applied to PowerPoint text
- Italic text (`*text*` or `_text_`) - Applied to PowerPoint text
- Regular text - Added as slide content

## Installation

1. Clone or download this repository
2. Create a virtual environment (recommended):
   ```bash
   python -m venv .venv
   source .venv/bin/activate  # On Windows: .venv\Scripts\activate
   ```
3. Install required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Basic Usage
```bash
python markdown_to_pptx.py input.md
```

### With Custom Output File
```bash
python markdown_to_pptx.py input.md -o presentation.pptx
```

### With Title and Subtitle
```bash
python markdown_to_pptx.py input.md -t "My Presentation" -s "Created from Markdown"
```

### Command Line Options
- `input_file`: Path to the input Markdown file (required)
- `-o, --output`: Output PPTX file path (default: output.pptx)
- `-t, --title`: Presentation title for title slide
- `-s, --subtitle`: Presentation subtitle for title slide

## Example

Create a sample Markdown file (`example.md`):

```markdown
# Welcome Slide

This is the introduction to our **important** presentation.

# Main Points

- **First** important point with *emphasis*
- Second key insight
- **Final** consideration

# Conclusion

Thank you for your attention!

- Questions?
- Contact information
```

Convert to PowerPoint:
```bash
source .venv/bin/activate
python markdown_to_pptx.py example.md -t "Sample Presentation" -o sample.pptx
```

## Significant Limitations (POC)

This is a basic proof of concept with substantial limitations:

### Markdown Parsing Limitations
- Only supports basic headers, bullets, and simple bold/italic formatting
- No support for nested bullet points or complex list structures
- No support for links, images, tables, or code blocks
- No support for blockquotes, horizontal rules, or other markdown elements
- Text formatting parsing is rudimentary and may fail on complex cases

### Presentation Limitations
- Very basic slide layouts (title + content only)
- Simple geometric shapes only (circles, rectangles, triangles)
- No real image support or advanced graphics
- Fixed font sizes and limited styling options
- No theme customization or branding options
- No slide transitions, animations, or speaker notes
- No master slide templates or consistent styling

### Technical Limitations
- Basic error handling - may fail on malformed markdown
- No validation of markdown syntax
- Performance not optimized for large files
- Single-threaded processing only
- No configuration options or extensibility

**This tool is suitable only for very basic presentations and should not be used for professional or complex presentation needs.**

## Requirements

- Python 3.6+
- python-pptx library (installed via requirements.txt)

## License

This is a proof of concept project. Use at your own discretion.