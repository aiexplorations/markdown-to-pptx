# Markdown to PPTX Converter - Project Documentation

## Project Overview

This is a **Proof of Concept (POC)** implementation for converting Markdown files to PowerPoint presentations. The project consists of a single Python script that handles basic markdown parsing and PPTX generation.

## Architecture

### Core Components

1. **MarkdownToPPTX Class** (`markdown_to_pptx.py:25`)
   - Main converter class handling the conversion logic
   - Uses python-pptx library for PowerPoint generation

2. **Markdown Parser** (`markdown_to_pptx.py:32`)
   - Simple regex-based parsing for headers, bullets, and text
   - Converts markdown structure to slide data format

3. **Slide Generator** (`markdown_to_pptx.py:75`)
   - Creates PowerPoint slides from parsed data
   - Handles different slide layouts based on content type

## Technical Implementation

### Dependencies
- `python-pptx`: PowerPoint file generation
- `argparse`: Command-line interface
- `pathlib`: File path handling
- `re`: Regular expression matching for markdown parsing

### Markdown Processing Flow
1. Read markdown file content
2. Split content into lines
3. Parse headers (# ## ###) as slide titles
4. Parse bullet points (- *) as slide bullets  
5. Collect regular text as slide content
6. Generate slide data structure
7. Create PowerPoint slides using python-pptx

### Supported Elements
- **Headers**: `# ## ###` create new slides
- **Bullet Points**: `- *` convert to PowerPoint bullets
- **Regular Text**: Added as paragraph content

## POC Limitations

As this is a proof of concept, the following limitations exist:

### Parsing Limitations
- No support for nested bullet points
- No markdown tables support
- No image embedding
- No code block formatting
- No emphasis (bold/italic) processing
- No link processing

### Presentation Limitations
- Basic slide layouts only
- No theme customization
- Fixed font sizes and styles
- No master slide templates
- No slide transitions or animations

### Error Handling
- Minimal error handling implementation
- No validation of markdown syntax
- Limited file format checking

## Usage Patterns

### Command Line Interface
```bash
# Basic conversion
python markdown_to_pptx.py input.md

# With custom output and title
python markdown_to_pptx.py input.md -o output.pptx -t "Title" -s "Subtitle"
```

### Programmatic Usage
```python
converter = MarkdownToPPTX()
converter.convert("input.md", "output.pptx", "Title", "Subtitle")
```

## Development Notes

### Code Organization
- Single file implementation for POC simplicity
- Class-based design for potential extensibility
- Type hints included for maintainability
- Comprehensive docstrings for all methods

### Testing Requirements
- No formal tests included in POC
- Manual testing with sample markdown files recommended
- Validation against different markdown structures needed

### Future Enhancement Areas

#### Parsing Improvements
- Support for nested lists
- Table parsing and conversion
- Image embedding with file references
- Code block formatting with syntax highlighting
- Emphasis and link processing

#### Presentation Features
- Custom themes and templates
- Advanced slide layouts
- Chart generation from markdown tables
- Speaker notes from markdown comments
- Slide transitions and animations

#### Architecture Enhancements
- Separate parser and generator modules
- Plugin system for custom markdown extensions
- Configuration file support
- Batch processing capabilities
- Template system for consistent styling

## Installation and Setup

### Virtual Environment (Recommended)
```bash
python -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate
pip install python-pptx
```

### Direct Installation
```bash
pip install python-pptx
```

## File Structure
```
markdown-to-pptx/
├── markdown_to_pptx.py  # Main conversion script
├── README.md            # User documentation
└── CLAUDE.md           # This development documentation
```

## Performance Considerations

- Single-threaded processing suitable for POC
- Memory usage scales with markdown file size
- No optimization for large files
- PowerPoint generation is the performance bottleneck

## Security Considerations

- File path validation implemented
- No network operations
- Local file system access only
- No user input sanitization beyond path checking