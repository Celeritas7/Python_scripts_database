#!/usr/bin/env python3
"""
Jupyter Notebook to Print-Ready HTML Converter

Converts .ipynb files to beautifully styled HTML optimized for printing to PDF.

Features:
- Syntax highlighting (GitHub, Friendly, Monokai themes)
- Table styling with zebra striping, bold headers/first column
- Repeating table headers on each printed page
- Aspect-ratio-based image sizing
- External images embedded as base64
- A4/A3 page size support
- Optional code cell hiding
- Optional PDF generation via weasyprint

Usage:
    python notebook_to_html.py input.ipynb [options]
    
    Or import and use programmatically:
    from notebook_to_html import convert_notebook
    convert_notebook("input.ipynb", "output.html")
"""

import json
import base64
import re
import argparse
from pathlib import Path
from typing import Optional, Literal
from io import BytesIO

# Check for optional dependencies
try:
    from bs4 import BeautifulSoup
    HAS_BS4 = True
except ImportError:
    HAS_BS4 = False

try:
    from pygments import highlight
    from pygments.lexers import PythonLexer, get_lexer_by_name
    from pygments.formatters import HtmlFormatter
    from pygments.styles import get_style_by_name
    HAS_PYGMENTS = True
except ImportError:
    HAS_PYGMENTS = False

try:
    import requests
    HAS_REQUESTS = True
except ImportError:
    HAS_REQUESTS = False

try:
    from weasyprint import HTML as WeasyHTML
    HAS_WEASYPRINT = True
except ImportError:
    HAS_WEASYPRINT = False


# =============================================================================
# CSS TEMPLATES
# =============================================================================

def get_google_fonts_css() -> str:
    """Google Fonts import for JetBrains Mono, Fira Code, Source Code Pro, Inter"""
    return """
@import url('https://fonts.googleapis.com/css2?family=Fira+Code:wght@400;700&family=Inter:wght@400;600;700&family=JetBrains+Mono:wght@400;700&family=Source+Code+Pro:wght@400;700&display=swap');
"""


def get_base_css(page_size: str = "A4", margins: str = "narrow") -> str:
    """Base CSS for page layout and general styling"""
    
    margin_value = "10mm" if margins == "narrow" else "20mm"
    
    return f"""
/* Page setup for printing */
@page {{
    size: {page_size};
    margin: {margin_value};
}}

/* Base styles */
* {{
    box-sizing: border-box;
}}

html, body {{
    font-family: 'Inter', 'Source Serif 4', 'Charter', system-ui, -apple-system, sans-serif;
    font-size: 12pt;
    line-height: 1.6;
    color: #333;
    margin: 0;
    padding: 0;
}}

/* Container */
.notebook-container {{
    max-width: 100%;
    margin: 0 auto;
    padding: 20px;
}}

@media print {{
    .notebook-container {{
        padding: 0;
    }}
}}

/* Headings */
h1, h2, h3, h4, h5, h6 {{
    font-family: 'Inter', system-ui, -apple-system, sans-serif;
    font-weight: 600;
    margin-top: 1.5em;
    margin-bottom: 0.5em;
    color: #2b579a;
}}

h1 {{ font-size: 2em; }}
h2 {{ font-size: 1.6em; }}
h3 {{ font-size: 1.35em; }}
h4 {{ font-size: 1.15em; }}

/* Paragraphs and text */
p {{
    margin: 0.8em 0;
}}

/* Links */
a {{
    color: #0066cc;
    text-decoration: none;
}}

a:hover {{
    text-decoration: underline;
}}

@media print {{
    a {{
        color: #333;
        text-decoration: none;
    }}
}}

/* Lists */
ul, ol {{
    margin: 0.5em 0;
    padding-left: 1.5em;
}}

li {{
    margin: 0.25em 0;
}}

/* Horizontal rule */
hr {{
    border: none;
    border-top: 1px solid #ddd;
    margin: 1.5em 0;
}}
"""


def get_table_css() -> str:
    """CSS for table styling - full width, black header, bold index, zebra stripes"""
    return """
/* Table styles */
table {
    border-collapse: collapse;
    width: 100% !important;
    margin: 0.5em auto !important;
    font-size: 10pt;
}

table, th, td {
    border: 1px solid #999;
}

th, td {
    padding: 2mm 4mm !important;
    vertical-align: middle;
    line-height: 1.3;
    white-space: nowrap;
    text-align: center !important;
}

/* Header row styling - BLACK background, WHITE text */
thead {
    display: table-header-group; /* Repeat on each printed page */
}

thead th {
    background-color: #000 !important;
    color: #fff !important;
    font-weight: 700;
    text-align: center !important;
}

/* Index column (first column) - bold text, centered */
tbody th {
    font-weight: 700;
    text-align: center !important;
}

/* Data cells - centered */
tbody td {
    text-align: center !important;
}

/* Zebra striping - alternating grey/white */
tbody tr:nth-child(odd) {
    background-color: #ffffff;
}

tbody tr:nth-child(even) {
    background-color: #e8e8e8;
}

/* First header cell (empty corner) - black like header */
thead th:first-child {
    background-color: #000 !important;
}

/* Center the output div containing tables */
.code-output {
    text-align: center;
}

.code-output table {
    display: inline-table;
}

/* Prevent row breaking across pages */
@media print {
    tr, td, th {
        page-break-inside: avoid;
        break-inside: avoid;
    }
    
    table {
        page-break-inside: auto;
    }
}

/* Remove Colab/Jupyter interactive elements */
.colab-df-container,
.colab-df-buttons,
.colab-df-convert {
    display: none !important;
}
"""


def get_code_css(syntax_theme: str = "github") -> str:
    """CSS for code cells with syntax highlighting"""
    
    # Base code styling
    base = """
/* Code cell container */
.code-cell {
    margin: 1em 0;
}

.code-cell.hidden {
    display: none;
}

/* Code input area */
.code-input {
    background-color: #f6f8fa;
    border: 1px solid #e1e4e8;
    border-radius: 4px;
    padding: 12px;
    overflow-x: auto;
    margin-bottom: 0.5em;
}

.code-input pre {
    margin: 0;
    white-space: pre-wrap;
    word-wrap: break-word;
    font-family: 'JetBrains Mono', 'Fira Code', 'Source Code Pro', monospace;
    font-size: 9.5pt;
    line-height: 1.4;
}

/* Code output area */
.code-output {
    margin: 0.5em 0 1em 0;
}

.code-output pre {
    background-color: #fff;
    border: 1px solid #e1e4e8;
    border-radius: 4px;
    padding: 12px;
    overflow-x: auto;
    white-space: pre-wrap;
    word-wrap: break-word;
    font-family: 'JetBrains Mono', 'Fira Code', 'Source Code Pro', monospace;
    font-size: 9.5pt;
    line-height: 1.4;
    margin: 0;
}

/* Output tables get their own styling (inherits from table CSS) */
.code-output table {
    margin: 0.5em 0;
}

/* Stderr output */
.output-stderr {
    background-color: #fff5f5;
    border-color: #ffcccc;
    color: #cc0000;
}

@media print {
    .code-input, .code-output pre {
        overflow: visible;
        white-space: pre-wrap;
        word-break: break-word;
    }
    
    .code-cell {
        page-break-inside: avoid;
        break-inside: avoid;
    }
}
"""
    return base


def get_image_css() -> str:
    """CSS for image sizing based on aspect ratio"""
    return """
/* Image styles */
.notebook-image {
    display: block;
    margin: 1em auto;
    max-width: 100%;
    height: auto;
}

/* Aspect ratio based sizing */
.img-landscape {
    max-width: 100%;
}

.img-square {
    max-width: 70%;
}

.img-portrait {
    max-width: 50%;
}

/* Markdown images */
.markdown-cell img {
    display: block;
    margin: 1em auto;
    max-width: 80%;
    height: auto;
}

/* Plot outputs */
.plot-output {
    text-align: center;
    margin: 1em 0;
}

.plot-output img {
    max-width: 100%;
    height: auto;
}

@media print {
    .notebook-image,
    .markdown-cell img,
    .plot-output img {
        page-break-inside: avoid;
        break-inside: avoid;
    }
    
    /* Ensure images don't overflow page */
    img {
        max-width: 100% !important;
        max-height: 90vh;
        object-fit: contain;
    }
}
"""


def get_markdown_css() -> str:
    """CSS for markdown cells"""
    return """
/* Markdown cell */
.markdown-cell {
    margin: 1em 0;
}

/* Inline code in markdown */
.markdown-cell code {
    background-color: #f0f0f0;
    padding: 2px 6px;
    border-radius: 3px;
    font-family: 'JetBrains Mono', 'Fira Code', 'Source Code Pro', monospace;
    font-size: 0.9em;
}

/* Code blocks in markdown */
.markdown-cell pre {
    background-color: #f6f8fa;
    border: 1px solid #e1e4e8;
    border-radius: 4px;
    padding: 12px;
    overflow-x: auto;
}

.markdown-cell pre code {
    background: none;
    padding: 0;
    font-size: 9.5pt;
}

/* Blockquotes */
.markdown-cell blockquote {
    border-left: 4px solid #ddd;
    margin: 1em 0;
    padding-left: 1em;
    color: #666;
}
"""


def get_syntax_highlight_css(theme: str = "github") -> str:
    """Get Pygments CSS for the specified theme"""
    if not HAS_PYGMENTS:
        return ""
    
    # Map friendly names to pygments style names
    theme_map = {
        "github": "github-dark" if theme == "github" else "default",
        "friendly": "friendly",
        "monokai": "monokai",
    }
    
    # For GitHub light theme, we use a custom minimal style
    if theme == "github":
        return """
/* GitHub-style syntax highlighting */
.highlight .c, .highlight .c1, .highlight .cm { color: #6a737d; font-style: italic; } /* Comments */
.highlight .k, .highlight .kn, .highlight .kd { color: #d73a49; } /* Keywords */
.highlight .s, .highlight .s1, .highlight .s2, .highlight .sa { color: #032f62; } /* Strings */
.highlight .n { color: #24292e; } /* Names */
.highlight .nf, .highlight .fm { color: #6f42c1; } /* Function names */
.highlight .nb { color: #005cc5; } /* Builtins */
.highlight .nc { color: #6f42c1; } /* Class names */
.highlight .nn { color: #24292e; } /* Module names */
.highlight .mi, .highlight .mf { color: #005cc5; } /* Numbers */
.highlight .o { color: #d73a49; } /* Operators */
.highlight .p { color: #24292e; } /* Punctuation */
.highlight .bp { color: #005cc5; } /* Built-in pseudo (self, etc) */
.highlight .nd { color: #6f42c1; } /* Decorators */
"""
    elif theme == "friendly":
        return """
/* Friendly syntax highlighting */
.highlight .c, .highlight .c1, .highlight .cm { color: #60a0b0; font-style: italic; } /* Comments */
.highlight .k, .highlight .kn, .highlight .kd { color: #007020; font-weight: bold; } /* Keywords */
.highlight .s, .highlight .s1, .highlight .s2, .highlight .sa { color: #4070a0; } /* Strings */
.highlight .n { color: #000; } /* Names */
.highlight .nf, .highlight .fm { color: #06287e; } /* Function names */
.highlight .nb { color: #007020; } /* Builtins */
.highlight .nc { color: #0e84b5; font-weight: bold; } /* Class names */
.highlight .nn { color: #0e84b5; font-weight: bold; } /* Module names */
.highlight .mi, .highlight .mf { color: #40a070; } /* Numbers */
.highlight .o { color: #666; } /* Operators */
.highlight .bp { color: #007020; } /* Built-in pseudo */
.highlight .nd { color: #555; font-weight: bold; } /* Decorators */
"""
    elif theme == "monokai":
        return """
/* Monokai syntax highlighting */
.code-input { background-color: #272822 !important; }
.code-input pre { color: #f8f8f2; }
.highlight .c, .highlight .c1, .highlight .cm { color: #75715e; font-style: italic; } /* Comments */
.highlight .k, .highlight .kn, .highlight .kd { color: #66d9ef; } /* Keywords */
.highlight .s, .highlight .s1, .highlight .s2, .highlight .sa { color: #e6db74; } /* Strings */
.highlight .n { color: #f8f8f2; } /* Names */
.highlight .nf, .highlight .fm { color: #a6e22e; } /* Function names */
.highlight .nb { color: #f8f8f2; } /* Builtins */
.highlight .nc { color: #a6e22e; font-weight: bold; } /* Class names */
.highlight .nn { color: #f8f8f2; } /* Module names */
.highlight .mi, .highlight .mf { color: #ae81ff; } /* Numbers */
.highlight .o { color: #f92672; } /* Operators */
.highlight .p { color: #f8f8f2; } /* Punctuation */
.highlight .bp { color: #f8f8f2; } /* Built-in pseudo */
.highlight .nd { color: #a6e22e; } /* Decorators */
"""
    
    return ""


# =============================================================================
# HTML CONVERSION FUNCTIONS
# =============================================================================

def highlight_code(code: str, language: str = "python") -> str:
    """Apply syntax highlighting to code"""
    if not HAS_PYGMENTS:
        return f'<pre><code>{escape_html(code)}</code></pre>'
    
    try:
        lexer = get_lexer_by_name(language, stripall=True)
    except:
        lexer = PythonLexer()
    
    formatter = HtmlFormatter(nowrap=True, cssclass="highlight")
    highlighted = highlight(code, lexer, formatter)
    return f'<pre><code class="highlight">{highlighted}</code></pre>'


def escape_html(text: str) -> str:
    """Escape HTML special characters"""
    return (text
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;"))


def convert_markdown_to_html(md_text: str) -> str:
    """Convert markdown to HTML (basic conversion)"""
    # This is a simplified markdown converter
    # For full markdown support, you'd want to use a library like markdown or mistune
    
    html = md_text
    
    # Headers
    html = re.sub(r'^######\s+(.+)$', r'<h6>\1</h6>', html, flags=re.MULTILINE)
    html = re.sub(r'^#####\s+(.+)$', r'<h5>\1</h5>', html, flags=re.MULTILINE)
    html = re.sub(r'^####\s+(.+)$', r'<h4>\1</h4>', html, flags=re.MULTILINE)
    html = re.sub(r'^###\s+(.+)$', r'<h3>\1</h3>', html, flags=re.MULTILINE)
    html = re.sub(r'^##\s+(.+)$', r'<h2>\1</h2>', html, flags=re.MULTILINE)
    html = re.sub(r'^#\s+(.+)$', r'<h1>\1</h1>', html, flags=re.MULTILINE)
    
    # Bold and italic
    html = re.sub(r'\*\*\*(.+?)\*\*\*', r'<strong><em>\1</em></strong>', html)
    html = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', html)
    html = re.sub(r'\*(.+?)\*', r'<em>\1</em>', html)
    
    # Inline code
    html = re.sub(r'`([^`]+)`', r'<code>\1</code>', html)
    
    # Links
    html = re.sub(r'\[([^\]]+)\]\(([^)]+)\)', r'<a href="\2">\1</a>', html)
    
    # Images (markdown style)
    html = re.sub(r'!\[([^\]]*)\]\(([^)]+)\)', r'<img src="\2" alt="\1">', html)
    
    # Horizontal rules
    html = re.sub(r'^---+$', r'<hr>', html, flags=re.MULTILINE)
    html = re.sub(r'^\*\*\*+$', r'<hr>', html, flags=re.MULTILINE)
    
    # Unordered lists (simple)
    lines = html.split('\n')
    in_list = False
    result = []
    for line in lines:
        if re.match(r'^\s*[\*\-]\s+', line):
            if not in_list:
                result.append('<ul>')
                in_list = True
            content = re.sub(r'^\s*[\*\-]\s+', '', line)
            result.append(f'<li>{content}</li>')
        else:
            if in_list:
                result.append('</ul>')
                in_list = False
            result.append(line)
    if in_list:
        result.append('</ul>')
    html = '\n'.join(result)
    
    # Ordered lists (simple)
    lines = html.split('\n')
    in_list = False
    result = []
    for line in lines:
        if re.match(r'^\s*\d+\.\s+', line):
            if not in_list:
                result.append('<ol>')
                in_list = True
            content = re.sub(r'^\s*\d+\.\s+', '', line)
            result.append(f'<li>{content}</li>')
        else:
            if in_list:
                result.append('</ol>')
                in_list = False
            result.append(line)
    if in_list:
        result.append('</ol>')
    html = '\n'.join(result)
    
    # Paragraphs (wrap loose text)
    # Split by double newlines for paragraphs
    parts = re.split(r'\n\n+', html)
    wrapped = []
    for part in parts:
        part = part.strip()
        if not part:
            continue
        # Don't wrap if already has block-level element
        if re.match(r'^<(h[1-6]|ul|ol|li|blockquote|hr|table|div|pre|img)', part):
            wrapped.append(part)
        else:
            wrapped.append(f'<p>{part}</p>')
    
    return '\n'.join(wrapped)


def get_image_dimensions_from_base64(data: str) -> tuple:
    """Try to get image dimensions from base64 PNG data"""
    try:
        import struct
        img_data = base64.b64decode(data)
        # PNG header check
        if img_data[:8] == b'\x89PNG\r\n\x1a\n':
            # IHDR chunk contains dimensions
            width = struct.unpack('>I', img_data[16:20])[0]
            height = struct.unpack('>I', img_data[20:24])[0]
            return width, height
    except:
        pass
    return None, None


def get_image_class(width: Optional[int], height: Optional[int]) -> str:
    """Determine CSS class based on aspect ratio"""
    if width is None or height is None:
        return "img-square"  # Default
    
    ratio = width / height
    if ratio > 1.5:
        return "img-landscape"
    elif ratio < 0.67:
        return "img-portrait"
    else:
        return "img-square"


def process_html_output(html_content: str) -> str:
    """Clean up HTML output from notebook (remove Colab widgets, etc.)"""
    if not HAS_BS4:
        # Basic cleanup without BeautifulSoup
        # Remove Colab-specific button elements
        html_content = re.sub(r'<button[^>]*class="[^"]*colab-df[^"]*"[^>]*>.*?</button>', '', html_content, flags=re.DOTALL)
        html_content = re.sub(r'<script[^>]*>.*?</script>', '', html_content, flags=re.DOTALL)
        html_content = re.sub(r'<style[^>]*>.*?</style>', '', html_content, flags=re.DOTALL)
        return html_content
    
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # First, extract and preserve tables
    tables = soup.find_all('table')
    
    # Remove Colab interactive elements (buttons, not containers with tables)
    # Only remove elements that don't contain tables
    for element in soup.find_all(class_=re.compile(r'colab-df-buttons|colab-df-convert')):
        element.decompose()
    
    # Remove script tags
    for script in soup.find_all('script'):
        script.decompose()
    
    # Remove inline styles that might conflict (but keep scoped styles in the HTML)
    for style in soup.find_all('style'):
        # Check if it's defining useful table styles
        style_content = style.string or ''
        if 'dataframe' not in style_content:
            style.decompose()
    
    # If there are tables, extract just the tables from any colab containers
    if tables:
        # Find table and return just the table HTML
        result_parts = []
        for table in soup.find_all('table'):
            result_parts.append(str(table))
        return '\n'.join(result_parts)
    
    return str(soup)


def embed_external_image(url: str) -> Optional[str]:
    """Download external image and convert to base64 data URI"""
    if not HAS_REQUESTS:
        return None
    
    try:
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            content_type = response.headers.get('content-type', 'image/png')
            if 'image' in content_type:
                b64 = base64.b64encode(response.content).decode('utf-8')
                return f"data:{content_type};base64,{b64}"
    except:
        pass
    return None


def convert_cell_to_html(cell: dict, show_code: bool = True, embed_images: bool = True, syntax_theme: str = "github") -> str:
    """Convert a single notebook cell to HTML"""
    cell_type = cell.get('cell_type', '')
    source = cell.get('source', [])
    if isinstance(source, list):
        source = ''.join(source)
    
    html_parts = []
    
    if cell_type == 'markdown':
        # Process markdown cell
        md_html = convert_markdown_to_html(source)
        
        # Embed external images if requested
        if embed_images and HAS_REQUESTS:
            # Find all image URLs
            img_pattern = r'<img[^>]+src="(https?://[^"]+)"'
            for match in re.finditer(img_pattern, md_html):
                url = match.group(1)
                embedded = embed_external_image(url)
                if embedded:
                    md_html = md_html.replace(f'src="{url}"', f'src="{embedded}"')
        
        html_parts.append(f'<div class="markdown-cell">{md_html}</div>')
    
    elif cell_type == 'code':
        # Code input
        if show_code and source.strip():
            highlighted = highlight_code(source.strip())
            html_parts.append(f'<div class="code-cell"><div class="code-input">{highlighted}</div>')
        else:
            html_parts.append('<div class="code-cell hidden">')
        
        # Code outputs
        outputs = cell.get('outputs', [])
        for output in outputs:
            output_type = output.get('output_type', '')
            
            if output_type == 'stream':
                # Text output (stdout/stderr)
                text = output.get('text', [])
                if isinstance(text, list):
                    text = ''.join(text)
                stream_name = output.get('name', 'stdout')
                css_class = 'output-stderr' if stream_name == 'stderr' else ''
                html_parts.append(f'<div class="code-output"><pre class="{css_class}">{escape_html(text)}</pre></div>')
            
            elif output_type in ('execute_result', 'display_data'):
                data = output.get('data', {})
                
                # Prefer HTML output (for DataFrames)
                if 'text/html' in data:
                    html_out = data['text/html']
                    if isinstance(html_out, list):
                        html_out = ''.join(html_out)
                    html_out = process_html_output(html_out)
                    html_parts.append(f'<div class="code-output">{html_out}</div>')
                
                # Image output (plots)
                elif 'image/png' in data:
                    img_data = data['image/png']
                    if isinstance(img_data, list):
                        img_data = ''.join(img_data)
                    
                    # Get dimensions for aspect-ratio class
                    width, height = get_image_dimensions_from_base64(img_data)
                    img_class = get_image_class(width, height)
                    
                    html_parts.append(f'<div class="plot-output"><img class="notebook-image {img_class}" src="data:image/png;base64,{img_data}"></div>')
                
                # Plain text fallback
                elif 'text/plain' in data:
                    text = data['text/plain']
                    if isinstance(text, list):
                        text = ''.join(text)
                    html_parts.append(f'<div class="code-output"><pre>{escape_html(text)}</pre></div>')
            
            elif output_type == 'error':
                # Error output
                traceback = output.get('traceback', [])
                if isinstance(traceback, list):
                    traceback = '\n'.join(traceback)
                # Remove ANSI color codes
                traceback = re.sub(r'\x1b\[[0-9;]*m', '', traceback)
                html_parts.append(f'<div class="code-output"><pre class="output-stderr">{escape_html(traceback)}</pre></div>')
        
        html_parts.append('</div>')  # Close code-cell
    
    return '\n'.join(html_parts)


def convert_notebook(
    input_path: str,
    output_path: Optional[str] = None,
    page_size: Literal["A4", "A3"] = "A4",
    margins: Literal["narrow", "normal"] = "narrow",
    show_code: bool = True,
    embed_images: bool = True,
    syntax_theme: Literal["github", "friendly", "monokai"] = "github",
    generate_pdf: bool = False,
) -> str:
    """
    Convert a Jupyter notebook to print-ready HTML.
    
    Args:
        input_path: Path to the .ipynb file
        output_path: Path for output HTML (default: same name with .html extension)
        page_size: Paper size for printing ("A4" or "A3")
        margins: Margin size ("narrow" ~10mm, "normal" ~20mm)
        show_code: Whether to show code cells (True) or hide them (False)
        embed_images: Whether to download and embed external images as base64
        syntax_theme: Code syntax highlighting theme ("github", "friendly", "monokai")
        generate_pdf: Whether to also generate PDF using weasyprint (if available)
    
    Returns:
        Path to the generated HTML file
    """
    
    # Read notebook
    input_path = Path(input_path)
    with open(input_path, 'r', encoding='utf-8') as f:
        notebook = json.load(f)
    
    # Determine output path
    if output_path is None:
        output_path = input_path.with_suffix('.html')
    output_path = Path(output_path)
    
    # Get notebook title from first H1 or filename
    title = input_path.stem
    cells = notebook.get('cells', [])
    for cell in cells:
        if cell.get('cell_type') == 'markdown':
            source = cell.get('source', [])
            if isinstance(source, list):
                source = ''.join(source)
            match = re.search(r'^#\s+(.+)$', source, re.MULTILINE)
            if match:
                title = match.group(1).strip()
                break
    
    # Build CSS
    all_css = '\n'.join([
        get_google_fonts_css(),
        get_base_css(page_size, margins),
        get_table_css(),
        get_code_css(syntax_theme),
        get_image_css(),
        get_markdown_css(),
        get_syntax_highlight_css(syntax_theme),
    ])
    
    # Convert cells
    body_parts = []
    for cell in cells:
        cell_html = convert_cell_to_html(cell, show_code, embed_images, syntax_theme)
        if cell_html:
            body_parts.append(cell_html)
    
    # Build full HTML document
    html_document = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{escape_html(title)}</title>
    <style>
{all_css}
    </style>
</head>
<body>
    <div class="notebook-container">
{chr(10).join(body_parts)}
    </div>
</body>
</html>
"""
    
    # Write HTML
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_document)
    
    print(f"✓ HTML generated: {output_path}")
    
    # Generate PDF if requested
    if generate_pdf:
        if HAS_WEASYPRINT:
            pdf_path = output_path.with_suffix('.pdf')
            try:
                WeasyHTML(filename=str(output_path)).write_pdf(str(pdf_path))
                print(f"✓ PDF generated: {pdf_path}")
            except Exception as e:
                print(f"✗ PDF generation failed: {e}")
        else:
            print("✗ PDF generation skipped: weasyprint not installed")
            print("  Install with: pip install weasyprint")
            print("  Or open the HTML in a browser and print to PDF")
    
    return str(output_path)


def batch_convert(
    input_dir: str,
    output_dir: Optional[str] = None,
    page_size: Literal["A4", "A3"] = "A4",
    margins: Literal["narrow", "normal"] = "narrow",
    show_code: bool = True,
    embed_images: bool = True,
    syntax_theme: Literal["github", "friendly", "monokai"] = "github",
    generate_pdf: bool = False,
) -> list:
    """
    Batch convert all .ipynb files in a directory to HTML.
    
    Args:
        input_dir: Directory containing .ipynb files
        output_dir: Directory for output HTML files (default: same as input)
        page_size: Paper size for printing ("A4" or "A3")
        margins: Margin size ("narrow" ~10mm, "normal" ~20mm)
        show_code: Whether to show code cells (True) or hide them (False)
        embed_images: Whether to download and embed external images as base64
        syntax_theme: Code syntax highlighting theme ("github", "friendly", "monokai")
        generate_pdf: Whether to also generate PDF using weasyprint (if available)
    
    Returns:
        List of generated HTML file paths
    """
    input_dir = Path(input_dir)
    
    if output_dir is None:
        output_dir = input_dir
    else:
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
    
    # Find all .ipynb files
    notebook_files = list(input_dir.glob("*.ipynb"))
    
    if not notebook_files:
        print(f"No .ipynb files found in {input_dir}")
        return []
    
    print(f"Found {len(notebook_files)} notebook(s) to convert...")
    print("-" * 50)
    
    converted_files = []
    for notebook_path in notebook_files:
        output_path = output_dir / notebook_path.with_suffix('.html').name
        try:
            result = convert_notebook(
                input_path=str(notebook_path),
                output_path=str(output_path),
                page_size=page_size,
                margins=margins,
                show_code=show_code,
                embed_images=embed_images,
                syntax_theme=syntax_theme,
                generate_pdf=generate_pdf,
            )
            converted_files.append(result)
        except Exception as e:
            print(f"✗ Failed to convert {notebook_path.name}: {e}")
    
    print("-" * 50)
    print(f"Converted {len(converted_files)} of {len(notebook_files)} notebooks")
    
    return converted_files


# =============================================================================
# COMMAND LINE INTERFACE
# =============================================================================

def main():
    parser = argparse.ArgumentParser(
        description='Convert Jupyter notebooks to print-ready HTML',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
    # Auto-convert all .ipynb files in script's folder (no arguments needed)
    python notebook_to_html.py
    
    # Single file conversion
    python notebook_to_html.py notebook.ipynb
    python notebook_to_html.py notebook.ipynb -o output.html --page-size A3
    python notebook_to_html.py notebook.ipynb --theme monokai --no-code
    python notebook_to_html.py notebook.ipynb --pdf
    
    # Batch conversion (all .ipynb files in a directory)
    python notebook_to_html.py --batch /path/to/notebooks/
    python notebook_to_html.py --batch /path/to/notebooks/ -o /path/to/output/
    python notebook_to_html.py --batch . --theme friendly
        """
    )
    
    parser.add_argument('input', nargs='?', help='Input .ipynb file (or directory with --batch)')
    parser.add_argument('-o', '--output', help='Output HTML file or directory (for --batch)')
    parser.add_argument('--batch', action='store_true',
                        help='Batch convert all .ipynb files in the input directory')
    parser.add_argument('--page-size', choices=['A4', 'A3'], default='A4',
                        help='Page size for printing (default: A4)')
    parser.add_argument('--margins', choices=['narrow', 'normal'], default='narrow',
                        help='Margin size (default: narrow)')
    parser.add_argument('--no-code', action='store_true',
                        help='Hide code cells, show only outputs')
    parser.add_argument('--no-embed', action='store_true',
                        help='Keep external image URLs instead of embedding')
    parser.add_argument('--theme', choices=['github', 'friendly', 'monokai'], default='github',
                        help='Syntax highlighting theme (default: github)')
    parser.add_argument('--pdf', action='store_true',
                        help='Also generate PDF (requires weasyprint)')
    
    args = parser.parse_args()
    
    # If no input provided, auto-batch convert all .ipynb in script's directory
    if not args.input:
        script_dir = Path(__file__).parent.resolve()
        work_dir = Path.cwd()
        if work_dir != script_dir:
            script_dir = work_dir

        print(f"No input specified. Auto-converting all .ipynb files in: {script_dir}")
        print("=" * 60)
        batch_convert(
            input_dir=str(script_dir),
            output_dir=args.output,
            page_size=args.page_size,
            margins=args.margins,
            show_code=not args.no_code,
            embed_images=not args.no_embed,
            syntax_theme=args.theme,
            generate_pdf=args.pdf,
        )
    elif args.batch:
        batch_convert(
            input_dir=args.input,
            output_dir=args.output,
            page_size=args.page_size,
            margins=args.margins,
            show_code=not args.no_code,
            embed_images=not args.no_embed,
            syntax_theme=args.theme,
            generate_pdf=args.pdf,
        )
    else:
        convert_notebook(
            input_path=args.input,
            output_path=args.output,
            page_size=args.page_size,
            margins=args.margins,
            show_code=not args.no_code,
            embed_images=not args.no_embed,
            syntax_theme=args.theme,
            generate_pdf=args.pdf,
        )


if __name__ == '__main__':
    main()
