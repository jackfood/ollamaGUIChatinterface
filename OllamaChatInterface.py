import ctypes
from ctypes import wintypes
import os
import threading
import subprocess
import time
import re
import json
import tkinter as tk
from tkinter import font as tkfont
import customtkinter as ctk
import requests
import uuid
import atexit
import signal
from datetime import datetime
from openai import OpenAI
from typing import Optional, Dict, List, Callable
from tkinter import filedialog, messagebox
import html

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SETTINGS_FILE = os.path.join(BASE_DIR, "settings.json")
SESSIONS_FILE = os.path.join(BASE_DIR, "sessions.json")

DEFAULT_OLLAMA_PATH = r"C:\Users\Peter-Susan\Desktop\ollama-ipex-llm-2.3.0b20250630-win\start-ollama.bat"
OLLAMA_API_BASE = "http://localhost:11434"
OLLAMA_CHAT_URL = f"{OLLAMA_API_BASE}/v1"
API_KEY = "ollama"


def copy_html_to_clipboard(html_content: str, plain_text: str) -> bool:
    try:
        # 1. Setup Windows API types (Crucial for 64-bit Python)
        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32

        # Define argtypes and restypes to prevent pointer truncation
        user32.RegisterClipboardFormatW.argtypes = [wintypes.LPCWSTR]
        user32.RegisterClipboardFormatW.restype = wintypes.UINT
        user32.OpenClipboard.argtypes = [wintypes.HWND]
        user32.OpenClipboard.restype = wintypes.BOOL
        user32.EmptyClipboard.argtypes = []
        user32.EmptyClipboard.restype = wintypes.BOOL
        user32.SetClipboardData.argtypes = [wintypes.UINT, wintypes.HANDLE]
        user32.SetClipboardData.restype = wintypes.HANDLE
        user32.CloseClipboard.argtypes = []
        user32.CloseClipboard.restype = wintypes.BOOL

        kernel32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
        kernel32.GlobalAlloc.restype = wintypes.HGLOBAL
        kernel32.GlobalLock.argtypes = [wintypes.HGLOBAL]
        kernel32.GlobalLock.restype = ctypes.c_void_p
        kernel32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]
        kernel32.GlobalUnlock.restype = wintypes.BOOL
        
        # Constants
        GMEM_MOVEABLE = 0x0002
        GMEM_ZEROINIT = 0x0040
        CF_UNICODETEXT = 13
        
        # 2. Prepare HTML Format
        # Microsoft Office requires a specific header with byte offsets
        CF_HTML = user32.RegisterClipboardFormatW("HTML Format")
        
        header_template = (
            "Version:0.9\r\n"
            "StartHTML:{:08d}\r\n"
            "EndHTML:{:08d}\r\n"
            "StartFragment:{:08d}\r\n"
            "EndFragment:{:08d}\r\n"
        )
        html_prefix = (
            "<!DOCTYPE html>\r\n"
            "<html>\r\n"
            "<body>\r\n"
            "<!--StartFragment-->"
        )
        html_suffix = "<!--EndFragment-->\r\n</body>\r\n</html>"
        
        # Encode content to UTF-8
        content_bytes = html_content.encode('utf-8')
        prefix_bytes = html_prefix.encode('utf-8')
        suffix_bytes = html_suffix.encode('utf-8')
        
        # Calculate dummy header length to determine offsets
        dummy_header = header_template.format(0, 0, 0, 0).encode('utf-8')
        header_len = len(dummy_header)
        
        start_html = header_len
        start_fragment = start_html + len(prefix_bytes)
        end_fragment = start_fragment + len(content_bytes)
        end_html = end_fragment + len(suffix_bytes)
        
        # Create final header and payload
        final_header = header_template.format(start_html, end_html, start_fragment, end_fragment).encode('utf-8')
        final_payload = final_header + prefix_bytes + content_bytes + suffix_bytes + b'\x00'

        # 3. Open Clipboard
        if not user32.OpenClipboard(None):
            return False
            
        try:
            user32.EmptyClipboard()

            # 4. Write Plain Text (CF_UNICODETEXT)
            text_bytes = plain_text.encode('utf-16le') + b'\x00\x00'
            h_text = kernel32.GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, len(text_bytes))
            if h_text:
                p_text = kernel32.GlobalLock(h_text)
                if p_text:
                    ctypes.memmove(p_text, text_bytes, len(text_bytes))
                    kernel32.GlobalUnlock(h_text)
                    user32.SetClipboardData(CF_UNICODETEXT, h_text)

            # 5. Write HTML (CF_HTML)
            h_html = kernel32.GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, len(final_payload))
            if h_html:
                p_html = kernel32.GlobalLock(h_html)
                if p_html:
                    ctypes.memmove(p_html, final_payload, len(final_payload))
                    kernel32.GlobalUnlock(h_html)
                    user32.SetClipboardData(CF_HTML, h_html)
                    
        finally:
            user32.CloseClipboard()
            
        return True
    except Exception as e:
        print(f"Clipboard error: {e}")
        return False

def markdown_to_html(text: str) -> str:
    # Clean thinking tags
    text = re.sub(r'<think>.*?</think>', '', text, flags=re.DOTALL)
    text = re.sub(r'<think>.*$', '', text, flags=re.DOTALL)
    text = re.sub(r'</think>', '', text)
    text = re.sub(r'<think>', '', text)
    
    lines = text.split('\n')
    html_parts = []
    i = 0
    in_list = False
    list_type = None
    
    while i < len(lines):
        line = lines[i].rstrip()
        
        # Code Blocks
        if line.strip().startswith('```'):
            if in_list:
                html_parts.append(f"</{list_type}>")
                in_list = False
            i += 1
            code_lines = []
            while i < len(lines) and not lines[i].strip().startswith('```'):
                code_lines.append(html.escape(lines[i]))
                i += 1
            code_content = '<br>'.join(code_lines)
            # Word friendly code block style
            html_parts.append(
                f'<div style="background-color:#f0f0f0; padding:10px; border:1px solid #ccc; '
                f'font-family:Consolas,monospace; font-size:10pt; white-space:pre-wrap;">'
                f'{code_content}</div>'
            )
            i += 1
            continue
            
        # Tables
        if line.strip().startswith('|') and line.strip().endswith('|'):
            if in_list:
                html_parts.append(f"</{list_type}>")
                in_list = False
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                table_lines.append(lines[i])
                i += 1
            html_parts.append(parse_table_to_html(table_lines))
            continue
            
        # Horizontal Rules
        if line.strip() in ['---', '***', '___']:
            if in_list:
                html_parts.append(f"</{list_type}>")
                in_list = False
            html_parts.append('<hr>')
            i += 1
            continue
            
        # Headers
        header_match = re.match(r'^(#{1,6})\s+(.+)$', line)
        if header_match:
            if in_list:
                html_parts.append(f"</{list_type}>")
                in_list = False
            level = len(header_match.group(1))
            content = process_inline_formatting_html(header_match.group(2))
            html_parts.append(f'<h{level}>{content}</h{level}>')
            i += 1
            continue
            
        # Unordered Lists
        bullet_match = re.match(r'^(\s*)([-*+])\s+(.+)$', line)
        if bullet_match:
            content = process_inline_formatting_html(bullet_match.group(3))
            if not in_list or list_type != "ul":
                if in_list:
                    html_parts.append(f"</{list_type}>")
                html_parts.append("<ul>")
                in_list = True
                list_type = "ul"
            html_parts.append(f"<li>{content}</li>")
            i += 1
            continue
            
        # Ordered Lists
        num_match = re.match(r'^(\s*)(\d+\.)\s+(.+)$', line)
        if num_match:
            content = process_inline_formatting_html(num_match.group(3))
            if not in_list or list_type != "ol":
                if in_list:
                    html_parts.append(f"</{list_type}>")
                html_parts.append("<ol>")
                in_list = True
                list_type = "ol"
            html_parts.append(f"<li>{content}</li>")
            i += 1
            continue
            
        # Empty Lines
        if not line.strip():
            if in_list:
                html_parts.append(f"</{list_type}>")
                in_list = False
            i += 1
            continue
            
        # Close list if regular text
        if in_list:
            html_parts.append(f"</{list_type}>")
            in_list = False
            
        content = process_inline_formatting_html(line)
        html_parts.append(f"<p>{content}</p>")
        i += 1
        
    if in_list:
        html_parts.append(f"</{list_type}>")
        
    return ''.join(html_parts)

def parse_table_to_html(lines: List[str]) -> str:
    rows = []
    for line in lines:
        stripped = line.strip()
        if stripped.startswith('|') and stripped.endswith('|'):
            cells = [c.strip() for c in stripped.split('|')]
            cells = cells[1:-1]
            if cells and all(re.match(r'^[\-:\s]+$', cell) and '-' in cell for cell in cells):
                continue
            rows.append(cells)
    if not rows:
        return ""
        
    # Inline CSS is critical for Word/PowerPoint
    html_parts = ['<table style="border-collapse: collapse; width: 100%; border: 1px solid black;">']
    for idx, row in enumerate(rows):
        html_parts.append("<tr>")
        is_header = idx == 0
        tag = "th" if is_header else "td"
        bg_style = "background-color: #d3d3d3;" if is_header else ""
        font_style = "font-weight: bold;" if is_header else ""
        
        for cell in row:
            content = process_inline_formatting_html(cell)
            html_parts.append(
                f'<{tag} style="border: 1px solid black; padding: 5px; {bg_style} {font_style}">'
                f'{content}</{tag}>'
            )
        html_parts.append("</tr>")
    html_parts.append("</table>")
    return ''.join(html_parts)

def parse_table_to_html(lines: List[str]) -> str:
    rows = []
    for line in lines:
        stripped = line.strip()
        if stripped.startswith('|') and stripped.endswith('|'):
            cells = [c.strip() for c in stripped.split('|')]
            cells = cells[1:-1]
            if cells and all(re.match(r'^[\-:\s]+$', cell) and '-' in cell for cell in cells):
                continue
            rows.append(cells)
    if not rows:
        return ""
        
    html_parts = ['<table style="border-collapse: collapse; width: 100%; border: 1px solid black;">']
    for idx, row in enumerate(rows):
        html_parts.append("<tr>")
        is_header = idx == 0
        tag = "th" if is_header else "td"
        bg_style = "background-color: #f2f2f2;" if is_header else ""
        font_style = "font-weight: bold;" if is_header else ""
        
        for cell in row:
            content = process_inline_formatting_html(cell)
            html_parts.append(
                f'<{tag} style="border: 1px solid black; padding: 8px; {bg_style} {font_style}">'
                f'{content}</{tag}>'
            )
        html_parts.append("</tr>")
    html_parts.append("</table>")
    return ''.join(html_parts)

def process_inline_formatting_html(text: str) -> str:
    result = html.escape(text)
    result = re.sub(r'\*\*\*(.+?)\*\*\*', r'<strong><em>\1</em></strong>', result)
    result = re.sub(r'___(.+?)___', r'<strong><em>\1</em></strong>', result)
    def handle_bold(m):
        inner = m.group(1)
        inner = re.sub(r'\*([^*]+)\*', r'<em>\1</em>', inner)
        inner = re.sub(r'_([^_]+)_', r'<em>\1</em>', inner)
        return f'<strong>{inner}</strong>'
    result = re.sub(r'\*\*(.+?)\*\*', handle_bold, result)
    result = re.sub(r'__(.+?)__', handle_bold, result)
    result = re.sub(r'(?<!\*)\*([^*\n]+?)\*(?!\*)', r'<em>\1</em>', result)
    result = re.sub(r'(?<!_)_([^_\n]+?)_(?!_)', r'<em>\1</em>', result)
    result = re.sub(r'~~(.+?)~~', r'<del>\1</del>', result)
    result = re.sub(r'`([^`]+)`', r'<code>\1</code>', result)
    return result

def parse_table_to_html(lines: List[str]) -> str:
    rows = []
    for line in lines:
        stripped = line.strip()
        if stripped.startswith('|') and stripped.endswith('|'):
            cells = [c.strip() for c in stripped.split('|')]
            cells = cells[1:-1]
            if cells and all(re.match(r'^[\-:\s]+$', cell) and '-' in cell for cell in cells):
                continue
            rows.append(cells)
    if not rows:
        return ""
    html_parts = ['<table border="1" cellpadding="5" cellspacing="0">']
    for idx, row in enumerate(rows):
        html_parts.append("<tr>")
        tag = "th" if idx == 0 else "td"
        for cell in row:
            content = process_inline_formatting_html(cell)
            html_parts.append(f'<{tag}>{content}</{tag}>')
        html_parts.append("</tr>")
    html_parts.append("</table>")
    return ''.join(html_parts)

def has_markdown_formatting(text: str) -> bool:
    patterns = [
        r'\*\*\*.+?\*\*\*',
        r'___.+?___',
        r'\*\*.+?\*\*',
        r'__.+?__',
        r'(?<![*\\])\*[^*\n]+?\*(?!\*)',
        r'(?<![_\\])_[^_\n]+?_(?!_)',
        r'~~.+?~~',
        r'<u>.+?</u>',
        r'`.+?`',
        r'^#{1,6}\s+',
        r'^\s*[-*+]\s+',
        r'^\s*\d+\.\s+',
        r'```',
        r'^\|.+\|$',
    ]
    for pattern in patterns:
        if re.search(pattern, text, re.MULTILINE):
            return True
    return False


def strip_markdown(text: str) -> str:
    text = re.sub(r'<think>.*?</think>', '', text, flags=re.DOTALL)
    text = re.sub(r'<think>.*$', '', text, flags=re.DOTALL)
    text = re.sub(r'</think>', '', text)
    text = re.sub(r'<think>', '', text)
    return text


class Theme:
    DARK = {
        "bg_primary": "#0d1117",
        "bg_secondary": "#161b22",
        "bg_tertiary": "#21262d",
        "bg_hover": "#30363d",
        "accent": "#238636",
        "accent_hover": "#2ea043",
        "accent_blue": "#1f6feb",
        "accent_purple": "#8957e5",
        "text_primary": "#f0f6fc",
        "text_secondary": "#8b949e",
        "text_muted": "#6e7681",
        "border": "#30363d",
        "user_bubble": "#1f3a5f",
        "ai_bubble": "#21262d",
        "success": "#238636",
        "warning": "#d29922",
        "error": "#f85149",
        "code_bg": "#161b22",
        "table_header": "#2d333b",
        "table_border": "#444c56",
    }
    LIGHT = {
        "bg_primary": "#ffffff",
        "bg_secondary": "#f6f8fa",
        "bg_tertiary": "#eaeef2",
        "bg_hover": "#d0d7de",
        "accent": "#1a7f37",
        "accent_hover": "#2da44e",
        "accent_blue": "#0969da",
        "accent_purple": "#8250df",
        "text_primary": "#1f2328",
        "text_secondary": "#656d76",
        "text_muted": "#8c959f",
        "border": "#d0d7de",
        "user_bubble": "#ddf4ff",
        "ai_bubble": "#f6f8fa",
        "success": "#1a7f37",
        "warning": "#9a6700",
        "error": "#cf222e",
        "code_bg": "#f6f8fa",
        "table_header": "#f6f8fa",
        "table_border": "#d0d7de",
    }

    @classmethod
    def get(cls) -> Dict[str, str]:
        return cls.DARK if ctk.get_appearance_mode() == "Dark" else cls.LIGHT


class MarkdownParser:
    @staticmethod
    def clean_text(text: str) -> str:
        text = re.sub(r'<think>.*?</think>', '', text, flags=re.DOTALL)
        text = re.sub(r'<think>.*$', '', text, flags=re.DOTALL)
        text = re.sub(r'</think>', '', text)
        text = re.sub(r'<think>', '', text)
        return text.strip()

    @staticmethod
    def parse_table_rows(lines: List[str]) -> List[List[str]]:
        rows = []
        for line in lines:
            # More robust separator detection - matches |---|, |:---|, |---:|, |:---:|
            # Check if all cells contain only dashes, colons, and spaces
            stripped = line.strip()
            if stripped.startswith('|') and stripped.endswith('|'):
                # Extract cells and check if it's a separator row
                cells = [c.strip() for c in stripped.split('|')]
                cells = cells[1:-1]  # Remove empty first/last from split
                if cells and all(re.match(r'^[:\-\s]+$', cell) and '-' in cell for cell in cells):
                    continue  # Skip separator row
            # Regular row processing
            cells = [c.strip() for c in line.strip().split('|')]
            if len(cells) > 2:
                rows.append(cells[1:-1])
        return rows

class InlineFormatter:
    PATTERNS = [
        (re.compile(r'\*\*_(.+?)_\*\*'), 'underline'),
        (re.compile(r'__\*(.+?)\*__'), 'underline'),
        (re.compile(r'\*\*\*(.+?)\*\*\*'), 'bold_italic'),
        (re.compile(r'___(.+?)___'), 'bold_italic'),
        (re.compile(r'\*\*(.+?)\*\*'), 'bold'),
        (re.compile(r'__(.+?)__'), 'bold'),
        (re.compile(r'(?<!\*)\*([^*]+)\*(?!\*)'), 'italic'),
        (re.compile(r'(?<!_)_([^_]+)_(?!_)'), 'italic'),
        (re.compile(r'~~(.+?)~~'), 'strikethrough'),
        (re.compile(r'<u>(.+?)</u>'), 'underline'),
        (re.compile(r'`([^`]+)`'), 'inline_code'),
    ]
    @staticmethod
    def render_line(text_widget: tk.Text, text: str, base_tags: tuple = ()):
        if not text:
            return
        segments = InlineFormatter._parse_segments(text)
        for segment_text, segment_tags in segments:
            combined_tags = base_tags + segment_tags if segment_tags else base_tags
            if not combined_tags:
                combined_tags = ("normal",)
            text_widget.insert(tk.END, segment_text, combined_tags)
    @staticmethod
    def _parse_segments(text: str) -> List[tuple]:
        segments = []
        current_pos = 0
        events = []
        for pattern, tag in InlineFormatter.PATTERNS:
            for match in pattern.finditer(text):
                events.append((match.start(), match.end(), match.group(1), tag, match.group(0)))
        events.sort(key=lambda x: (x[0], -x[1]))
        used_ranges = []
        filtered_events = []
        for start, end, content, tag, full_match in events:
            overlaps = False
            for used_start, used_end in used_ranges:
                if start < used_end and end > used_start:
                    overlaps = True
                    break
            if not overlaps:
                filtered_events.append((start, end, content, tag, full_match))
                used_ranges.append((start, end))
        filtered_events.sort(key=lambda x: x[0])
        for start, end, content, tag, full_match in filtered_events:
            if start > current_pos:
                segments.append((text[current_pos:start], ()))
            if tag == 'bold':
                nested_segments = InlineFormatter._parse_nested_in_bold(content)
                for nested_text, nested_tags in nested_segments:
                    if nested_tags:
                        combined = ('bold',) + nested_tags
                        segments.append((nested_text, combined))
                    else:
                        segments.append((nested_text, ('bold',)))
            else:
                segments.append((content, (tag,)))
            current_pos = end
        if current_pos < len(text):
            segments.append((text[current_pos:], ()))
        return segments
    @staticmethod
    def _parse_nested_in_bold(text: str) -> List[tuple]:
        segments = []
        italic_pattern = re.compile(r'\*([^*]+)\*|_([^_]+)_')
        current_pos = 0
        for match in italic_pattern.finditer(text):
            if match.start() > current_pos:
                segments.append((text[current_pos:match.start()], ()))
            content = match.group(1) or match.group(2)
            segments.append((content, ('italic',)))
            current_pos = match.end()
        if current_pos < len(text):
            segments.append((text[current_pos:], ()))
        return segments if segments else [(text, ())]

class CodeBlockWidget(tk.Frame):
    def __init__(self, master, theme: Dict, code: str, language: str = ""):
        super().__init__(master, bg=theme["code_bg"], highlightbackground=theme["border"], highlightthickness=1)
        self.theme = theme
        self.code = code
        self.columnconfigure(0, weight=1)
        header = tk.Frame(self, bg=theme["bg_tertiary"])
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)
        tk.Label(
            header, text=language if language else "code", font=("Consolas", 9),
            fg=theme["text_muted"], bg=theme["bg_tertiary"], anchor="w", padx=8
        ).grid(row=0, column=0, sticky="w", pady=2)
        self.copy_btn = tk.Button(
            header, text="Copy", font=("Segoe UI", 8), bg=theme["accent"], fg="white",
            bd=0, padx=8, pady=1, cursor="hand2", activebackground=theme["accent_hover"],
            command=self._copy_code
        )
        self.copy_btn.grid(row=0, column=1, sticky="e", padx=4, pady=2)
        line_count = code.count('\n') + 1
        display_height = min(max(line_count, 1), 20)
        self.text_widget = tk.Text(
            self, wrap=tk.NONE, font=("Consolas", 10), bg=theme["code_bg"], fg=theme["text_primary"],
            relief=tk.FLAT, bd=0, padx=8, pady=6, height=display_height, cursor="arrow", highlightthickness=0
        )
        self.text_widget.grid(row=1, column=0, sticky="ew")
        self.text_widget.insert("1.0", code)
        self.text_widget.configure(state=tk.DISABLED)

    def _copy_code(self):
        self.clipboard_clear()
        self.clipboard_append(self.code)
        self.copy_btn.configure(text="Copied!")
        self.after(1500, lambda: self.copy_btn.configure(text="Copy"))


class TableWidget(tk.Frame):
    def __init__(self, master, theme: Dict, rows: List[List[str]]):
        super().__init__(master, bg=theme["table_border"])
        self.theme = theme
        if not rows:
            return
        num_cols = max(len(r) for r in rows) if rows else 0
        if num_cols == 0:
            return
        for col in range(num_cols):
            self.columnconfigure(col, weight=1)
        for row_idx, row in enumerate(rows):
            is_header = row_idx == 0
            for col_idx in range(num_cols):
                cell_text = row[col_idx] if col_idx < len(row) else ""
                bg_color = theme["table_header"] if is_header else theme["bg_tertiary"]
                font_style = ("Segoe UI", 10, "bold") if is_header else ("Segoe UI", 10)
                cell_frame = tk.Frame(self, bg=bg_color, highlightbackground=theme["table_border"], highlightthickness=1)
                cell_frame.grid(row=row_idx, column=col_idx, sticky="nsew", padx=0, pady=0)
                cell_label = tk.Label(
                    cell_frame, text=cell_text, font=font_style, bg=bg_color, fg=theme["text_primary"],
                    padx=10, pady=6, anchor="w", justify="left"
                )
                cell_label.pack(fill="both", expand=True)


class MessageRenderer(tk.Frame):
    RENDER_THRESHOLD = 30
    RERENDER_INTERVAL = 100  # Re-render every 100 tokens during streaming

    def __init__(self, master, theme: Dict, role: str):
        bubble_bg = theme["user_bubble"] if role == "user" else theme["ai_bubble"]
        super().__init__(master, bg=bubble_bg)
        self.theme = theme
        self.role = role
        self.bubble_bg = bubble_bg
        self._raw_content = ""
        self._is_streaming = False
        self._token_count = 0
        self._last_render_token_count = 0  # Track when we last rendered
        self._last_rendered_content = ""
        self._formatted_mode = False
        self.content_frame = tk.Frame(self, bg=bubble_bg)
        self.content_frame.pack(fill="both", expand=True, padx=12, pady=10)
        self._widgets = []
        self._plain_text_widget = None

    def _create_text_widget(self, parent) -> tk.Text:
        text = tk.Text(
            parent, wrap=tk.WORD, relief=tk.FLAT, bd=0, bg=self.bubble_bg, fg=self.theme["text_primary"],
            font=("Segoe UI", 11), cursor="arrow", padx=0, pady=0, highlightthickness=0, height=1
        )
        self._configure_tags(text)
        return text

    def _configure_tags(self, text_widget: tk.Text):
        theme = self.theme
        text_widget.tag_configure("normal", font=("Segoe UI", 11))
        text_widget.tag_configure("bold", font=("Segoe UI", 11, "bold"))
        text_widget.tag_configure("italic", font=("Segoe UI", 11, "italic"))
        text_widget.tag_configure("bold_italic", font=("Segoe UI", 11, "bold", "italic"))
        text_widget.tag_configure("underline", font=("Segoe UI", 11, "underline"))
        text_widget.tag_configure("strikethrough", font=("Segoe UI", 11, "overstrike"))
        text_widget.tag_configure("inline_code", font=("Consolas", 10), background=theme["code_bg"])
        for level in range(1, 7):
            sizes = {1: 18, 2: 16, 3: 14, 4: 13, 5: 12, 6: 11}
            text_widget.tag_configure(f"h{level}", font=("Segoe UI", sizes[level], "bold"), foreground=theme["accent_blue"])
        text_widget.tag_configure("bullet", lmargin1=15, lmargin2=30)
        text_widget.tag_configure("bullet1", lmargin1=35, lmargin2=50)
        text_widget.tag_configure("bullet2", lmargin1=55, lmargin2=70)
        text_widget.tag_configure("numbered", lmargin1=15, lmargin2=30)

    def _clear_widgets(self):
        for w in self._widgets:
            try:
                w.destroy()
            except:
                pass
        self._widgets.clear()
        self._plain_text_widget = None

    def _auto_height(self, text_widget: tk.Text):
        text_widget.update_idletasks()
        content = text_widget.get("1.0", "end-1c")
        line_count = content.count('\n') + 1
        text_widget.configure(height=max(1, line_count))

    def set_streaming(self, streaming: bool):
        was_streaming = self._is_streaming
        self._is_streaming = streaming
        if was_streaming and not streaming:
            # Final render when streaming ends
            self._token_count = 0
            self._last_render_token_count = 0
            self._render_formatted(self._raw_content)
            self._formatted_mode = True

    def update_content(self, content: str):
        if content == self._raw_content:
            return
        old_len = len(self._raw_content)
        self._raw_content = content
        if self._is_streaming:
            new_tokens = len(content) - old_len
            self._token_count += max(1, new_tokens // 4)
            
            # Check if we should switch to formatted mode
            if not self._formatted_mode and self._token_count >= self.RENDER_THRESHOLD:
                if has_markdown_formatting(content):
                    self._formatted_mode = True
                    self._render_formatted(content)
                    self._last_rendered_content = content
                    self._last_render_token_count = self._token_count
                else:
                    self._update_plain_text(content)
            elif self._formatted_mode:
                # Check if it's time to re-render (every RERENDER_INTERVAL tokens)
                tokens_since_last_render = self._token_count - self._last_render_token_count
                if tokens_since_last_render >= self.RERENDER_INTERVAL:
                    self._render_formatted(content)
                    self._last_rendered_content = content
                    self._last_render_token_count = self._token_count
                else:
                    # Incremental update between full re-renders
                    self._incremental_update(content)
            else:
                self._update_plain_text(content)
        else:
            # Not streaming - always render formatted
            self._render_formatted(content)
            self._formatted_mode = True

    def _incremental_update(self, content: str):
        new_content = content[len(self._last_rendered_content):]
        if not new_content:
            return
        has_special = bool(re.search(r'```|\n#{1,6}\s|\n\s*[-*+]\s|\n\s*\d+\.\s|\n\|', new_content))
        if has_special:
            self._render_formatted(content)
            self._last_rendered_content = content
            self._last_render_token_count = self._token_count
            return
        if self._widgets and isinstance(self._widgets[-1], tk.Text):
            last_text = self._widgets[-1]
            last_text.configure(state=tk.NORMAL)
            last_text.insert(tk.END, new_content, ("normal",))
            last_text.configure(state=tk.DISABLED)
            self._auto_height(last_text)
            self._last_rendered_content = content
        else:
            self._render_formatted(content)
            self._last_rendered_content = content
            self._last_render_token_count = self._token_count

    def _update_plain_text(self, text: str):
        cleaned = strip_markdown(text)
        if self._plain_text_widget is None:
            self._clear_widgets()
            self._plain_text_widget = self._create_text_widget(self.content_frame)
            self._plain_text_widget.pack(fill="x", expand=True)
            self._widgets.append(self._plain_text_widget)
        self._plain_text_widget.configure(state=tk.NORMAL)
        self._plain_text_widget.delete("1.0", tk.END)
        self._plain_text_widget.insert("1.0", cleaned, ("normal",))
        self._plain_text_widget.configure(state=tk.DISABLED)
        self._auto_height(self._plain_text_widget)

    def _render_formatted(self, text: str):
        self._clear_widgets()
        cleaned = MarkdownParser.clean_text(text)
        if not cleaned:
            return
        lines = cleaned.split('\n')
        i = 0
        current_text = None
        while i < len(lines):
            line = lines[i]
            if line.strip().startswith('```'):
                if current_text:
                    current_text.configure(state=tk.DISABLED)
                    self._auto_height(current_text)
                    current_text = None
                lang = line.strip()[3:].strip()
                code_lines = []
                i += 1
                while i < len(lines) and not lines[i].strip().startswith('```'):
                    code_lines.append(lines[i])
                    i += 1
                code_content = '\n'.join(code_lines)
                code_widget = CodeBlockWidget(self.content_frame, self.theme, code_content, lang)
                code_widget.pack(fill="x", pady=4)
                self._widgets.append(code_widget)
                i += 1
                continue
            if line.strip().startswith('|') and line.strip().endswith('|'):
                if current_text:
                    current_text.configure(state=tk.DISABLED)
                    self._auto_height(current_text)
                    current_text = None
                table_lines = []
                while i < len(lines) and lines[i].strip().startswith('|'):
                    table_lines.append(lines[i])
                    i += 1
                rows = MarkdownParser.parse_table_rows(table_lines)
                if rows:
                    table_widget = TableWidget(self.content_frame, self.theme, rows)
                    table_widget.pack(fill="x", pady=4)
                    self._widgets.append(table_widget)
                continue
            if current_text is None:
                current_text = self._create_text_widget(self.content_frame)
                current_text.pack(fill="x", expand=True)
                self._widgets.append(current_text)
                current_text.configure(state=tk.NORMAL)
            else:
                current_text.insert(tk.END, "\n")
            header_match = re.match(r'^(#{1,6})\s+(.+)$', line)
            if header_match:
                level = len(header_match.group(1))
                InlineFormatter.render_line(current_text, header_match.group(2), (f"h{level}",))
                i += 1
                continue
            bullet_match = re.match(r'^(\s*)([-*+])\s+(.+)$', line)
            if bullet_match:
                indent_level = len(bullet_match.group(1)) // 2
                tag = "bullet" if indent_level == 0 else f"bullet{min(indent_level, 2)}"
                bullets = ["•", "◦", "▪"]
                bullet_char = bullets[min(indent_level, 2)]
                current_text.insert(tk.END, f"{bullet_char} ", (tag,))
                InlineFormatter.render_line(current_text, bullet_match.group(3), (tag,))
                i += 1
                continue
            num_match = re.match(r'^(\s*)(\d+\.)\s+(.+)$', line)
            if num_match:
                current_text.insert(tk.END, f"{num_match.group(2)} ", ("numbered",))
                InlineFormatter.render_line(current_text, num_match.group(3), ("numbered",))
                i += 1
                continue
            if line.strip():
                InlineFormatter.render_line(current_text, line)
            i += 1
        if current_text:
            current_text.configure(state=tk.DISABLED)
            self._auto_height(current_text)
        self._plain_text_widget = None

    def get_raw_content(self) -> str:
        return self._raw_content

    def get_plain_text(self) -> str:
        return MarkdownParser.clean_text(self._raw_content)


class SessionManager:
    def __init__(self):
        self.sessions: Dict[str, Dict] = {}
        self.current_session_id: Optional[str] = None
        self._load_sessions()

    def _load_sessions(self):
        if os.path.exists(SESSIONS_FILE):
            try:
                with open(SESSIONS_FILE, 'r', encoding='utf-8') as f:
                    self.sessions = json.load(f)
            except:
                self.sessions = {}

    def _save_sessions(self):
        try:
            with open(SESSIONS_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.sessions, f, ensure_ascii=False, indent=2)
        except:
            pass

    def create_new_session(self) -> str:
        session_id = str(uuid.uuid4())
        self.sessions[session_id] = {
            "title": "New Chat",
            "created_at": datetime.now().isoformat(),
            "updated_at": datetime.now().isoformat(),
            "messages": [],
            "model": None
        }
        self.current_session_id = session_id
        self._save_sessions()
        return session_id

    def add_message(self, role: str, content: str):
        if not self.current_session_id or self.current_session_id not in self.sessions:
            self.create_new_session()
        session = self.sessions[self.current_session_id]
        session["messages"].append({
            "role": role, "content": content, "timestamp": datetime.now().isoformat()
        })
        session["updated_at"] = datetime.now().isoformat()
        if role == "user" and len(session["messages"]) == 1:
            session["title"] = content[:40].strip() + ("..." if len(content) > 40 else "")
        self._save_sessions()

    def delete_message(self, index: int):
        if self.current_session_id in self.sessions:
            msgs = self.sessions[self.current_session_id]["messages"]
            if 0 <= index < len(msgs):
                msgs.pop(index)
                self._save_sessions()

    def get_conversation_history(self, system_prompt: str) -> List[Dict]:
        if not self.current_session_id:
            return [{"role": "system", "content": system_prompt}]
        history = [{"role": "system", "content": system_prompt}]
        for msg in self.sessions[self.current_session_id]["messages"]:
            history.append({"role": msg["role"], "content": msg["content"]})
        return history

    def delete_session(self, session_id: str):
        if session_id in self.sessions:
            del self.sessions[session_id]
            if self.current_session_id == session_id:
                self.current_session_id = None
            self._save_sessions()

    def get_current_messages(self) -> List[Dict]:
        if self.current_session_id in self.sessions:
            return self.sessions[self.current_session_id]["messages"]
        return []


class MessageWidget(tk.Frame):
    def __init__(self, master, role: str, content: str, index: int, theme: Dict,
                 on_delete: Callable, on_regenerate: Callable, is_last: bool = False):
        super().__init__(master, bg=theme["bg_primary"])
        self.role = role
        self.content = content
        self.index = index
        self.theme = theme
        self.on_delete = on_delete
        self.on_regenerate = on_regenerate
        self.is_last = is_last
        self.renderer = None
        self._build()

    def _build(self):
        is_user = self.role == "user"
        row = tk.Frame(self, bg=self.theme["bg_primary"])
        row.pack(fill="x", padx=16, pady=6)
        if is_user:
            tk.Frame(row, bg=self.theme["bg_primary"]).pack(side="left", fill="x", expand=True)
        container = tk.Frame(row, bg=self.theme["bg_primary"])
        container.pack(side="right" if is_user else "left", anchor="e" if is_user else "w")
        header = tk.Frame(container, bg=self.theme["bg_primary"])
        header.pack(fill="x", pady=(0, 4))
        avatar_text = "You" if is_user else "AI"
        avatar_bg = self.theme["accent_blue"] if is_user else self.theme["accent_purple"]
        avatar = tk.Label(header, text=avatar_text, font=("Segoe UI", 9, "bold"), fg="white", bg=avatar_bg, padx=6, pady=2)
        avatar.pack(side="right" if is_user else "left")
        bubble_bg = self.theme["user_bubble"] if is_user else self.theme["ai_bubble"]
        bubble = tk.Frame(container, bg=bubble_bg)
        bubble.pack(fill="x")
        self.renderer = MessageRenderer(bubble, self.theme, self.role)
        self.renderer.pack(fill="both", expand=True)
        self.renderer.update_content(self.content)
        toolbar = tk.Frame(container, bg=self.theme["bg_primary"])
        toolbar.pack(fill="x", pady=(2, 0))
        btn_frame = tk.Frame(toolbar, bg=self.theme["bg_primary"])
        btn_frame.pack(side="right" if is_user else "left")
        copy_btn = tk.Button(
            btn_frame, text="📋", font=("Segoe UI", 10), bg=self.theme["bg_primary"], fg=self.theme["text_muted"],
            bd=0, padx=4, pady=0, cursor="hand2", activebackground=self.theme["bg_hover"], command=self._copy
        )
        copy_btn.pack(side="left", padx=2)
        if not is_user and self.is_last:
            regen_btn = tk.Button(
                btn_frame, text="🔄", font=("Segoe UI", 10), bg=self.theme["bg_primary"], fg=self.theme["text_muted"],
                bd=0, padx=4, pady=0, cursor="hand2", activebackground=self.theme["bg_hover"], command=self.on_regenerate
            )
            regen_btn.pack(side="left", padx=2)
        del_btn = tk.Button(
            btn_frame, text="🗑️", font=("Segoe UI", 10), bg=self.theme["bg_primary"], fg=self.theme["text_muted"],
            bd=0, padx=4, pady=0, cursor="hand2", activebackground=self.theme["bg_hover"],
            command=lambda: self.on_delete(self.index)
        )
        del_btn.pack(side="left", padx=2)
        if not is_user:
            tk.Frame(row, bg=self.theme["bg_primary"]).pack(side="right", fill="x", expand=True)

    def _copy(self):
        raw_content = self.renderer.get_raw_content() if self.renderer else self.content
        plain_text = MarkdownParser.clean_text(raw_content)
        html_content = markdown_to_html(raw_content)
        
        success = copy_html_to_clipboard(html_content, plain_text)
        
        if not success:
            self.clipboard_clear()
            self.clipboard_append(plain_text)
            self.update()

    def update_content(self, new_content: str):
        self.content = new_content
        if self.renderer:
            self.renderer.update_content(new_content)

    def set_streaming(self, streaming: bool):
        if self.renderer:
            self.renderer.set_streaming(streaming)


class ChatArea(tk.Frame):
    def __init__(self, master, session_manager: SessionManager, theme: Dict):
        super().__init__(master, bg=theme["bg_primary"])
        self.session_manager = session_manager
        self.theme = theme
        self.message_widgets: List[MessageWidget] = []
        self.on_send_callback: Optional[Callable] = None
        self.on_regenerate_callback: Optional[Callable] = None
        self.on_stop_callback: Optional[Callable] = None 
        self._is_streaming = False
        self._auto_scroll_enabled = True
        self._scroll_job_id = None
        self._build_ui()

    def _build_ui(self):
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        scroll_container = tk.Frame(self, bg=self.theme["bg_primary"])
        scroll_container.grid(row=0, column=0, sticky="nsew")
        scroll_container.grid_rowconfigure(0, weight=1)
        scroll_container.grid_columnconfigure(0, weight=1)
        self.canvas = tk.Canvas(scroll_container, bg=self.theme["bg_primary"], highlightthickness=0, bd=0)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar = tk.Scrollbar(scroll_container, orient="vertical", command=self._on_scrollbar)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.messages_frame = tk.Frame(self.canvas, bg=self.theme["bg_primary"])
        self.canvas_window = self.canvas.create_window((0, 0), window=self.messages_frame, anchor="nw")
        self.messages_frame.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        input_container = tk.Frame(self, bg=self.theme["bg_secondary"], height=100)
        input_container.grid(row=1, column=0, sticky="ew")
        input_container.grid_propagate(False)
        inner = tk.Frame(input_container, bg=self.theme["bg_secondary"])
        inner.pack(fill="x", padx=40, pady=12)
        entry_frame = tk.Frame(inner, bg=self.theme["bg_tertiary"], highlightbackground=self.theme["border"], highlightthickness=1)
        entry_frame.pack(fill="x")
        self.input_box = tk.Text(
            entry_frame, height=2, wrap="word", font=("Segoe UI", 12), bg=self.theme["bg_tertiary"],
            fg=self.theme["text_primary"], relief="flat", bd=0, insertbackground=self.theme["text_primary"], padx=12, pady=8
        )
        self.input_box.pack(side="left", fill="both", expand=True)
        self.input_box.bind("<Control-Return>", self._on_send)
        self.send_btn = tk.Button(
            entry_frame, text="➤", font=("Segoe UI", 16), bg=self.theme["accent"], fg="white", bd=0,
            padx=16, pady=8, cursor="hand2", activebackground=self.theme["accent_hover"], command=self._on_send
        )
        self.send_btn.pack(side="right", padx=8, pady=8)
        hint = tk.Label(
            input_container, text="Ctrl+Enter to send", font=("Segoe UI", 9),
            fg=self.theme["text_muted"], bg=self.theme["bg_secondary"]
        )
        hint.pack(pady=(0, 8))

    def _on_scrollbar(self, *args):
        """Handle scrollbar interaction with bounds checking"""
        self.canvas.yview(*args)
        self._clamp_scroll()

    def _on_frame_configure(self, event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self._clamp_scroll()

    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self.canvas_window, width=event.width)
        self._clamp_scroll()

    def _on_mousewheel(self, event):
        # During streaming, user scroll disables auto-scroll temporarily
        if self._is_streaming:
            # Check if user is scrolling up (away from bottom)
            if event.delta > 0:  # Scrolling up
                self._auto_scroll_enabled = False
            else:  # Scrolling down
                # Check if near bottom, re-enable auto-scroll
                self._check_near_bottom()
        
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        self._clamp_scroll()

    def _clamp_scroll(self):
        """Prevent scrolling beyond content bounds"""
        self.canvas.update_idletasks()
        
        # Get scroll region and canvas info
        bbox = self.canvas.bbox("all")
        if not bbox:
            return
            
        content_height = bbox[3] - bbox[1]
        canvas_height = self.canvas.winfo_height()
        
        if content_height <= canvas_height:
            # Content fits in canvas, reset to top
            self.canvas.yview_moveto(0)
            return
        
        # Get current scroll position
        current = self.canvas.yview()
        
        # Clamp to valid range [0, 1]
        if current[0] < 0:
            self.canvas.yview_moveto(0)
        elif current[1] > 1:
            self.canvas.yview_moveto(1.0 - (canvas_height / content_height))

    def _check_near_bottom(self):
        """Check if scroll position is near bottom and re-enable auto-scroll"""
        current = self.canvas.yview()
        if current[1] >= 0.95:  # Within 5% of bottom
            self._auto_scroll_enabled = True

    def _on_send(self, event=None):
        if self._is_streaming:
            if self.on_stop_callback:
                self.on_stop_callback()
            return "break"

        text = self.input_box.get("1.0", "end").strip()
        if text and self.on_send_callback:
            self.input_box.delete("1.0", "end")
            self.on_send_callback(text)
        return "break"

    def clear_messages(self):
        for w in self.messages_frame.winfo_children():
            w.destroy()
        self.message_widgets.clear()

    def add_message(self, role: str, content: str, index: int, is_last: bool = False) -> MessageWidget:
        widget = MessageWidget(
            self.messages_frame, role=role, content=content, index=index, theme=self.theme,
            on_delete=self._on_delete, on_regenerate=self._on_regenerate, is_last=is_last
        )
        widget.pack(fill="x", pady=2)
        self.message_widgets.append(widget)
        self.after_idle(self._smart_scroll_to_bottom)
        return widget

    def _on_delete(self, index: int):
        self.session_manager.delete_message(index)
        self.reload_messages()

    def _on_regenerate(self):
        if self.on_regenerate_callback:
            self.on_regenerate_callback()

    def reload_messages(self):
        self.clear_messages()
        messages = self.session_manager.get_current_messages()
        for idx, msg in enumerate(messages):
            self.add_message(msg["role"], msg["content"], idx, is_last=(idx == len(messages) - 1))

    def scroll_to_bottom(self):
        """Immediate scroll to bottom"""
        self.update_idletasks()
        self.canvas.yview_moveto(1.0)

    def _smart_scroll_to_bottom(self):
        """Smart scroll that respects user interaction during streaming"""
        if self._is_streaming and not self._auto_scroll_enabled:
            return
        
        # Cancel any pending scroll job to avoid buildup
        if self._scroll_job_id:
            self.after_cancel(self._scroll_job_id)
            self._scroll_job_id = None
        
        self.update_idletasks()
        self.canvas.yview_moveto(1.0)

    def request_scroll_to_bottom(self):
        """Request scroll to bottom with debouncing for performance"""
        if self._is_streaming and not self._auto_scroll_enabled:
            return
        
        # Debounce: cancel previous job and schedule new one
        if self._scroll_job_id:
            self.after_cancel(self._scroll_job_id)
        
        self._scroll_job_id = self.after(50, self._do_scroll_to_bottom)

    def _do_scroll_to_bottom(self):
        """Actual scroll execution"""
        self._scroll_job_id = None
        if self._is_streaming and not self._auto_scroll_enabled:
            return
        self.update_idletasks()
        self.canvas.yview_moveto(1.0)

    def set_streaming_mode(self, streaming: bool):
        self._is_streaming = streaming
        if streaming:
            self._auto_scroll_enabled = True  # Reset auto-scroll when starting
            self.send_btn.configure(text="⏹", bg=self.theme["error"])
        else:
            self._auto_scroll_enabled = True  # Ensure it's enabled when done
            self.send_btn.configure(text="➤", bg=self.theme["accent"])
            # Final scroll to bottom
            self.after_idle(self.scroll_to_bottom)


class Sidebar(tk.Frame):
    def __init__(self, master, session_manager: SessionManager, theme: Dict):
        super().__init__(master, bg=theme["bg_secondary"], width=280)
        self.pack_propagate(False)
        self.session_manager = session_manager
        self.theme = theme
        self.on_new_chat: Optional[Callable] = None
        self.on_select_session: Optional[Callable] = None
        self.on_delete_session: Optional[Callable] = None
        self._build_ui()

    def _build_ui(self):
        header = tk.Frame(self, bg=self.theme["bg_secondary"])
        header.pack(fill="x", padx=15, pady=15)
        tk.Label(
            header, text="💬 Chats", font=("Segoe UI", 16, "bold"),
            fg=self.theme["text_primary"], bg=self.theme["bg_secondary"]
        ).pack(side="left")
        new_btn = tk.Button(
            header, text="+", font=("Segoe UI", 14, "bold"), bg=self.theme["accent"], fg="white", bd=0,
            padx=10, pady=2, cursor="hand2", activebackground=self.theme["accent_hover"],
            command=lambda: self.on_new_chat() if self.on_new_chat else None
        )
        new_btn.pack(side="right")
        self.sessions_canvas = tk.Canvas(self, bg=self.theme["bg_secondary"], highlightthickness=0)
        self.sessions_scrollbar = tk.Scrollbar(self, command=self.sessions_canvas.yview)
        self.sessions_frame = tk.Frame(self.sessions_canvas, bg=self.theme["bg_secondary"])
        self.sessions_canvas.pack(side="left", fill="both", expand=True, padx=10)
        self.sessions_scrollbar.pack(side="right", fill="y")
        self.sessions_canvas.configure(yscrollcommand=self.sessions_scrollbar.set)
        self.sessions_window = self.sessions_canvas.create_window((0, 0), window=self.sessions_frame, anchor="nw")
        self.sessions_frame.bind("<Configure>", lambda e: self.sessions_canvas.configure(scrollregion=self.sessions_canvas.bbox("all")))
        self.sessions_canvas.bind("<Configure>", lambda e: self.sessions_canvas.itemconfig(self.sessions_window, width=e.width))
        self.server_frame = tk.Frame(self, bg=self.theme["bg_tertiary"])
        self.server_frame.pack(fill="x", side="bottom", padx=10, pady=10)
        status_row = tk.Frame(self.server_frame, bg=self.theme["bg_tertiary"])
        status_row.pack(fill="x", padx=10, pady=5)
        self.status_dot = tk.Label(status_row, text="●", font=("Segoe UI", 14), fg=self.theme["error"], bg=self.theme["bg_tertiary"])
        self.status_dot.pack(side="left")
        self.status_text = tk.Label(status_row, text="Offline", font=("Segoe UI", 11), fg=self.theme["text_secondary"], bg=self.theme["bg_tertiary"])
        self.status_text.pack(side="left", padx=8)
        self.server_btn = tk.Button(
            self.server_frame, text="Start Server", font=("Segoe UI", 10), bg=self.theme["accent"], fg="white", bd=0,
            padx=12, pady=6, cursor="hand2", activebackground=self.theme["accent_hover"]
        )
        self.server_btn.pack(fill="x", padx=10, pady=(0, 8))

    def refresh_sessions(self):
        for w in self.sessions_frame.winfo_children():
            w.destroy()
        sorted_sessions = sorted(
            self.session_manager.sessions.items(),
            key=lambda x: x[1].get('updated_at', x[1].get('created_at', '')), reverse=True
        )
        for sid, data in sorted_sessions:
            is_current = sid == self.session_manager.current_session_id
            bg = self.theme["accent"] if is_current else self.theme["bg_secondary"]
            frame = tk.Frame(self.sessions_frame, bg=bg)
            frame.pack(fill="x", pady=2)
            title = data["title"][:22] + "..." if len(data["title"]) > 22 else data["title"]
            btn = tk.Button(
                frame, text=title, font=("Segoe UI", 10), bg=bg, fg=self.theme["text_primary"], bd=0,
                anchor="w", padx=8, pady=6, cursor="hand2", activebackground=self.theme["bg_hover"],
                command=lambda s=sid: self.on_select_session(s) if self.on_select_session else None
            )
            btn.pack(side="left", fill="x", expand=True)
            del_btn = tk.Button(
                frame, text="×", font=("Segoe UI", 12), bg=bg, fg=self.theme["text_muted"], bd=0,
                padx=6, cursor="hand2", activebackground=self.theme["error"],
                command=lambda s=sid: self.on_delete_session(s) if self.on_delete_session else None
            )
            del_btn.pack(side="right")

    def update_server_status(self, online: bool, starting: bool = False):
        if starting:
            self.status_dot.configure(fg=self.theme["warning"])
            self.status_text.configure(text="Starting...")
            self.server_btn.configure(text="Starting...", state="disabled")
        elif online:
            self.status_dot.configure(fg=self.theme["success"])
            self.status_text.configure(text="Online")
            self.server_btn.configure(text="Stop Server", bg=self.theme["error"], state="normal")
        else:
            self.status_dot.configure(fg=self.theme["error"])
            self.status_text.configure(text="Offline")
            self.server_btn.configure(text="Start Server", bg=self.theme["accent"], state="normal")


class ConfigPanel(tk.Frame):
    def __init__(self, master, theme: Dict):
        super().__init__(master, bg=theme["bg_secondary"], width=300)
        self.pack_propagate(False)
        self.theme = theme
        self.on_theme_changed: Optional[Callable] = None
        self.on_refresh_models: Optional[Callable] = None
        self.on_browse_ollama: Optional[Callable] = None
        self.on_model_changed: Optional[Callable] = None
        self._build_ui()

    def _build_ui(self):
        canvas = tk.Canvas(self, bg=self.theme["bg_secondary"], highlightthickness=0)
        scrollbar = tk.Scrollbar(self, command=canvas.yview)
        scroll_frame = tk.Frame(canvas, bg=self.theme["bg_secondary"])
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas_window = canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(canvas_window, width=e.width - 20))
        
        header = tk.Frame(scroll_frame, bg=self.theme["bg_secondary"])
        header.pack(fill="x", padx=15, pady=15)
        
        tk.Label(header, text="⚙️ Settings", font=("Segoe UI", 16, "bold"), 
                 fg=self.theme["text_primary"], bg=self.theme["bg_secondary"]).pack(side="left")
        
        self.theme_var = tk.IntVar(value=1 if ctk.get_appearance_mode() == "Dark" else 0)
        theme_chk = tk.Checkbutton(
            header, text="Dark", font=("Segoe UI", 10), fg=self.theme["text_secondary"], 
            bg=self.theme["bg_secondary"], selectcolor=self.theme["bg_tertiary"], 
            variable=self.theme_var, command=self._toggle_theme
        )
        theme_chk.pack(side="right")
        
        self._section(scroll_frame, "Ollama Path")
        self.ollama_label = tk.Label(scroll_frame, text="Not configured", font=("Segoe UI", 9), 
                                     fg=self.theme["text_muted"], bg=self.theme["bg_secondary"], 
                                     anchor="w", wraplength=250)
        self.ollama_label.pack(fill="x", padx=15)
        
        browse_btn = tk.Button(
            scroll_frame, text="Browse Ollama...", font=("Segoe UI", 10), 
            bg=self.theme["accent_blue"], fg="white", bd=0, padx=12, pady=6, 
            cursor="hand2", command=lambda: self.on_browse_ollama() if self.on_browse_ollama else None
        )
        browse_btn.pack(fill="x", padx=15, pady=(5, 10))
        
        self._section(scroll_frame, "Model")
        model_row = tk.Frame(scroll_frame, bg=self.theme["bg_secondary"])
        model_row.pack(fill="x", padx=15, pady=(0, 10))
        
        self.model_var = tk.StringVar(value="No models")
        self.model_combo = tk.OptionMenu(model_row, self.model_var, "No models")
        self.model_combo.config(font=("Segoe UI", 10), bg=self.theme["bg_tertiary"], 
                                fg=self.theme["text_primary"], highlightthickness=0, bd=0)
        self.model_combo.pack(side="left", fill="x", expand=True)
        
        refresh_btn = tk.Button(
            model_row, text="🔄", font=("Segoe UI", 12), bg=self.theme["accent_blue"], 
            fg="white", bd=0, padx=8, cursor="hand2",
            command=lambda: self.on_refresh_models() if self.on_refresh_models else None
        )
        refresh_btn.pack(side="right", padx=(5, 0))
        
        self._section(scroll_frame, "System Prompt")
        self.system_prompt = tk.Text(
            scroll_frame, height=5, wrap="word", font=("Segoe UI", 10), 
            bg=self.theme["bg_tertiary"], fg=self.theme["text_primary"], 
            relief="flat", bd=1, padx=8, pady=6
        )
        self.system_prompt.pack(fill="x", padx=15, pady=(0, 10))
        self.system_prompt.insert("1.0", "You are a helpful AI assistant.")
        
        self._section(scroll_frame, "Prefix")
        self.prefix_entry = tk.Entry(scroll_frame, font=("Segoe UI", 10), 
                                     bg=self.theme["bg_tertiary"], fg=self.theme["text_primary"], 
                                     relief="flat", bd=1)
        self.prefix_entry.pack(fill="x", padx=15, pady=(0, 10))
        
        self._section(scroll_frame, "Suffix")
        self.suffix_entry = tk.Entry(scroll_frame, font=("Segoe UI", 10), 
                                     bg=self.theme["bg_tertiary"], fg=self.theme["text_primary"], 
                                     relief="flat", bd=1)
        self.suffix_entry.pack(fill="x", padx=15, pady=(0, 10))
        
        self._section(scroll_frame, "Temperature")
        temp_row = tk.Frame(scroll_frame, bg=self.theme["bg_secondary"])
        temp_row.pack(fill="x", padx=15)
        
        self.temp_var = tk.DoubleVar(value=0.7)
        self.temp_scale = tk.Scale(
            temp_row, from_=0, to=2, resolution=0.1, orient="horizontal", 
            variable=self.temp_var, bg=self.theme["bg_secondary"], 
            fg=self.theme["text_primary"], highlightthickness=0, 
            troughcolor=self.theme["bg_tertiary"]
        )
        self.temp_scale.pack(fill="x")
        
        self._section(scroll_frame, "Context Length")
        self.ctx_var = tk.IntVar(value=4096)
        # CHANGED: Increased max to 131072 (128k)
        self.ctx_scale = tk.Scale(
            scroll_frame, from_=512, to=131072, orient="horizontal", 
            variable=self.ctx_var, bg=self.theme["bg_secondary"], 
            fg=self.theme["text_primary"], highlightthickness=0, 
            troughcolor=self.theme["bg_tertiary"]
        )
        self.ctx_scale.pack(fill="x", padx=15, pady=(0, 15))

    def _section(self, parent, title: str):
        tk.Label(parent, text=title, font=("Segoe UI", 11, "bold"), fg=self.theme["text_primary"], bg=self.theme["bg_secondary"], anchor="w").pack(fill="x", padx=15, pady=(10, 5))

    def _toggle_theme(self):
        mode = "Dark" if self.theme_var.get() == 1 else "Light"
        ctk.set_appearance_mode(mode)
        if self.on_theme_changed:
            self.on_theme_changed(mode)

    def update_ollama_path(self, path: str):
        self.ollama_label.configure(text=os.path.basename(path) if path else "Not configured")

    def set_max_context(self, max_val: int):
        self.ctx_scale.configure(to=max_val)
        if self.ctx_var.get() > max_val:
            self.ctx_var.set(max_val)

    def get_settings(self) -> Dict:
        return {
            "system_prompt": self.system_prompt.get("1.0", "end").strip(),
            "prefix": self.prefix_entry.get().strip(),
            "suffix": self.suffix_entry.get().strip(),
            "model": self.model_var.get(),
            "temperature": self.temp_var.get(),
            "context_length": self.ctx_var.get(),
            "theme": "Dark" if self.theme_var.get() == 1 else "Light"
        }

    def load_settings(self, settings: Dict):
        if "system_prompt" in settings:
            self.system_prompt.delete("1.0", "end")
            self.system_prompt.insert("1.0", settings["system_prompt"])
        if "prefix" in settings:
            self.prefix_entry.delete(0, "end")
            self.prefix_entry.insert(0, settings["prefix"])
        if "suffix" in settings:
            self.suffix_entry.delete(0, "end")
            self.suffix_entry.insert(0, settings["suffix"])
        if "model" in settings:
            self.model_var.set(settings["model"])
        if "temperature" in settings:
            self.temp_var.set(settings["temperature"])
        if "context_length" in settings:
            self.ctx_var.set(settings["context_length"])

    def update_models(self, models: List[str]):
        if models:
            menu = self.model_combo["menu"]
            menu.delete(0, "end")
            for m in models:
                menu.add_command(label=m, command=lambda v=m: self._set_model(v))
            if self.model_var.get() == "No models":
                self.model_var.set(models[0])
                if self.on_model_changed:
                    self.on_model_changed(models[0])

    def _set_model(self, value: str):
        self.model_var.set(value)
        if self.on_model_changed:
            self.on_model_changed(value)


class OllamaManager:
    def __init__(self):
        self.process: Optional[subprocess.Popen] = None
        atexit.register(self.cleanup)
        signal.signal(signal.SIGINT, lambda s, f: (self.cleanup(), exit(0)))
        signal.signal(signal.SIGTERM, lambda s, f: (self.cleanup(), exit(0)))

    def start(self, path: str) -> bool:
        if not path or not os.path.exists(path):
            return False
        try:
            self.process = subprocess.Popen([path], cwd=os.path.dirname(path), creationflags=subprocess.CREATE_NEW_CONSOLE, shell=True)
            return True
        except:
            return False

    def stop(self):
        if self.process:
            try:
                self.process.terminate()
                self.process.wait(timeout=5)
            except:
                try:
                    self.process.kill()
                except:
                    pass
            self.process = None

    def cleanup(self):
        self.stop()
        try:
            if os.name == 'nt':
                subprocess.run(['taskkill', '/F', '/IM', 'ollama.exe'], capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW)
        except:
            pass


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Ollama Chat Interface")
        self.geometry("1400x900")
        self.minsize(1000, 700)
        self.session_manager = SessionManager()
        self.settings = self._load_settings()
        self.ollama_manager = OllamaManager()
        self.is_streaming = False
        self.abort_stream = False
        self.current_ai_widget: Optional[MessageWidget] = None
        self.ollama_path = self.settings.get("ollama_path", DEFAULT_OLLAMA_PATH)
        ctk.set_appearance_mode(self.settings.get("theme", "Dark"))
        self.theme = Theme.get()
        self._build_layout()
        self._connect_events()
        self._apply_settings()
        self.sidebar.refresh_sessions()
        self._check_ollama()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _on_close(self):
        self._save_settings()
        self.ollama_manager.cleanup()
        self.destroy()

    def _load_settings(self) -> Dict:
        defaults = {
            "system_prompt": "You are a helpful AI assistant.",
            "prefix": "", "suffix": "", "model": "qwen3:1.7b",
            "temperature": 0.7, "context_length": 4096, "theme": "Dark",
            "ollama_path": DEFAULT_OLLAMA_PATH
        }
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
                    return {**defaults, **json.load(f)}
            except:
                pass
        return defaults

    def _save_settings(self):
        try:
            settings = self.config_panel.get_settings()
            settings["ollama_path"] = self.ollama_path
            with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
            self.settings = settings
        except:
            pass

    def _apply_settings(self):
        self.config_panel.load_settings(self.settings)
        self.config_panel.update_ollama_path(self.ollama_path)

    def _build_layout(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.sidebar = Sidebar(self, self.session_manager, self.theme)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.chat_area = ChatArea(self, self.session_manager, self.theme)
        self.chat_area.grid(row=0, column=1, sticky="nsew")
        self.config_panel = ConfigPanel(self, self.theme)
        self.config_panel.grid(row=0, column=2, sticky="nsew")

    def _connect_events(self):
        self.sidebar.on_new_chat = self._new_chat
        self.sidebar.on_select_session = self._select_session
        self.sidebar.on_delete_session = self._delete_session
        self.sidebar.server_btn.configure(command=self._toggle_server)
        self.chat_area.on_send_callback = self._send_message
        self.chat_area.on_stop_callback = self._stop_generation
        self.chat_area.on_regenerate_callback = self._regenerate
        self.config_panel.on_theme_changed = self._theme_changed
        self.config_panel.on_refresh_models = self._refresh_models
        self.config_panel.on_browse_ollama = self._browse_ollama
        self.config_panel.on_model_changed = self._fetch_model_details

    def _stop_generation(self):
        if self.is_streaming:
            self.abort_stream = True

    def _check_ollama(self):
        try:
            if requests.get(f"{OLLAMA_API_BASE}/api/tags", timeout=2).status_code == 200:
                self._server_ready()
                return
        except:
            pass
        if self.ollama_path and os.path.exists(self.ollama_path):
            self._start_server()
        else:
            self.after(100, self._prompt_ollama)

    def _prompt_ollama(self):
        if messagebox.askyesno("Ollama Not Found", f"Ollama not found at:\n{self.ollama_path}\n\nBrowse for it?"):
            self._browse_ollama()
        else:
            self.sidebar.update_server_status(False)

    def _browse_ollama(self):
        path = filedialog.askopenfilename(title="Select Ollama start file", filetypes=[("Batch files", "*.bat"), ("Shell scripts", "*.sh"), ("All files", "*.*")])
        if path:
            self.ollama_path = path
            self.config_panel.update_ollama_path(path)
            self._save_settings()
            if messagebox.askyesno("Start Ollama?", "Start the server now?"):
                self._start_server()

    def _new_chat(self):
        self.session_manager.create_new_session()
        self.sidebar.refresh_sessions()
        self.chat_area.clear_messages()

    def _select_session(self, sid: str):
        if sid not in self.session_manager.sessions:
            self.sidebar.refresh_sessions()
            return
        self.session_manager.current_session_id = sid
        self.sidebar.refresh_sessions()
        self.chat_area.reload_messages()

    def _delete_session(self, sid: str):
        self.session_manager.delete_session(sid)
        self.sidebar.refresh_sessions()
        if not self.session_manager.current_session_id:
            self.chat_area.clear_messages()

    def _send_message(self, text: str):
        settings = self.config_panel.get_settings()
        final = f"{settings.get('prefix', '')} {text} {settings.get('suffix', '')}".strip()
        self.session_manager.add_message("user", final)
        self.chat_area.reload_messages()
        self._start_stream()

    def _regenerate(self):
        messages = self.session_manager.get_current_messages()
        if messages and messages[-1]["role"] == "assistant":
            self.session_manager.delete_message(len(messages) - 1)
            self.chat_area.reload_messages()
            self._start_stream()

    def _start_stream(self):
        self.is_streaming = True
        self.abort_stream = False
        self.chat_area.set_streaming_mode(True)
        messages = self.session_manager.get_current_messages()
        self.current_ai_widget = self.chat_area.add_message("assistant", "", len(messages), is_last=True)
        self.current_ai_widget.set_streaming(True)
        self.chat_area.scroll_to_bottom()
        threading.Thread(target=self._stream_worker, daemon=True).start()

    def _stream_worker(self):
        try:
            settings = self.config_panel.get_settings()
            client = OpenAI(base_url=OLLAMA_CHAT_URL, api_key=API_KEY)
            history = self.session_manager.get_conversation_history(settings.get("system_prompt", ""))
            
            stream = client.chat.completions.create(
                model=settings.get("model", "qwen3:1.7b"), 
                messages=history,
                stream=True, 
                temperature=settings.get("temperature", 0.7)
            )
            
            full_response = ""
            last_update = time.time()
            last_scroll = time.time()
            
            for chunk in stream:
                if self.abort_stream:
                    break
                content = chunk.choices[0].delta.content or ""
                if content:
                    full_response += content
                    now = time.time()
                    if now - last_update >= 0.05:
                        last_update = now
                        self.after(0, lambda t=full_response: self._update_response(t))
                        
                        if now - last_scroll >= 0.15:
                            last_scroll = now
                            self.after(0, self.chat_area.request_scroll_to_bottom)
            if full_response:
                self.after(0, lambda t=full_response: self._update_response(t))
                self.session_manager.add_message("assistant", full_response)

            self.after(0, self._finish_stream)
        except Exception as e:
            if not self.abort_stream:
                self.after(0, lambda: self._update_response(f"Error: {str(e)}"))
            self.after(0, self._finish_stream)

    def _update_response(self, content: str):
        if self.abort_stream:
            return
        if self.current_ai_widget:
            self.current_ai_widget.update_content(content)

    def _finish_stream(self):
        if self.current_ai_widget:
            self.current_ai_widget.set_streaming(False)
            self.current_ai_widget.is_last = True
            self.current_ai_widget.index = len(self.session_manager.get_current_messages()) - 1
        self.is_streaming = False
        self.chat_area.set_streaming_mode(False)

    def _toggle_server(self):
        try:
            if requests.get(f"{OLLAMA_API_BASE}/api/tags", timeout=1).status_code == 200:
                self._stop_server()
                return
        except:
            pass
        self._start_server()

    def _start_server(self):
        if not self.ollama_path or not os.path.exists(self.ollama_path):
            self._prompt_ollama()
            return
        self.sidebar.update_server_status(False, starting=True)
        if self.ollama_manager.start(self.ollama_path):
            threading.Thread(target=self._wait_server, daemon=True).start()
        else:
            self.sidebar.update_server_status(False)

    def _stop_server(self):
        self.ollama_manager.stop()
        self.sidebar.update_server_status(False)

    def _wait_server(self):
        for _ in range(30):
            try:
                if requests.get(f"{OLLAMA_API_BASE}/api/tags", timeout=1).status_code == 200:
                    self.after(0, self._server_ready)
                    return
            except:
                pass
            time.sleep(1)
        self.after(0, lambda: self.sidebar.update_server_status(False))

    def _server_ready(self):
        self.sidebar.update_server_status(True)
        self._refresh_models()

    def _refresh_models(self):
        try:
            resp = requests.get(f"{OLLAMA_API_BASE}/api/tags", timeout=5)
            if resp.status_code == 200:
                models = [m["name"] for m in resp.json().get("models", [])]
                if models:
                    self.config_panel.update_models(models)
                    self._fetch_model_details(models[0])
        except:
            pass

    def _fetch_model_details(self, model_name: str):
        def run():
            try:
                resp = requests.post(f"{OLLAMA_API_BASE}/api/show", json={"name": model_name}, timeout=5)
                if resp.status_code == 200:
                    data = resp.json()
                    # Default to 128k if detection fails (optimistic), so we don't restrict the user
                    ctx = 131072
                    
                    info = data.get("model_info", {})
                    
                    # List of known GGUF keys for context length
                    keys_to_check = [
                        "llama.context_length",
                        "qwen.context_length",
                        "qwen2.context_length",
                        "phi3.context_length",
                        "gemma.context_length",
                        "gemma2.context_length",
                        "context_length"
                    ]
                    
                    found = False
                    for key in keys_to_check:
                        if key in info:
                            try:
                                val = int(info[key])
                                if val > 0:
                                    ctx = val
                                    found = True
                                    break
                            except:
                                pass
                    
                    # If not found in model_info, check if there is a default in parameters
                    if not found:
                        params = data.get("parameters", "")
                        # Simple check for num_ctx in the Modelfile parameters
                        import re
                        match = re.search(r'num_ctx\s+(\d+)', params)
                        if match:
                            ctx = int(match.group(1))

                    self.after(0, lambda: self.config_panel.set_max_context(ctx))
            except Exception as e:
                print(f"Error fetching model details: {e}")
                # On error, ensure we allow high context
                self.after(0, lambda: self.config_panel.set_max_context(131072))
                
        threading.Thread(target=run, daemon=True).start()

    def _theme_changed(self, mode: str):
        self._save_settings()
        self.destroy()
        new_app = App()
        new_app.mainloop()


if __name__ == "__main__":
    app = App()
    app.mainloop()
