"""
Script: AI-Assisted GDL Object Editor
Description: Extract GDL object to text format, modify with AI, and recompile
Features:
    - Extract all scripts and parameters to a single copyable text
    - Tkinter interface for easy copy/paste workflow
    - Parse AI-modified text back to HSF structure
    - Compile and reload automatically
Requirements:
    - archicad-api package
    - Tapir Add-On installed
    - tkinter (included with Python)
"""

import os
import sys
import shutil
import tempfile
import subprocess
import platform
from pathlib import Path
from tkinter import Tk, filedialog, messagebox, scrolledtext, ttk
import tkinter as tk

# Import functions from existing scripts
try:
    from archicad import ACConnection
except ImportError:
    print("Error: archicad package not installed. Run: pip install archicad")
    sys.exit(1)


class AIGDLEditor:
    def __init__(self, root):
        self.root = root
        self.root.title("AI-Assisted GDL Object Editor")
        self.root.geometry("1200x800")

        # Variables
        self.gsm_file = None
        self.temp_dir = None
        self.hsf_path = None
        self.archicad_path = None
        self.converter_path = None

        # Initialize Archicad connection
        self.init_archicad()

        # Create UI
        self.create_ui()

    def init_archicad(self):
        """Initialize connection to Archicad and get paths"""
        try:
            conn = ACConnection.connect()
            acc = conn.commands
            act = conn.types

            response = acc.ExecuteAddOnCommand(
                act.AddOnCommandId('TapirCommand', 'GetArchicadLocation')
            )

            if 'archicadLocation' in response:
                self.archicad_path = response['archicadLocation']
                system = platform.system()
                
                print(f"Archicad location: {self.archicad_path}")
                print(f"System: {system}")

                if system == "Windows":
                    archicad_dir = os.path.dirname(self.archicad_path)
                    self.converter_path = os.path.join(
                        archicad_dir, "LP_XMLConverter.exe")
                    print(f"Trying Windows path: {self.converter_path}")
                    
                elif system == "Darwin":  # macOS
                    # On Mac, archicadLocation returns the .app bundle:
                    # /Applications/Graphisoft/Archicad 27/Archicad 27.app
                    
                    # Determine the MacOS directory
                    if self.archicad_path.endswith('.app'):
                        # archicadLocation is the .app bundle itself
                        macos_dir = os.path.join(self.archicad_path, "Contents", "MacOS")
                    else:
                        # archicadLocation is the executable inside MacOS
                        archicad_dir = os.path.dirname(self.archicad_path)
                        if archicad_dir.endswith("MacOS"):
                            macos_dir = archicad_dir
                        else:
                            macos_dir = archicad_dir
                    
                    # LP_XMLConverter is inside LP_XMLConverter.app bundle
                    # Path: .../MacOS/LP_XMLConverter.app/Contents/MacOS/LP_XMLConverter
                    self.converter_path = os.path.join(
                        macos_dir, "LP_XMLConverter.app", "Contents", "MacOS", "LP_XMLConverter")
                    
                    print(f"LP_XMLConverter path: {self.converter_path}")
                    
                    if os.path.exists(self.converter_path):
                        print(f"✓ LP_XMLConverter found!")
                    else:
                        print(f"✗ LP_XMLConverter not found at: {self.converter_path}")
                        self.converter_path = None

                # Verify the path exists
                if self.converter_path and not os.path.exists(self.converter_path):
                    print(f"Warning: Converter path set but file doesn't exist: {self.converter_path}")
                    self.converter_path = None
                elif self.converter_path:
                    print(f"✓ LP_XMLConverter found: {self.converter_path}")

        except Exception as e:
            print(f"Error initializing Archicad connection: {e}")
            self.archicad_path = None
            self.converter_path = None

    def create_ui(self):
        """Create the user interface"""
        # Top frame - controls
        top_frame = ttk.Frame(self.root, padding="10")
        top_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N))

        # Object selection
        ttk.Label(top_frame, text="1. Select GDL Object:").grid(
            row=0, column=0, sticky=tk.W, pady=5)
        self.file_label = ttk.Label(
            top_frame, text="No file selected", foreground="gray")
        self.file_label.grid(row=0, column=1, sticky=tk.W, padx=10)

        ttk.Button(top_frame, text="Browse...", command=self.select_file).grid(
            row=0, column=2, padx=5)
        ttk.Button(top_frame, text="Extract to Text",
                   command=self.extract_to_text).grid(row=0, column=3, padx=5)
        
        # Converter path (for troubleshooting on Mac)
        if platform.system() == "Darwin" and not self.converter_path:
            ttk.Button(top_frame, text="⚙️ Set Converter Path", 
                      command=self.select_converter_path).grid(row=0, column=4, padx=5)

        # Middle frame - extracted content (READ-ONLY for copy)
        middle_frame = ttk.LabelFrame(
            self.root, text="2. Extracted Content (Copy this to AI)", padding="10")
        middle_frame.grid(row=1, column=0, sticky=(
            tk.W, tk.E, tk.N, tk.S), padx=10, pady=5)

        self.extracted_text = scrolledtext.ScrolledText(
            middle_frame, wrap=tk.WORD, height=15)
        self.extracted_text.grid(row=0, column=0, sticky=(
            tk.W, tk.E, tk.N, tk.S), pady=5)

        btn_frame1 = ttk.Frame(middle_frame)
        btn_frame1.grid(row=1, column=0, sticky=tk.W)

        ttk.Button(btn_frame1, text="📋 Copy to Clipboard",
                   command=self.copy_to_clipboard).grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame1, text="💾 Save to File",
                   command=self.save_to_file).grid(row=0, column=1, padx=5)
        ttk.Button(btn_frame1, text="🔍 Show Format Debug",
                   command=self.show_format_debug).grid(row=0, column=2, padx=5)

        # Bottom frame - modified content (EDITABLE for paste)
        bottom_frame = ttk.LabelFrame(
            self.root, text="3. Modified Content (Paste AI response here)", padding="10")
        bottom_frame.grid(row=2, column=0, sticky=(
            tk.W, tk.E, tk.N, tk.S), padx=10, pady=5)

        self.modified_text = scrolledtext.ScrolledText(
            bottom_frame, wrap=tk.WORD, height=15)
        self.modified_text.grid(row=0, column=0, sticky=(
            tk.W, tk.E, tk.N, tk.S), pady=5)

        btn_frame2 = ttk.Frame(bottom_frame)
        btn_frame2.grid(row=1, column=0, sticky=tk.W)

        ttk.Button(btn_frame2, text="📥 Paste from Clipboard",
                   command=self.paste_from_clipboard).grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame2, text="📂 Load from File",
                   command=self.load_from_file).grid(row=0, column=1, padx=5)
        ttk.Button(btn_frame2, text="🔍 Test Parse", command=self.test_parse,
                   style="Accent.TButton").grid(row=0, column=2, padx=5)
        ttk.Button(btn_frame2, text="🔨 Compile & Reload", command=self.compile_and_reload,
                   style="Accent.TButton").grid(row=0, column=3, padx=5)

        # Status bar
        self.status_bar = ttk.Label(
            self.root, text="Ready. Select a GDL object to start.", relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.grid(row=3, column=0, sticky=(
            tk.W, tk.E), padx=5, pady=5)

        # Configure grid weights for resizing
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)
        self.root.rowconfigure(2, weight=1)
        middle_frame.columnconfigure(0, weight=1)
        middle_frame.rowconfigure(0, weight=1)
        bottom_frame.columnconfigure(0, weight=1)
        bottom_frame.rowconfigure(0, weight=1)

    def update_status(self, message, color="black"):
        """Update status bar"""
        self.status_bar.config(text=message, foreground=color)
        self.root.update()

    def select_file(self):
        """Select a .gsm file"""
        file_path = filedialog.askopenfilename(
            title="Select GDL Object",
            filetypes=[("Archicad GDL Objects", "*.gsm"), ("All files", "*.*")]
        )

        if file_path:
            self.gsm_file = file_path
            self.file_label.config(text=os.path.basename(
                file_path), foreground="black")
            self.update_status(
                f"Selected: {os.path.basename(file_path)}", "blue")

    def select_converter_path(self):
        """Manually select LP_XMLConverter path (mainly for Mac troubleshooting)"""
        system = platform.system()
        
        if system == "Darwin":
            # On Mac, we're looking for the executable inside the .app bundle
            file_path = filedialog.askopenfilename(
                title="Select LP_XMLConverter (inside .app/Contents/MacOS/)",
                filetypes=[("All files", "*.*")]
            )
        else:
            file_path = filedialog.askopenfilename(
                title="Select LP_XMLConverter.exe",
                filetypes=[("Executable files", "*.exe"), ("All files", "*.*")]
            )
        
        if file_path and os.path.exists(file_path):
            self.converter_path = file_path
            print(f"✓ Manually set converter path: {self.converter_path}")
            messagebox.showinfo("Success", 
                              f"LP_XMLConverter path set successfully:\n\n{os.path.basename(file_path)}")
            self.update_status("✓ Converter path set manually", "green")
        elif file_path:
            messagebox.showerror("Error", "Selected file does not exist")

    def extract_to_text(self):
        """Extract GDL object to text format"""
        if not self.gsm_file:
            messagebox.showerror("Error", "Please select a GDL object first")
            return

        if not self.converter_path:
            system = platform.system()
            if system == "Darwin":
                error_msg = ("LP_XMLConverter not found on macOS.\n\n"
                           f"Expected location: {os.path.join(self.archicad_path, 'Contents', 'MacOS', 'LP_XMLConverter') if self.archicad_path else 'Unknown'}\n\n"
                           "Make sure:\n"
                           "1. Archicad is running\n"
                           "2. LP_XMLConverter is installed with Archicad\n\n"
                           "You can use the '⚙️ Set Converter Path' button to set it manually.")
            else:
                error_msg = "LP_XMLConverter not found. Make sure Archicad is running."
            
            messagebox.showerror("Error", error_msg)
            return

        self.update_status("Extracting object...", "blue")

        try:
            # Create temporary directory
            self.temp_dir = tempfile.mkdtemp(prefix="gdl_ai_")

            # Extract to HSF
            cmd = [
                self.converter_path,
                "libpart2hsf",
                self.gsm_file,
                self.temp_dir
            ]

            result = subprocess.run(
                cmd, capture_output=True, text=True, check=True)

            # DEBUG: List all contents of temp_dir
            temp_path = Path(self.temp_dir)
            all_items = list(temp_path.iterdir())
            debug_info = f"Temp dir: {self.temp_dir}\n"
            debug_info += f"Contents ({len(all_items)} items):\n"
            for item in all_items:
                debug_info += f"  - {item.name} ({'DIR' if item.is_dir() else 'FILE'})\n"
                if item.is_dir():
                    subitems = list(item.iterdir())[:5]  # First 5 items
                    for subitem in subitems:
                        debug_info += f"    - {subitem.name}\n"

            # Find the HSF folder
            # Try 1: Look for subfolder with scripts/ or paramlist.xml
            subfolders = [f for f in temp_path.iterdir()
                          if f.is_dir() and
                          ((f / "scripts").exists() or (f / "paramlist.xml").exists())]

            if subfolders:
                self.hsf_path = str(subfolders[0])
            else:
                # Try 2: Maybe HSF content is directly in temp_dir?
                if (temp_path / "scripts").exists() or (temp_path / "paramlist.xml").exists():
                    self.hsf_path = str(temp_path)
                else:
                    # Try 3: Look for any folder with .gdl files inside scripts
                    subfolders = []
                    for f in temp_path.iterdir():
                        if f.is_dir():
                            scripts_dir = f / "scripts"
                            if scripts_dir.exists() and scripts_dir.is_dir():
                                gdl_files = list(scripts_dir.glob("*.gdl"))
                                if gdl_files:
                                    subfolders.append(f)

                    if subfolders:
                        self.hsf_path = str(subfolders[0])
                    else:
                        raise Exception(
                            f"No valid HSF folder found.\n\n{debug_info}")

            # Extract to text format
            text_content = self.hsf_to_text(self.hsf_path)

            # Display in text area
            self.extracted_text.delete(1.0, tk.END)
            self.extracted_text.insert(1.0, text_content)

            self.update_status(
                "✓ Extraction complete! Copy the text and send it to AI.", "green")

        except Exception as e:
            self.update_status(f"✗ Extraction failed: {e}", "red")
            messagebox.showerror("Extraction Error", str(e))

    def hsf_to_text(self, hsf_path):
        """Convert HSF folder structure to a single text format"""
        hsf_folder = Path(hsf_path)
        lines = []

        # Header
        lines.append("="*70)
        lines.append("GDL OBJECT - AI-EDITABLE FORMAT")
        lines.append("="*70)
        lines.append("")
        lines.append("INSTRUCTIONS:")
        lines.append("- Modify any section below as needed")
        lines.append("- Keep the section markers (70 equal signs '=') intact")
        lines.append("- Do NOT remove or modify the section headers")
        lines.append("- Return the entire modified content")
        lines.append("")

        # Object info
        lines.append("="*70)
        lines.append("OBJECT INFO")
        lines.append("="*70)
        lines.append(f"Name: {hsf_folder.name}")
        lines.append(f"Source: {self.gsm_file}")
        lines.append("")

        # Parameters (paramlist.xml)
        paramlist_file = hsf_folder / "paramlist.xml"
        if paramlist_file.exists():
            lines.append("="*70)
            lines.append("PARAMETERS (paramlist)")
            lines.append("="*70)
            lines.append(
                "You can modify parameter types, names, values, and descriptions below")
            lines.append("")
            with open(paramlist_file, 'r', encoding='utf-8', errors='ignore') as f:
                lines.append(f.read())
            lines.append("")

        # Scripts
        scripts_dir = hsf_folder / "scripts"
        if scripts_dir.exists():
            script_files = sorted(scripts_dir.iterdir())

            # Define script order and descriptions
            script_info = {
                'm.gdl': ('Master Script', 'Main script, always executed'),
                'p.gdl': ('Parameters Script', 'Calculate and adjust parameter values'),
                '3d.gdl': ('3D Script', 'Generate 3D geometry'),
                '2d.gdl': ('2D Script', 'Generate 2D representation'),
                '1d.gdl': ('1D Script', 'Generate linear representation'),
                'ui.gdl': ('User Interface Script', 'Create parameter interface'),
                'vl.gdl': ('Value List Script', 'Define dropdown lists'),
                'pr.gdl': ('Properties Script', 'Object properties'),
                'if.gdl': ('Interface Script', 'Advanced interface'),
                'fw.gdl': ('Forward Migration Script', 'Upgrade to newer versions'),
                'bw.gdl': ('Backward Migration Script', 'Downgrade compatibility'),
                'mi.gdl': ('Migration Script', 'General migration'),
            }

            for script_file in script_files:
                if script_file.suffix == '.gdl':
                    script_name = script_file.name
                    script_title, script_desc = script_info.get(
                        script_name, (script_name, ''))

                    lines.append("="*70)
                    lines.append(f"SCRIPT: {script_title} ({script_name})")
                    if script_desc:
                        lines.append(f"Purpose: {script_desc}")
                    lines.append("="*70)

                    try:
                        with open(script_file, 'r', encoding='utf-8', errors='ignore') as f:
                            content = f.read()
                            if content.strip():
                                lines.append(content)
                            else:
                                lines.append("! Empty script")
                    except Exception as e:
                        lines.append(f"! Error reading script: {e}")

                    lines.append("")

        # Metadata info (informational only, not editable)
        lines.append("="*70)
        lines.append("METADATA (Informational - Not Editable)")
        lines.append("="*70)

        for meta_file in ['ancestry.xml', 'calledmacros.xml', 'libpartdata.xml']:
            meta_path = hsf_folder / meta_file
            if meta_path.exists():
                lines.append(f"\n--- {meta_file} ---")
                try:
                    with open(meta_path, 'r', encoding='utf-8', errors='ignore') as f:
                        content = f.read()
                        lines.append(content)
                except:
                    lines.append("(binary or unreadable)")

        lines.append("")
        lines.append("="*70)
        lines.append("END OF OBJECT")
        lines.append("="*70)

        return "\n".join(lines)

    def text_to_hsf(self, text_content, output_hsf_path):
        """Parse text format back to HSF folder structure"""
        hsf_folder = Path(output_hsf_path)
        scripts_dir = hsf_folder / "scripts"
        scripts_dir.mkdir(parents=True, exist_ok=True)

        lines = text_content.split('\n')
        current_section = None
        current_content = []
        sections_saved = []  # Track what we save for debugging

        print(f"\n=== PARSING DEBUG ===")
        print(f"Total lines to parse: {len(lines)}")

        # Debug: Look for section markers
        marker_pattern = "="*70
        marker_lines = []
        for idx, line in enumerate(lines[:50]):  # First 50 lines
            stripped = line.strip()
            if stripped.startswith("="):
                marker_lines.append((idx, len(stripped), stripped[:80]))

        print(
            f"Found {len(marker_lines)} lines starting with '=' in first 50 lines:")
        for idx, length, preview in marker_lines:
            print(f"  Line {idx}: {length} chars: {preview}")

        i = 0
        while i < len(lines):
            line = lines[i]
            stripped_line = line.strip()

            # Detect section markers - be more flexible
            # Check if line is mostly equal signs (at least 60 '=' chars)
            if stripped_line.startswith("=") and len(stripped_line) >= 60:
                print(
                    f"\n[Line {i}] Found marker with {len(stripped_line)} '=' chars")

                # Save previous section before starting new one
                if current_section and current_content:
                    content_to_save = '\n'.join(current_content).strip()
                    if content_to_save:  # Only save non-empty content
                        try:
                            self.save_section(
                                hsf_folder, current_section, content_to_save)
                            sections_saved.append(
                                f"{current_section[0]}: {current_section[1]}")
                            print(
                                f"  ✓ Saved previous section: {current_section}")
                        except Exception as e:
                            print(
                                f"  ✗ Warning: Could not save section {current_section}: {e}")
                    current_content = []

                # Look ahead for section title
                if i + 1 < len(lines):
                    next_line = lines[i + 1].strip()
                    print(f"  Next line: {next_line[:80]}")

                    # Skip non-editable sections
                    skip_sections = ["METADATA", "END OF OBJECT", "OBJECT INFO",
                                     "GDL OBJECT - AI-EDITABLE FORMAT", "INSTRUCTIONS"]
                    if any(skip_word in next_line for skip_word in skip_sections):
                        print(f"  -> SKIPPING (non-editable section)")
                        current_section = None
                        i += 1
                        continue

                    # Detect SCRIPT sections
                    if next_line.startswith("SCRIPT:"):
                        # Format: "SCRIPT: Title (filename.gdl)"
                        if '(' in next_line and ')' in next_line:
                            start_paren = next_line.index('(')
                            end_paren = next_line.index(')', start_paren)
                            filename = next_line[start_paren +
                                                 1:end_paren].strip()
                            current_section = ('script', filename)
                            print(f"  -> SCRIPT: {filename}")

                            # Skip: title line, Purpose line if present, then the next marker line
                            i += 1  # Skip to title line
                            if i + 1 < len(lines) and lines[i + 1].strip().startswith("Purpose:"):
                                i += 1  # Skip Purpose line
                                print(f"    Skipped Purpose line")
                            # Now skip the next separator line (====)
                            if i + 1 < len(lines):
                                next_stripped = lines[i + 1].strip()
                                if next_stripped.startswith("=") and len(next_stripped) >= 60:
                                    i += 1  # Skip separator
                                    print(f"    Skipped separator at line {i}")
                            i += 1  # Move to first content line
                            continue

                    # Detect PARAMETERS section
                    elif "PARAMETERS" in next_line and "paramlist" in next_line:
                        current_section = ('paramlist', None)
                        print(f"  -> PARAMETERS")

                        # Skip: title line, separator, description line
                        i += 1  # Skip PARAMETERS line
                        if i + 1 < len(lines):
                            # Skip the next separator (====)
                            next_stripped = lines[i + 1].strip()
                            if next_stripped.startswith("=") and len(next_stripped) >= 60:
                                i += 1  # Skip separator
                                print(f"    Skipped separator at line {i}")
                            # Skip description line if present
                            if i + 1 < len(lines) and "modify parameter" in lines[i + 1].lower():
                                i += 1  # Skip description
                                print(f"    Skipped description at line {i}")
                        i += 1  # Move to first content line
                        continue

                    else:
                        print(f"  -> UNKNOWN section, skipping")
                        current_section = None

                i += 1
                continue

            # Collect content for current section
            if current_section:
                current_content.append(line)

            i += 1

        # Save last section
        if current_section and current_content:
            content_to_save = '\n'.join(current_content).strip()
            if content_to_save:
                try:
                    self.save_section(
                        hsf_folder, current_section, content_to_save)
                    sections_saved.append(
                        f"{current_section[0]}: {current_section[1]}")
                    print(f"\n✓ Saved final section: {current_section}")
                except Exception as e:
                    print(
                        f"\n✗ Warning: Could not save final section {current_section}: {e}")

        # Debug output
        print(f"\n=== PARSING COMPLETE ===")
        print(f"Sections parsed and saved: {len(sections_saved)}")
        for section in sections_saved:
            print(f"  - {section}")

        if not sections_saved:
            print("\n!!! NO SECTIONS FOUND !!!")
            print("Possible issues:")
            print("  1. Section markers might not have exactly 70 '=' signs")
            print("  2. Text format might be corrupted")
            print("  3. Copy/paste might have altered the format")
            raise Exception(
                "No sections were parsed! Check that the text format has correct section markers (at least 60 equal signs)")

    def save_section(self, hsf_folder, section, content):
        """Save a section to appropriate file"""
        section_type, filename = section

        # Clean up content (remove leading/trailing empty lines but preserve internal structure)
        content = content.strip()

        if not content:
            print(f"Warning: Empty content for section {section}")
            return

        if section_type == 'script':
            script_file = hsf_folder / "scripts" / filename
            print(f"Saving script: {filename} ({len(content)} chars)")
            with open(script_file, 'w', encoding='utf-8') as f:
                f.write(content)
            # Verify it was written
            if script_file.exists():
                written_size = script_file.stat().st_size
                print(f"  ✓ Written: {written_size} bytes")
            else:
                print(f"  ✗ Failed to write file")

        elif section_type == 'paramlist':
            paramlist_file = hsf_folder / "paramlist.xml"
            print(f"Saving paramlist ({len(content)} chars)")
            with open(paramlist_file, 'w', encoding='utf-8') as f:
                f.write(content)
            # Verify it was written
            if paramlist_file.exists():
                written_size = paramlist_file.stat().st_size
                print(f"  ✓ Written: {written_size} bytes")
            else:
                print(f"  ✗ Failed to write file")

    def copy_to_clipboard(self):
        """Copy extracted text to clipboard"""
        content = self.extracted_text.get(1.0, tk.END)
        if content.strip():
            self.root.clipboard_clear()
            self.root.clipboard_append(content)
            self.update_status("✓ Copied to clipboard!", "green")
        else:
            messagebox.showwarning(
                "Warning", "No content to copy. Extract an object first.")

    def show_format_debug(self):
        """Show format debugging information"""
        content = self.extracted_text.get(1.0, tk.END)
        if not content.strip():
            messagebox.showwarning("Warning", "No content to analyze")
            return

        lines = content.split('\n')

        # Analyze section markers
        markers = []
        for idx, line in enumerate(lines[:100]):  # First 100 lines
            stripped = line.strip()
            if stripped.startswith("=") and len(stripped) >= 60:
                markers.append({
                    'line': idx,
                    'length': len(stripped),
                    'preview': stripped[:80]
                })

        # Build debug info
        debug_info = "FORMAT DEBUG ANALYSIS\n"
        debug_info += "="*50 + "\n\n"
        debug_info += f"Total lines: {len(lines)}\n"
        debug_info += f"Section markers found (first 100 lines): {len(markers)}\n\n"

        if markers:
            debug_info += "Section markers:\n"
            for m in markers[:10]:  # First 10 markers
                debug_info += f"  Line {m['line']}: {m['length']} chars\n"
                debug_info += f"    {m['preview']}\n"
        else:
            debug_info += "⚠ NO SECTION MARKERS FOUND!\n"
            debug_info += "This means the text format is not correct.\n\n"
            debug_info += "First 10 lines of content:\n"
            for idx, line in enumerate(lines[:10]):
                debug_info += f"  {idx}: {line[:80]}\n"

        # Show in messagebox
        messagebox.showinfo("Format Debug", debug_info)

        # Also print to console
        print("\n" + debug_info)

    def paste_from_clipboard(self):
        """Paste from clipboard to modified text area"""
        try:
            content = self.root.clipboard_get()
            self.modified_text.delete(1.0, tk.END)
            self.modified_text.insert(1.0, content)
            self.update_status("✓ Pasted from clipboard", "green")
        except:
            messagebox.showwarning("Warning", "Clipboard is empty")

    def save_to_file(self):
        """Save extracted content to file"""
        content = self.extracted_text.get(1.0, tk.END)
        if not content.strip():
            messagebox.showwarning("Warning", "No content to save")
            return

        file_path = filedialog.asksaveasfilename(
            title="Save Content",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )

        if file_path:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(content)
            self.update_status(
                f"✓ Saved to {os.path.basename(file_path)}", "green")

    def load_from_file(self):
        """Load modified content from file"""
        file_path = filedialog.askopenfilename(
            title="Load Modified Content",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )

        if file_path:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            self.modified_text.delete(1.0, tk.END)
            self.modified_text.insert(1.0, content)
            self.update_status(
                f"✓ Loaded from {os.path.basename(file_path)}", "green")

    def test_parse(self):
        """Test parsing without compiling - for debugging"""
        if not self.hsf_path:
            messagebox.showerror("Error", "Please extract an object first")
            return

        modified_content = self.modified_text.get(1.0, tk.END)
        if not modified_content.strip():
            messagebox.showerror("Error", "No content to parse")
            return

        self.update_status("Testing parse...", "blue")

        # Create a temporary test directory
        test_dir = tempfile.mkdtemp(prefix="gdl_parse_test_")

        try:
            # Try to parse
            self.text_to_hsf(modified_content, test_dir)

            # Count what was created
            test_path = Path(test_dir)
            scripts_dir = test_path / "scripts"

            script_count = 0
            if scripts_dir.exists():
                script_count = len(list(scripts_dir.glob("*.gdl")))

            paramlist_exists = (test_path / "paramlist.xml").exists()

            msg = f"Parse test successful!\n\n"
            msg += f"Found {script_count} script(s)\n"
            msg += f"Paramlist: {'YES' if paramlist_exists else 'NO'}\n\n"

            if script_count > 0:
                msg += "Scripts:\n"
                for script_file in sorted(scripts_dir.glob("*.gdl")):
                    size = script_file.stat().st_size
                    msg += f"  - {script_file.name}: {size} bytes\n"

            self.update_status("✓ Parse test successful", "green")
            messagebox.showinfo("Parse Test Results", msg)

        except Exception as e:
            self.update_status(f"✗ Parse test failed: {e}", "red")
            messagebox.showerror("Parse Test Failed",
                                 f"Could not parse the modified content:\n\n{str(e)}\n\n"
                                 f"Check the console output for details.")
        finally:
            # Clean up test directory
            try:
                shutil.rmtree(test_dir)
            except:
                pass

    def compile_and_reload(self):
        """Compile modified content and reload in Archicad"""
        if not self.gsm_file or not self.hsf_path:
            messagebox.showerror("Error", "Please extract an object first")
            return

        modified_content = self.modified_text.get(1.0, tk.END)
        if not modified_content.strip():
            messagebox.showerror("Error", "No modified content to compile")
            return

        self.update_status("Parsing modified content...", "blue")
        print("\n" + "="*50)
        print("COMPILATION STARTED")
        print("="*50)

        try:
            # Parse text back to HSF
            try:
                print("\n[1/4] Parsing text to HSF format...")
                self.text_to_hsf(modified_content, self.hsf_path)
                print("✓ Parsing complete")
            except Exception as e:
                raise Exception(f"Failed to parse modified content: {str(e)}")

            self.update_status("Compiling to .gsm...", "blue")

            # Determine output path
            gsm_path = Path(self.gsm_file)
            output_gsm = gsm_path.parent / gsm_path.name

            # Backup existing file (copy, don't rename)
            print(f"\n[2/4] Creating backup of {gsm_path.name}...")
            if output_gsm.exists():
                backup_path = Path(str(output_gsm) + ".bak")
                counter = 1
                while backup_path.exists():
                    backup_path = Path(f"{output_gsm}.bak.{counter}")
                    counter += 1
                try:
                    # Copy the file instead of renaming it
                    # This way Archicad keeps the reference to the original file
                    shutil.copy2(output_gsm, backup_path)
                    self.update_status(
                        f"Created backup: {backup_path.name}", "blue")
                    print(f"✓ Backup created: {backup_path.name}")
                except Exception as e:
                    raise Exception(f"Failed to create backup: {str(e)}")

            # Compile HSF to GSM
            print(f"\n[3/4] Compiling HSF to GSM...")
            cmd = [
                self.converter_path,
                "hsf2libpart",
                self.hsf_path,
                str(output_gsm)
            ]
            print(f"Command: {' '.join(cmd)}")

            try:
                result = subprocess.run(
                    cmd, capture_output=True, text=True, check=True)
                print(f"✓ Compilation successful")
                if result.stdout:
                    print(f"Converter output: {result.stdout}")
            except subprocess.CalledProcessError as e:
                error_msg = f"LP_XMLConverter failed:\n"
                error_msg += f"Command: {' '.join(cmd)}\n"
                error_msg += f"Return code: {e.returncode}\n"
                if e.stdout:
                    error_msg += f"Stdout: {e.stdout}\n"
                if e.stderr:
                    error_msg += f"Stderr: {e.stderr}\n"
                raise Exception(error_msg)

            self.update_status("Reloading libraries in Archicad...", "blue")

            # Reload libraries
            print(f"\n[4/4] Reloading libraries in Archicad...")
            try:
                conn = ACConnection.connect()
                acc = conn.commands
                act = conn.types

                response = acc.ExecuteAddOnCommand(
                    act.AddOnCommandId('TapirCommand', 'ReloadLibraries')
                )

                print(f"✓ Libraries reloaded successfully")
                print("\n" + "="*50)
                print("COMPILATION COMPLETED SUCCESSFULLY")
                print("="*50 + "\n")

                self.update_status(
                    "✓ Compilation complete! Object reloaded in Archicad.", "green")
                messagebox.showinfo("Success",
                                    f"Object compiled successfully!\n\n"
                                    f"File: {output_gsm.name}\n"
                                    f"Libraries reloaded in Archicad.\n\n"
                                    f"You can now test your modifications!")

            except Exception as e:
                print(f"⚠ Library reload failed: {e}")
                print("\n" + "="*50)
                print("COMPILATION COMPLETED (manual reload needed)")
                print("="*50 + "\n")

                self.update_status(
                    f"✓ Compiled, but reload failed: {e}", "orange")
                messagebox.showwarning("Partial Success",
                                       f"Object compiled successfully to:\n{output_gsm.name}\n\n"
                                       f"But library reload failed. You can reload manually in Archicad.")

        except Exception as e:
            print(f"\n✗ COMPILATION FAILED")
            print(f"Error: {e}")
            print("="*50 + "\n")
            self.update_status(f"✗ Compilation failed: {e}", "red")
            messagebox.showerror("Compilation Error", str(e))

    def cleanup(self):
        """Clean up temporary files"""
        if self.temp_dir and os.path.exists(self.temp_dir):
            try:
                shutil.rmtree(self.temp_dir)
            except:
                pass

    def on_closing(self):
        """Handle window closing"""
        self.cleanup()
        self.root.destroy()


def main():
    root = Tk()
    app = AIGDLEditor(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()


if __name__ == "__main__":
    main()