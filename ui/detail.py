import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import json
from core.gemini import call_gemini
from core.agent import execute_commands
from core.app_state import AppState
from .choice import ChoiceScreen, ShiningButton
import logging
import os
from PIL import Image, ImageTk, ImageFilter
import tempfile

class LoadingOverlay(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        
        # Configure the overlay window
        self.overrideredirect(True)
        self.attributes('-alpha', 0.85)  # Slightly more transparent
        self.attributes('-topmost', True)
        
        # Get parent window dimensions and position
        x = parent.winfo_x()
        y = parent.winfo_y()
        width = parent.winfo_width()
        height = parent.winfo_height()
        
        # Position overlay over parent window
        self.geometry(f"{width}x{height}+{x}+{y}")
        
        # Create main container with dark background
        self.frame = tk.Frame(self, bg='#1a1a1a')
        self.frame.place(relwidth=1, relheight=1)
        
        # Create a semi-transparent overlay for blur effect
        self.blur_frame = tk.Frame(self.frame, bg='#000000')
        self.blur_frame.place(relwidth=1, relheight=1)
        self.blur_frame.configure(bg='#000000')
        
        # Create content container with a border
        self.content_frame = tk.Frame(
            self.frame,
            bg='#1a1a1a',
            bd=2,
            highlightthickness=2,
            highlightbackground='#4CAF50'
        )
        self.content_frame.place(relx=0.5, rely=0.5, anchor="center", relwidth=0.4, relheight=0.3)
        
        # Add loading text with animated dots
        self.loading_text = "Agent is Working"
        self.dots = ""
        self.loading_label = tk.Label(
            self.content_frame,
            text=f"{self.loading_text}",
            font=("Helvetica", 20, "bold"),
            fg="#4CAF50",
            bg='#1a1a1a'
        )
        self.loading_label.place(relx=0.5, rely=0.4, anchor="center")
        
        # Add status text
        self.status_label = tk.Label(
            self.content_frame,
            text="Generating your presentation...",
            font=("Helvetica", 12),
            fg="#BDBDBD",
            bg='#1a1a1a'
        )
        self.status_label.place(relx=0.5, rely=0.6, anchor="center")
        
        # Add a modern progress bar
        self.progress_frame = tk.Frame(
            self.content_frame,
            bg='#1a1a1a',
            bd=0,
            highlightthickness=0
        )
        self.progress_frame.place(relx=0.5, rely=0.75, anchor="center", relwidth=0.8, relheight=0.1)
        
        # Custom progress bar style
        self.style = ttk.Style()
        self.style.configure(
            "Custom.Horizontal.TProgressbar",
            troughcolor='#2d2d2d',
            background='#4CAF50',
            thickness=8,
            borderwidth=0
        )
        
        self.progress = ttk.Progressbar(
            self.progress_frame,
            style="Custom.Horizontal.TProgressbar",
            mode='indeterminate',
            length=300
        )
        self.progress.place(relx=0.5, rely=0.5, anchor="center", relwidth=1, relheight=1)
        self.progress.start(15)  # Faster animation
        
        # Start the dots animation
        self.animate_dots()
        
        # Make sure the overlay is visible
        self.lift()
        self.update()
    
    def animate_dots(self):
        if not self.winfo_exists():
            return
            
        self.dots = self.dots + "." if len(self.dots) < 3 else ""
        self.loading_label.configure(text=f"{self.loading_text}{self.dots}")
        self.after(500, self.animate_dots)
    
    def destroy(self):
        self.progress.stop()
        super().destroy()

class DetailScreen:
    def __init__(self, root, doc_type):
        self.root = root
        self.doc_type = doc_type
        self.frame = tk.Frame(root, bg="#1a1a1a")  # Dark background
        self.app_state = AppState()
        
        # Configure style
        self.style = ttk.Style()
        self.style.configure(
            "Custom.TButton",
            font=("Helvetica", 10),
            padding=10,
            background="#4CAF50"
        )
        self.style.configure(
            "Custom.TCheckbutton",
            background="#1a1a1a",
            foreground="#BDBDBD"
        )
        
        # Create header frame
        self.header_frame = tk.Frame(self.frame, bg="#2d2d2d", height=60)  # Darker header
        self.header_frame.pack(fill="x", pady=(0, 20))
        self.header_frame.pack_propagate(False)
        
        # Back button with 3D effect
        self.back_btn = ShiningButton(
            self.header_frame,
            text="← Back",
            command=self._go_back,
            width=100,
            height=35
        )
        self.back_btn.pack(side="left", padx=20)
        
        # Title in header
        title = tk.Label(
            self.header_frame,
            text=f"Create {doc_type.upper()} Document",
            font=("Helvetica", 16, "bold"),
            bg="#2d2d2d",
            fg="#4CAF50"  # Green text like choice screen
        )
        title.pack(side="left", padx=20)
        
        # Main content frame
        self.content_frame = tk.Frame(self.frame, bg="#1a1a1a", padx=40, pady=20)  # Dark background
        self.content_frame.pack(fill="both", expand=True)
        
        # Create options frame based on document type
        if doc_type.lower() == "ppt":
            self._create_ppt_options()
        elif doc_type.lower() == "word":
            self._create_word_options()
        
        # Instructions label
        instructions = tk.Label(
            self.content_frame,
            text="Describe what you want to create in detail. Be specific about the content, structure, and any special requirements.",
            font=("Helvetica", 11),
            fg="#BDBDBD",  # Light gray text
            bg="#1a1a1a",
            wraplength=600,
            justify="left"
        )
        instructions.pack(pady=(0, 20))
        
        # Text area with custom styling
        self.text_frame = tk.Frame(self.content_frame, bg="#2d2d2d", bd=1, relief="solid")  # Darker frame
        self.text_frame.pack(fill="both", expand=True, pady=(0, 20))
        
        self.prompt_text = tk.Text(
            self.text_frame,
            wrap="word",
            font=("Helvetica", 11),
            padx=15,
            pady=15,
            bg="#2d2d2d",  # Dark background
            fg="#BDBDBD",  # Light gray text
            relief="flat",
            height=12,
            insertbackground="#BDBDBD"  # Cursor color
        )
        self.prompt_text.pack(fill="both", expand=True, padx=1, pady=1)
        
        # Add placeholder text
        self.prompt_text.insert("1.0", "Enter your document requirements here...")
        self.prompt_text.config(fg="#666666")  # Darker gray for placeholder
        
        # Bind focus events for placeholder
        self.prompt_text.bind("<FocusIn>", self._on_focus_in)
        self.prompt_text.bind("<FocusOut>", self._on_focus_out)
        
        # Button frame
        self.button_frame = tk.Frame(self.content_frame, bg="#1a1a1a")
        self.button_frame.pack(fill="x", pady=(0, 20))
        
        # Generate button with 3D effect
        self.generate_btn = ShiningButton(
            self.button_frame,
            text="Generate Document",
            command=self._on_generate,
            width=220,
            height=45
        )
        self.generate_btn.pack(side="right")
        
        # Status label
        self.status_label = tk.Label(
            self.content_frame,
            text="",
            font=("Helvetica", 10),
            fg="#BDBDBD",  # Light gray text
            bg="#1a1a1a"
        )
        self.status_label.pack(pady=(0, 10))
        
        # Add loading overlay reference
        self.loading_overlay = None

        # Initialize document style variable
        self.doc_style_var = tk.StringVar(value="Professional")  # Initialize doc_style_var

        # Ensure page_size_var is initialized in the constructor
        self.page_size_var = tk.StringVar(value="A4")
        self.page_size_var.set("A4")

        # Ensure alignment_var is initialized in the constructor
        self.alignment_var = tk.StringVar(value="None")

    def _create_ppt_options(self):
        # Options frame
        options_frame = tk.LabelFrame(
            self.content_frame,
            text="Presentation Options",
            font=("Helvetica", 11, "bold"),
            bg="#2d2d2d",  # Darker background
            fg="#4CAF50",  # Green text
            padx=20,
            pady=15
        )
        options_frame.pack(fill="x", pady=(0, 20))
        
        # Grid layout for options
        options_frame.grid_columnconfigure(1, weight=1)
        options_frame.grid_columnconfigure(3, weight=1)
        
        # Style for labels
        label_style = {"font": ("Helvetica", 10), "bg": "#2d2d2d", "fg": "#BDBDBD"}
        
        # Row 1: Slide Count and Theme
        # Slide Count
        tk.Label(
            options_frame,
            text="Number of Slides:",
            **label_style
        ).grid(row=0, column=0, padx=(0, 10), pady=5, sticky="e")
        
        self.slide_count = ttk.Spinbox(
            options_frame,
            from_=1,
            to=50,
            width=5,
            font=("Helvetica", 10)
        )
        self.slide_count.set(5)
        self.slide_count.grid(row=0, column=1, padx=(0, 20), pady=5, sticky="w")
        
        # Theme Selection
        tk.Label(
            options_frame,
            text="Theme:",
            **label_style
        ).grid(row=0, column=2, padx=(0, 10), pady=5, sticky="e")
        
        self.theme_var = tk.StringVar(value="Professional")
        themes = ["Professional", "Creative", "Corporate", "Modern", "Elegant"]
        self.theme_combo = ttk.Combobox(
            options_frame,
            textvariable=self.theme_var,
            values=themes,
            state="readonly",
            width=15,
            font=("Helvetica", 10)
        )
        self.theme_combo.grid(row=0, column=3, padx=(0, 0), pady=5, sticky="w")
        
        # Row 2: Font and Color Scheme
        # Font Selection
        tk.Label(
            options_frame,
            text="Font:",
            **label_style
        ).grid(row=1, column=0, padx=(0, 10), pady=5, sticky="e")
        
        self.font_var = tk.StringVar(value="Calibri")
        fonts = ["Calibri", "Arial", "Helvetica", "Times New Roman", "Verdana"]
        self.font_combo = ttk.Combobox(
            options_frame,
            textvariable=self.font_var,
            values=fonts,
            state="readonly",
            width=15,
            font=("Helvetica", 10)
        )
        self.font_combo.grid(row=1, column=1, padx=(0, 20), pady=5, sticky="w")
        
        # Color Scheme
        tk.Label(
            options_frame,
            text="Color Scheme:",
            **label_style
        ).grid(row=1, column=2, padx=(0, 10), pady=5, sticky="e")
        
        self.color_scheme_var = tk.StringVar(value="Default")
        color_schemes = ["Default", "Blue", "Green", "Red", "Purple", "Orange"]
        self.color_scheme_combo = ttk.Combobox(
            options_frame,
            textvariable=self.color_scheme_var,
            values=color_schemes,
            state="readonly",
            width=15,
            font=("Helvetica", 10)
        )
        self.color_scheme_combo.grid(row=1, column=3, padx=(0, 0), pady=5, sticky="w")
        
        # Row 3: Additional Options
        # Include Table of Contents
        # self.toc_var = tk.BooleanVar(value=True)
        # ttk.Checkbutton(
        #     options_frame,
        #     text="Include Table of Contents",
        #     variable=self.toc_var,
        #     style="Custom.TCheckbutton"
        # ).grid(row=2, column=0, columnspan=2, padx=(0, 20), pady=5, sticky="w")
        
        # Include Speaker Notes
        self.notes_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            options_frame,
            text="Include Speaker Notes",
            variable=self.notes_var,
            style="Custom.TCheckbutton"
        ).grid(row=2, column=2, columnspan=2, padx=(0, 0), pady=5, sticky="w")

    def _create_word_options(self):
        # Options frame
        options_frame = tk.LabelFrame(
            self.content_frame,
            text="Document Options",
            font=("Helvetica", 11, "bold"),
            bg="#2d2d2d",
            fg="#4CAF50",
            padx=20,
            pady=15
        )
        options_frame.pack(fill="x", pady=(0, 20))
        
        # Grid layout for options
        options_frame.grid_columnconfigure(1, weight=1)
        options_frame.grid_columnconfigure(3, weight=1)
        
        # Style for labels
        label_style = {"font": ("Helvetica", 10), "bg": "#2d2d2d", "fg": "#BDBDBD"}
        
        # Row 1: Document Style and Font
        # Document Style
        tk.Label(
            options_frame,
            text="Document Style:",
            **label_style
        ).grid(row=0, column=0, padx=(0, 10), pady=5, sticky="e")
        
        self.doc_style_var = tk.StringVar(value="Professional")
        doc_styles = ["Professional", "Academic", "Business", "Creative", "Technical"]
        self.doc_style_combo = ttk.Combobox(
            options_frame,
            textvariable=self.doc_style_var,
            values=doc_styles,
            state="readonly",
            width=15,
            font=("Helvetica", 10)
        )
        self.doc_style_combo.grid(row=0, column=1, padx=(0, 20), pady=5, sticky="w")
        
        # Font Selection
        tk.Label(
            options_frame,
            text="Font:",
            **label_style
        ).grid(row=0, column=2, padx=(0, 10), pady=5, sticky="e")
        
        self.font_var = tk.StringVar(value="None")
        fonts = ["None", "Calibri", "Arial", "Times New Roman", "Georgia", "Verdana"]
        self.font_combo = ttk.Combobox(
            options_frame,
            textvariable=self.font_var,
            values=fonts,
            state="readonly",
            width=15,
            font=("Helvetica", 10)
        )
        self.font_combo.grid(row=0, column=3, padx=(0, 0), pady=5, sticky="w")
        
        # Row 2: Page Size and Text Alignment
        # Page Size
        tk.Label(
            options_frame,
            text="Page Size:",
            **label_style
        ).grid(row=1, column=0, padx=(0, 10), pady=5, sticky="e")
        
        self.page_size_var = tk.StringVar(value="A4")
        # Ensure page_size_var is initialized in the constructor
        self.page_size_var.set("A4")
        page_sizes = ["A4", "Letter", "Legal", "A3", "A5"]
        self.page_size_combo = ttk.Combobox(
            options_frame,
            textvariable=self.page_size_var,
            values=page_sizes,
            state="readonly",
            width=15,
            font=("Helvetica", 10)
        )
        self.page_size_combo.grid(row=1, column=1, padx=(0, 20), pady=5, sticky="w")
        
        # Text Alignment
        tk.Label(
            options_frame,
            text="Text Alignment:",
            **label_style
        ).grid(row=1, column=2, padx=(0, 10), pady=5, sticky="e")
        
        self.alignment_var = tk.StringVar(value="None")
        alignments = ["None", "Left", "Center", "Right", "Justify", "Distribute"]
        self.alignment_combo = ttk.Combobox(
            options_frame,
            textvariable=self.alignment_var,
            values=alignments,
            state="readonly",
            width=15,
            font=("Helvetica", 10)
        )
        self.alignment_combo.grid(row=1, column=3, padx=(0, 0), pady=5, sticky="w")
        
        # Row 3: Additional Options
        # Include Table of Contents
        # self.toc_var = tk.BooleanVar(value=True)
        # ttk.Checkbutton(
        #     options_frame,
        #     text="Include Table of Contents",
        #     variable=self.toc_var,
        #     style="Custom.TCheckbutton"
        # ).grid(row=2, column=0, columnspan=2, padx=(0, 20), pady=5, sticky="w")
        
        # Include Page Numbers
        self.page_num_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            options_frame,
            text="Include Page Numbers",
            variable=self.page_num_var,
            style="Custom.TCheckbutton"
        ).grid(row=2, column=2, columnspan=2, padx=(0, 0), pady=5, sticky="w")
        
        # Row 4: More Options
        # Header and Footer
        self.header_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            options_frame,
            text="Include Header",
            variable=self.header_var,
            style="Custom.TCheckbutton"
        ).grid(row=3, column=0, columnspan=2, padx=(0, 20), pady=5, sticky="w")
        
        self.footer_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            options_frame,
            text="Include Footer",
            variable=self.footer_var,
            style="Custom.TCheckbutton"
        ).grid(row=3, column=2, columnspan=2, padx=(0, 0), pady=5, sticky="w")

    def show(self):
        self.frame.pack(fill="both", expand=True)

    def _go_back(self):
        # Clear current screen
        for widget in self.root.winfo_children():
            widget.destroy()
        
        # Create and show choice screen
        choice_screen = ChoiceScreen(self.root, on_select=self._handle_selection)
        choice_screen.pack(fill="both", expand=True)
    
    def _handle_selection(self, doc_type):
        # Clear current screen
        for widget in self.root.winfo_children():
            widget.destroy()
        
        # Create and show new detail screen
        detail_screen = DetailScreen(self.root, doc_type)
        detail_screen.show()

    def _on_focus_in(self, event):
        if self.prompt_text.get("1.0", "end-1c") == "Enter your document requirements here...":
            self.prompt_text.delete("1.0", "end")
            self.prompt_text.config(fg="#BDBDBD")  # Light gray text when focused

    def _on_focus_out(self, event):
        if not self.prompt_text.get("1.0", "end-1c").strip():
            self.prompt_text.insert("1.0", "Enter your document requirements here...")
            self.prompt_text.config(fg="#666666")  # Darker gray for placeholder

    def _show_loading_overlay(self):
        if self.loading_overlay is None:
            self.loading_overlay = LoadingOverlay(self.root)
            self.root.update()
    
    def _hide_loading_overlay(self):
        if self.loading_overlay:
            self.loading_overlay.destroy()
            self.loading_overlay = None
            self.root.update()

    def _on_generate(self):
        prompt = self.prompt_text.get("1.0", tk.END).strip()
        if prompt == "Enter your document requirements here...":
            prompt = ""
            
        if not prompt:
            messagebox.showerror("Error", "Please enter your document requirements.")
            return

        # Show loading overlay
        self._show_loading_overlay()
        
        # Disable the generate button
        self.generate_btn.configure(state="disabled")
        
        # Start processing in a separate thread
        threading.Thread(target=self._process_prompt, args=(prompt,), daemon=True).start()

    def _process_prompt(self, prompt):
        try:
            # Suppress all logging output
            logging.getLogger().setLevel(logging.ERROR)
            
            # Redirect stdout temporarily to suppress print statements
            import sys
            from io import StringIO
            old_stdout = sys.stdout
            sys.stdout = StringIO()
            
            try:
                gen_prompt = self._construct_prompt(prompt)
                gemini_full_response = call_gemini(gen_prompt)
                
                if not gemini_full_response or "candidates" not in gemini_full_response or not gemini_full_response["candidates"]:
                    raise RuntimeError("Gemini API returned an unexpected or empty response structure.")

                extracted_text = gemini_full_response["candidates"][0].get("content", {}).get("parts", [{}])[0].get("text", "")
                
                if not extracted_text.strip():
                    raise RuntimeError("Gemini API returned empty text content in the expected field.")

                # Clean and parse JSON
                cleaned_json_string = extracted_text.strip()
                if cleaned_json_string.startswith('```json'):
                    cleaned_json_string = cleaned_json_string[len('```json'):].strip()
                if cleaned_json_string.endswith('```'):
                    cleaned_json_string = cleaned_json_string[:-len('```')].strip()

                command_json = json.loads(cleaned_json_string)
                
                self.app_state.update("prompt", prompt)
                self.app_state.update("commands", command_json)
                
                # Generate document
                path = execute_commands(command_json, self.doc_type)
                self.app_state.update("output_file", path)
                self.app_state.save()

            finally:
                # Restore stdout
                sys.stdout = old_stdout
                
            # Restore logging level
            logging.getLogger().setLevel(logging.INFO)
            
            # Show success message and reset UI
            self.root.after(0, self._hide_loading_overlay)
            self.root.after(0, lambda: self.generate_btn.configure(state="normal"))
            self.root.after(0, lambda: messagebox.showinfo("Success", f"Document generated successfully!\nSaved at: {path}"))
            self.root.after(0, self._reset_ui)
            
        except Exception as e:
            # Restore stdout in case of error
            if 'sys' in locals():
                sys.stdout = old_stdout

            # Restore logging level
            logging.getLogger().setLevel(logging.INFO)

            # Hide overlay and show error
            self.root.after(0, self._hide_loading_overlay)
            self.root.after(0, lambda: self.generate_btn.configure(state="normal"))
            self.root.after(0, lambda e=e: messagebox.showerror("Error", f"An error occurred: {str(e)}"))

    def _construct_prompt(self, user_prompt):
        if self.doc_type.lower() == "ppt":
            # Include PowerPoint-specific options in the prompt
            options = {
                "slide_count": self.slide_count.get(),
                "theme": self.theme_var.get(),
                "font": self.font_var.get(),
                "color_scheme": self.color_scheme_var.get(),
                "include_notes": self.notes_var.get()
            }
            
            # Define theme-specific requirements
            theme_requirements = {
                "Professional": "Use simple text layouts with clean transitions. Focus on clear, professional content presentation with minimal design elements.",
                "Creative": "Include creative transitions and shapes. Use visual elements and drawings to make the presentation engaging and artistic.",
                "Corporate": "Incorporate charts, drawings, and smooth transitions. Focus on data visualization and professional corporate style.",
                "Modern": "Use modern design elements, drawings, and transitions. Focus on contemporary aesthetics and visual appeal.",
                "Elegant": "Include sophisticated design elements, charts, drawings, and smooth transitions. Focus on premium, elegant presentation style."
            }
            
            return f"""
You are a PowerPoint presentation generation assistant. Create a professional presentation based on the following requirements.

Document Type: {self.doc_type.upper()}
User Prompt: "{user_prompt}"

Presentation Options:
- Number of Slides: {options['slide_count']}
- Theme: {options['theme']}
- Font: {options['font']}
- Color Scheme: {options['color_scheme']}
- Include Speaker Notes: {options['include_notes']}

Theme Requirements:
{theme_requirements[options['theme']]}

Return a JSON with the structure:
{{
    "title": "Presentation Title",
    "subtitle": "Presentation Subtitle",
    "theme": "{options['theme']}",
    "font": "{options['font']}",
    "color_scheme": "{options['color_scheme']}",
    "include_notes": {str(options['include_notes']).lower()},
    "slides": [
        {{
            "title": "Slide Title",
            "content": "Slide Content",
            "notes": "Speaker Notes (if include_notes is true)",
            "formatting": {{
                "title_font_size": 44,  // Font size for title in points
                "content_font_size": 24,  // Font size for content in points
                "elements": [  // Theme-specific elements to include
                    {{
                        "type": "shape",  // Can be "shape", "chart", "transition", "design"
                        "description": "Description of the element"
                    }}
                ]
            }}
        }},
        ...
    ]
}}

Note: For each slide, specify appropriate font sizes and include theme-specific elements (shapes, charts, transitions, design elements) based on the selected theme.
For Professional theme: Use simple text layouts with clean transitions
For Creative theme: Include creative transitions and shapes
For Corporate theme: Include charts, drawings, and smooth transitions
For Modern theme: Use modern design elements, drawings, and transitions
For Elegant theme: Include sophisticated design elements, charts, drawings, and smooth transitions
"""
        elif self.doc_type.lower() == "word":
            # Include Word-specific options in the prompt
            options = {
                "doc_style": self.doc_style_var.get(),
                "font": self.font_var.get(),
                "page_size": self.page_size_var.get(),
                "text_alignment": self.alignment_var.get(),
                "include_page_numbers": self.page_num_var.get(),
                "include_header": self.header_var.get(),
                "include_footer": self.footer_var.get()
            }
            
            return f"""
You are a Word document generation assistant. Create a professional document based on the following requirements.

Document Type: {self.doc_type.upper()}
User Prompt: "{user_prompt}"

Document Options:
- Document Style: {options['doc_style']}
- Font: {options['font']}
- Page Size: {options['page_size']}
- Text Alignment: {options['text_alignment']}
- Include Page Numbers: {options['include_page_numbers']}
- Include Header: {options['include_header']}
- Include Footer: {options['include_footer']}

Return a JSON with the structure:
{{
    "title": "Document Title",
    "style": "{options['doc_style']}",
    "formatting": {{
        "font": "{options['font']}",
        "page_size": "{options['page_size']}",
        "text_alignment": "{options['text_alignment']}",
        "include_page_numbers": {str(options['include_page_numbers']).lower()},
        "include_header": {str(options['include_header']).lower()},
        "include_footer": {str(options['include_footer']).lower()}
    }},
    "sections": [
        {{
            "heading": "Section Title",
            "content": "Section Content",
            "level": 1
        }},
        ...
    ]
}}
"""
        else:
            # Original prompt for other document types
            return f"""
You are a document generation assistant. Convert the following user request into a JSON structure.
Document Type: {self.doc_type.upper()}
User Prompt: "{user_prompt}"

Return a JSON with the structure:
For Word:
{{"title": ..., "sections": [{{"heading": ..., "content": ...}}, ...]}}

For PPT:
{{"title": ..., "subtitle": ..., "slides": [{{"title": ..., "content": ...}}, ...]}}
"""

    def _reset_ui(self):
        """Reset all UI fields to their default values."""
        self.prompt_text.delete("1.0", tk.END)
        self.prompt_text.insert("1.0", "Enter your document requirements here...")
        self.prompt_text.config(fg="#666666")

        # Reset Word-specific fields if they exist
        if hasattr(self, 'doc_style_var'):
            self.doc_style_var.set("Professional")
        if hasattr(self, 'font_var'):
            self.font_var.set("None")
        if hasattr(self, 'page_size_var'):
            self.page_size_var.set("A4")
        if hasattr(self, 'alignment_var'):
            self.alignment_var.set("None")
        if hasattr(self, 'page_num_var'):
            self.page_num_var.set(True)
        if hasattr(self, 'header_var'):
            self.header_var.set(True)
        if hasattr(self, 'footer_var'):
            self.footer_var.set(True)

        # Reset PowerPoint-specific fields if they exist
        if hasattr(self, 'slide_count'):
            self.slide_count.set(5)
        if hasattr(self, 'theme_var'):
            self.theme_var.set("Professional")
        if hasattr(self, 'color_scheme_var'):
            self.color_scheme_var.set("Default")
        if hasattr(self, 'notes_var'):
            self.notes_var.set(True)


def create_detail_screen(root, doc_type):
    # Clear existing widgets
    for widget in root.winfo_children():
        widget.destroy()
    
    # Create and show detail screen
    detail_screen = DetailScreen(root, doc_type)
    detail_screen.show()
    return detail_screen
