import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk
import webbrowser
import math

class ShiningButton(tk.Canvas):
    def __init__(self, parent, text, command, width=200, height=40, **kwargs):
        super().__init__(parent, width=width, height=height, 
                        bg="#2d2d2d", highlightthickness=0, **kwargs)
        self.command = command
        self.text = text
        self.width = width
        self.height = height
        
        # Button colors with enhanced contrast for 3D effect
        self.normal_color = "#4CAF50"      # Base green
        self.hover_color = "#45a049"       # Darker green for hover
        self.pressed_color = "#388E3C"     # Even darker for press
        self.border_color = "#2E7D32"      # Dark green for 3D border
        self.highlight_color = "#66BB6A"   # Lighter green for top/left edges
        self.shadow_color = "#1B5E20"      # Darker green for bottom/right edges
        
        # Button state
        self.is_pressed = False
        self.is_hovered = False
        
        # Create button elements
        self._create_button()
        
        # Bind events
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<Button-1>", self._on_press)
        self.bind("<ButtonRelease-1>", self._on_release)
    
    def _create_button(self):
        try:
            # Clear previous elements
            self.delete("all")
            
            # Calculate 3D effect with increased depth
            depth = 6  # Increased depth for more pronounced 3D effect
            if self.is_pressed:
                depth = -6
            
            # Draw button shadow (bottom and right edges)
            self.create_rectangle(
                depth, depth, self.width, self.height,
                fill=self.shadow_color,
                outline=""
            )
            
            # Draw button highlight (top and left edges)
            self.create_rectangle(
                0, 0, self.width - abs(depth), self.height - abs(depth),
                fill=self.highlight_color,
                outline=""
            )
            
            # Draw button face with gradient effect
            face_offset = 3
            # Main face
            self.create_rectangle(
                face_offset, face_offset, 
                self.width - abs(depth) - face_offset, 
                self.height - abs(depth) - face_offset,
                fill=self.normal_color if not self.is_hovered else self.hover_color,
                outline=""
            )
            
            # Add subtle gradient effect using multiple rectangles
            gradient_steps = 4
            for i in range(gradient_steps):
                offset = face_offset + i
                color = self._adjust_color(
                    self.normal_color if not self.is_hovered else self.hover_color,
                    -i * 5  # Darken each step
                )
                self.create_rectangle(
                    offset, offset,
                    self.width - abs(depth) - offset,
                    self.height - abs(depth) - offset,
                    fill=color,
                    outline=""
                )
            
            # Draw text with enhanced shadow for better visibility
            # Text shadow
            shadow_color = "#1a1a1a"
            for offset in [(1,1), (2,2)]:
                self.create_text(
                    self.width/2 + offset[0], 
                    self.height/2 + offset[1],
                    text=self.text,
                    font=("Helvetica", 12, "bold"),
                    fill=shadow_color
                )
            
            # Main text
            self.create_text(
                self.width/2, 
                self.height/2,
                text=self.text,
                font=("Helvetica", 12, "bold"),
                fill="white"
            )
            
        except Exception as e:
            print(f"Error in _create_button: {e}")
            # Fallback to a simple button if there's an error
            self.create_rectangle(0, 0, self.width, self.height, fill=self.normal_color, outline="")
            self.create_text(self.width/2, self.height/2, text=self.text, fill="white")
    
    def _adjust_color(self, color, amount):
        """Helper function to adjust color brightness"""
        # Convert hex to RGB
        r = int(color[1:3], 16)
        g = int(color[3:5], 16)
        b = int(color[5:7], 16)
        
        # Adjust each component
        r = max(0, min(255, r + amount))
        g = max(0, min(255, g + amount))
        b = max(0, min(255, b + amount))
        
        # Convert back to hex
        return f"#{r:02x}{g:02x}{b:02x}"
    
    def _on_enter(self, event):
        self.is_hovered = True
        self._create_button()
    
    def _on_leave(self, event):
        self.is_hovered = False
        self.is_pressed = False
        self._create_button()
    
    def _on_press(self, event):
        self.is_pressed = True
        self._create_button()
    
    def _on_release(self, event):
        if self.is_pressed:
            self.is_pressed = False
            self._create_button()
            if self.command:
                self.command()

class ChoiceScreen(ttk.Frame):
    def __init__(self, master, on_select):
        super().__init__(master)
        self.master = master
        self.on_select = on_select
        
        # Configure style
        self.style = ttk.Style()
        self.style.configure(
            "Choice.TFrame",
            background="#1a1a1a"
        )
        self.style.configure(
            "Card.TFrame",
            background="#2d2d2d",
            relief="flat",
            borderwidth=0
        )
        self.style.configure(
            "Title.TLabel",
            font=("Helvetica", 28, "bold"),
            background="#1a1a1a",
            foreground="#4CAF50"
        )
        self.style.configure(
            "Subtitle.TLabel",
            font=("Helvetica", 14),
            background="#1a1a1a",
            foreground="#BDBDBD"
        )
        
        # Configure the frame to expand
        self.pack(fill="both", expand=True)
        
        # Create a canvas with scrollbar for the main content
        self.canvas = tk.Canvas(self, bg="#1a1a1a", highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas, style="Choice.TFrame")
        
        # Configure canvas
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw", width=self.master.winfo_width())
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # Pack the canvas and scrollbar
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        # Bind canvas resize
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        
        # Header with more informative content
        header_frame = ttk.Frame(self.scrollable_frame, style="Choice.TFrame")
        header_frame.pack(fill="x", pady=(40, 20), padx=40)
        
        title_label = ttk.Label(
            header_frame,
            text="Choose Document Type",
            style="Title.TLabel"
        )
        title_label.pack()
        
        subtitle_label = ttk.Label(
            header_frame,
            text="Select the type of document you want to create with AI Agent",
            style="Subtitle.TLabel"
        )
        subtitle_label.pack(pady=(10, 0))
        
        # Additional info label
        info_label = ttk.Label(
            header_frame,
            text="Our AI Agent will help you create professional documents with proper formatting and structure",
            font=("Helvetica", 12),
            background="#1a1a1a",
            foreground="#BDBDBD",
            wraplength=600
        )
        info_label.pack(pady=(10, 0))
        
        # Cards container
        cards_frame = ttk.Frame(self.scrollable_frame, style="Choice.TFrame")
        cards_frame.pack(fill="both", expand=True, padx=40, pady=20)
        
        # Configure grid for cards
        cards_frame.grid_columnconfigure(0, weight=1)
        cards_frame.grid_columnconfigure(1, weight=1)
        cards_frame.grid_rowconfigure(0, weight=1)
        
        # Word Document Card with more detailed description
        word_card = self._create_document_card(
            cards_frame,
            "Microsoft Word",
            "Create professional documents with AI assistance:\n\n" +
            "• Formatted text and headings\n" +
            "• Professional sections and layouts\n" +
            "• Smart content suggestions\n" +
            "• Automatic formatting\n" +
            "• Export to DOCX format",
            "📝",
            lambda: self._handle_selection("word")
        )
        word_card.grid(row=0, column=0, padx=(0, 20), sticky="nsew")
        
        # PowerPoint Card with more detailed description
        ppt_card = self._create_document_card(
            cards_frame,
            "PowerPoint",
            "Create engaging presentations with AI assistance:\n\n" +
            "• Professional slide layouts\n" +
            "• Smart content organization\n" +
            "• Visual design suggestions\n" +
            "• Automatic formatting\n" +
            "• Export to PPTX format",
            "📊",
            lambda: self._handle_selection("ppt")
        )
        ppt_card.grid(row=0, column=1, padx=(20, 0), sticky="nsew")
        
        # Footer with more information
        footer_frame = ttk.Frame(self.scrollable_frame, style="Choice.TFrame")
        footer_frame.pack(fill="x", pady=(20, 40), padx=40)
        
        # Powered by text
        powered_text = ttk.Label(
            footer_frame,
            text="Powered by Google Gemini AI",
            font=("Helvetica", 12),
            foreground="#BDBDBD",
            background="#1a1a1a"
        )
        powered_text.pack()
        
        # Additional footer info
        footer_info = ttk.Label(
            footer_frame,
            text="Your documents will be generated using advanced AI technology for the best results",
            font=("Helvetica", 10),
            foreground="#888888",
            background="#1a1a1a",
            wraplength=400
        )
        footer_info.pack(pady=(5, 0))

    def _on_canvas_configure(self, event):
        # Update the width of the scrollable frame when the canvas is resized
        self.canvas.itemconfig("all", width=event.width)

    def _create_document_card(self, parent, title, description, icon, command):
        # Card container with no border
        card = ttk.Frame(parent, style="Card.TFrame", padding=30)
        
        # Icon with larger size
        icon_label = ttk.Label(
            card,
            text=icon,
            font=("Segoe UI Emoji", 64),
            background="#2d2d2d",
            foreground="#4CAF50"
        )
        icon_label.pack(pady=(20, 20))
        
        # Title with larger font
        title_label = ttk.Label(
            card,
            text=title,
            font=("Helvetica", 24, "bold"),
            background="#2d2d2d",
            foreground="#FFFFFF"
        )
        title_label.pack(pady=(0, 15))
        
        # Description with better formatting
        desc_label = ttk.Label(
            card,
            text=description,
            font=("Helvetica", 12),
            background="#2d2d2d",
            foreground="#BDBDBD",
            wraplength=300,
            justify="left"
        )
        desc_label.pack(pady=(0, 25))
        
        # Create 3D shining button
        button = ShiningButton(
            card,
            text="Create Document",
            command=command,
            width=220,
            height=45
        )
        button.pack(pady=(0, 10))
        
        return card

    def _handle_selection(self, doc_type):
        # Add a small delay for visual feedback
        self.after(100, lambda: self.on_select(doc_type))

    def _show_help(self, event):
        # Open help documentation in default browser
        help_url = "https://github.com/yourusername/lmhu/wiki/Document-Types"
        webbrowser.open(help_url)