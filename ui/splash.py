import tkinter as tk
from tkinter import ttk
import time

class SplashScreen(tk.Frame):
    def __init__(self, master, on_complete):
        super().__init__(master)
        self.master = master
        self.on_complete = on_complete
        
        # Get screen dimensions
        screen_width = self.master.winfo_screenwidth()
        screen_height = self.master.winfo_screenheight()
        
        # Configure the main window
        self.master.title("LMHU - Language Model Helper Utility")
        self.master.configure(bg="#1a1a1a")
        
        # Configure the frame to expand
        self.configure(bg="#1a1a1a")
        self.pack(fill="both", expand=True)
        
        # Create main frame with dynamic padding
        padding_x = int(screen_width * 0.1)  # 10% of screen width
        padding_y = int(screen_height * 0.1)  # 10% of screen height
        self.main_frame = tk.Frame(self, bg="#1a1a1a", padx=padding_x, pady=padding_y)
        self.main_frame.pack(fill="both", expand=True)
        
        # Calculate dynamic font sizes based on screen size
        title_font_size = min(72, int(screen_height * 0.1))  # 10% of screen height, max 72
        subtitle_font_size = min(24, int(screen_height * 0.03))  # 3% of screen height, max 24
        status_font_size = min(18, int(screen_height * 0.025))  # 2.5% of screen height, max 18
        
        # Logo/Title with custom font and gradient-like effect
        self.title_frame = tk.Frame(self.main_frame, bg="#1a1a1a")
        self.title_frame.pack(pady=(int(screen_height * 0.15), int(screen_height * 0.05)))
        
        title_label = tk.Label(
            self.title_frame,
            text="LMHU",
            font=("Helvetica", title_font_size, "bold"),
            fg="#4CAF50",  # Material Design Green
            bg="#1a1a1a"
        )
        title_label.pack()
        
        subtitle_label = tk.Label(
            self.title_frame,
            text="Language Model Helper Utility",
            font=("Helvetica", subtitle_font_size),
            fg="#BDBDBD",  # Light Gray
            bg="#1a1a1a"
        )
        subtitle_label.pack(pady=(int(screen_height * 0.02), 0))
        
        # Custom styled progress bar
        self.progress_frame = tk.Frame(self.main_frame, bg="#1a1a1a")
        self.progress_frame.pack(pady=int(screen_height * 0.1), fill="x", padx=int(screen_width * 0.2))
        
        # Progress bar container
        self.progress_container = tk.Frame(
            self.progress_frame,
            bg="#2d2d2d",
            height=int(screen_height * 0.02),  # 2% of screen height
            highlightthickness=0
        )
        self.progress_container.pack(fill="x")
        
        # Actual progress bar
        self.progress = tk.Frame(
            self.progress_container,
            bg="#4CAF50",
            width=0,
            height=int(screen_height * 0.02)  # 2% of screen height
        )
        self.progress.pack(side="left")
        
        # Loading text
        self.loading_label = tk.Label(
            self.main_frame,
            text="Initializing...",
            font=("Helvetica", status_font_size),
            fg="#BDBDBD",
            bg="#1a1a1a"
        )
        self.loading_label.pack(pady=(int(screen_height * 0.03), 0))
        
        # Status messages
        self.status_messages = [
            "Loading components...",
            "Preparing interface...",
            "Almost ready...",
            "Starting application..."
        ]
        self.current_message = 0
        
        # Start animation
        self.animate()

    def animate(self):
        if not hasattr(self, 'start_time'):
            self.start_time = time.time()
        
        # Calculate progress based on time
        elapsed = time.time() - self.start_time
        progress = min(100, (elapsed / 2.5) * 100)  # 2.5 seconds total animation
        
        # Update progress bar
        width = int((progress / 100) * self.progress_container.winfo_width())
        self.progress.configure(width=width)
        
        # Update loading text
        if progress < 100:
            message_index = int((progress / 100) * len(self.status_messages))
            if message_index != self.current_message and message_index < len(self.status_messages):
                self.current_message = message_index
                self.loading_label.configure(text=self.status_messages[message_index])
            self.after(16, self.animate)  # ~60 FPS
        else:
            self.loading_label.configure(text="Ready!")
            self.after(200, self._finish)  # Small delay before transitioning

    def _finish(self):
        # Clear the splash screen
        self.pack_forget()
        # Call the completion callback
        self.on_complete()