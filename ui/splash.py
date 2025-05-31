import tkinter as tk
import time
import math
from PIL import Image, ImageDraw, ImageTk  # <-- Add this import

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
        self.master.configure(bg="#000000")  # Pure black

        # Configure the frame to expand
        self.configure(bg="#000000")
        self.pack(fill="both", expand=True)

        # Create main frame with dynamic padding
        padding_x = int(screen_width * 0.1)
        padding_y = int(screen_height * 0.1)
        self.main_frame = tk.Frame(self, bg="#000000", padx=padding_x, pady=padding_y)
        self.main_frame.pack(fill="both", expand=True)

        # Calculate dynamic font sizes based on screen size
        title_font_size = min(80, int(screen_height * 0.12))
        subtitle_font_size = min(28, int(screen_height * 0.035))
        status_font_size = min(20, int(screen_height * 0.025))

        # Title with glowing effect
        self.title_frame = tk.Frame(self.main_frame, bg="#000000")
        self.title_frame.pack(pady=(int(screen_height * 0.12), int(screen_height * 0.03)))

        # --- Add geometric figure (hexagon) above the title ---
        hex_size = min(int(screen_height * 0.09), 70)
        hex_img = Image.new("RGBA", (hex_size, hex_size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(hex_img)
        angle = math.pi / 3
        center = hex_size // 2
        radius = hex_size // 2 - 4
        points = [
            (
                center + radius * math.cos(angle * i - math.pi/2),
                center + radius * math.sin(angle * i - math.pi/2)
            )
            for i in range(6)
        ]
        draw.polygon(points, outline="#00FF41", fill=(0, 0, 0, 0), width=4)
        draw.line(points + [points[0]], fill="#00FF41", width=4)
        # Optional: add a glowing effect
        for r in range(radius-2, radius+6, 2):
            glow_points = [
                (
                    center + r * math.cos(angle * i - math.pi/2),
                    center + r * math.sin(angle * i - math.pi/2)
                )
                for i in range(6)
            ]
            draw.polygon(glow_points, outline="#00FF41", fill=None, width=1)
        self.hex_photo = ImageTk.PhotoImage(hex_img)
        self.hex_label = tk.Label(self.title_frame, image=self.hex_photo, bg="#000000")
        self.hex_label.pack(pady=(0, int(screen_height * 0.01)))
        # --- End geometric figure ---

        # Glowing shadow effect using multiple labels
        for i, color in enumerate(["#003311", "#005522", "#00FF41", "#00FF41"]):
            tk.Label(
                self.title_frame,
                text="LMHU",
                font=("Helvetica", title_font_size, "bold"),
                fg=color,
                bg="#000000"  # Ensure background is black
            ).place(x=0, y=0, relx=0.5, rely=0.5, anchor="center")

        self.title_label = tk.Label(
            self.title_frame,
            text="LMHU",
            font=("Helvetica", title_font_size, "bold"),
            fg="#00FF41",  # Matrix Green
            bg="#000000"   # Ensure background is black
        )
        self.title_label.pack()

        self.subtitle_label = tk.Label(
            self.title_frame,
            text="Language Model Helper Utility",
            font=("Helvetica", subtitle_font_size),
            fg="#00FF41",  # Matrix Green
            bg="#000000"   # Ensure background is black
        )
        self.subtitle_label.pack(pady=(int(screen_height * 0.01), 0))

        # Circular progress indicator
        self.circle_frame = tk.Frame(self.main_frame, bg="#000000")
        self.circle_frame.pack(pady=int(screen_height * 0.04))
        circle_size = min(int(screen_height * 0.13), 120)
        self.canvas = tk.Canvas(
            self.circle_frame,
            width=circle_size,
            height=circle_size,
            bg="#000000",
            highlightthickness=0
        )
        self.canvas.pack()
        self.circle_size = circle_size

        # Loading text
        self.loading_label = tk.Label(
            self.main_frame,
            text="Initializing...",
            font=("Helvetica", status_font_size, "bold"),
            fg="#00FF41",
            bg="#000000"
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

        # Green highlight bar at the bottom
        self.bottom_bar = tk.Frame(self, bg="#00FF41", height=4)
        self.bottom_bar.pack(side="bottom", fill="x")

        # Start animation
        self.animate()

    def animate(self):
        if not hasattr(self, 'start_time'):
            self.start_time = time.time()

        elapsed = time.time() - self.start_time
        progress = min(100, (elapsed / 2.5) * 100)  # 2.5 seconds total animation

        # Draw circular progress
        self.canvas.delete("all")
        angle = (progress / 100) * 360
        # Draw background circle
        self.canvas.create_oval(
            6, 6, self.circle_size-6, self.circle_size-6,
            outline="#003311", width=8
        )
        # Draw progress arc
        self.canvas.create_arc(
            6, 6, self.circle_size-6, self.circle_size-6,
            start=90, extent=-angle,
            style="arc", outline="#00FF41", width=10
        )
        # Draw glowing dot at the end of arc
        if progress > 0:
            radians = math.radians(90 - angle)
            r = (self.circle_size-12) // 2
            cx, cy = self.circle_size//2, self.circle_size//2
            x = cx + r * math.cos(radians)
            y = cy - r * math.sin(radians)
            self.canvas.create_oval(
                x-8, y-8, x+8, y+8,
                fill="#00FF41", outline="#00FF41", width=2
            )

        # Update loading text
        if progress < 100:
            message_index = int((progress / 100) * len(self.status_messages))
            if message_index != self.current_message and message_index < len(self.status_messages):
                self.current_message = message_index
                self.loading_label.configure(text=self.status_messages[message_index])
            self.after(16, self.animate)
        else:
            self.loading_label.configure(text="Ready!")
            self.after(300, self._finish)

    def _finish(self):
        self.pack_forget()
        self.on_complete()