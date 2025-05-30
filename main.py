import tkinter as tk
from tkinter import ttk
import os
from dotenv import load_dotenv
from ui.splash import SplashScreen
from ui.choice import ChoiceScreen
from ui.detail import DetailScreen, create_detail_screen
from core.logger import setup_logging
from config.settings import load_config


def show_choice_screen(root):
    # Clear existing widgets
    for widget in root.winfo_children():
        widget.destroy()
    
    # Create and show choice screen
    choice_screen = ChoiceScreen(root, on_select=lambda doc_type: create_detail_screen(root, doc_type))
    choice_screen.pack(fill="both", expand=True)


def show_detail_screen(root, doc_type):
    # Clear existing widgets
    for widget in root.winfo_children():
        widget.destroy()
    
    # Create and show detail screen
    detail_screen = DetailScreen(root, doc_type)
    return detail_screen


def main():
    # Load environment variables
    load_dotenv()
    setup_logging()
    load_config()
    
    # Create the main window
    root = tk.Tk()
    root.title("AI Document Generator")
    
    # Get screen dimensions
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    
    # Set window state to maximized
    root.state('zoomed')  # For Windows
    
    # Create and show splash screen
    splash = SplashScreen(root, on_complete=lambda: show_choice_screen(root))
    splash.pack(fill="both", expand=True)
    
    # Start the application
    root.mainloop()


if __name__ == "__main__":
    main()
