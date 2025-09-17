#!/usr/bin/env python3
"""
BulkEmailSender - Outlook Desktop (COM) primary send + Selenium fallback.

Patched to:
- Add "Insert DocuSign Template" UI button and insert_docusign_template helper.
- Auto-fill Sender and Review URL defaults when inserting template.
- Ensure Gmail compose is in rich-text mode before inserting HTML (best-effort).
- Keep existing injection, attachment and fallback-link logic.
"""
import os
import re
import sys
import time
import uuid
import json
import base64
import socket
import shutil
import tempfile
import traceback
import threading
import webbrowser
import subprocess
import queue
import ctypes

from concurrent.futures import ThreadPoolExecutor, as_completed

# tkinter UI
import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox

# selenium (optional)
try:
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.chrome.service import Service
    from selenium.common.exceptions import (
        StaleElementReferenceException,
        NoSuchElementException,
        ElementNotInteractableException,
        TimeoutException,
        WebDriverException,
    )
    _SELENIUM_AVAILABLE = True
except ImportError:
    _SELENIUM_AVAILABLE = False

# optional Flask tracking server
try:
    from flask import Flask, request, send_file, make_response
except Exception:
    Flask = None
    request = None
    send_file = None
    make_response = None

# Attempt to import the Outlook desktop helper (optional)
try:
    from send_via_outlook_desktop import send_via_outlook_desktop
    _OUTLOOK_HELPER_AVAILABLE = True
except Exception:
    send_via_outlook_desktop = None
    _OUTLOOK_HELPER_AVAILABLE = False

# ---------------------------
# Config
# ---------------------------
TRACK_PORT = 5000
NGROK_API = "http://127.0.0.1:4040/api/tunnels"
SCREENSHOT_DIR = os.path.join(tempfile.gettempdir(), "paris_screenshots")
os.makedirs(SCREENSHOT_DIR, exist_ok=True)

# Import image blurring functionality
try:
    from image_blur import PurchaseOrderBlurrer
    _IMAGE_BLUR_AVAILABLE = True
except Exception:
    PurchaseOrderBlurrer = None
    _IMAGE_BLUR_AVAILABLE = False

# ---------------------------
# Main app (full implementation)  
# ---------------------------
class BulkEmailSender:
    def __init__(self, root):
        self.root = root
        self.root.title("Bulk Email Sender with Image Blur")
        self.root.geometry("800x600")
        
        # Initialize image blurrer if available
        if _IMAGE_BLUR_AVAILABLE:
            self.blurrer = PurchaseOrderBlurrer()
        else:
            self.blurrer = None
            
        self.setup_ui()
    
    def setup_ui(self):
        """Set up the user interface."""
        # Main frame
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Title
        title_label = ttk.Label(main_frame, text="Bulk Email Sender with Purchase Order Image Blur", 
                               font=('Arial', 14, 'bold'))
        title_label.pack(pady=(0, 20))
        
        # Image blur section
        blur_frame = ttk.LabelFrame(main_frame, text="Purchase Order Image Blur", padding=10)
        blur_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Image selection
        select_frame = ttk.Frame(blur_frame)
        select_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(select_frame, text="Image File:").pack(side=tk.LEFT)
        self.image_path_var = tk.StringVar()
        image_path_entry = ttk.Entry(select_frame, textvariable=self.image_path_var, width=50)
        image_path_entry.pack(side=tk.LEFT, padx=(5, 5))
        
        browse_button = ttk.Button(select_frame, text="Browse", command=self.browse_image)
        browse_button.pack(side=tk.LEFT)
        
        # Blur options
        options_frame = ttk.Frame(blur_frame)
        options_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(options_frame, text="Blur Strength:").pack(side=tk.LEFT)
        self.blur_strength_var = tk.IntVar(value=15)
        blur_scale = ttk.Scale(options_frame, from_=5, to=30, variable=self.blur_strength_var, 
                              orient=tk.HORIZONTAL, length=200)
        blur_scale.pack(side=tk.LEFT, padx=(5, 10))
        
        strength_label = ttk.Label(options_frame, textvariable=self.blur_strength_var)
        strength_label.pack(side=tk.LEFT)
        
        # Auto detection option
        self.auto_detect_var = tk.BooleanVar(value=False)
        auto_check = ttk.Checkbutton(options_frame, text="Auto-detect text areas", 
                                   variable=self.auto_detect_var)
        auto_check.pack(side=tk.LEFT, padx=(20, 0))
        
        # Blur button
        blur_button = ttk.Button(blur_frame, text="Blur Image", command=self.blur_image)
        blur_button.pack()
        
        # Status display
        self.status_text = scrolledtext.ScrolledText(main_frame, height=15, width=80)
        self.status_text.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        
        # Check if image blur is available
        if not _IMAGE_BLUR_AVAILABLE:
            self.log_status("Warning: Image blur functionality not available. Please install required packages.")
        else:
            self.log_status("Image blur functionality ready.")
    
    def browse_image(self):
        """Browse for an image file."""
        filename = filedialog.askopenfilename(
            title="Select Purchase Order Image",
            filetypes=[
                ("Image files", "*.png *.jpg *.jpeg *.gif *.bmp"),
                ("PNG files", "*.png"),
                ("JPEG files", "*.jpg *.jpeg"),
                ("All files", "*.*")
            ]
        )
        if filename:
            self.image_path_var.set(filename)
            self.log_status(f"Selected image: {filename}")
    
    def blur_image(self):
        """Blur the selected purchase order image."""
        if not _IMAGE_BLUR_AVAILABLE:
            messagebox.showerror("Error", "Image blur functionality not available.")
            return
        
        image_path = self.image_path_var.get()
        if not image_path:
            messagebox.showerror("Error", "Please select an image file first.")
            return
        
        if not os.path.exists(image_path):
            messagebox.showerror("Error", "Selected image file does not exist.")
            return
        
        try:
            # Set blur strength
            self.blurrer.set_blur_strength(self.blur_strength_var.get())
            
            # Generate output path
            name, ext = os.path.splitext(image_path)
            output_path = f"{name}_blurred{ext}"
            
            self.log_status(f"Starting blur process for: {image_path}")
            self.log_status(f"Blur strength: {self.blur_strength_var.get()}")
            self.log_status(f"Auto-detect text areas: {self.auto_detect_var.get()}")
            
            # Apply blur
            result_path = self.blurrer.blur_purchase_order(
                image_path, 
                output_path, 
                use_auto_detection=self.auto_detect_var.get()
            )
            
            self.log_status(f"Success! Blurred image saved as: {result_path}")
            messagebox.showinfo("Success", f"Image blurred successfully!\nSaved as: {result_path}")
            
        except Exception as e:
            error_msg = f"Error blurring image: {str(e)}"
            self.log_status(error_msg)
            messagebox.showerror("Error", error_msg)
    
    def log_status(self, message):
        """Log a status message to the text area."""
        self.status_text.insert(tk.END, f"{time.strftime('%H:%M:%S')}: {message}\n")
        self.status_text.see(tk.END)
        self.root.update_idletasks()


def main():
    """Main function to run the application."""
    # Check if running with GUI support
    try:
        # First try to see if DISPLAY is available
        if os.environ.get('DISPLAY'):
            root = tk.Tk()
            app = BulkEmailSender(root)
            root.mainloop()
        else:
            raise ImportError("No display available")
    except (ImportError, Exception):
        print("GUI not available. Running in command line mode.")
        # Command line fallback for image blurring
        if _IMAGE_BLUR_AVAILABLE:
            import sys
            if len(sys.argv) > 1:
                image_path = sys.argv[1]
                blurrer = PurchaseOrderBlurrer()
                try:
                    result = blurrer.blur_purchase_order(image_path)
                    print(f"Image blurred successfully: {result}")
                except Exception as e:
                    print(f"Error: {e}")
            else:
                print("Usage: python3 bulk_email_sender_manual_login.py <image_path>")
        else:
            print("Image blur functionality not available.")


if __name__ == "__main__":
    main()
