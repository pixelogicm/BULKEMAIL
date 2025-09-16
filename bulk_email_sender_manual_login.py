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

# selenium
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

# ---------------------------
# Main app (full implementation)
# ---------------------------
class BulkEmailSender:
    ...
