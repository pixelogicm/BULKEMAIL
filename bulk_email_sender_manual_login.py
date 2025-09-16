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
    def __init__(self, root):
        self.root = root
        self.root.title("PARIS SENDER - Outlook Desktop + Selenium Fallback")
        self.root.geometry("980x760")

        # Selenium state
        self.driver = None
        self.driver_lock = threading.Lock()

        # queues & threading
        self.send_queue = queue.Queue()
        self.executor = None
        self.max_workers = 1  # default 1 to avoid races with shared driver

        # app state
        self.email_list = []
        self.tree_items = {}
        self.tracking_map = {}
        self.action_queue = queue.Queue()
        self.gui_update_queue = queue.Queue()
        self.recent_replies = {}
        self.running = False
        self.paused = False
        self.sent_count = 0
        self.failed_count = 0

        # hosting
        self.TRACK_SERVER_PORT = TRACK_PORT
        self.hosted_dir = os.path.join(tempfile.gettempdir(), "paris_hosted_files")
        os.makedirs(self.hosted_dir, exist_ok=True)
        self.hosted_files = {}
        self.ngrok_process = None
        self.ngrok_public_url = None

        # chromedriver
        self.chromedriver_path = None
        self.attachment_path = None

        # additional user inputs
        self.sender_var = tk.StringVar(value="")          # for replacing %SENDER%
        self.review_var = tk.StringVar(value="")          # for replacing REVIEW_DOCUMENT_URL

        # Informative log about helper availability
        if _OUTLOOK_HELPER_AVAILABLE:
            self.log("Outlook desktop helper module found. Using Outlook desktop (if available) or Selenium fallback.")
        else:
            self.log("Outlook desktop helper NOT found or failed to import. Install pywin32 and add send_via_outlook_desktop.py for Outlook COM sending. Will use Selenium fallback for Outlook sends.")

        self._build_ui()
        threading.Thread(target=self.action_worker, daemon=True).start()
        self.root.after(500, self.process_gui_updates)
        self.log("App initialized. Use 'OPEN CHROME & LOGIN', 'Start Tracking Server', then prepare your HTML and START.")
        self.log(f"Screenshot folder: {SCREENSHOT_DIR}")
        self.log("Tip: Use 'Insert DocuSign Template', set Review URL to your HTTPS link, Choose File, OPEN CHROME & LOGIN, then START to test.")

    # ---------------------------
    # UI
    # ---------------------------
    def _build_ui(self):
        tk.Label(self.root, text="PARIS SENDER", fg="darkred", font=("Segoe UI", 14, "bold")).pack(pady=6)

        top_row = tk.Frame(self.root)
        top_row.pack(fill='x', anchor='w', pady=2)

        provider_frame = tk.Frame(top_row)
        provider_frame.pack(side='left', padx=6)
        tk.Label(provider_frame, text="Email Provider:").pack(side=tk.LEFT)
        self.provider_var = tk.StringVar(value="Gmail")
        provider_options = ["Gmail", "Outlook.com", "Comcast"]
        self.provider_box = ttk.Combobox(provider_frame, textvariable=self.provider_var, values=provider_options, state="readonly", width=16)
        self.provider_box.pack(side=tk.LEFT, padx=5)

        subject_frame = tk.Frame(top_row)
        subject_frame.pack(side='left', padx=6)
        tk.Label(subject_frame, text="Subject:").pack(side=tk.LEFT)
        self.subject_var = tk.StringVar(value="Bulk Email")
        tk.Entry(subject_frame, textvariable=self.subject_var, width=32).pack(side=tk.LEFT, padx=4)

        # New: sender and review link inputs
        meta_frame = tk.Frame(top_row)
        meta_frame.pack(side='left', padx=6)
        tk.Label(meta_frame, text="Sender:").pack(side=tk.LEFT)
        tk.Entry(meta_frame, textvariable=self.sender_var, width=18).pack(side=tk.LEFT, padx=4)
        tk.Label(meta_frame, text="Review URL:").pack(side=tk.LEFT)
        tk.Entry(meta_frame, textvariable=self.review_var, width=26).pack(side=tk.LEFT, padx=4)

        batch_frame = tk.Frame(top_row)
        batch_frame.pack(side='left', padx=6)
        tk.Label(batch_frame, text="Emails per batch:").pack(side=tk.LEFT)
        self.batch_var = tk.IntVar(value=1)
        tk.Entry(batch_frame, textvariable=self.batch_var, width=5).pack(side=tk.LEFT)

        threads_frame = tk.Frame(top_row)
        threads_frame.pack(side='left', padx=6)
        tk.Label(threads_frame, text="Concurrent threads:").pack(side=tk.LEFT)
        self.threads_var = tk.IntVar(value=1)
        threads_spinbox = tk.Spinbox(threads_frame, from_=1, to=10, textvariable=self.threads_var, width=5)
        threads_spinbox.pack(side=tk.LEFT)

        driver_frame = tk.Frame(top_row)
        driver_frame.pack(side='right', padx=6)
        tk.Button(driver_frame, text="Set Chromedriver", command=self.set_chromedriver_path).pack(side=tk.LEFT)
        self.driver_path_label = tk.Label(driver_frame, text="(using PATH)")
        self.driver_path_label.pack(side=tk.LEFT, padx=4)

        recipients_frame = tk.Frame(self.root)
        recipients_frame.pack(anchor='w', pady=6, padx=6, fill='x')
        tk.Label(recipients_frame, text="Recipients:").pack(side=tk.LEFT)
        tk.Button(recipients_frame, text="Load Email List (CSV)", command=self.load_emails).pack(side=tk.LEFT, padx=6)
        tk.Label(recipients_frame, text="Add recipient:").pack(side=tk.LEFT, padx=6)
        self.new_recipient_var = tk.StringVar()
        tk.Entry(recipients_frame, textvariable=self.new_recipient_var, width=40).pack(side=tk.LEFT, padx=2)
        tk.Button(recipients_frame, text="Add", command=self.add_recipient).pack(side=tk.LEFT, padx=2)
        tk.Button(recipients_frame, text="Remove Selected", command=self.remove_selected).pack(side=tk.LEFT, padx=2)
        tk.Label(recipients_frame, text="(Load or add recipients before editing message)").pack(side=tk.LEFT, padx=8)

        tk.Label(self.root, text="Message (edit HTML source):").pack(anchor='w')
        self.message_box = scrolledtext.ScrolledText(self.root, height=12, width=100)
        self.message_box.pack(padx=5, pady=5)
        # Start the editor empty to avoid accidentally sending a leftover test message
        self.message_box.insert(tk.END, "")

        attach_frame = tk.Frame(self.root)
        attach_frame.pack(anchor='w', pady=5)
        tk.Label(attach_frame, text="Attachment:").pack(side=tk.LEFT)
        self.attach_label = tk.Label(attach_frame, text="No file selected")
        self.attach_label.pack(side=tk.LEFT, padx=5)
        tk.Button(attach_frame, text="Choose File (Attach)", command=self.choose_file).pack(side=tk.LEFT)
        tk.Button(attach_frame, text="Load HTML File", command=self.load_html_file).pack(side=tk.LEFT, padx=5)
        tk.Button(attach_frame, text="Preview HTML", command=self.preview_html).pack(side=tk.LEFT, padx=5)
        tk.Button(attach_frame, text="Host Document & Insert Link", command=self.host_document_and_insert_link).pack(side=tk.LEFT, padx=5)
        tk.Button(attach_frame, text="Start/Stop ngrok", command=self.toggle_ngrok).pack(side=tk.LEFT, padx=5)

        # New: Insert DocuSign Template button
        tk.Button(attach_frame, text="Insert DocuSign Template", command=self.insert_docusign_template).pack(side=tk.LEFT, padx=5)

        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=5)
        self.open_chrome_btn = tk.Button(btn_frame, text="OPEN CHROME & LOGIN", command=self.open_chrome_and_login)
        self.open_chrome_btn.pack(side=tk.LEFT, padx=5)
        self.start_tracking_btn = tk.Button(btn_frame, text="Start Tracking Server", command=self.start_tracking_server_thread)
        self.start_tracking_btn.pack(side=tk.LEFT, padx=5)
        self.start_btn = tk.Button(btn_frame, text="START", command=self.start_sending)
        self.start_btn.pack(side=tk.LEFT, padx=5)
        self.pause_btn = tk.Button(btn_frame, text="PAUSE", command=self.pause)
        self.pause_btn.pack(side=tk.LEFT, padx=5)
        self.stop_btn = tk.Button(btn_frame, text="STOP", command=self.stop)
        self.stop_btn.pack(side=tk.LEFT, padx=5)

        tk.Label(self.root, text="Progress:").pack(anchor='w')
        self.progress = ttk.Progressbar(self.root, orient='horizontal', length=900, mode='determinate')
        self.progress.pack(pady=2)
        self.status_label = tk.Label(self.root, text="Ready. Sent: 0 | Failed: 0")
        self.status_label.pack(anchor='w')

        tk.Label(self.root, text="Activity Log:").pack(anchor='w')
        self.log_box = scrolledtext.ScrolledText(self.root, height=8, width=120)
        self.log_box.pack(padx=5, pady=5)

        tk.Label(self.root, text="Recipients Status:").pack(anchor='w')
        tree_frame = tk.Frame(self.root)
        tree_frame.pack(padx=5, pady=5, fill='both', expand=False)
        columns = ("email", "status", "sent_at", "opened_at", "replied")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=8)
        for col in columns:
            self.tree.heading(col, text=col.capitalize())
            self.tree.column(col, width=180 if col == "email" else 110, anchor='w')
        self.tree.pack(side='left', fill='x', expand=True)
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side='right', fill='y')
        self.tree.configure(yscrollcommand=scrollbar.set)

    # ---------------------------
    # Logging and GUI updates
    # ---------------------------
    def log(self, msg):
        ts = time.strftime("[%Y-%m-%d %H:%M:%S] ")
        try:
            self.log_box.insert(tk.END, ts + msg + "\n")
            self.log_box.see(tk.END)
        except Exception:
            pass
        print(ts + msg)

    def process_gui_updates(self):
        while True:
            try:
                item = self.gui_update_queue.get_nowait()
            except queue.Empty:
                break
            if not item:
                continue
            if item[0] == "update_status":
                _, email, status, sent_at, opened_at, replied = item
                iid = self.tree_items.get(email)
                if iid:
                    self.tree.item(iid, values=(email, status, sent_at, opened_at, replied))
            elif item[0] == "update_replied":
                _, email, replied = item
                iid = self.tree_items.get(email)
                if iid:
                    vals = list(self.tree.item(iid, "values"))
                    if len(vals) < 5:
                        vals = vals + [""]*(5-len(vals))
                    vals[4] = replied
                    self.tree.item(iid, values=tuple(vals))
            self.gui_update_queue.task_done()
        self.root.after(500, self.process_gui_updates)

    # ---------------------------
    # Recipient management
    # ---------------------------
    def load_emails(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if not file_path:
            return
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            lines = [line.strip() for line in f if line.strip()]
        self.email_list = lines
        self.tree.delete(*self.tree.get_children())
        self.tree_items.clear()
        for email in self.email_list:
            iid = self.tree.insert("", "end", values=(email, "Pending", "", "", "No"))
            self.tree_items[email] = iid
        self.log(f"Loaded {len(self.email_list)} recipients from {file_path}")

    def add_recipient(self):
        email = self.new_recipient_var.get().strip()
        if not email:
            messagebox.showwarning("Input required", "Please enter an email address to add.")
            return
        if not re.match(r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$", email):
            if not messagebox.askyesno("Confirm add", f"'{email}' doesn't look like a valid email. Add anyway?"):
                return
        if email in self.email_list:
            messagebox.showinfo("Already added", f"{email} is already in the recipient list.")
            return
        self.email_list.append(email)
        iid = self.tree.insert("", "end", values=(email, "Pending", "", "", "No"))
        self.tree_items[email] = iid
        self.new_recipient_var.set("")
        self.log(f"Added recipient: {email}")

    def remove_selected(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo("No selection", "Select one or more recipients in the table to remove.")
            return
        for iid in selected:
            vals = self.tree.item(iid, "values")
            if not vals:
                continue
            email = vals[0]
            if email in self.email_list:
                try:
                    self.email_list.remove(email)
                except ValueError:
                    pass
            if email in self.tree_items:
                del self.tree_items[email]
            for tid, entry in list(self.tracking_map.items()):
                if entry.get("email") == email:
                    del self.tracking_map[tid]
            self.tree.delete(iid)
            self.log(f"Removed recipient: {email}")

    # ---------------------------
    # HTML load/preview
    # ---------------------------
    def load_html_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("HTML Files", "*.htm;*.html")])
        if not file_path:
            return
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                html_content = f.read()
        except Exception:
            with open(file_path, 'r', encoding='latin-1') as f:
                html_content = f.read()
        self.message_box.delete("1.0", tk.END)
        self.message_box.insert(tk.END, html_content)
        self.log(f"Loaded HTML file into message editor: {file_path}")

    def preview_html(self):
        html = self.message_box.get("1.0", tk.END).strip()
        if not html:
            messagebox.showinfo("Nothing to preview", "Message body is empty.")
            return
        if "<html" not in html.lower():
            html = f"<html><head><meta charset='utf-8'></head><body>{html}</body></html>"
        tf = tempfile.NamedTemporaryFile(delete=False, suffix=".html", prefix="paris_preview_", mode="w", encoding="utf-8")
        tf.write(html)
        tf.close()
        webbrowser.open(f"file://{tf.name}")
        self.log(f"Preview opened in browser: {tf.name}")

    # ---------------------------
    # Hosting & ngrok
    # ---------------------------
    def choose_file(self):
        file_path = filedialog.askopenfilename()
        if not file_path:
            return
        self.attachment_path = file_path
        self.attach_label.config(text=os.path.basename(file_path))
        self.log(f"Attachment selected: {file_path}")

    def host_document_and_insert_link(self):
        if Flask is None:
            messagebox.showwarning("Flask missing", "Install Flask (pip install flask) to enable hosting local documents.")
            return
        file_path = filedialog.askopenfilename(title="Select document to host")
        if not file_path:
            return
        try:
            fid = str(uuid.uuid4())
            basename = os.path.basename(file_path)
            dest_name = f"{fid}_{basename}"
            dest_path = os.path.join(self.hosted_dir, dest_name)
            shutil.copyfile(file_path, dest_path)
            self.hosted_files[fid] = {"path": dest_path, "name": basename}
            ip = self._detect_local_ip()
            local_url = f"http://{ip}:{self.TRACK_SERVER_PORT}/file/{fid}/{basename}"
            insert_url = local_url
            if self.ngrok_public_url:
                insert_url = f"{self.ngrok_public_url}/file/{fid}/{basename}"
            anchor_html = f'<p><a href="{insert_url}" target="_blank" rel="noopener">View Document: {basename}</a></p>'
            try:
                self.message_box.insert(tk.INSERT, "\n" + anchor_html + "\n")
            except Exception:
                self.message_box.insert(tk.END, "\n" + anchor_html + "\n")
            # Optional: auto-fill Review URL so placeholders get replaced
            try:
                self.review_var.set(insert_url)
            except Exception:
                pass
            self.log(f"Hosted document {basename} as id {fid} -> {insert_url}")
            messagebox.showinfo("Hosted link created", "Inserted link into message and set Review URL.")
        except Exception as e:
            self.log(f"Failed to host document: {e}\n{traceback.format_exc()}")
            messagebox.showerror("Host error", f"Could not host the selected document: {e}")

    def toggle_ngrok(self):
        if self.ngrok_process:
            try:
                self.ngrok_process.terminate()
            except Exception:
                pass
            self.ngrok_process = None
            self.ngrok_public_url = None
            messagebox.showinfo("ngrok", "Stopped ngrok process.")
            self.log("ngrok stopped.")
            return
        try:
            self.ngrok_process = subprocess.Popen(["ngrok", "http", str(self.TRACK_SERVER_PORT)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            self.log("Started ngrok process (waiting for tunnels)...")
            for _ in range(15):
                time.sleep(0.5)
                public = self._get_ngrok_public_url()
                if public:
                    self.ngrok_public_url = public.rstrip("/")
                    break
            if not self.ngrok_public_url:
                messagebox.showinfo("ngrok", "ngrok started but public URL not detected.")
                self.log("ngrok started but public URL not detected.")
            else:
                self.log(f"ngrok public URL: {self.ngrok_public_url}")
        except FileNotFoundError:
            self.ngrok_process = None
            messagebox.showinfo("ngrok missing", "ngrok binary not found on PATH.")
            self.log("ngrok binary not found on PATH.")
        except Exception as e:
            self.ngrok_process = None
            messagebox.showerror("ngrok error", f"Failed to start ngrok: {e}")
            self.log(f"Failed to start ngrok: {e}\n{traceback.format_exc()}")

    def _get_ngrok_public_url(self):
        try:
            import urllib.request
            with urllib.request.urlopen(NGROK_API, timeout=2) as resp:
                data = json.load(resp)
            tunnels = data.get("tunnels", [])
            for t in tunnels:
                public_url = t.get("public_url")
                if public_url and public_url.startswith("http"):
                    return public_url
        except Exception:
            return None

    def set_chromedriver_path(self):
        path = filedialog.askopenfilename(title="Select chromedriver executable")
        if not path:
            return
        self.chromedriver_path = path
        display = os.path.basename(path) if path else "(using PATH)"
        self.driver_path_label.config(text=display)
        self.log(f"Chromedriver path set: {path}")

    def open_chrome_and_login(self):
        provider = self.provider_var.get()
        options = webdriver.ChromeOptions()
        options.add_argument("--log-level=3")
        options.debugger_address = "127.0.0.1:9222"

        try:
            if self.chromedriver_path:
                service = Service(executable_path=self.chromedriver_path)
            else:
                service = Service()
            service.log_path = os.devnull
            if sys.platform.startswith("win"):
                try:
                    service.creationflags = subprocess.CREATE_NO_WINDOW
                except Exception:
                    pass
        except Exception:
            service = None

        try:
            if service:
                self.driver = webdriver.Chrome(service=service, options=options)
            else:
                self.driver = webdriver.Chrome(options=options)
            self.log("Attached to existing Chrome via remote debugging at 127.0.0.1:9222")
        except Exception as e:
            self.log(f"Attach failed: {e}\n{traceback.format_exc()}")
            try:
                temp_profile = tempfile.mkdtemp()
                opts2 = webdriver.ChromeOptions()
                opts2.add_argument(f"user-data-dir={temp_profile}")
                opts2.add_argument("--log-level=3")
                if self.chromedriver_path:
                    service2 = Service(executable_path=self.chromedriver_path)
                else:
                    service2 = Service()
                service2.log_path = os.devnull
                if sys.platform.startswith("win"):
                    try:
                        service2.creationflags = subprocess.CREATE_NO_WINDOW
                    except Exception:
                        pass
                self.driver = webdriver.Chrome(service=service2, options=opts2)
                self.log("Started fallback automated Chrome.")
            except Exception as e2:
                self.log(f"Failed to start fallback Chrome: {e2}\n{traceback.format_exc()}")
                messagebox.showerror("Chrome error", "Could not attach or start Chrome.")
                return

        try:
            with self.driver_lock:
                self.driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
                    "source": "Object.defineProperty(navigator,'webdriver',{get:() => undefined});"
                })
        except Exception:
            pass

        if provider == "Gmail":
            url = "https://mail.google.com/"
        elif provider == "Outlook.com":
            url = "https://outlook.live.com/mail/0/"
        elif provider == "Comcast":
            url = "https://connect.xfinity.com/appsuite/"
        else:
            url = "https://www.google.com/"

        try:
            with self.driver_lock:
                self.driver.get(url)
            self.log(f"Please complete manual sign-in in the attached Chrome at {url}. After signing in click START.")
        except Exception as e:
            self.log(f"Navigation error: {e}")

    # ---------------------------
    # Sending architecture
    # ---------------------------
    def start_sending(self):
        if not self.email_list:
            messagebox.showerror("Error", "No recipients loaded!")
            return
        # Raw message and subject
        raw_message = self.message_box.get("1.0", tk.END).strip()

        # Debug/confirm: log and save the exact message that will be sent
        try:
            preview = raw_message.strip()
            # Log a short preview to the activity log so you can confirm what will be sent
            self.log("Message preview (first 400 chars):\n" + (preview[:400] + ("..." if len(preview) > 400 else "")))
            # Save the full HTML to a temp file for inspection (helps debugging)
            tf = tempfile.NamedTemporaryFile(delete=False, suffix=".html", prefix="paris_message_preview_", mode="w", encoding="utf-8")
            tf.write(preview)
            tf.close()
            self.log(f"Saved message body to preview file: {tf.name}")
            # Confirmation dialog - prevents accidental sends
            if not messagebox.askyesno("Confirm send", f"Send this message? Preview saved to:\n{tf.name}\n\nSend anyway?"):
                self.log("Send cancelled by user (message confirmation).")
                return
        except Exception as _e:
            self.log(f"Could not write preview/confirm: {_e}")

        subject_text = (self.subject_var.get() or "Bulk Email").strip()
        total = len(self.email_list)
        self.sent_count = 0
        self.failed_count = 0
        self.progress['maximum'] = total
        self.progress['value'] = 0
        self.running = True
        self.paused = False

        self.max_workers = max(1, int(self.threads_var.get()))
        if self.max_workers != 1:
            self.log("Warning: concurrent threads >1 with a single shared driver can be unreliable. Prefer 1 or use separate driver instances per worker.")

        # warm compose for selenium fallback if necessary (always attempt)
        provider = self.provider_var.get()
        provider_lc = (provider or "").lower()
        if provider_lc and "outlook" in provider_lc:
            try:
                with self.driver_lock:
                    try:
                        self.driver.get("https://outlook.live.com/mail/0/")
                    except Exception:
                        pass
                    try:
                        self._ensure_compose_open(provider)
                    except Exception:
                        pass
            except Exception:
                pass

        email_tasks = []
        for email in self.email_list:
            track_id = str(uuid.uuid4())
            sent_time = time.strftime("%Y-%m-%d %H:%M:%S")
            self.tracking_map[track_id] = {"email": email, "status": "Queued", "sent_time": sent_time, "opened_time": "", "replied": False}

            # New: extract the body fragment (avoid sending full html/head wrappers)
            fragment = self._extract_body_fragment(raw_message)

            # Inject placeholders: %SENDER% and REVIEW_DOCUMENT_URL (and audio variants)
            fragment = self._inject_placeholders(fragment)
            # Ensure visible fallback link is appended so recipient always sees an absolute URL
            fragment = self._ensure_review_fallback(fragment)

            pixel_src = f"http://{self._detect_local_ip()}:{self.TRACK_SERVER_PORT}/track?id={track_id}"
            if self.ngrok_public_url:
                pixel_src = f"{self.ngrok_public_url}/track?id={track_id}"
            pixel_tag = f'<img src="{pixel_src}" width="1" height="1" style="display:none;" alt="">'
            html_message = fragment + pixel_tag

            email_tasks.append((track_id, email, html_message, subject_text))
            self.gui_update_queue.put(("update_status", email, "Queued", sent_time, "", "No"))

        self.log(f"Starting concurrent sending with {self.max_workers} threads for {len(email_tasks)} emails")
        threading.Thread(target=self._concurrent_sender, args=(email_tasks,), daemon=True).start()

    def _concurrent_sender(self, email_tasks):
        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            future_to_email = {
                executor.submit(self._send_single_email, track_id, email, html_message, subject): (track_id, email)
                for (track_id, email, html_message, subject) in email_tasks
            }
            for future in as_completed(future_to_email):
                track_id, email = future_to_email[future]
                try:
                    success = future.result()
                    if success:
                        self.sent_count += 1
                        self.tracking_map[track_id]["status"] = "Sent"
                        self.gui_update_queue.put(("update_status", email, "Sent", self.tracking_map[track_id].get("sent_time", ""), "", "No"))
                        self.log(f"Successfully sent to {email} (track id {track_id})")
                    else:
                        self.failed_count += 1
                        self.tracking_map[track_id]["status"] = "Failed"
                        self.gui_update_queue.put(("update_status", email, "Failed", "", "", "No"))
                except Exception as e:
                    self.failed_count += 1
                    self.tracking_map[track_id]["status"] = "Failed"
                    self.log(f"Failed to send to {email}: {e}")
                    self.gui_update_queue.put(("update_status", email, "Failed", "", "", "No"))
                finally:
                    try:
                        total = int(self.progress['maximum']) if self.progress['maximum'] else (self.sent_count + self.failed_count)
                        self.progress['value'] = self.sent_count + self.failed_count
                        self.status_label.config(text=f"Total: {total} | Sent: {self.sent_count} | Failed: {self.failed_count}")
                    except Exception:
                        pass

    def _send_single_email(self, track_id, email, html_message, subject):
        try:
            # Try Outlook desktop first for Outlook provider
            provider = self.provider_var.get()
            provider_lc = (provider or "").lower()
            if "outlook" in provider_lc:
                if _OUTLOOK_HELPER_AVAILABLE and send_via_outlook_desktop is not None:
                    try:
                        self.log(f"Attempting Outlook desktop send for {email}")
                        attachments = [getattr(self, "attachment_path", None)] if getattr(self, "attachment_path", None) else None
                        send_via_outlook_desktop(email, subject, html_message, attachments=attachments)
                        self.log(f"Outlook desktop send succeeded for {email}")
                        return True
                    except Exception as e:
                        self.log(f"Outlook desktop send failed for {email}: {e}. Falling back to Selenium.")
                else:
                    self.log("Outlook desktop helper not available; ensure send_via_outlook_desktop.py is present and pywin32 is installed. Falling back to Selenium.")

            # Selenium fallback (no Graph)
            with self.driver_lock:
                if not self.driver:
                    raise RuntimeError("Driver not available")
                try:
                    ss_path = os.path.join(tempfile.gettempdir(), f"paris_sending_{track_id}_{int(time.time())}.png")
                    self.driver.save_screenshot(ss_path)
                    self.log(f"Saved pre-send screenshot: {ss_path}")
                except Exception:
                    pass
                self.send_email(email, html_message, self.provider_var.get(), subject)
                return True
        except Exception as e:
            try:
                with self.driver_lock:
                    if self.driver:
                        fname = os.path.join(SCREENSHOT_DIR, f"paris_failure_{track_id}_{int(time.time())}.png")
                        self.driver.save_screenshot(fname)
                        self.log(f"Saved failure screenshot: {fname}")
            except Exception:
                pass
            self.log(f"Failed to send to {email}: {e}\n{traceback.format_exc()}")
            return False

    def pause(self):
        if not self.running:
            return
        self.paused = not self.paused
        self.log("Paused." if self.paused else "Resumed.")

    def stop(self):
        self.running = False
        if self.executor:
            self.executor.shutdown(wait=False)
        try:
            while not self.send_queue.empty():
                self.send_queue.get_nowait()
                self.send_queue.task_done()
        except Exception:
            pass
        self.log("Stopped sending and cleared queue.")

    # ---------------------------
    # Body & DOM helper methods
    # ---------------------------
    def _strip_tags(self, html):
        try:
            return re.sub(r'<[^>]+>', '', html).strip()
        except Exception:
            return html

    def _inject_placeholders(self, html_fragment):
        """
        Replace %SENDER% and REVIEW_DOCUMENT_URL placeholders in the fragment
        with UI-provided values. Also replace common audio placeholder variants.
        """
        try:
            res = html_fragment or ""
            sender = (self.sender_var.get() or "").strip()
            review = (self.review_var.get() or "").strip()
            if sender:
                res = res.replace("%SENDER%", sender)
            if review:
                # Replace both token names commonly used in our templates
                res = res.replace("REVIEW_DOCUMENT_URL", review)
                res = res.replace("REPLACE_WITH_AUDIO_URL", review)
                res = res.replace("%REVIEW_DOCUMENT_URL%", review)
            return res
        except Exception:
            return html_fragment

    def _ensure_review_fallback(self, fragment):
        """
        Ensure there is a visible absolute link fallback for the review document.
        Call after placeholder injection and before insertion.
        """
        try:
            review = (self.review_var.get() or "").strip()
            if not review:
                return fragment
            if review in (fragment or ""):
                return fragment
            fallback = (
                f'<p style="margin-top:12px;font-size:14px;color:#222;">'
                f'If the button above does not work, open this link to review the document: '
                f'<a href="{review}" target="_blank" rel="noopener" style="color:#1f61c3;word-break:break-all;">{review}</a>'
                f'</p>'
            )
            return (fragment or "") + fallback
        except Exception:
            return fragment

    def _extract_body_fragment(self, html):
        """
        Return just the fragment that should go into the message body.
        If html contains a <body>...</body> block, return its contents.
        Otherwise remove <!doctype> and <html>/<head> wrappers if present,
        and return the remainder unchanged.
        """
        try:
            if not html:
                return ""
            s = html.strip()
            m = re.search(r"<body[^>]*>(.*?)</body\s*>", s, flags=re.IGNORECASE | re.DOTALL)
            if m:
                return m.group(1).strip()
            # strip typical wrappers
            s = re.sub(r"(?is)<!doctype[^>]*>", "", s)
            s = re.sub(r"(?is)</?html[^>]*>", "", s)
            s = re.sub(r"(?is)</?head[^>]*>", "", s)
            s = re.sub(r"(?is)<meta[^>]*>", "", s)
            s = re.sub(r"(?is)<title[^>]*>.*?</title\s*>", "", s)
            return s.strip()
        except Exception:
            return html

    def _force_set_body_element(self, webelement, html):
        driver = self.driver
        try:
            try:
                driver.execute_script("arguments[0].focus(); arguments[0].innerHTML = arguments[1];", webelement, html)
            except Exception:
                try:
                    driver.execute_script("arguments[0].innerHTML = arguments[1];", webelement, html)
                except Exception:
                    pass
            time.sleep(0.12)
            try:
                inner = webelement.get_attribute("innerHTML") or ""
                wanted_plain = self._strip_tags(html)[:60]
                got_plain = self._strip_tags(inner)[:60]
                if wanted_plain and wanted_plain == got_plain:
                    return True
                if len(self._strip_tags(inner).strip()) > 0:
                    return True
            except Exception:
                pass
            try:
                plain = self._strip_tags(html)
                try:
                    webelement.click()
                except Exception:
                    pass
                try:
                    webelement.clear()
                except Exception:
                    try:
                        driver.execute_script("arguments[0].innerHTML = '';", webelement)
                    except Exception:
                        pass
                chunk_size = 500
                if plain:
                    for i in range(0, len(plain), chunk_size):
                        part = plain[i:i+chunk_size]
                        try:
                            webelement.send_keys(part)
                        except Exception:
                            try:
                                driver.execute_script("arguments[0].textContent += arguments[1];", webelement, part)
                            except Exception:
                                pass
                        time.sleep(0.02)
                time.sleep(0.12)
                inner2 = webelement.get_attribute("innerHTML") or ""
                if len(self._strip_tags(inner2)) > 0:
                    return True
            except Exception:
                pass
            try:
                plain = self._strip_tags(html)
                driver.execute_script("arguments[0].textContent = arguments[1];", webelement, plain)
                time.sleep(0.08)
                inner3 = webelement.get_attribute("innerHTML") or ""
                if len(self._strip_tags(inner3)) > 0:
                    return True
            except Exception:
                pass
        except Exception:
            pass
        return False

    def _set_body_html_generic(self, wait, html, selectors=None, iframe_ok=False, plain_fallback=True, container=None):
        """
        Attempt to place HTML into a contenteditable / body element.

        Strategy:
         - Try direct innerHTML assignment.
         - Try document.execCommand('insertHTML', ...) into the element.
         - On Windows, set CF_HTML on the clipboard and send Ctrl+V into the element (pastes HTML fragment).
         - Fallback to typed/plain insertion logic already present (_force_set_body_element).
        """
        driver = self.driver
        if selectors is None:
            selectors = []
        wanted_plain = self._strip_tags(html)

        # Helper: try execCommand('insertHTML') on an element
        def try_execcommand_insert(el, html_fragment):
            try:
                js = """
                var el = arguments[0];
                var html = arguments[1];
                try{ el.focus(); }catch(e){}
                var sel = window.getSelection();
                try{
                    var range = document.createRange();
                    range.selectNodeContents(el);
                    range.collapse(false);
                    sel.removeAllRanges();
                    sel.addRange(range);
                }catch(e){}
                try{
                    var ok = document.execCommand('insertHTML', false, html);
                    return !!ok;
                }catch(e){
                    try{
                        el.innerHTML = el.innerHTML + html;
                        return true;
                    }catch(e2){
                        return false;
                    }
                }
                """
                res = driver.execute_script(js, el, html_fragment)
                return bool(res)
            except Exception:
                return False

        # Windows-only: set CF_HTML clipboard content
        def set_clipboard_html_windows(html_fragment):
            try:
                fragment = "<!--StartFragment-->" + html_fragment + "<!--EndFragment-->"
                full_html = "<html><body>" + fragment + "</body></html>"
                header = "Version:0.9\r\nStartHTML:aaaaaaaaaa\r\nEndHTML:bbbbbbbbbb\r\nStartFragment:cccccccccc\r\nEndFragment:dddddddddd\r\n"
                header_bytes = header.encode("utf-8")
                html_bytes = full_html.encode("utf-8")
                start_html = len(header_bytes)
                end_html = start_html + len(html_bytes)
                frag_start_in_html = html_bytes.find(b"<!--StartFragment-->") + len(b"<!--StartFragment-->")
                frag_end_in_html = html_bytes.find(b"<!--EndFragment-->")
                if frag_start_in_html == -1 or frag_end_in_html == -1:
                    frag_start_in_html = 0
                    frag_end_in_html = len(html_bytes)
                start_fragment = start_html + frag_start_in_html
                end_fragment = start_html + frag_end_in_html
                header_filled = header.replace("aaaaaaaaaa", str(start_html).zfill(10))
                header_filled = header_filled.replace("bbbbbbbbbb", str(end_html).zfill(10))
                header_filled = header_filled.replace("cccccccccc", str(start_fragment).zfill(10))
                header_filled = header_filled.replace("dddddddddd", str(end_fragment).zfill(10))
                final = header_filled.encode("utf-8") + html_bytes

                user32 = ctypes.windll.user32
                kernel32 = ctypes.windll.kernel32

                CF_HTML = user32.RegisterClipboardFormatW("HTML Format")
                GMEM_MOVEABLE = 0x0002
                hGlobal = kernel32.GlobalAlloc(GMEM_MOVEABLE, len(final) + 1)
                if not hGlobal:
                    return False
                lpGlobal = kernel32.GlobalLock(hGlobal)
                if not lpGlobal:
                    kernel32.GlobalFree(hGlobal)
                    return False
                ctypes.memmove(lpGlobal, final, len(final))
                kernel32.GlobalUnlock(hGlobal)

                if not user32.OpenClipboard(None):
                    kernel32.GlobalFree(hGlobal)
                    return False
                try:
                    user32.EmptyClipboard()
                    if not user32.SetClipboardData(CF_HTML, hGlobal):
                        try:
                            kernel32.GlobalFree(hGlobal)
                        except Exception:
                            pass
                        return False
                finally:
                    user32.CloseClipboard()
                return True
            except Exception:
                return False

        # container-scoped attempt
        if container is not None:
            for sel in selectors:
                try:
                    rel = sel
                    if rel.startswith(".//"):
                        els = container.find_elements(By.XPATH, rel)
                    elif rel.startswith("/"):
                        els = container.find_elements(By.XPATH, "." + rel)
                    else:
                        els = container.find_elements(By.XPATH, ".//" + rel)
                except Exception:
                    els = []
                for el in els:
                    try:
                        try:
                            driver.execute_script("arguments[0].focus(); arguments[0].innerHTML = arguments[1];", el, html)
                        except Exception:
                            try:
                                driver.execute_script("arguments[0].innerHTML = arguments[1];", el, html)
                            except Exception:
                                pass
                        time.sleep(0.12)
                        try:
                            inner = el.get_attribute("innerHTML") or ""
                            got_plain = self._strip_tags(inner)
                            if (wanted_plain and wanted_plain[:60] == got_plain[:60]) or len(got_plain.strip()) > 0:
                                return True
                        except Exception:
                            return True

                        try:
                            ok = try_execcommand_insert(el, html)
                            if ok:
                                time.sleep(0.12)
                                try:
                                    inner2 = el.get_attribute("innerHTML") or ""
                                    if len(self._strip_tags(inner2)) > 0:
                                        return True
                                except Exception:
                                    return True
                        except Exception:
                            pass

                        try:
                            if sys.platform.startswith("win"):
                                ok_clip = set_clipboard_html_windows(html)
                                if ok_clip:
                                    try:
                                        el.click()
                                    except Exception:
                                        pass
                                    time.sleep(0.08)
                                    try:
                                        el.send_keys(Keys.CONTROL, "v")
                                    except Exception:
                                        try:
                                            driver.execute_script("arguments[0].focus();", el)
                                            el.send_keys(Keys.CONTROL, "v")
                                        except Exception:
                                            pass
                                    time.sleep(0.2)
                                    try:
                                        inner3 = el.get_attribute("innerHTML") or ""
                                        if len(self._strip_tags(inner3)) > 0:
                                            return True
                                    except Exception:
                                        pass
                        except Exception:
                            pass

                        if plain_fallback:
                            try:
                                ok = self._force_set_body_element(el, html)
                                if ok:
                                    return True
                            except Exception:
                                pass
                    except Exception:
                        continue

        # global attempt
        for sel in selectors:
            try:
                el = wait.until(EC.presence_of_element_located((By.XPATH, sel)))
            except Exception:
                continue
            try:
                try:
                    driver.execute_script("arguments[0].focus(); arguments[0].innerHTML = arguments[1];", el, html)
                except Exception:
                    try:
                        driver.execute_script("arguments[0].innerHTML = arguments[1];", el, html)
                    except Exception:
                        pass
                time.sleep(0.12)
                try:
                    inner = el.get_attribute("innerHTML") or ""
                    got_plain = self._strip_tags(inner)
                    if (wanted_plain and wanted_plain[:60] == got_plain[:60]) or len(got_plain.strip()) > 0:
                        return True
                except Exception:
                    return True

                try:
                    ok = try_execcommand_insert(el, html)
                    if ok:
                        time.sleep(0.12)
                        try:
                            inner2 = el.get_attribute("innerHTML") or ""
                            if len(self._strip_tags(inner2)) > 0:
                                return True
                        except Exception:
                            return True
                except Exception:
                    pass

                try:
                    if sys.platform.startswith("win"):
                        ok_clip = set_clipboard_html_windows(html)
                        if ok_clip:
                            try:
                                el.click()
                            except Exception:
                                pass
                            time.sleep(0.08)
                            try:
                                el.send_keys(Keys.CONTROL, "v")
                            except Exception:
                                try:
                                    driver.execute_script("arguments[0].focus();", el)
                                    el.send_keys(Keys.CONTROL, "v")
                                except Exception:
                                    pass
                            time.sleep(0.2)
                            try:
                                inner3 = el.get_attribute("innerHTML") or ""
                                if len(self._strip_tags(inner3)) > 0:
                                    return True
                            except Exception:
                                pass
                except Exception:
                    pass

                if plain_fallback:
                    try:
                        ok = self._force_set_body_element(el, html)
                        if ok:
                            return True
                    except Exception:
                        pass
            except (TimeoutException, StaleElementReferenceException):
                continue

        # iframe attempt
        if iframe_ok:
            try:
                iframes = driver.find_elements(By.TAG_NAME, "iframe")
                for fr in iframes:
                    try:
                        driver.switch_to.frame(fr)
                        try:
                            body = driver.find_element(By.TAG_NAME, "body")
                            ok = try_execcommand_insert(body, html)
                            driver.switch_to.default_content()
                            if ok:
                                return True
                        except Exception:
                            driver.switch_to.default_content()
                            continue
                    except Exception:
                        try:
                            driver.switch_to.default_content()
                        except Exception:
                            pass
                        continue
            except Exception:
                pass

        # final typed/paste fallback over any editable found
        if plain_fallback:
            try:
                elems = driver.find_elements(By.XPATH, "//*[(@contenteditable='true') or (@role='textbox') or @aria-label='Message Body']")
                for e in elems:
                    try:
                        ok = self._force_set_body_element(e, html)
                        if ok:
                            return True
                    except Exception:
                        continue
            except Exception:
                pass

        return False

    def _find_compose_container(self, provider):
        driver = self.driver
        try:
            elems = driver.find_elements(By.XPATH, "//div[@role='dialog'] | //div[contains(@class,'Compose')] | //div[contains(@aria-label,'New message') or contains(@aria-label,'Compose')] | //div[contains(@class,'ms-Dialog')]")
            if elems:
                return elems[-1]
        except Exception:
            pass
        try:
            panes = driver.find_elements(By.XPATH, "//div[contains(@class,'Compose') and .//div[@role='textbox']] | //div[@aria-label='Message body']/ancestor::div[1]")
            if panes:
                return panes[-1]
        except Exception:
            pass
        return None

    def _wait_for_attachment_visible(self, driver, filename, timeout=12):
        """
        Wait until an element inside the compose shows the attachment filename.
        Returns True if found, False on timeout.
        """
        try:
            end = time.time() + timeout
            short_name = os.path.basename(filename).strip()
            if not short_name:
                return False
            try:
                container = self._find_compose_container(self.provider_var.get()) or driver
            except Exception:
                container = driver
            while time.time() < end:
                try:
                    xpath = f".//*[contains(text(), '{short_name}')]"
                    els = container.find_elements(By.XPATH, xpath)
                    if els:
                        return True
                    els2 = driver.find_elements(By.XPATH, xpath)
                    if els2:
                        return True
                except Exception:
                    pass
                time.sleep(0.4)
        except Exception:
            pass
        return False

    def _attach_file_to_compose(self, driver):
        """
        Attempt to attach the file at self.attachment_path to the currently open compose.
        Returns True if we found a file input and attempted to send_keys the file.
        """
        try:
            if not getattr(self, "attachment_path", None):
                return False
            attach_path = os.path.abspath(self.attachment_path)
            self.log(f"Attempting to attach file: {attach_path}")
            container = None
            try:
                container = self._find_compose_container(self.provider_var.get())
            except Exception:
                container = None
            file_inputs = []
            try:
                if container:
                    file_inputs = container.find_elements(By.XPATH, ".//input[@type='file']")
                if not file_inputs:
                    file_inputs = driver.find_elements(By.XPATH, "//input[@type='file']")
            except Exception:
                try:
                    file_inputs = driver.find_elements(By.XPATH, "//input[@type='file']")
                except Exception:
                    file_inputs = []
            file_input = None
            for fi in reversed(file_inputs):
                try:
                    if fi.is_displayed():
                        file_input = fi
                        break
                except Exception:
                    file_input = fi
            if not file_input and file_inputs:
                file_input = file_inputs[-1]
            if file_input:
                try:
                    file_input.send_keys(attach_path)
                except Exception as e:
                    self.log(f"Attachment send_keys failed: {e}")
                    return False
                # Wait for visible upload indicator (filename) to appear before returning
                appeared = self._wait_for_attachment_visible(driver, attach_path, timeout=14)
                if appeared:
                    self.log("Attachment visible in compose (upload likely finished).")
                else:
                    self.log("Attachment upload indicator not detected within timeout; proceeding anyway.")
                return True
            else:
                self.log("No file input found on compose to attach file. Attachment not added automatically.")
                return False
        except Exception as e:
            self.log(f"Attachment attempt error: {e}")
            return False

    def _set_recipient_general(self, email):
        """
        Multiple attempts to set recipient:
         - Prefer combobox / tokenizer input (role='combobox' or ms picker).
         - Type the address character‑by‑character to trigger provider handlers.
         - Use clipboard paste fallback if needed.
        """
        driver = self.driver
        self.log(f"_set_recipient_general: attempting to set recipient to {email}")
        lower_email = email.lower()

        def tokenized_in_container(container):
            try:
                inner = (container.get_attribute("innerHTML") or "").lower()
                if lower_email in inner:
                    return True
                tokens = container.find_elements(By.XPATH, f".//*[contains(text(), '{email}')]")
                if tokens and len(tokens) > 0:
                    return True
            except Exception:
                pass
            return False

        # attempt 1: container-scoped combobox typing
        try:
            container = self._find_compose_container(self.provider_var.get())
            if container:
                combobox_xpaths = [
                    ".//div[@role='combobox']//input",
                    ".//div[contains(@class,'ms-BasePicker')]//input",
                    ".//input[@aria-label='To']",
                    ".//input[contains(@aria-label,'To')]",
                    ".//textarea[@name='to']",
                    ".//div[contains(@class,'recipient-row')]//input"
                ]
                for xp in combobox_xpaths:
                    try:
                        inputs = container.find_elements(By.XPATH, xp)
                        if not inputs:
                            continue
                        inp = inputs[-1]
                        try:
                            inp.click()
                        except Exception:
                            pass
                        try:
                            inp.clear()
                        except Exception:
                            try:
                                driver.execute_script("arguments[0].value='';", inp)
                            except Exception:
                                pass
                        for ch in email:
                            try:
                                inp.send_keys(ch)
                            except Exception:
                                try:
                                    driver.execute_script("arguments[0].value = (arguments[0].value || '') + arguments[1]; arguments[0].dispatchEvent(new Event('input',{bubbles:true}));", inp, ch)
                                except Exception:
                                    pass
                            time.sleep(0.02)
                        try:
                            inp.send_keys(Keys.ENTER)
                        except Exception:
                            try:
                                inp.send_keys("\n")
                            except Exception:
                                pass
                        end = time.time() + 4.0
                        while time.time() < end:
                            if tokenized_in_container(container):
                                self.log("_set_recipient_general: tokenization detected inside container")
                                return True
                            try:
                                val = (inp.get_attribute("value") or "").strip()
                                if val == "":
                                    if tokenized_in_container(container):
                                        return True
                            except Exception:
                                pass
                            time.sleep(0.12)
                    except Exception:
                        continue
        except Exception as e:
            self.log(f"_set_recipient_general: container combobox attempt error: {e}")

        # attempt 2: global inputs
        try:
            global_xps = [
                "//div[@role='combobox']//input",
                "//input[@aria-label='To']",
                "//input[contains(@aria-label,'To')]",
                "//textarea[@name='to']",
                "//input[@type='email']",
                "//div[contains(@class,'ms-BasePicker')]//input"
            ]
            for xp in global_xps:
                try:
                    candidate = None
                    els = driver.find_elements(By.XPATH, xp)
                    if not els:
                        continue
                    candidate = els[-1]
                    try:
                        candidate.click()
                    except Exception:
                        pass
                    try:
                        candidate.clear()
                    except Exception:
                        try:
                            driver.execute_script("arguments[0].value='';", candidate)
                        except Exception:
                            pass
                    for ch in email:
                        try:
                            candidate.send_keys(ch)
                        except Exception:
                            try:
                                driver.execute_script("arguments[0].value = (arguments[0].value || '') + arguments[1]; arguments[0].dispatchEvent(new Event('input',{bubbles:true}));", candidate, ch)
                            except Exception:
                                pass
                        time.sleep(0.02)
                    try:
                        candidate.send_keys(Keys.ENTER)
                    except Exception:
                        try:
                            candidate.send_keys("\n")
                        except Exception:
                            pass
                    end = time.time() + 4.0
                    while time.time() < end:
                        try:
                            val = (candidate.get_attribute("value") or "").strip()
                            if val == "":
                                page_html = (driver.execute_script("return document.body.innerHTML;") or "").lower()
                                if lower_email in page_html:
                                    self.log("_set_recipient_general: tokenization detected globally")
                                    return True
                        except Exception:
                            pass
                        time.sleep(0.12)
                except Exception:
                    continue
        except Exception as e:
            self.log(f"_set_recipient_general: global combobox attempt error: {e}")

        # attempt 3: JS-assisted set
        try:
            js = """
            var val = arguments[0];
            var inputs = Array.from(document.querySelectorAll('input, textarea'));
            for(var i=0;i<inputs.length;i++){
                var a = (inputs[i].getAttribute('aria-label')||'').toLowerCase();
                var p = (inputs[i].getAttribute('placeholder')||'').toLowerCase();
                if(a.indexOf('to')!==-1 || p.indexOf('to')!==-1 || inputs[i].type==='email'){
                    try{ inputs[i].focus(); inputs[i].value = val; inputs[i].dispatchEvent(new Event('input',{bubbles:true})); inputs[i].dispatchEvent(new KeyboardEvent('keydown',{key:'Enter'})); return true; }catch(e){}
                }
            }
            return false;
            """
            ok = driver.execute_script(js, email)
            if ok:
                time.sleep(0.18)
                container2 = self._find_compose_container(self.provider_var.get())
                if container2 and (lower_email in (container2.get_attribute("innerHTML") or "").lower()):
                    self.log("_set_recipient_general: set via JS fallback and detected in container")
                    return True
        except Exception as e:
            self.log(f"_set_recipient_general: js fallback error: {e}")

        # attempt 4: clipboard paste fallback
        try:
            try:
                self.root.clipboard_clear()
                self.root.clipboard_append(email)
                self.root.update()
            except Exception:
                pass

            paste_key = Keys.COMMAND if sys.platform.startswith("darwin") else Keys.CONTROL
            try:
                el = driver.find_element(By.XPATH, "//div[@role='combobox']//input | //input[@aria-label='To'] | //input[contains(@aria-label,'To')]")
                try:
                    el.click()
                except Exception:
                    pass
                time.sleep(0.12)
                try:
                    el.send_keys(paste_key, "v")
                except Exception:
                    try:
                        driver.execute_script("arguments[0].focus(); arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('input',{bubbles:true}));", el, email)
                    except Exception:
                        pass
                try:
                    el.send_keys(Keys.TAB)
                except Exception:
                    pass
                end = time.time() + 3.0
                while time.time() < end:
                    try:
                        if tokenized_in_container(self._find_compose_container(self.provider_var.get()) or driver):
                            self.log("_set_recipient_general: tokenization detected after clipboard paste")
                            return True
                        if lower_email in (driver.page_source or "").lower():
                            return True
                    except Exception:
                        pass
                    time.sleep(0.12)
            except Exception:
                pass
        except Exception as e:
            self.log(f"_set_recipient_general: clipboard fallback error: {e}")

        # final: screenshot and return False
        try:
            fname = os.path.join(SCREENSHOT_DIR, f"recipient_fail_{int(time.time())}.png")
            driver.save_screenshot(fname)
            self.log(f"_set_recipient_general: failed to set recipient, screenshot saved: {fname}")
        except Exception:
            pass
        return False

    def _ensure_compose_open(self, provider, max_attempts=1):
        driver = self.driver
        provider_lc = (provider or "").lower()
        try:
            if "gmail" in provider_lc:
                try:
                    elems = driver.find_elements(By.XPATH, "//div[@role='dialog']")
                    if elems:
                        return elems[-1]
                except Exception:
                    pass
                try:
                    btn = driver.find_element(By.XPATH, "//div[text()='Compose'] | //button[@aria-label='Compose']")
                    try:
                        driver.execute_script("arguments[0].click();", btn)
                    except Exception:
                        btn.click()
                    WebDriverWait(driver, 3).until(lambda d: len(d.find_elements(By.XPATH, "//div[@role='dialog']")) > 0)
                    elems = driver.find_elements(By.XPATH, "//div[@role='dialog']")
                    if elems:
                        return elems[-1]
                except Exception:
                    try:
                        elems = driver.find_elements(By.XPATH, "//div[@role='dialog']")
                        if elems:
                            return elems[-1]
                    except Exception:
                        return None
            elif "outlook" in provider_lc:
                try:
                    panes = driver.find_elements(By.XPATH, "//div[@aria-label='Message body']/ancestor::div[1]")
                    if panes:
                        return panes[-1]
                except Exception:
                    pass
                try:
                    btn = driver.find_element(By.XPATH, "//button[@aria-label='New mail']")
                    try:
                        driver.execute_script("arguments[0].click();", btn)
                    except Exception:
                        btn.click()
                    WebDriverWait(driver, 3).until(lambda d: len(d.find_elements(By.XPATH, "//div[@role='textbox'] | //div[@aria-label='Message body']"))>0)
                    return None
                except Exception:
                    return None
            elif "comcast" in provider_lc or "xfinity" in provider_lc:
                try:
                    btn = driver.find_element(By.XPATH, "//button[contains(@title,'Compose') or contains(@title,'compose')]")
                    try:
                        driver.execute_script("arguments[0].click();", btn)
                    except Exception:
                        btn.click()
                    WebDriverWait(driver, 3).until(lambda d: len(d.find_elements(By.XPATH, "//div[contains(@aria-label,'Message body')] | //div[@role='textbox']"))>0)
                    return None
                except Exception:
                    return None
        except Exception:
            return None
        return None

    def _wait_for_send_toast(self, provider, timeout=8):
        driver = self.driver
        provider_lc = (provider or "").lower()
        end = time.time() + timeout
        self.log(f"_wait_for_send_toast: checking for {provider} toast/confirmation")
        try:
            if "gmail" in provider_lc:
                while time.time() < end:
                    try:
                        selectors = [
                            "//*[@role='alert' and (contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'message sent') or contains(., 'Message sent'))]",
                            "//span[contains(., 'Message sent') or contains(., 'message sent')]",
                        ]
                        for selector in selectors:
                            elements = driver.find_elements(By.XPATH, selector)
                            if elements:
                                return True
                        compose_dialogs = driver.find_elements(By.XPATH, "//div[@role='dialog']")
                        if not compose_dialogs:
                            return True
                    except Exception:
                        pass
                    time.sleep(0.12)
                return False
            elif "outlook" in provider_lc:
                while time.time() < end:
                    try:
                        selectors = [
                            "//*[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'sent') and (contains(@role,'status') or contains(@class,'Toast') or contains(@class,'ms-Toast') or contains(@class,'ms-MessageBar'))]",
                            "//*[contains(text(), 'Your message was sent')]",
                        ]
                        for selector in selectors:
                            elements = driver.find_elements(By.XPATH, selector)
                            if elements:
                                return True
                        try:
                            compose_elements = driver.find_elements(By.XPATH, "//div[@role='textbox'] | //div[contains(@aria-label,'Message body') or contains(@class,'Compose')]")
                            if not compose_elements:
                                return True
                        except Exception:
                            pass
                        current_url = driver.current_url or ""
                        if "sentitems" in current_url.lower():
                            return True
                    except Exception:
                        pass
                    time.sleep(0.12)
                return False
            elif "comcast" in provider_lc or "xfinity" in provider_lc:
                while time.time() < end:
                    try:
                        selectors = [
                            "//*[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'sent')]",
                        ]
                        for selector in selectors:
                            elements = driver.find_elements(By.XPATH, selector)
                            if elements:
                                return True
                        page_source = (driver.page_source or "").lower()
                        if "message sent" in page_source or "email sent" in page_source:
                            return True
                    except Exception:
                        pass
                    time.sleep(0.12)
                return False
            # generic
            while time.time() < end:
                try:
                    generic_selectors = [
                        "//*[@role='alert']",
                        "//*[contains(@class, 'toast')]",
                    ]
                    for selector in generic_selectors:
                        alerts = driver.find_elements(By.XPATH, selector)
                        for a in alerts:
                            text = (a.text or "").lower()
                            if "sent" in text or "message sent" in text:
                                return True
                except Exception:
                    pass
                time.sleep(0.12)
        except Exception:
            pass
        return False

    def _confirm_sent(self, provider, recipient_email, subject="Bulk Email", timeout=15):
        driver = self.driver
        end = time.time() + timeout
        provider = (provider or "").lower()
        self.log(f"_confirm_sent: checking sent folder for {provider}")
        if "outlook" in provider:
            sent_urls = [
                "https://outlook.live.com/mail/0/sentitems",
                "https://outlook.live.com/mail/0/sent",
                "https://outlook.live.com/mail/0/"
            ]
            try:
                current_url = driver.current_url or ""
                current_page = (driver.page_source or "").lower()
                if ("sentitems" in current_url or "sent" in current_url) and recipient_email.lower() in current_page:
                    return True
            except Exception:
                pass
            while time.time() < end:
                for url in sent_urls:
                    try:
                        driver.get(url)
                        time.sleep(1.2)
                        try:
                            WebDriverWait(driver, 5).until(
                                EC.any_of(
                                    EC.presence_of_element_located((By.XPATH, "//div[contains(@role,'list')]")),
                                    EC.presence_of_element_located((By.XPATH, "//div[contains(@class,'sentItems')]")),
                                )
                            )
                        except TimeoutException:
                            pass
                        page = (driver.page_source or "").lower()
                        if recipient_email.lower() in page:
                            return True
                        if subject and subject.lower() in page:
                            return True
                    except Exception:
                        pass
                time.sleep(1.0)
            return False
        elif "gmail" in provider:
            sent_url = "https://mail.google.com/mail/u/0/#sent"
            while time.time() < end:
                try:
                    driver.get(sent_url)
                    time.sleep(1.2)
                    page = (driver.page_source or "").lower()
                    if recipient_email.lower() in page:
                        return True
                    if subject and subject.lower() in page:
                        return True
                except Exception:
                    pass
                time.sleep(1.0)
            return False
        elif "comcast" in provider_lc or "xfinity" in provider_lc:
            sent_url = "https://connect.xfinity.com/appsuite/#!!&app=mail&folder=sent"
            while time.time() < end:
                try:
                    driver.get(sent_url)
                    time.sleep(1.5)
                    page = (driver.page_source or "").lower()
                    if recipient_email.lower() in page:
                        return True
                    if subject and subject.lower() in page:
                        return True
                except Exception:
                    pass
                time.sleep(1.0)
            return False
        while time.time() < end:
            try:
                page = (driver.page_source or "").lower()
                if recipient_email.lower() in page:
                    return True
                if subject and subject.lower() in page:
                    return True
            except Exception:
                pass
            time.sleep(1.0)
        return False

    # ---------------------------
    # Gmail rich-text helper (best-effort)
    # ---------------------------
    def _ensure_gmail_rich_text(self, driver, timeout=3):
        """
        Best-effort: open Compose "More options" and ensure Plain text mode is OFF.
        Returns True if action performed or seems unnecessary, False on failure.
        """
        try:
            end = time.time() + timeout
            container = None
            try:
                container = self._find_compose_container("Gmail")
            except Exception:
                container = None
            # try common selectors for the "More options" (three-dot) button in Gmail compose
            cand_xpaths = [
                ".//div[@aria-label='More options' and @role='button']",
                ".//div[@aria-label='More options']",
                "//div[@aria-label='More options' and @role='button']",
                "//div[contains(@aria-label,'More options') and @role='button']",
                "//span[@aria-label='More options']"
            ]
            menu_btn = None
            for xp in cand_xpaths:
                try:
                    if container:
                        els = container.find_elements(By.XPATH, xp)
                    else:
                        els = driver.find_elements(By.XPATH, xp)
                    if els:
                        menu_btn = els[-1]
                        break
                except Exception:
                    continue
            if not menu_btn:
                return True  # can't find button, assume compose is already in rich mode
            try:
                driver.execute_script("arguments[0].click();", menu_btn)
            except Exception:
                try:
                    menu_btn.click()
                except Exception:
                    return True
            time.sleep(0.35)
            # look for menu item that mentions "Plain text mode"
            try:
                items = driver.find_elements(By.XPATH, "//div[@role='menuitem'] | //div[contains(@role,'option')]")
                for it in items:
                    try:
                        txt = (it.text or "").strip().lower()
                        if "plain text" in txt:
                            # If the option is present, ensure it is OFF.
                            # The menu item toggles the state; safest approach: if it contains an indicator that it's active, click it to toggle off.
                            if ("✓" in txt) or ("on" in txt and "plain text" in txt):
                                try:
                                    driver.execute_script("arguments[0].click();", it)
                                except Exception:
                                    try:
                                        it.click()
                                    except Exception:
                                        pass
                                time.sleep(0.18)
                            # close menu
                            try:
                                body = driver.find_element(By.TAG_NAME, "body")
                                body.click()
                            except Exception:
                                pass
                            return True
            except Exception:
                pass
            # fallback: click body to close menu
            try:
                body = driver.find_element(By.TAG_NAME, "body")
                body.click()
            except Exception:
                pass
            return True
        except Exception:
            return True

    # ---------------------------
    # insert docuSign template helper
    # ---------------------------
    def insert_docusign_template(self):
        """
        Inserts a DocuSign HTML template into the message editor and
        populates Sender and Review URL fields with sensible defaults.
        The template uses %SENDER% and REVIEW_DOCUMENT_URL placeholders
        so the existing injection logic will replace them on send.
        """
        try:
            template = """<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>DocuSign - Please DocuSign: Document</title>
</head>
<body style="margin:0; padding:0; background:#f4f4f4; -webkit-text-size-adjust:none; font-family:Arial,Helvetica,sans-serif; color:#222222;">
  <div style="display:none;max-height:0px;overflow:hidden;mso-hide:all;font-size:1px;color:#f4f4f4;line-height:1px;opacity:0;">
    %SENDER% sent you a document to review and sign. Click the button to review the document.
  </div>

  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#f4f4f4" style="width:100%;background:#f4f4f4;">
    <tr>
      <td align="center" style="padding:24px 12px;">
        <table role="presentation" width="600" cellpadding="0" cellspacing="0" border="0" style="width:100%;max-width:600px;background:#ffffff;border-radius:6px;overflow:hidden;border:1px solid #e9eef5;">
          <tr>
            <td style="padding:28px 28px 18px 28px;">
              <img src="" alt="" width="1" height="1" style="display:block;border:0;outline:none;text-decoration:none;" />
            </td>
          </tr>

          <tr>
            <td align="center" bgcolor="#1f61c3" style="padding:36px 28px 36px 28px;background:#1f61c3;color:#ffffff;text-align:center;">
              <h1 style="margin:0;font-size:20px;line-height:1.35;font-weight:600;color:#ffffff;font-family:Arial,Helvetica,sans-serif;">
                %SENDER% sent you a document to review and sign.
              </h1>
              <p style="margin:12px 0 28px 0;font-size:15px;color:#ffffff;line-height:1.4;">
                Electronic signature required
              </p>
              <a href="REVIEW_DOCUMENT_URL" target="_blank" rel="noopener"
                 style="display:inline-block;padding:12px 24px;background:#ffc72c;color:#183153;font-weight:700;text-decoration:none;border-radius:6px;font-size:14px;box-shadow:0 2px 8px rgba(0,0,0,0.08);">
                REVIEW DOCUMENT
              </a>
            </td>
          </tr>

          <tr>
            <td style="padding:24px 28px 0 28px;">
              <div style="font-size:15px;color:#222222;line-height:1.6;">
                <strong>Summary:</strong><br>
                You have been sent a copy of a completed document for your records. Please review the document as needed.
              </div>
            </td>
          </tr>

          <tr>
            <td style="padding:18px 28px 0 28px;">
              <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;">
                <tr>
                  <td style="padding:12px;background:#fafbff;border:1px solid #eef2ff;border-radius:6px;font-size:14px;color:#333;">
                    Document: <strong>Document Title or Reference</strong><br/>
                    Sender: <strong>%SENDER%</strong>
                  </td>
                </tr>
              </table>
            </td>
          </tr>

          <tr>
            <td style="padding:22px 28px 28px 28px;">
              <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">
                <tr>
                  <td style="font-size:13px;color:#666666;line-height:1.6;padding-bottom:12px;">
                    <strong style="color:#222222;font-size:14px;display:block;margin-bottom:6px;">Do Not Share This Email</strong>
                    This email contains a secure link to DocuSign. Please do not share this email, link, or access code with others.
                  </td>
                </tr>
                <tr>
                  <td style="font-size:13px;color:#666666;line-height:1.6;padding-bottom:12px;">
                    <strong style="color:#222222;font-size:14px;display:block;margin-bottom:6px;">About DocuSign</strong>
                    Sign documents electronically in just minutes. It's safe, secure, and legally binding.
                  </td>
                </tr>
                <tr>
                  <td style="font-size:13px;color:#666666;line-height:1.6;padding-bottom:10px;">
                    <strong style="color:#222222;font-size:14px;display:block;margin-bottom:6px;">Questions about the Document?</strong>
                    If you need to modify the document or have questions, please reach out to the sender.
                  </td>
                </tr>
                <tr>
                  <td style="padding-top:10px;border-top:1px solid #eef2f7;font-size:12px;color:#666666;">
                    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">
                      <tr>
                        <td style="padding-top:10px;font-size:12px;color:#666666;">&#10003; SSL Encrypted</td>
                        <td style="padding-top:10px;font-size:12px;color:#666666;">&#10003; Legally Binding</td>
                        <td style="padding-top:10px;font-size:12px;color:#666666;">&#10003; ESIGN Act Compliant</td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>

        </table>
      </td>
    </tr>
  </table>
</body>
</html>"""
            try:
                self.message_box.delete("1.0", tk.END)
            except Exception:
                pass
            self.message_box.insert(tk.END, template)
            # Pre-fill Sender + Review URL defaults so _inject_placeholders will do its job
            try:
                self.sender_var.set("Sabaheta")
                # set a placeholder review URL; replace with hosted URL or your document link before send
                self.review_var.set("https://example.com/review")
            except Exception:
                pass
            self.log("Inserted DocuSign template into editor and set Sender + Review URL defaults. Edit REVIEW URL as needed.")
        except Exception as e:
            self.log(f"insert_docusign_template failed: {e}")

    # ---------------------------
    # misc helpers (continued)
    # ---------------------------
    def _set_input_value_generic(self, wait, selectors, value, allow_enter=False):
        driver = self.driver
        for sel in selectors:
            try:
                el = wait.until(EC.presence_of_element_located((By.XPATH, sel)))
                try:
                    el.click()
                except Exception:
                    pass
                try:
                    el.clear()
                except Exception:
                    pass
                try:
                    el.send_keys(value)
                    if allow_enter:
                        el.send_keys("\n")
                except Exception:
                    try:
                        # fixed JS: dispatchEvent call must be a function call
                        driver.execute_script("arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('input',{bubbles:true}));", el, value)
                        if allow_enter:
                            driver.execute_script("arguments[0].dispatchEvent(new KeyboardEvent('keydown', {'key':'Enter'}));", el)
                    except Exception:
                        continue
                return True
            except (TimeoutException, NoSuchElementException, StaleElementReferenceException):
                continue
        try:
            js = "var v = arguments[0]; var sel = arguments[1]; var el = document.querySelector(sel); if(el){ el.value = v; el.dispatchEvent(new Event('input',{bubbles:true})); return true; } return false;"
            for q in ["input[aria-label*='To']", "textarea[aria-label*='To']", "input[placeholder*='To']"]:
                try:
                    ok = self.driver.execute_script(js, value, q)
                    if ok:
                        return True
                except Exception:
                    pass
        except Exception:
            pass
        return False

    def action_worker(self):
        while True:
            try:
                action = self.action_queue.get(timeout=1)
            except queue.Empty:
                time.sleep(0.2)
                continue
            if not action:
                self.action_queue.task_done()
                continue
            if action[0] == "auto_reply":
                _, email, track_id = action
                now = time.time()
                last = self.recent_replies.get(email, 0)
                if now - last < 3600:
                    self.log(f"Skipping auto-reply to {email}: recently replied.")
                    self.action_queue.task_done()
                    continue
                reply_html = "<html><body><p>Hi,</p><p>Thanks for opening the message. This is an automated reply.</p></body></html>"
                try:
                    with self.driver_lock:
                        if self.driver:
                            self.send_email(email, reply_html, self.provider_var.get(), subject_text="Auto-reply")
                            self.recent_replies[email] = now
                            entry = self.tracking_map.get(track_id)
                            if entry:
                                entry["replied"] = True
                                self.gui_update_queue.put(("update_replied", entry["email"], "Yes"))
                        else:
                            self.log("Cannot send auto-reply: driver not available.")
                except Exception as e:
                    self.log(f"Auto-reply failed: {e}\n{traceback.format_exc()}")
            self.action_queue.task_done()

    def start_tracking_server_thread(self):
        if Flask is None:
            messagebox.showwarning("Flask missing", "Install Flask (pip install flask) to enable tracking server.")
            return
        t = threading.Thread(target=self._run_tracking_server, daemon=True)
        t.start()

    def _run_tracking_server(self):
        app = Flask("tracking_server")
        @app.route("/track")
        def track():
            track_id = request.args.get("id", "")
            ua = request.headers.get("User-Agent", "")
            ip = request.remote_addr
            ts = time.strftime("%Y-%m-%d %H:%M:%S")
            if not track_id:
                return self._one_by_one_gif_response()
            entry = self.tracking_map.get(track_id)
            if entry:
                entry["opened_time"] = ts
                entry["status"] = "Opened"
                self.log(f"Tracking pixel hit for {entry['email']} (id {track_id}) from {ip} ua:{ua}")
                self.gui_update_queue.put(("update_status", entry["email"], "Opened", entry.get("sent_time", ""), ts, "No"))
                if not entry.get("replied", False):
                    self.action_queue.put(("auto_reply", entry["email"], track_id))
                    entry["replied"] = True
                    self.gui_update_queue.put(("update_replied", entry["email"], "Yes"))
            else:
                self.log(f"Unknown tracking id hit: {track_id} from {ip} ua:{ua}")
            return self._one_by_one_gif_response()

        @app.route("/file/<fid>/<filename>")
        def serve_file(fid, filename):
            meta = self.hosted_files.get(fid)
            if not meta:
                return ("Not found", 404)
            path = meta.get("path")
            if not path or not os.path.exists(path):
                return ("Not found", 404)
            try:
                return send_file(path)
            except Exception as e:
                self.log(f"Error serving hosted file {fid}: {e}")
                return ("Server error", 500)

        server_ip = self._detect_local_ip()
        self.log(f"Starting tracking server at http://{server_ip}:{self.TRACK_SERVER_PORT}/ - pixel /track and files /file/<id>/<name>")
        try:
            from waitress import serve
            serve(app, host="0.0.0.0", port=self.TRACK_SERVER_PORT)
        except Exception as e:
            self.log(f"Waitress not available or failed: {e}. Falling back to Flask dev server.")
            app.run(host="0.0.0.0", port=self.TRACK_SERVER_PORT, threaded=True, debug=False, use_reloader=False)

    def _one_by_one_gif_response(self):
        gif_bytes = base64.b64decode(b'R0lGODlhAQABAIABAP///wAAACwAAAAAAQABAAACAkQBADs=')
        resp = make_response(gif_bytes)
        resp.headers.set('Content-Type', 'image/gif')
        resp.headers.set('Cache-Control', 'no-cache, no-store, must-revalidate')
        return resp

    def _find_any(self, *locators):
        driver = self.driver
        for loc in locators:
            try:
                by, expr = loc
                el = driver.find_element(by, expr)
                if el:
                    return el
            except Exception:
                continue
        return None

    def _detect_local_ip(self):
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.connect(("8.8.8.8", 80))
            ip = s.getsockname()[0]
            s.close()
            return ip
        except Exception:
            return "127.0.0.1"

    # ---------------------------
    # send_email implementation (Selenium automation for providers)
    # ---------------------------
    def send_email(self, recipient, html_body, provider, subject_text=""):
        """
        Provider-specific send via Selenium (fills To, Subject, Body, attachments and sends).
        """
        driver = self.driver
        wait = WebDriverWait(driver, 12)
        provider_lc = (provider or "").lower()
        subject = subject_text or self.subject_var.get() or "Bulk Email"
        try:
            # For Gmail
            if "gmail" in provider_lc:
                # open compose
                try:
                    btn = self._find_any((By.XPATH, "//div[text()='Compose']"), (By.XPATH, "//button[@aria-label='Compose']"))
                    if btn:
                        try:
                            driver.execute_script("arguments[0].click();", btn)
                        except Exception:
                            btn.click()
                        time.sleep(0.5)
                except Exception:
                    pass

                # set recipient
                ok = self._set_recipient_general(recipient)
                if not ok:
                    raise RuntimeError("Failed to set recipient in Gmail")

                # set subject
                self._set_input_value_generic(wait, ["//input[@name='subjectbox']", "//input[@placeholder='Subject']"], subject, allow_enter=False)

                # Prepare fragment and ensure placeholders are replaced (we already injected in start_sending but guard here)
                fragment = self._extract_body_fragment(html_body)
                fragment = self._inject_placeholders(fragment)
                fragment = self._ensure_review_fallback(fragment)

                # Ensure Gmail is in rich-text mode (best effort)
                try:
                    self._ensure_gmail_rich_text(driver)
                except Exception:
                    pass

                # set body
                body_selectors = [
                    "//div[@aria-label='Message Body']",
                    "//div[@role='textbox']",
                    "//div[contains(@class,'Am') and @contenteditable='true']",
                    "//div[@aria-label='Message Body' and @contenteditable='true']",
                ]
                ok_body = self._set_body_html_generic(wait, fragment, selectors=body_selectors, iframe_ok=False, plain_fallback=True)
                if not ok_body:
                    raise RuntimeError("Failed to set body in Gmail")

                # DEBUG: read compose innerHTML and save to a temp file for inspection
                try:
                    compose_el = driver.find_element(By.XPATH, "//div[@aria-label='Message Body'] | //div[@role='textbox']")
                    inner = driver.execute_script("return arguments[0].innerHTML;", compose_el) or ""
                    preview = (inner[:400].replace("\n", "\\n") if inner else "<empty>")
                    self.log("DEBUG: compose innerHTML (first 400 chars): " + preview)
                    tfc = tempfile.NamedTemporaryFile(delete=False, suffix=".html", prefix="paris_compose_preview_", mode="w", encoding="utf-8")
                    tfc.write(inner)
                    tfc.close()
                    self.log(f"DEBUG: saved compose preview to {tfc.name}")
                except Exception as e:
                    self.log(f"DEBUG: could not read compose innerHTML: {e}")

                # Robust attach: attempt to attach the chosen file automatically if available
                try:
                    self._attach_file_to_compose(driver)
                except Exception:
                    pass

                # click send
                try:
                    send_btns = driver.find_elements(By.XPATH, "//div[text()='Send'] | //button[@aria-label='Send']")
                    if send_btns:
                        try:
                            driver.execute_script("arguments[0].click();", send_btns[-1])
                        except Exception:
                            send_btns[-1].click()
                    else:
                        # fallback: press Ctrl+Enter
                        try:
                            body_el = driver.find_element(By.XPATH, "//div[@role='textbox'] | //div[@aria-label='Message Body']")
                            body_el.send_keys(Keys.CONTROL, Keys.ENTER)
                        except Exception:
                            pass
                except Exception:
                    pass

                # confirm
                ok = self._wait_for_send_toast("gmail", timeout=8)
                if not ok:
                    confirmed = self._confirm_sent("gmail", recipient, subject=subject, timeout=8)
                    if not confirmed:
                        raise RuntimeError("Send not confirmed for Gmail")
                return True

            # For Outlook.com (web)
            elif "outlook" in provider_lc:
                # open compose if needed
                try:
                    try:
                        btn = driver.find_element(By.XPATH, "//button[@aria-label='New mail'] | //button[contains(., 'New message')]")
                        try:
                            driver.execute_script("arguments[0].click();", btn)
                        except Exception:
                            btn.click()
                        time.sleep(0.6)
                    except Exception:
                        pass
                except Exception:
                    pass

                # set recipient
                ok = self._set_recipient_general(recipient)
                if not ok:
                    raise RuntimeError("Failed to set recipient in Outlook.com")

                # set subject
                self._set_input_value_generic(wait, [
                    "//input[@aria-label='Add a subject']",
                    "//input[@placeholder='Add a subject']",
                    "//input[contains(@aria-label,'Subject')]",
                    "//input[@name='subject']"
                ], subject, allow_enter=False)

                # set body (outlook uses role='textbox' or aria-label message body)
                body_selectors = [
                    "//div[@role='textbox']",
                    "//div[@aria-label='Message body']",
                    "//div[contains(@class,'ms-TextField') and @role='textbox']"
                ]
                fragment = self._extract_body_fragment(html_body)
                fragment = self._inject_placeholders(fragment)
                fragment = self._ensure_review_fallback(fragment)
                ok_body = self._set_body_html_generic(wait, fragment, selectors=body_selectors, iframe_ok=False, plain_fallback=True)
                if not ok_body:
                    raise RuntimeError("Failed to set body in Outlook.com")

                # attachments: attempt robust attach as well
                try:
                    self._attach_file_to_compose(driver)
                except Exception:
                    pass

                # click send
                try:
                    send_btn = self._find_any((By.XPATH, "//button[@title='Send']"), (By.XPATH, "//button[contains(., 'Send')]"))
                    if send_btn:
                        try:
                            driver.execute_script("arguments[0].click();", send_btn)
                        except Exception:
                            send_btn.click()
                    else:
                        # fallback: press Ctrl+Enter
                        try:
                            body_el = driver.find_element(By.XPATH, "//div[@role='textbox'] | //div[@aria-label='Message body']")
                            body_el.send_keys(Keys.CONTROL, Keys.ENTER)
                        except Exception:
                            pass
                except Exception:
                    pass

                ok = self._wait_for_send_toast("outlook", timeout=8)
                if not ok:
                    confirmed = self._confirm_sent("outlook", recipient, subject=subject, timeout=8)
                    if not confirmed:
                        raise RuntimeError("Send not confirmed for Outlook.com")
                return True

            # For Comcast / Xfinity
            elif "comcast" in provider_lc or "xfinity" in provider_lc:
                try:
                    btn = self._find_any((By.XPATH, "//button[contains(@title,'Compose') or contains(., 'Compose')]"), (By.XPATH, "//button[contains(@class,'compose')]"))
                    if btn:
                        try:
                            driver.execute_script("arguments[0].click();", btn)
                        except Exception:
                            btn.click()
                        time.sleep(0.6)
                except Exception:
                    pass

                ok = self._set_recipient_general(recipient)
                if not ok:
                    raise RuntimeError("Failed to set recipient in Comcast/Xfinity")

                self._set_input_value_generic(wait, ["//input[contains(@placeholder,'Subject')]", "//input[@name='subject']"], subject, allow_enter=False)

                body_selectors = [
                    "//div[@role='textbox']",
                    "//div[contains(@class,'editor') and @contenteditable='true']"
                ]
                fragment = self._extract_body_fragment(html_body)
                fragment = self._inject_placeholders(fragment)
                fragment = self._ensure_review_fallback(fragment)
                ok_body = self._set_body_html_generic(wait, fragment, selectors=body_selectors, iframe_ok=False, plain_fallback=True)
                if not ok_body:
                    raise RuntimeError("Failed to set body in Comcast")

                # attachments
                try:
                    self._attach_file_to_compose(driver)
                except Exception:
                    pass

                try:
                    send_btn = self._find_any((By.XPATH, "//button[contains(., 'Send')]"), (By.XPATH, "//button[@title='Send']"))
                    if send_btn:
                        try:
                            driver.execute_script("arguments[0].click();", send_btn)
                        except Exception:
                            send_btn.click()
                    else:
                        try:
                            body_el = driver.find_element(By.XPATH, "//div[@role='textbox']")
                            body_el.send_keys(Keys.CONTROL, Keys.ENTER)
                        except Exception:
                            pass
                except Exception:
                    pass

                ok = self._wait_for_send_toast("comcast", timeout=8)
                if not ok:
                    confirmed = self._confirm_sent("comcast", recipient, subject=subject, timeout=8)
                    if not confirmed:
                        raise RuntimeError("Send not confirmed for Comcast")
                return True

            # Generic fallback
            else:
                try:
                    ok = self._set_recipient_general(recipient)
                    if not ok:
                        raise RuntimeError("Failed to set recipient (generic)")
                except Exception:
                    pass
                try:
                    body_selectors = ["//div[@role='textbox']", "//div[@contenteditable='true']", "//body"]
                    fragment = self._extract_body_fragment(html_body)
                    fragment = self._inject_placeholders(fragment)
                    fragment = self._ensure_review_fallback(fragment)
                    ok_body = self._set_body_html_generic(wait, fragment, selectors=body_selectors, iframe_ok=True, plain_fallback=True)
                    if not ok_body:
                        raise RuntimeError("Failed to set body (generic)")
                except Exception:
                    pass

                # Try attach
                try:
                    self._attach_file_to_compose(driver)
                except Exception:
                    pass

                try:
                    body_el = driver.find_element(By.XPATH, "//div[@role='textbox'] | //div[@contenteditable='true'] | //body")
                    body_el.send_keys(Keys.CONTROL, Keys.ENTER)
                except Exception:
                    pass
                time.sleep(1.0)
                return True
        except Exception as e:
            self.log(f"send_email error for {recipient} on {provider}: {e}\n{traceback.format_exc()}")
            raise

# ---------------------------
# Run
# ---------------------------
if __name__ == "__main__":
    root = tk.Tk()
    app = BulkEmailSender(root)

    root.mainloop()