"""IIL Invoice Extractor v2.3 - Simplified & Stable"""

import os
import shutil
import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime
import re

class IILInvoiceExtractor:
    def __init__(self, root):
        self.root = root
        self.root.title("IIL Invoice Extractor v2.3")
        self.root.geometry("820x750")
        self.root.resizable(False, False)
        
        # OFFICIAL IIL BRAND COLORS
        self.colors = {
            'green': '#01783f',        # IIL Primary Green
            'yellow': '#c2d501',       # IIL Primary Yellow-Green
            'accent': '#76690f',       # IIL Accent Olive
            'white': '#ffffff',        # Primary White
            'bg': '#f8f9fa',           # Light Background
            'text': '#2c3e50',         # Dark Text
            'gray': '#6c757d'          # Secondary Gray
        }
        
        self.root.configure(bg=self.colors['white'])
        self.setup_ui()
    
    def setup_ui(self):
        # ═══════════════════════════════════════════════════════════
        # HEADER - IIL GREEN WITH WHITE TEXT
        # ═══════════════════════════════════════════════════════════
        header = tk.Frame(self.root, bg=self.colors['green'], height=110)
        header.pack(fill=tk.X)
        header.pack_propagate(False)
        
        # Company Name
        tk.Label(
            header,
            text="INTERNATIONAL INDUSTRIES LIMITED",
            font=('Arial', 20, 'bold'),
            bg=self.colors['green'],
            fg=self.colors['white']
        ).pack(pady=(22, 2))
        
        # Tagline with Yellow-Green
        tk.Label(
            header,
            text="Building Pakistan's Future | Invoice Management System",
            font=('Arial', 10),
            bg=self.colors['green'],
            fg=self.colors['yellow']
        ).pack(pady=2)
        
        # Version Badge
        version_frame = tk.Frame(header, bg=self.colors['yellow'], padx=12, pady=3)
        version_frame.place(relx=0.95, rely=0.15, anchor='ne')
        tk.Label(
            version_frame,
            text="v2.3",
            font=('Arial', 8, 'bold'),
            bg=self.colors['yellow'],
            fg=self.colors['text']
        ).pack()
        
        # ═══════════════════════════════════════════════════════════
        # MAIN CONTENT AREA
        # ═══════════════════════════════════════════════════════════
        main = tk.Frame(self.root, bg=self.colors['white'])
        main.pack(fill=tk.BOTH, expand=True, padx=40, pady=25)
        
        # ───────────────────────────────────────────────────────────
        # SOURCE FOLDER
        # ───────────────────────────────────────────────────────────
        self.create_section_header(main, "📁 SOURCE FOLDER (Invoices)", 0)
        self.source_entry = self.create_folder_input(main, 1, self.browse_source)
        
        # ───────────────────────────────────────────────────────────
        # DESTINATION FOLDER
        # ───────────────────────────────────────────────────────────
        self.create_section_header(main, "📂 DESTINATION FOLDER (Output)", 2)
        self.dest_entry = self.create_folder_input(main, 3, self.browse_destination)
        
        # Separator
        self.create_separator(main, 4)
        
        # ───────────────────────────────────────────────────────────
        # FILTER TYPE
        # ───────────────────────────────────────────────────────────
        self.create_section_header(main, "🔍 FILTER TYPE", 5)
        
        filter_frame = tk.Frame(main, bg=self.colors['white'])
        filter_frame.grid(row=6, column=0, columnspan=3, sticky='w', pady=5)
        
        self.filter_var = tk.StringVar(value="name")
        
        tk.Radiobutton(
            filter_frame,
            text="Customer Name",
            variable=self.filter_var,
            value="name",
            font=('Arial', 10),
            bg=self.colors['white'],
            fg=self.colors['text'],
            selectcolor=self.colors['yellow'],
            activebackground=self.colors['white']
        ).pack(side=tk.LEFT, padx=(0, 30))
        
        tk.Radiobutton(
            filter_frame,
            text="Account ID",
            variable=self.filter_var,
            value="account",
            font=('Arial', 10),
            bg=self.colors['white'],
            fg=self.colors['text'],
            selectcolor=self.colors['yellow'],
            activebackground=self.colors['white']
        ).pack(side=tk.LEFT)
        
        # ───────────────────────────────────────────────────────────
        # SEARCH VALUE
        # ───────────────────────────────────────────────────────────
        self.create_section_header(main, "🔎 SEARCH VALUE", 7)
        
        self.search_entry = tk.Entry(
            main,
            font=('Arial', 11),
            bd=0,
            relief=tk.FLAT,
            highlightthickness=2,
            highlightbackground=self.colors['bg'],
            highlightcolor=self.colors['green']
        )
        self.search_entry.grid(row=8, column=0, columnspan=3, sticky='ew', pady=5, ipady=10)
        
        # Separator
        self.create_separator(main, 9)
        
        # ───────────────────────────────────────────────────────────
        # DATE RANGE - SIMPLIFIED (No tkcalendar dependency)
        # ───────────────────────────────────────────────────────────
        self.create_section_header(main, "📅 DATE RANGE FILTER (Optional - DD/MM/YYYY)", 10)
        
        date_frame = tk.Frame(main, bg=self.colors['white'])
        date_frame.grid(row=11, column=0, columnspan=3, sticky='w', pady=8)
        
        # From Date
        tk.Label(
            date_frame,
            text="From:",
            font=('Arial', 10, 'bold'),
            bg=self.colors['white'],
            fg=self.colors['text']
        ).pack(side=tk.LEFT, padx=(0, 8))
        
        self.from_date_entry = tk.Entry(
            date_frame,
            font=('Arial', 10),
            width=15,
            bd=1,
            relief=tk.SOLID,
            highlightthickness=1,
            highlightbackground=self.colors['green']
        )
        self.from_date_entry.pack(side=tk.LEFT, padx=(0, 35))
        self.from_date_entry.insert(0, datetime.now().strftime("%d/%m/%Y"))
        
        # To Date
        tk.Label(
            date_frame,
            text="To:",
            font=('Arial', 10, 'bold'),
            bg=self.colors['white'],
            fg=self.colors['text']
        ).pack(side=tk.LEFT, padx=(0, 8))
        
        self.to_date_entry = tk.Entry(
            date_frame,
            font=('Arial', 10),
            width=15,
            bd=1,
            relief=tk.SOLID,
            highlightthickness=1,
            highlightbackground=self.colors['green']
        )
        self.to_date_entry.pack(side=tk.LEFT)
        self.to_date_entry.insert(0, datetime.now().strftime("%d/%m/%Y"))
        
        # Date format hint
        hint_label = tk.Label(
            main,
            text="💡 Format: DD/MM/YYYY (e.g., 31/01/2025)",
            font=('Arial', 8, 'italic'),
            bg=self.colors['white'],
            fg=self.colors['gray']
        )
        hint_label.grid(row=12, column=0, columnspan=3, sticky='w', pady=(0, 5))
        
        # ───────────────────────────────────────────────────────────
        # OPTIONS CHECKBOXES
        # ───────────────────────────────────────────────────────────
        options_frame = tk.Frame(main, bg=self.colors['white'])
        options_frame.grid(row=13, column=0, columnspan=3, sticky='w', pady=10)
        
        # Date Filter Toggle
        self.use_date_filter = tk.BooleanVar(value=False)
        tk.Checkbutton(
            options_frame,
            text="✓ Enable Date Range Filter",
            variable=self.use_date_filter,
            font=('Arial', 9, 'bold'),
            bg=self.colors['white'],
            fg=self.colors['green'],
            selectcolor=self.colors['yellow'],
            activebackground=self.colors['white']
        ).pack(anchor='w', pady=2)
        
        # Overwrite Toggle
        self.overwrite_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            options_frame,
            text="⚠️  Overwrite existing files (uncheck to skip duplicates)",
            variable=self.overwrite_var,
            font=('Arial', 9),
            bg=self.colors['white'],
            fg=self.colors['accent'],
            selectcolor=self.colors['yellow'],
            activebackground=self.colors['white']
        ).pack(anchor='w', pady=2)
        
        # Separator
        self.create_separator(main, 14)
        
        # ───────────────────────────────────────────────────────────
        # PROGRESS SECTION
        # ───────────────────────────────────────────────────────────
        progress_container = tk.Frame(main, bg=self.colors['bg'], bd=0)
        progress_container.grid(row=15, column=0, columnspan=3, sticky='ew', pady=15, ipady=15, ipadx=20)
        
        self.progress_label = tk.Label(
            progress_container,
            text="Ready to extract invoices",
            font=('Arial', 11, 'bold'),
            bg=self.colors['bg'],
            fg=self.colors['text']
        )
        self.progress_label.pack(pady=(5, 8))
        
        # Progress Bar
        style = ttk.Style()
        style.theme_use('clam')
        style.configure(
            "
