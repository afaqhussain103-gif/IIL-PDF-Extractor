"""
IIL Invoice Extractor v1.0
Extracts invoices by customer name or account ID
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import re
import threading

try:
    import fitz  # PyMuPDF
except ImportError:
    import subprocess
    import sys
    subprocess.check_call([sys.executable, "-m", "pip", "install", "PyMuPDF"])
    import fitz

class InvoiceExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("IIL Invoice Extractor v1.0")
        self.root.geometry("700x550")
        self.root.resizable(False, False)
        
        self.source_folder = tk.StringVar()
        self.dest_folder = tk.StringVar()
        self.customer_filter = tk.StringVar()
        self.filter_type = tk.StringVar(value="name")
        
        self.create_widgets()
        
    def create_widgets(self):
        header = tk.Label(
            self.root, 
            text="IIL Invoice Extractor", 
            font=("Arial", 16, "bold"),
            bg="#2c3e50",
            fg="white",
            pady=10
        )
        header.pack(fill=tk.X)
        
        main_frame = tk.Frame(self.root, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(main_frame, text="📁 Invoice Folder (Source):", font=("Arial", 10, "bold")).grid(
            row=0, column=0, sticky=tk.W, pady=(0, 5)
        )
        
        source_frame = tk.Frame(main_frame)
        source_frame.grid(row=1, column=0, columnspan=3, sticky=tk.EW, pady=(0, 15))
        
        tk.Entry(source_frame, textvariable=self.source_folder, width=60, state="readonly").pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10)
        )
        tk.Button(source_frame, text="Browse", command=self.browse_source, width=10).pack(side=tk.LEFT)
        
        tk.Label(main_frame, text="📂 Delivery Folder (Destination):", font=("Arial", 10, "bold")).grid(
            row=2, column=0, sticky=tk.W, pady=(0, 5)
        )
        
        dest_frame = tk.Frame(main_frame)
        dest_frame.grid(row=3, column=0, columnspan=3, sticky=tk.EW, pady=(0, 15))
        
        tk.Entry(dest_frame, textvariable=self.dest_folder, width=60, state="readonly").pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10)
        )
        tk.Button(dest_frame, text="Browse", command=self.browse_dest, width=10).pack(side=tk.LEFT)
        
        tk.Label(main_frame, text="🔍 Customer Filter:", font=("Arial", 10, "bold")).grid(
            row=4, column=0, sticky=tk.W, pady=(0, 5)
        )
        
        tk.Entry(main_frame, textvariable=self.customer_filter, width=60).grid(
            row=5, column=0, columnspan=3, sticky=tk.EW, pady=(0, 15)
        )
        
        filter_frame = tk.Frame(main_frame)
        filter_frame.grid(row=6, column=0, columnspan=3, sticky=tk.W, pady=(0, 20))
        
        tk.Label(filter_frame, text="Filter By:", font=("Arial", 10)).pack(side=tk.LEFT, padx=(0, 10))
        tk.Radiobutton(filter_frame, text="Customer Name", variable=self.filter_type, value="name").pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(filter_frame, text="Account ID", variable=self.filter_type, value="id").pack(side=tk.LEFT, padx=5)
        
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=7, column=0, columnspan=3, sticky=tk.EW, pady=(0, 10))
        
        self.status_label = tk.Label(main_frame, text="Ready to extract invoices", font=("Arial", 9), fg="gray")
        self.status_label.grid(row=8, column=0, columnspan=3, sticky=tk.W, pady=(0, 15))
        
        self.extract_btn = tk.Button(
            main_frame,
            text="🚀 Extract Invoices",
            command=self.start_extraction,
            font=("Arial", 12, "bold"),
            bg="#27ae60",
            fg="white",
            height=2,
            cursor="hand2"
        )
        self.extract_btn.grid(row=9, column=0, columnspan=3, sticky=tk.EW)
        
        main_frame.columnconfigure(0, weight=1)
        
    def browse_source(self):
        folder = filedialog.askdirectory(title="Select Invoice Folder")
        if folder:
            self.source_folder.set(folder)
            
    def browse_dest(self):
        folder = filedialog.askdirectory(title="Select Delivery Folder")
        if folder:
            self.dest_folder.set(folder)
            
    def start_extraction(self):
        if not self.source_folder.get():
            messagebox.showerror("Error", "Please select an invoice folder!")
            return
        if not self.dest_folder.get():
            messagebox.showerror("Error", "Please select a delivery folder!")
            return
        if not self.customer_filter.get():
            messagebox.showerror("Error", "Please enter a customer name or account ID!")
            return
            
        self.extract_btn.config(state=tk.DISABLED)
        self.progress.start()
        self.status_label.config(text="Processing invoices...", fg="blue")
        
        thread = threading.Thread(target=self.extract_invoices)
        thread.daemon = True
        thread.start()
        
    def extract_invoices(self):
        try:
            source = self.source_folder.get()
            dest = self.dest_folder.get()
            customer = self.customer_filter.get().strip()
            filter_by = self.filter_type.get()
            
            pdf_files = [f for f in os.listdir(source) if f.lower().endswith('.pdf')]
            
            if not pdf_files:
                self.show_result("No PDF files found in source folder!", "error")
                return
            
            matched_invoices = []
            total_invoices = 0
            customer_upper = customer.upper()
            
            for pdf_file in pdf_files:
                pdf_path = os.path.join(source, pdf_file)
                
                try:
                    doc = fitz.open(pdf_path)
                    current_invoice_pages = []
                    current_invoice_data = None
                    
                    for page_num in range(len(doc)):
                        text = doc[page_num].get_text()
                        
                        if "Invoice Details" in text and "Number" in text:
                            total_invoices += 1
                            
                            if current_invoice_pages and current_invoice_data:
                                if self.is_match(current_invoice_data, customer_upper, filter_by):
                                    matched_invoices.append({
                                        'data': current_invoice_data,
                                        'pages': current_invoice_pages.copy(),
                                        'source_doc': pdf_path
                                    })
                            
                            current_invoice_pages = [page_num]
                            current_invoice_data = self.extract_metadata(text)
                        
                        elif current_invoice_pages:
                            current_invoice_pages.append(page_num)
                    
                    if current_invoice_pages and current_invoice_data:
                        if self.is_match(current_invoice_data, customer_upper, filter_by):
                            matched_invoices.append({
                                'data': current_invoice_data,
                                'pages': current_invoice_pages.copy(),
                                'source_doc': pdf_path
                            })
                    
                    doc.close()
                    
                except Exception as e:
                    print(f"Error processing {pdf_file}: {str(e)}")
            
            created_files = []
            
            for invoice_info in matched_invoices:
                metadata = invoice_info['data']
                pages = invoice_info['pages']
                source_doc = invoice_info['source_doc']
                
                clean_customer = re.sub(r'[<>:"/\\|?*]', '', metadata['customer_name'])[:50]
                filename = f"INV-{metadata['invoice_num']}-{metadata['date']}-{clean_customer}.pdf"
                output_path = os.path.join(dest, filename)
                
                source = fitz.open(source_doc)
                output = fitz.open()
                
                for page_num in pages:
                    output.insert_pdf(source, from_page=page_num, to_page=page_num)
                
                output.save(output_path)
                output.close()
                source.close()
                
                created_files.append(filename)
            
            self.show_summary(len(pdf_files), total_invoices, len(matched_invoices), created_files, dest)
            
        except Exception as e:
            self.show_result(f"Error: {str(e)}", "error")
    
    def extract_metadata(self, text):
        invoice_num_match = re.search(r'Number\s+([A-Z]{2,5}\d{2}-\d{7})', text)
        if not invoice_num_match:
            invoice_num_match = re.search(r'Number\s+([A-Z0-9\-]+)', text)
        invoice_num = invoice_num_match.group(1) if invoice_num_match else "UNKNOWN"
        
        date_match = re.search(r'Date\s+(\d{2}-[A-Z]{3}-\d{4})', text)
        invoice_date = date_match.group(1) if date_match else "NODATE"
        
        customer_name_match = re.search(r'Name\s+(.+?)(?:\n|Address)', text, re.DOTALL)
        if customer_name_match:
            customer_name = customer_name_match.group(1).strip()
            customer_name = ' '.join(customer_name.split())
        else:
            customer_name = "UNKNOWN_CUSTOMER"
        
        account_match = re.search(r'Account\s+(\d+)', text)
        account_id = account_match.group(1) if account_match else "NO_ID"
        
        return {
            'invoice_num': invoice_num,
            'date': invoice_date,
            'customer_name': customer_name,
            'account_id': account_id
        }
    
    def is_match(self, invoice_data, customer_filter, filter_by):
        if filter_by == "name":
            return customer_filter in invoice_data['customer_name'].upper()
        else:
            return customer_filter == invoice_data['account_id']
    
    def show_summary(self, total_files, total_invoices, matched, created_files, dest_folder):
        self.root.after(0, lambda: self._show_summary_window(total_files, total_invoices, matched, created_files, dest_folder))
    
    def _show_summary_window(self, total_files, total_invoices, matched, created_files, dest_folder):
        self.progress.stop()
        self.extract_btn.config(state=tk.NORMAL)
        self.status_label.config(text="Ready to extract invoices", fg="gray")
        
        summary_win = tk.Toplevel(self.root)
        summary_win.title("Extraction Complete")
        summary_win.geometry("600x500")
        summary_win.resizable(False, False)
        
        header = tk.Label(
            summary_win,
            text="✅ EXTRACTION COMPLETE!",
            font=("Arial", 14, "bold"),
            bg="#27ae60",
            fg="white",
            pady=15
        )
        header.pack(fill=tk.X)
        
        summary_frame = tk.Frame(summary_win, padx=20, pady=20)
        summary_frame.pack(fill=tk.BOTH, expand=True)
        
        summary_text = f"""
📊 Summary:
────────────────────────────────
• Processed: {total_files} PDF file(s)
• Total invoices found: {total_invoices}
• Extracted for {self.customer_filter.get()}: {matched}

📄 Created Files:
────────────────────────────────
"""
        
        tk.Label(summary_frame, text=summary_text, font=("Courier", 10), justify=tk.LEFT).pack(anchor=tk.W)
        
        list_frame = tk.Frame(summary_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        files_listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, font=("Courier", 9))
        files_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=files_listbox.yview)
        
        for idx, filename in enumerate(created_files, 1):
            files_listbox.insert(tk.END, f"✓ {filename}")
        
        btn_frame = tk.Frame(summary_frame)
        btn_frame.pack(pady=(20, 0))
        
        tk.Button(
            btn_frame,
            text="📂 Open Folder",
            command=lambda: os.startfile(dest_folder),
            font=("Arial", 11, "bold"),
            bg="#3498db",
            fg="white",
            width=15,
            height=2,
            cursor="hand2"
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            btn_frame,
            text="Close",
            command=summary_win.destroy,
            font=("Arial", 11),
            width=15,
            height=2,
            cursor="hand2"
        ).pack(side=tk.LEFT, padx=5)
        
        if not created_files:
            messagebox.showwarning(
                "No Matches",
                f"No invoices found for: {self.customer_filter.get()}\n\nTry adjusting the customer name or account ID."
            )
    
    def show_result(self, message, msg_type):
        self.root.after(0, lambda: self._show_result(message, msg_type))
    
    def _show_result(self, message, msg_type):
        self.progress.stop()
        self.extract_btn.config(state=tk.NORMAL)
        self.status_label.config(text="Ready to extract invoices", fg="gray")
        
        if msg_type == "error":
            messagebox.showerror("Error", message)
        else:
            messagebox.showinfo("Success", message)

def main():
    root = tk.Tk()
    app = InvoiceExtractorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
