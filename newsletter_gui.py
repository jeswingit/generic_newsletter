#!/usr/bin/env python3
"""
newsletter_gui.py

Graphical user interface for the newsletter generator created in Claude 2.
Provides a simple GUI to upload Excel files and configure newsletter parameters.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext, colorchooser
from pathlib import Path
import sys
from datetime import datetime

# Import functions from generate_newsletter.py
from generate_newsletter import (
    read_excel_rows,
    build_html_email,
    build_eml_message,
    _load_image_part,
    EMAIL_CONFIG,
    DEFAULT_XLSX,
    DEFAULT_OUT,
    DEFAULT_MONTH,
)


class NewsletterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("ADIA EMEA Newsletter Generator")
        self.root.geometry("700x650")
        self.root.resizable(True, True)
        
        # Variables
        self.xlsx_path = tk.StringVar(value="")
        self.output_path = tk.StringVar(value=DEFAULT_OUT)
        self.month = tk.StringVar(value=DEFAULT_MONTH)
        self.from_email = tk.StringVar(value=EMAIL_CONFIG["from"])
        self.to_email = tk.StringVar(value=EMAIL_CONFIG["to"])
        self.subject = tk.StringVar(value="")

        # Layout customization
        self.available_blocks = [
            "Month News",
            "Save the Date",
            "General Information",
            "General",
        ]
        self.enabled_blocks: list[str] = self.available_blocks[:]
        self.block_bg_colors: dict[str, str] = {
            "Month News": EMAIL_CONFIG["colors"].get("white", "#ffffff"),
            "Save the Date": EMAIL_CONFIG["colors"].get("save_date_bg", "#E5EFF0"),
            "General Information": EMAIL_CONFIG["colors"].get("white", "#ffffff"),
            "General": EMAIL_CONFIG["colors"].get("white", "#ffffff"),
        }
        
        self.setup_ui()
        
    def setup_ui(self):
        # Main container with padding
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        row = 0
        
        # Title
        title_label = ttk.Label(
            main_frame,
            text="ADIA EMEA Newsletter Generator",
            font=("Arial", 16, "bold")
        )
        title_label.grid(row=row, column=0, columnspan=3, pady=(0, 20))
        row += 1
        
        # Excel File Selection
        ttk.Label(main_frame, text="Excel File:", font=("Arial", 10, "bold")).grid(
            row=row, column=0, sticky=tk.W, pady=5
        )
        ttk.Entry(main_frame, textvariable=self.xlsx_path, width=50).grid(
            row=row, column=1, sticky=(tk.W, tk.E), padx=5, pady=5
        )
        ttk.Button(main_frame, text="Browse...", command=self.browse_xlsx).grid(
            row=row, column=2, padx=5, pady=5
        )
        row += 1
        
        # Output File Selection
        ttk.Label(main_frame, text="Output File:", font=("Arial", 10, "bold")).grid(
            row=row, column=0, sticky=tk.W, pady=5
        )
        ttk.Entry(main_frame, textvariable=self.output_path, width=50).grid(
            row=row, column=1, sticky=(tk.W, tk.E), padx=5, pady=5
        )
        ttk.Button(main_frame, text="Browse...", command=self.browse_output).grid(
            row=row, column=2, padx=5, pady=5
        )
        row += 1
        
        # Month Selection
        ttk.Label(main_frame, text="Month:", font=("Arial", 10, "bold")).grid(
            row=row, column=0, sticky=tk.W, pady=5
        )
        month_frame = ttk.Frame(main_frame)
        month_frame.grid(row=row, column=1, sticky=tk.W, padx=5, pady=5)
        
        months = [
            "January", "February", "March", "April", "May", "June",
            "July", "August", "September", "October", "November", "December"
        ]
        month_combo = ttk.Combobox(
            month_frame,
            textvariable=self.month,
            values=months,
            width=20,
            state="readonly"
        )
        month_combo.grid(row=0, column=0)
        row += 1
        
        # From Email
        ttk.Label(main_frame, text="From Email:", font=("Arial", 10, "bold")).grid(
            row=row, column=0, sticky=tk.W, pady=5
        )
        ttk.Entry(main_frame, textvariable=self.from_email, width=50).grid(
            row=row, column=1, columnspan=2, sticky=(tk.W, tk.E), padx=5, pady=5
        )
        row += 1
        
        # To Email
        ttk.Label(main_frame, text="To Email:", font=("Arial", 10, "bold")).grid(
            row=row, column=0, sticky=tk.W, pady=5
        )
        ttk.Entry(main_frame, textvariable=self.to_email, width=50).grid(
            row=row, column=1, columnspan=2, sticky=(tk.W, tk.E), padx=5, pady=5
        )
        row += 1
        
        # Subject (optional)
        ttk.Label(main_frame, text="Subject (optional):", font=("Arial", 10, "bold")).grid(
            row=row, column=0, sticky=tk.W, pady=5
        )
        ttk.Entry(main_frame, textvariable=self.subject, width=50).grid(
            row=row, column=1, columnspan=2, sticky=(tk.W, tk.E), padx=5, pady=5
        )
        ttk.Label(
            main_frame,
            text="Leave empty to use default: 'ADIA EMEA - Good to Know | {Month} {Year}'",
            font=("Arial", 8),
            foreground="gray"
        ).grid(row=row+1, column=1, columnspan=2, sticky=tk.W, padx=5)
        row += 2

        # Layout / Color customization
        layout_frame = ttk.LabelFrame(main_frame, text="Layout & Block Colors", padding="10")
        layout_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(5, 10))
        layout_frame.columnconfigure(0, weight=1)
        layout_frame.columnconfigure(1, weight=0)
        layout_frame.columnconfigure(2, weight=1)

        ttk.Label(layout_frame, text="Enabled blocks (order matters):").grid(
            row=0, column=0, sticky=tk.W, pady=(0, 5)
        )

        self.blocks_listbox = tk.Listbox(layout_frame, height=6, exportselection=False)
        self.blocks_listbox.grid(row=1, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        for b in self.enabled_blocks:
            self.blocks_listbox.insert(tk.END, b)

        buttons_col = ttk.Frame(layout_frame)
        buttons_col.grid(row=1, column=1, sticky=(tk.N))

        ttk.Button(buttons_col, text="Up", width=10, command=self.move_block_up).grid(row=0, column=0, pady=2)
        ttk.Button(buttons_col, text="Down", width=10, command=self.move_block_down).grid(row=1, column=0, pady=2)

        add_remove = ttk.Frame(buttons_col)
        add_remove.grid(row=2, column=0, pady=(10, 0))
        self.add_block_choice = tk.StringVar(value=self.available_blocks[0])
        ttk.Combobox(
            add_remove,
            textvariable=self.add_block_choice,
            values=self.available_blocks,
            width=18,
            state="readonly",
        ).grid(row=0, column=0, pady=2)
        ttk.Button(add_remove, text="Add", width=10, command=self.add_block).grid(row=1, column=0, pady=2)
        ttk.Button(add_remove, text="Remove", width=10, command=self.remove_block).grid(row=2, column=0, pady=2)

        colors_col = ttk.Frame(layout_frame)
        colors_col.grid(row=1, column=2, sticky=(tk.W, tk.E))
        ttk.Label(colors_col, text="Background colors (per block type):").grid(row=0, column=0, sticky=tk.W)

        self.color_buttons: dict[str, ttk.Button] = {}
        for i, block_id in enumerate(self.available_blocks, start=1):
            row_frame = ttk.Frame(colors_col)
            row_frame.grid(row=i, column=0, sticky=(tk.W, tk.E), pady=2)
            ttk.Label(row_frame, text=block_id, width=20).grid(row=0, column=0, sticky=tk.W)
            btn = ttk.Button(
                row_frame,
                text=self.block_bg_colors.get(block_id, "#ffffff"),
                width=12,
                command=lambda bid=block_id: self.choose_block_color(bid),
            )
            btn.grid(row=0, column=1, sticky=tk.W, padx=(6, 0))
            self.color_buttons[block_id] = btn

        row += 1
        
        # Status/Log Area
        ttk.Label(main_frame, text="Status:", font=("Arial", 10, "bold")).grid(
            row=row, column=0, sticky=(tk.W, tk.N), pady=(10, 5)
        )
        self.status_text = scrolledtext.ScrolledText(
            main_frame,
            height=12,
            width=70,
            wrap=tk.WORD,
            font=("Consolas", 9)
        )
        self.status_text.grid(
            row=row, column=1, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5
        )
        main_frame.rowconfigure(row, weight=1)
        row += 1
        
        # Generate Button
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=row, column=0, columnspan=3, pady=20)
        
        self.generate_button = ttk.Button(
            button_frame,
            text="Generate Newsletter",
            command=self.generate_newsletter,
            style="Accent.TButton"
        )
        self.generate_button.pack(side=tk.LEFT, padx=10)
        
        ttk.Button(
            button_frame,
            text="Clear Log",
            command=self.clear_log
        ).pack(side=tk.LEFT, padx=10)
        
        # Initial status message
        self.log("Welcome to ADIA EMEA Newsletter Generator")
        self.log("Please select an Excel file and configure the parameters.")

    def _sync_enabled_blocks_from_listbox(self):
        self.enabled_blocks = list(self.blocks_listbox.get(0, tk.END))

    def move_block_up(self):
        sel = self.blocks_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        if idx <= 0:
            return
        item = self.blocks_listbox.get(idx)
        self.blocks_listbox.delete(idx)
        self.blocks_listbox.insert(idx - 1, item)
        self.blocks_listbox.selection_set(idx - 1)
        self._sync_enabled_blocks_from_listbox()

    def move_block_down(self):
        sel = self.blocks_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        if idx >= self.blocks_listbox.size() - 1:
            return
        item = self.blocks_listbox.get(idx)
        self.blocks_listbox.delete(idx)
        self.blocks_listbox.insert(idx + 1, item)
        self.blocks_listbox.selection_set(idx + 1)
        self._sync_enabled_blocks_from_listbox()

    def add_block(self):
        block_id = self.add_block_choice.get()
        existing = set(self.blocks_listbox.get(0, tk.END))
        if block_id in existing:
            return
        self.blocks_listbox.insert(tk.END, block_id)
        self._sync_enabled_blocks_from_listbox()

    def remove_block(self):
        sel = self.blocks_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        self.blocks_listbox.delete(idx)
        self._sync_enabled_blocks_from_listbox()

    def choose_block_color(self, block_id: str):
        initial = self.block_bg_colors.get(block_id, "#ffffff")
        color = colorchooser.askcolor(color=initial, title=f"Choose background color for {block_id}")
        if not color or not color[1]:
            return
        self.block_bg_colors[block_id] = color[1]
        btn = self.color_buttons.get(block_id)
        if btn is not None:
            btn.config(text=color[1])
        
    def browse_xlsx(self):
        """Open file dialog to select Excel file."""
        filename = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.xlsx_path.set(filename)
            self.log(f"Selected Excel file: {Path(filename).name}")
    
    def browse_output(self):
        """Open file dialog to select output file."""
        filename = filedialog.asksaveasfilename(
            title="Save Newsletter As",
            defaultextension=".eml",
            filetypes=[("EML files", "*.eml"), ("All files", "*.*")]
        )
        if filename:
            self.output_path.set(filename)
            self.log(f"Output file: {Path(filename).name}")
    
    def log(self, message):
        """Add a message to the status log."""
        self.status_text.insert(tk.END, f"{message}\n")
        self.status_text.see(tk.END)
        self.root.update_idletasks()
    
    def clear_log(self):
        """Clear the status log."""
        self.status_text.delete(1.0, tk.END)
        self.log("Log cleared.")
    
    def generate_newsletter(self):
        """Generate the newsletter based on the current settings."""
        # Validate inputs
        if not self.xlsx_path.get():
            messagebox.showerror("Error", "Please select an Excel file.")
            return
        
        xlsx_path = Path(self.xlsx_path.get())
        if not xlsx_path.exists():
            messagebox.showerror("Error", f"Excel file not found: {xlsx_path}")
            return
        
        if not self.month.get():
            messagebox.showerror("Error", "Please select a month.")
            return
        
        # Determine output path
        script_dir = Path(__file__).parent
        out_path_str = self.output_path.get()
        out_path = Path(out_path_str) if Path(out_path_str).is_absolute() else script_dir / out_path_str
        
        # Disable generate button during processing
        self.generate_button.config(state="disabled")
        self.log("\n" + "="*60)
        self.log("Starting newsletter generation...")
        self.log("="*60)
        
        try:
            # 1. Read Excel data
            self.log(f"\nReading Excel file: {xlsx_path.name}")
            grouped = read_excel_rows(xlsx_path)
            self.log(f"Loaded data from {xlsx_path.name}:")
            for type_key, rows in grouped.items():
                self.log(f"  {type_key}: {len(rows)} row(s)")
            
            # 2. Prepare image attachments
            self.log("\nPreparing image attachments...")
            image_cids: dict[str, str] = {}
            image_parts: dict[str, object] = {}
            
            product_rows = grouped.get("Product", [])
            for row in product_rows:
                if row.get("image") and row["image"] not in image_cids:
                    img_part, cid = _load_image_part(row["image"], xlsx_path.parent)
                    if img_part is not None:
                        image_cids[row["image"]] = cid
                        image_parts[row["image"]] = img_part
                        self.log(f"    Prepared image: {row['image']} (CID: {cid})")
            
            # 3. Build HTML content
            self.log("\nBuilding HTML email structure...")
            layout = ["Header"] + (self.enabled_blocks[:] if self.enabled_blocks else []) + ["Footer"]
            # Only pass colors for enabled blocks
            block_bg_colors = {k: v for k, v in self.block_bg_colors.items() if k in self.enabled_blocks}
            html = build_html_email(
                grouped,
                self.month.get(),
                EMAIL_CONFIG,
                image_cids,
                layout=layout,
                block_bg_colors=block_bg_colors,
            )
            self.log(f"Built HTML email structure with {len(grouped)} section type(s).")
            
            # 4. Build EML message
            self.log("\nBuilding EML message...")
            if self.subject.get().strip():
                subject = self.subject.get().strip()
            else:
                current_year = datetime.now().year
                subject = EMAIL_CONFIG["subject"].format(
                    month=self.month.get(), year=current_year
                )
            
            self.log(f"Subject: {subject}")
            msg = build_eml_message(
                html, self.from_email.get(), self.to_email.get(), subject
            )
            
            # Attach images
            for image_path, img_part in image_parts.items():
                msg.attach(img_part)
                self.log(f"  Attached image: {image_path}")
            
            # 5. Write output EML
            self.log(f"\nWriting output file: {out_path}")
            with open(out_path, "wb") as fh:
                fh.write(msg.as_bytes())
            
            self.log("\n" + "="*60)
            self.log("✓ Newsletter generated successfully!")
            self.log(f"✓ Output written to: {out_path}")
            self.log("="*60)
            
            messagebox.showinfo(
                "Success",
                f"Newsletter generated successfully!\n\nOutput: {out_path}"
            )
            
        except Exception as e:
            error_msg = f"Error generating newsletter: {str(e)}"
            self.log("\n" + "="*60)
            self.log("✗ ERROR")
            self.log(error_msg)
            self.log("="*60)
            messagebox.showerror("Error", error_msg)
            import traceback
            self.log("\nTraceback:")
            self.log(traceback.format_exc())
        
        finally:
            # Re-enable generate button
            self.generate_button.config(state="normal")


def main():
    root = tk.Tk()
    app = NewsletterGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
