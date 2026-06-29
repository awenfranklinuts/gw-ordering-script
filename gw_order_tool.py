import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import openpyxl
import subprocess
import sys
import os
import threading


class GWOrderTool:
    def __init__(self, root):
        self.root = root
        self.root.title("GW Order Tool")
        self.root.geometry("1200x700")
        self.root.minsize(900, 500)

        self.order_pad_path = tk.StringVar()
        self.order_conf_path = tk.StringVar()
        self.order_pad_data = []
        self.conf_qty_lookup = {}

        self._build_file_selection_frame()
        self.table_container = None
        self.tree = None
        self.fetch_btn = None
        self.status_var = tk.StringVar()
        self.script_dir = os.path.dirname(os.path.abspath(__file__))

    def _build_file_selection_frame(self):
        frame = ttk.LabelFrame(self.root, text="Select Files", padding=10)
        frame.pack(fill="x", padx=10, pady=(10, 5))

        ttk.Label(frame, text="Last Week's Order Pad (.xlsx):").grid(row=0, column=0, sticky="w", pady=2)
        ttk.Entry(frame, textvariable=self.order_pad_path, width=80).grid(row=0, column=1, padx=5, pady=2)
        ttk.Button(frame, text="Browse...", command=self._browse_order_pad).grid(row=0, column=2, pady=2)

        ttk.Label(frame, text="Last Week's Order Confirmation (.xlsx):").grid(row=1, column=0, sticky="w", pady=2)
        ttk.Entry(frame, textvariable=self.order_conf_path, width=80).grid(row=1, column=1, padx=5, pady=2)
        ttk.Button(frame, text="Browse...", command=self._browse_order_conf).grid(row=1, column=2, pady=2)

        ttk.Button(frame, text="Reconcile Orders", command=self._on_compare).grid(row=2, column=1, pady=(10, 0))

        frame.columnconfigure(1, weight=1)

    def _create_treeview(self, parent):
        container = ttk.Frame(parent)
        container.pack(fill="both", expand=True)

        tree = ttk.Treeview(container, show="headings")

        vsb = ttk.Scrollbar(container, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(container, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        container.rowconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)

        return tree

    def _browse_order_pad(self):
        path = filedialog.askopenfilename(
            title="Select Order Pad",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if path:
            self.order_pad_path.set(path)

    def _browse_order_conf(self):
        path = filedialog.askopenfilename(
            title="Select Order Confirmation",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if path:
            self.order_conf_path.set(path)

    def _on_compare(self):
        pad_path = self.order_pad_path.get().strip()
        conf_path = self.order_conf_path.get().strip()

        if not pad_path or not conf_path:
            messagebox.showwarning("Missing Files", "Please select both files before comparing.")
            return

        self._load_order_pad(pad_path)
        self._load_order_conf(conf_path)
        self._refresh_view()

    def _load_order_pad(self, path):
        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            ws = wb[wb.sheetnames[0]]

            rows = list(ws.iter_rows(values_only=True))
            wb.close()

            headers = ["Product Code", "Qty Ordered"]
            data_rows = []
            for row in rows[4:]:
                product_code = row[5]
                if product_code is None or str(product_code).strip() == "":
                    continue
                order_qty = row[6]
                try:
                    qty = int(order_qty) if order_qty is not None and str(order_qty).strip() != "" else 0
                except (ValueError, TypeError):
                    qty = 0
                data_rows.append((product_code, qty))

            self.order_pad_data = data_rows

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load Order Pad:\n{e}")

    def _load_order_conf(self, path):
        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            ws = wb["Table 1"]

            rows = list(ws.iter_rows(values_only=True))
            wb.close()

            self.conf_qty_lookup = {}
            for row in rows[1:]:
                product_code = row[0]
                if product_code is None:
                    continue
                qty_available = row[4]
                try:
                    qty = int(qty_available) if qty_available is not None and str(qty_available).strip() != "" else 0
                except (ValueError, TypeError):
                    qty = 0
                self.conf_qty_lookup[str(product_code)] = qty

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load Order Confirmation:\n{e}")

    def _refresh_view(self):
        if self.table_container:
            self.table_container.destroy()
        if self.fetch_btn:
            self.fetch_btn.destroy()

        self.table_container = ttk.Frame(self.root)
        self.table_container.pack(fill="both", expand=True, padx=10, pady=(5, 5))
        self.tree = self._create_treeview(self.table_container)

        headers = ["Product Code", "Qty Ordered", "Qty Confirmed", "Qty Outstanding"]
        merged = []
        for product_code, qty_ordered in self.order_pad_data:
            qty_confirmed = self.conf_qty_lookup.get(str(product_code), 0)
            qty_outstanding = qty_ordered - qty_confirmed
            merged.append((product_code, qty_ordered, qty_confirmed, qty_outstanding))
        self._populate_tree(self.tree, headers, merged)

        bottom_frame = ttk.Frame(self.root)
        bottom_frame.pack(fill="x", padx=10, pady=(0, 10))
        self.fetch_btn = bottom_frame

        self.neto_btn = ttk.Button(bottom_frame, text="Fetch Stock from Neto", command=self._on_fetch_neto)
        self.neto_btn.pack(side="left", padx=(0, 10))

        ttk.Label(bottom_frame, textvariable=self.status_var).pack(side="left")

    def _on_fetch_neto(self):
        self.neto_btn.config(state="disabled")
        self.status_var.set("Launching Neto scraper...")

        thread = threading.Thread(target=self._run_neto_scraper, daemon=True)
        thread.start()

    def _run_neto_scraper(self):
        script_path = os.path.join(self.script_dir, "neto_scraper.py")
        try:
            process = subprocess.Popen(
                [sys.executable, script_path],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
            )

            self.root.after(0, lambda: self.status_var.set("Neto scraper is running..."))

            output = []
            for line in process.stdout:
                output.append(line.strip())
                display = line.strip()
                self.root.after(0, lambda d=display: self.status_var.set(d))

            process.wait()

            if process.returncode == 0:
                self.root.after(0, lambda: self.status_var.set("Neto scraper finished."))
            else:
                self.root.after(0, lambda: self.status_var.set(f"Scraper exited with error (code {process.returncode})"))

        except Exception as e:
            self.root.after(0, lambda: self.status_var.set(f"Error: {e}"))

        self.root.after(0, lambda: self.neto_btn.config(state="normal"))

    def _populate_tree(self, tree, headers, data_rows):
        tree.delete(*tree.get_children())

        tree["columns"] = headers
        for h in headers:
            tree.heading(h, text=h)
            tree.column(h, width=120, minwidth=60)

        for row in data_rows:
            values = [str(v) if v is not None else "" for v in row]
            tree.insert("", "end", values=values)


def main():
    root = tk.Tk()
    GWOrderTool(root)
    root.mainloop()


if __name__ == "__main__":
    main()
