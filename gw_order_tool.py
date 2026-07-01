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
        self.this_week_pad_path = tk.StringVar()
        self.order_pad_data = []
        self.conf_qty_lookup = {}
        self.outstanding_lookup = {}
        self.last_week_ordered_lookup = {}
        self.this_week_pad_data = []
        self.showing_outstanding_view = False
        self.unmatched_tree = None
        self.unmatched_row_data = {}
        self.main_row_iid = {}

        self._build_file_selection_frame()
        self.table_container = None
        self.tree = None
        self.bottom_frame = None
        self.status_var = tk.StringVar()
        self.script_dir = os.path.dirname(os.path.abspath(__file__))

    def _build_file_selection_frame(self):
        frame = ttk.LabelFrame(self.root, text="Select Files", padding=10)
        frame.pack(fill="x", padx=10, pady=(10, 5))

        self.order_pad_label = ttk.Label(frame, text="Last Week's Order Pad (.xlsx):")
        self.order_pad_entry = ttk.Entry(frame, textvariable=self.order_pad_path, width=80)
        self.order_pad_browse_btn = ttk.Button(frame, text="Browse...", command=self._browse_order_pad)
        self.order_pad_label.grid(row=0, column=0, sticky="w", pady=2)
        self.order_pad_entry.grid(row=0, column=1, padx=5, pady=2)
        self.order_pad_browse_btn.grid(row=0, column=2, pady=2)

        self.order_conf_label = ttk.Label(frame, text="Last Week's Order Confirmation (.xlsx):")
        self.order_conf_entry = ttk.Entry(frame, textvariable=self.order_conf_path, width=80)
        self.order_conf_browse_btn = ttk.Button(frame, text="Browse...", command=self._browse_order_conf)
        self.order_conf_label.grid(row=1, column=0, sticky="w", pady=2)
        self.order_conf_entry.grid(row=1, column=1, padx=5, pady=2)
        self.order_conf_browse_btn.grid(row=1, column=2, pady=2)

        self.reconcile_btn = ttk.Button(frame, text="Reconcile Orders", command=self._on_compare)
        self.reconcile_btn.grid(row=2, column=1, pady=(10, 0))

        # Rows hidden once Compare Orders has successfully run — see _hide_last_week_rows().
        self.last_week_row_widgets = [
            self.order_pad_label, self.order_pad_entry, self.order_pad_browse_btn,
            self.order_conf_label, self.order_conf_entry, self.order_conf_browse_btn,
            self.reconcile_btn,
        ]

        self.this_week_label = ttk.Label(frame, text="This Week's Order Pad (.xlsx):")
        self.this_week_entry = ttk.Entry(frame, textvariable=self.this_week_pad_path, width=80)
        self.this_week_browse_btn = ttk.Button(frame, text="Browse...", command=self._browse_this_week_pad)
        self.compare_btn = ttk.Button(frame, text="Compare Orders", command=self._on_compare_this_week)
        # Hidden until Reconcile Orders has run; revealed by _reveal_this_week_row().
        self.this_week_row_widgets = [
            (self.this_week_label, dict(row=3, column=0, sticky="w", pady=2)),
            (self.this_week_entry, dict(row=3, column=1, padx=5, pady=2)),
            (self.this_week_browse_btn, dict(row=3, column=2, pady=2)),
            (self.compare_btn, dict(row=4, column=1, pady=(10, 0))),
        ]

        frame.columnconfigure(1, weight=1)

    def _reveal_this_week_row(self):
        for widget, grid_opts in self.this_week_row_widgets:
            widget.grid(**grid_opts)

    def _hide_last_week_rows(self):
        for widget in self.last_week_row_widgets:
            widget.grid_remove()

    def _promote_compare_to_fetch_button(self):
        """Once Compare Orders has run, the same button slot becomes the Fetch Stock from Neto trigger."""
        self.compare_btn.config(text="Fetch Stock from Neto", command=self._on_fetch_neto, state="normal")
        self.neto_btn = self.compare_btn

    def _create_treeview(self, parent, expand=True, height=None):
        container = ttk.Frame(parent)
        container.pack(fill="both" if expand else "x", expand=expand)

        tree = ttk.Treeview(container, show="headings", height=height) if height else ttk.Treeview(container, show="headings")

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

    def _browse_this_week_pad(self):
        path = filedialog.askopenfilename(
            title="Select This Week's Order Pad",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if path:
            self.this_week_pad_path.set(path)

    def _on_compare_this_week(self):
        path = self.this_week_pad_path.get().strip()
        if not path:
            messagebox.showwarning("Missing File", "Please select this week's Order Pad before comparing.")
            return

        try:
            self.this_week_pad_data = self._parse_order_pad(path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load This Week's Order Pad:\n{e}")
            return

        self._show_outstanding_view()
        self._hide_last_week_rows()
        self._promote_compare_to_fetch_button()

    def _on_compare(self):
        pad_path = self.order_pad_path.get().strip()
        conf_path = self.order_conf_path.get().strip()

        if not pad_path or not conf_path:
            messagebox.showwarning("Missing Files", "Please select both files before comparing.")
            return

        self.showing_outstanding_view = False
        self._load_order_pad(pad_path)
        self._load_order_conf(conf_path)
        self._refresh_view()

        self.reconcile_btn.config(state="disabled")
        self._reveal_this_week_row()

    def _parse_order_pad(self, path):
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb[wb.sheetnames[0]]

        rows = list(ws.iter_rows(values_only=True))
        wb.close()

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
            product_name = row[7] if len(row) > 7 and row[7] is not None else ""
            data_rows.append((product_code, qty, product_name))

        return data_rows

    def _load_order_pad(self, path):
        try:
            self.order_pad_data = self._parse_order_pad(path)
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
        if self.bottom_frame:
            self.bottom_frame.destroy()

        self.table_container = ttk.Frame(self.root)
        self.table_container.pack(fill="both", expand=True, padx=10, pady=(5, 5))
        self.tree = self._create_treeview(self.table_container)

        headers = ["Product Code", "Product Name", "Qty Ordered", "Qty Confirmed", "Qty Outstanding"]
        merged = []
        last_week_ordered_lookup = {}
        for product_code, qty_ordered, product_name in self.order_pad_data:
            qty_confirmed = self.conf_qty_lookup.get(str(product_code), 0)
            qty_outstanding = qty_ordered - qty_confirmed
            merged.append((product_code, product_name, qty_ordered, qty_confirmed, qty_outstanding))
            if qty_ordered > 0:
                last_week_ordered_lookup[str(product_code)] = (product_code, product_name, qty_outstanding)
        self._populate_tree(self.tree, headers, merged)

        self.outstanding_lookup = {
            str(code): outstanding for code, _name, _qty_ordered, _qty_confirmed, outstanding in merged
        }
        self.last_week_ordered_lookup = last_week_ordered_lookup

        self.status_var.set("Reconciled. Select this week's Order Pad to continue.")
        self._build_bottom_frame()

    def _build_bottom_frame(self):
        if self.bottom_frame:
            self.bottom_frame.destroy()

        self.bottom_frame = ttk.Frame(self.root)
        self.bottom_frame.pack(fill="x", padx=10, pady=(0, 10))

        ttk.Label(self.bottom_frame, textvariable=self.status_var).pack(side="left")

    def _show_outstanding_view(self):
        if self.table_container:
            self.table_container.destroy()

        self.table_container = ttk.Frame(self.root)
        self.table_container.pack(fill="both", expand=True, padx=10, pady=(5, 5))

        this_week_codes = set()
        seen = set()
        matched_rows = []
        for product_code, _qty_ordered, product_name in self.this_week_pad_data:
            key = str(product_code)
            this_week_codes.add(key)
            if key in seen:
                continue
            seen.add(key)
            qty_outstanding = self.outstanding_lookup.get(key, 0)
            matched_rows.append((product_code, product_name, qty_outstanding))

        # Codes that had a real order last week but don't appear in this week's pad at all —
        # likely because the barcode/product code or name changed.
        unmatched_rows = [
            (orig_code, product_name, qty_outstanding)
            for key, (orig_code, product_name, qty_outstanding) in self.last_week_ordered_lookup.items()
            if key not in this_week_codes
        ]

        if unmatched_rows:
            paned = ttk.PanedWindow(self.table_container, orient="vertical")
            paned.pack(fill="both", expand=True)

            unmatched_frame = ttk.LabelFrame(
                paned,
                text="Unmatched Product Codes (ordered last week, not found in this week's pad — check for barcode/name changes)",
                padding=5,
            )
            self.unmatched_tree = self._create_treeview(
                unmatched_frame, expand=True, height=min(len(unmatched_rows), 15)
            )
            unmatched_iids = self._populate_tree(
                self.unmatched_tree,
                ["Product Code", "Product Name", "Qty Outstanding (Last Week)"],
                unmatched_rows,
            )
            self.unmatched_row_data = {
                iid: {"code": code, "name": name, "qty": qty}
                for iid, (code, name, qty) in zip(unmatched_iids, unmatched_rows)
            }

            action_frame = ttk.Frame(unmatched_frame)
            action_frame.pack(fill="x", pady=(6, 0))
            ttk.Button(action_frame, text="Ignore Selected", command=self._on_ignore_unmatched).pack(
                side="left", padx=(0, 8)
            )
            ttk.Button(
                action_frame, text="Reassign to Existing Code...", command=self._on_reassign_unmatched
            ).pack(side="left")

            self._enable_row_copy(self.unmatched_tree)

            paned.add(unmatched_frame, weight=1)

            main_frame = ttk.Frame(paned)
            self.tree = self._create_treeview(main_frame)
            main_iids = self._populate_tree(
                self.tree, ["Product Code", "Product Name", "Qty Outstanding"], matched_rows
            )
            self.main_row_iid = {str(code): iid for (code, _name, _qty), iid in zip(matched_rows, main_iids)}
            self._enable_row_copy(self.tree)
            paned.add(main_frame, weight=2)
        else:
            self.unmatched_tree = None
            self.unmatched_row_data = {}

            main_frame = ttk.Frame(self.table_container)
            main_frame.pack(fill="both", expand=True)
            self.tree = self._create_treeview(main_frame)
            main_iids = self._populate_tree(
                self.tree, ["Product Code", "Product Name", "Qty Outstanding"], matched_rows
            )
            self.main_row_iid = {str(code): iid for (code, _name, _qty), iid in zip(matched_rows, main_iids)}
            self._enable_row_copy(self.tree)

        self.showing_outstanding_view = True

        status = f"Showing outstanding qty for {os.path.basename(self.this_week_pad_path.get())}"
        if unmatched_rows:
            status += f" — {len(unmatched_rows)} unmatched product code(s)"
        self.status_var.set(status)
        self._build_bottom_frame()

    def _copy_selected_codes(self, tree):
        selected = tree.selection()
        if not selected:
            return

        codes = [tree.set(iid, "Product Code") for iid in selected]
        self.root.clipboard_clear()
        self.root.clipboard_append("\n".join(codes))

        if len(codes) == 1:
            self.status_var.set(f"Copied product code {codes[0]} to clipboard.")
        else:
            self.status_var.set(f"Copied {len(codes)} product codes to clipboard.")

    def _enable_row_copy(self, tree):
        menu = tk.Menu(tree, tearoff=0)
        menu.add_command(label="Copy Product Code", command=lambda: self._copy_selected_codes(tree))

        def show_context_menu(event):
            iid = tree.identify_row(event.y)
            if iid and iid not in tree.selection():
                tree.selection_set(iid)
            menu.tk_popup(event.x_root, event.y_root)

        tree.bind("<Control-c>", lambda e: self._copy_selected_codes(tree))
        tree.bind("<Command-c>", lambda e: self._copy_selected_codes(tree))
        tree.bind("<Button-3>", show_context_menu)
        tree.bind("<Button-2>", show_context_menu)
        tree.bind("<Control-Button-1>", show_context_menu)

    def _on_ignore_unmatched(self):
        if not self.unmatched_tree:
            return
        selected = self.unmatched_tree.selection()
        if not selected:
            messagebox.showinfo("No Selection", "Select one or more rows to ignore.")
            return

        for iid in selected:
            self.unmatched_tree.delete(iid)
            self.unmatched_row_data.pop(iid, None)

    def _on_reassign_unmatched(self):
        if not self.unmatched_tree:
            return
        selected = self.unmatched_tree.selection()
        if len(selected) != 1:
            messagebox.showinfo("Select One Row", "Select exactly one unmatched product code to reassign.")
            return

        iid = selected[0]
        row = self.unmatched_row_data.get(iid)
        if not row:
            return

        self._open_reassign_dialog(iid, row["code"], row["name"], row["qty"])

    def _bind_autocomplete(self, combo, all_values):
        def on_keyrelease(event):
            if event.keysym in ("Up", "Down", "Return", "Escape", "Tab"):
                return

            typed = combo.get()
            if typed == "":
                filtered = all_values
            else:
                typed_lower = typed.lower()
                filtered = [v for v in all_values if typed_lower in v.lower()]

            combo["values"] = filtered

            if typed and filtered:
                combo.event_generate("<Down>")
                combo.focus_set()
                combo.icursor(tk.END)

        combo.bind("<KeyRelease>", on_keyrelease)

    def _open_reassign_dialog(self, unmatched_iid, old_code, old_name, old_qty):
        codes = sorted(self.main_row_iid.keys())
        if not codes:
            messagebox.showinfo("No Product Codes", "There are no product codes in this week's table to reassign to.")
            return

        dialog = tk.Toplevel(self.root)
        dialog.title("Reassign Product Code")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)

        unmatched_label = f"Unmatched code: {old_code}"
        if old_name:
            unmatched_label += f" — {old_name}"
        ttk.Label(dialog, text=unmatched_label).grid(
            row=0, column=0, columnspan=2, sticky="w", padx=10, pady=(10, 6)
        )

        ttk.Label(dialog, text="New Product Code:").grid(row=1, column=0, sticky="w", padx=10, pady=4)
        code_var = tk.StringVar()
        combo = ttk.Combobox(dialog, textvariable=code_var, values=codes, width=28)
        combo.grid(row=1, column=1, padx=10, pady=4)
        self._bind_autocomplete(combo, codes)

        ttk.Label(dialog, text="Qty Outstanding:").grid(row=2, column=0, sticky="w", padx=10, pady=4)
        qty_var = tk.StringVar(value=str(old_qty))
        qty_entry = ttk.Entry(dialog, textvariable=qty_var, width=15)
        qty_entry.grid(row=2, column=1, sticky="w", padx=10, pady=4)

        btn_frame = ttk.Frame(dialog)
        btn_frame.grid(row=3, column=0, columnspan=2, pady=(10, 10))

        def on_confirm():
            new_code = code_var.get().strip()
            if not new_code:
                messagebox.showwarning("Missing Code", "Please select a product code.", parent=dialog)
                return
            try:
                new_qty = int(qty_var.get().strip())
            except ValueError:
                messagebox.showwarning("Invalid Qty", "Qty Outstanding must be a whole number.", parent=dialog)
                return

            main_iid = self.main_row_iid.get(new_code)
            if main_iid is None:
                messagebox.showwarning(
                    "Not Found", "Selected product code was not found in the table.", parent=dialog
                )
                return

            self.tree.set(main_iid, "Qty Outstanding", new_qty)

            self.unmatched_tree.delete(unmatched_iid)
            self.unmatched_row_data.pop(unmatched_iid, None)

            dialog.destroy()

        ttk.Button(btn_frame, text="Confirm", command=on_confirm).pack(side="left", padx=6)
        ttk.Button(btn_frame, text="Cancel", command=dialog.destroy).pack(side="left", padx=6)

    def _on_fetch_neto(self):
        if self.unmatched_row_data:
            count = len(self.unmatched_row_data)
            messagebox.showwarning(
                "Unmatched Product Codes",
                f"There {'is' if count == 1 else 'are'} still {count} unmatched product code"
                f"{'' if count == 1 else 's'} that haven't been resolved.\n\n"
                "Please use \"Ignore Selected\" or \"Reassign to Existing Code...\" to resolve them "
                "before fetching stock from Neto.",
            )
            return

        self.neto_btn.config(state="disabled")
        self.status_var.set("Launching Neto scraper...")

        thread = threading.Thread(target=self._run_neto_scraper, daemon=True)
        thread.start()

    def _run_neto_scraper(self):
        script_path = os.path.join(self.script_dir, "neto_scraper.py")
        login_prompt_shown = False
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
                stripped = line.strip()
                output.append(stripped)
                self.root.after(0, lambda d=stripped: self.status_var.set(d))

                if not login_prompt_shown and stripped.startswith("Not logged in to Neto"):
                    login_prompt_shown = True
                    self.root.after(0, lambda: messagebox.showinfo(
                        "Login Required",
                        "You're not logged in to Neto yet.\n\n"
                        "A Chrome window has opened — please log in there. This will "
                        "continue automatically once it detects you're logged in.",
                    ))

            process.wait()

            if process.returncode == 0:
                self.root.after(0, lambda: self.status_var.set("Neto scraper finished."))
            else:
                last_line = next((l for l in reversed(output) if l), "")
                self.root.after(0, lambda: self.status_var.set(f"Scraper exited with error (code {process.returncode})"))
                self.root.after(0, lambda msg=last_line, code=process.returncode: messagebox.showerror(
                    "Neto Scraper Failed",
                    msg or f"Scraper exited with error (code {code}).",
                ))

        except Exception as e:
            self.root.after(0, lambda: self.status_var.set(f"Error: {e}"))

        self.root.after(0, lambda: self.neto_btn.config(state="normal"))

    def _populate_tree(self, tree, headers, data_rows):
        tree.delete(*tree.get_children())

        tree["columns"] = headers
        for h in headers:
            tree.heading(h, text=h)
            if h == "Product Name":
                tree.column(h, width=300, minwidth=150)
            else:
                tree.column(h, width=120, minwidth=60)

        iids = []
        for row in data_rows:
            values = [str(v) if v is not None else "" for v in row]
            iids.append(tree.insert("", "end", values=values))
        return iids


def main():
    root = tk.Tk()
    GWOrderTool(root)
    root.mainloop()


if __name__ == "__main__":
    main()
