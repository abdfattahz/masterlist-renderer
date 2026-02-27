import os
import sys
import queue
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from render_masterlist import parse_rgb, run_render_process


class MasterlistGuiApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Masterlist Renderer")
        self.root.geometry("760x620")
        self.root.minsize(700, 560)

        self.events = queue.Queue()
        self.is_running = False

        self.excel_var = tk.StringVar()
        self.output_var = tk.StringVar(value=os.path.abspath("output"))
        self.background_var = tk.StringVar()
        self.font_var = tk.StringVar()

        self.pairs_var = tk.StringVar(value="3")
        self.rows_var = tk.StringVar(value="18")
        self.alpha_var = tk.StringVar(value="175")
        self.font_size_var = tk.StringVar(value="14")
        self.header_font_size_var = tk.StringVar(value="18")
        self.text_color_var = tk.StringVar(value="20,0,0")
        self.header_text_color_var = tk.StringVar(value="255,255,255")
        self.match_table_to_bg_var = tk.BooleanVar(value=False)

        self.status_var = tk.StringVar(value="Ready.")

        self._build_ui()

    def _build_ui(self):
        main = ttk.Frame(self.root, padding=14)
        main.pack(fill="both", expand=True)

        basic = ttk.LabelFrame(main, text="Main Inputs", padding=12)
        basic.pack(fill="x")
        basic.columnconfigure(1, weight=1)

        self._path_field(
            parent=basic,
            row=0,
            label="Excel file",
            variable=self.excel_var,
            browse_callback=self._browse_excel,
        )
        self._path_field(
            parent=basic,
            row=1,
            label="Output folder",
            variable=self.output_var,
            browse_callback=self._browse_output,
        )
        self._path_field(
            parent=basic,
            row=2,
            label="Background image",
            variable=self.background_var,
            browse_callback=self._browse_background,
        )
        self._path_field(
            parent=basic,
            row=3,
            label="Font (.ttf)",
            variable=self.font_var,
            browse_callback=self._browse_font,
        )

        advanced = ttk.LabelFrame(main, text="Advanced Options", padding=12)
        advanced.pack(fill="x", pady=(12, 0))

        fields = [
            ("Pairs per row", self.pairs_var),
            ("Rows per page", self.rows_var),
            ("Cell alpha (0-255)", self.alpha_var),
            ("Body font size", self.font_size_var),
            ("Header font size", self.header_font_size_var),
            ("Body text RGB", self.text_color_var),
            ("Header text RGB", self.header_text_color_var),
        ]

        for idx, (label, var) in enumerate(fields):
            r = idx // 2
            c = (idx % 2) * 2
            ttk.Label(advanced, text=label).grid(
                row=r, column=c, sticky="w", padx=(0, 8), pady=4
            )
            ttk.Entry(advanced, textvariable=var, width=22).grid(
                row=r, column=c + 1, sticky="ew", pady=4
            )

        advanced.columnconfigure(1, weight=1)
        advanced.columnconfigure(3, weight=1)

        ttk.Checkbutton(
            advanced,
            text="Auto match table colors to background",
            variable=self.match_table_to_bg_var,
        ).grid(row=4, column=0, columnspan=4, sticky="w", pady=(8, 0))

        actions = ttk.Frame(main)
        actions.pack(fill="x", pady=(12, 0))

        self.generate_button = ttk.Button(
            actions, text="Generate PNG Pages", command=self._start_render
        )
        self.generate_button.pack(side="left")

        ttk.Button(
            actions, text="Open Output Folder", command=self._open_output_folder
        ).pack(side="left", padx=(8, 0))

        self.progress = ttk.Progressbar(main, mode="determinate", maximum=1, value=0)
        self.progress.pack(fill="x", pady=(12, 0))

        ttk.Label(main, textvariable=self.status_var).pack(anchor="w", pady=(8, 0))

    def _path_field(self, parent, row, label, variable, browse_callback):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(parent, textvariable=variable).grid(
            row=row, column=1, sticky="ew", padx=(8, 8), pady=4
        )
        ttk.Button(parent, text="Browse", command=browse_callback).grid(
            row=row, column=2, pady=4
        )

    def _browse_excel(self):
        path = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if path:
            self.excel_var.set(path)

    def _browse_output(self):
        path = filedialog.askdirectory(title="Select output folder")
        if path:
            self.output_var.set(path)

    def _browse_background(self):
        path = filedialog.askopenfilename(
            title="Select background image",
            filetypes=[
                ("Image files", "*.png *.jpg *.jpeg *.bmp"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self.background_var.set(path)

    def _browse_font(self):
        path = filedialog.askopenfilename(
            title="Select TTF font",
            filetypes=[("TTF fonts", "*.ttf"), ("All files", "*.*")],
        )
        if path:
            self.font_var.set(path)

    def _validate_inputs(self):
        excel_path = self.excel_var.get().strip()
        if not excel_path:
            raise ValueError("Please select an Excel file.")
        if not os.path.isfile(excel_path):
            raise ValueError("Selected Excel file does not exist.")

        output_path = self.output_var.get().strip()
        if not output_path:
            raise ValueError("Please choose an output folder.")

        bg_path = self.background_var.get().strip() or None
        if bg_path and not os.path.isfile(bg_path):
            raise ValueError("Selected background image does not exist.")

        font_path = self.font_var.get().strip() or None
        if font_path and not os.path.isfile(font_path):
            raise ValueError("Selected font file does not exist.")

        try:
            pairs = int(self.pairs_var.get().strip())
            rows = int(self.rows_var.get().strip())
            alpha = int(self.alpha_var.get().strip())
            font_size = int(self.font_size_var.get().strip())
            header_font_size = int(self.header_font_size_var.get().strip())
        except ValueError:
            raise ValueError(
                "Pairs, rows, alpha, and font sizes must be whole numbers."
            )

        if pairs <= 0:
            raise ValueError("Pairs per row must be greater than 0.")
        if rows <= 0:
            raise ValueError("Rows per page must be greater than 0.")
        if not (0 <= alpha <= 255):
            raise ValueError("Cell alpha must be between 0 and 255.")
        if font_size <= 0:
            raise ValueError("Body font size must be greater than 0.")
        if header_font_size <= 0:
            raise ValueError("Header font size must be greater than 0.")

        text_color = self.text_color_var.get().strip()
        header_text_color = self.header_text_color_var.get().strip()

        parse_rgb(text_color)
        parse_rgb(header_text_color)

        return {
            "excel_path": excel_path,
            "out_dir": output_path,
            "bg_path": bg_path,
            "font_path": font_path,
            "pairs": pairs,
            "rows": rows,
            "alpha": alpha,
            "font_size": font_size,
            "header_font_size": header_font_size,
            "text_color": text_color,
            "header_text_color": header_text_color,
            "match_table_to_bg": self.match_table_to_bg_var.get(),
        }

    def _start_render(self):
        if self.is_running:
            return

        try:
            options = self._validate_inputs()
        except Exception as exc:
            messagebox.showerror("Invalid input", str(exc))
            return

        self.is_running = True
        self.generate_button.configure(state="disabled")
        self.progress.configure(maximum=1, value=0)
        self.status_var.set("Preparing to render...")

        worker = threading.Thread(
            target=self._render_worker, args=(options,), daemon=True
        )
        worker.start()
        self.root.after(100, self._poll_events)

    def _render_worker(self, options):
        try:

            def on_progress(page, total_pages):
                self.events.put(("progress", page, total_pages))

            pages, total_rows = run_render_process(
                excel_path=options["excel_path"],
                out_dir=options["out_dir"],
                bg_path=options["bg_path"],
                font_path=options["font_path"],
                pairs=options["pairs"],
                rows=options["rows"],
                alpha=options["alpha"],
                font_size=options["font_size"],
                header_font_size=options["header_font_size"],
                text_color=options["text_color"],
                header_text_color=options["header_text_color"],
                match_table_to_bg=options["match_table_to_bg"],
                progress_callback=on_progress,
            )
            self.events.put(("done", pages, total_rows, options["out_dir"]))
        except Exception as exc:
            self.events.put(("error", str(exc)))

    def _poll_events(self):
        while True:
            try:
                event = self.events.get_nowait()
            except queue.Empty:
                break

            kind = event[0]
            if kind == "progress":
                _, page, total_pages = event
                self.progress.configure(maximum=max(1, total_pages), value=page)
                self.status_var.set(f"Rendering page {page}/{total_pages}...")
            elif kind == "done":
                _, pages, total_rows, out_dir = event
                self.is_running = False
                self.generate_button.configure(state="normal")
                self.status_var.set(
                    f"Done. Generated {pages} page(s) from {total_rows} row(s)."
                )
                messagebox.showinfo(
                    "Rendering complete",
                    f"Generated {pages} page(s) into:\n{out_dir}",
                )
            elif kind == "error":
                _, message = event
                self.is_running = False
                self.generate_button.configure(state="normal")
                self.status_var.set("Failed.")
                messagebox.showerror("Rendering failed", message)

        if self.is_running:
            self.root.after(100, self._poll_events)

    def _open_output_folder(self):
        path = self.output_var.get().strip()
        if not path:
            messagebox.showerror("Output folder", "Please set an output folder first.")
            return

        if not os.path.isdir(path):
            messagebox.showerror(
                "Output folder",
                "Output folder does not exist yet. Generate once to create it.",
            )
            return

        try:
            if sys.platform.startswith("win"):
                startfile = getattr(os, "startfile", None)
                if callable(startfile):
                    startfile(path)
                else:
                    subprocess.Popen(["explorer", path])
            elif sys.platform == "darwin":
                subprocess.Popen(["open", path])
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception as exc:
            messagebox.showerror("Open folder failed", str(exc))


def main():
    root = tk.Tk()
    MasterlistGuiApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
