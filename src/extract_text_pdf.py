from pathlib import Path
import re
import json
import threading
import queue
import tkinter as tk
from tkinter import filedialog, messagebox, ttk



# File-type support
PLAIN_TEXT_EXTS = {
    ".txt", ".md", ".rst", ".log",
    ".csv", ".tsv",
    ".json", ".ndjson", ".jsonl",
    ".xml", ".xhtml",
    ".yaml", ".yml",
    ".toml", ".ini", ".cfg", ".env", ".properties",
    ".py", ".js", ".ts", ".jsx", ".tsx",
    ".java", ".c", ".cc", ".cpp", ".h", ".hpp",
    ".cs", ".go", ".rs", ".rb", ".php", ".sql", ".tex",
    ".bat", ".ps1", ".sh",
    ".html", ".htm",
    ".ipynb",
    ".srt",
}

# "Container" formats that need optional libraries
CONTAINER_EXTS = {".pdf", ".docx", ".rtf"}


class CancelledError(Exception):
    """Raised when the user cancels the current scan/run."""


def _pick_unique_report_path(output_folder: Path, base_name: str = "report", ext: str = ".txt") -> Path:
    output_folder.mkdir(parents=True, exist_ok=True)
    report_file = output_folder / f"{base_name}{ext}"
    if not report_file.exists():
        return report_file
    i = 1
    while True:
        candidate = output_folder / f"{base_name}({i}){ext}"
        if not candidate.exists():
            return candidate
        i += 1


def _extract_text_from_plain_text(path: Path) -> str:
    """
    Reads a file as text without being too picky about encoding.

    Most files are UTF-8 these days; for anything else, we replace characters we can't decode
    instead of failing the whole run.
    """
    data = path.read_bytes()
    try:
        return data.decode("utf-8-sig", errors="replace")
    except Exception:
        return data.decode("utf-8", errors="replace")


def _extract_text_from_pdf(path: Path, cancel_event: threading.Event | None = None) -> str:
    """
    Extract text from a PDF. If cancellation is requested, we bail out between pages.
    """
    try:
        from pypdf import PdfReader
    except ModuleNotFoundError as e:
        import sys
        raise ModuleNotFoundError(
            "Missing dependency 'pypdf' for PDF extraction.\n"
            f"Python executable: {sys.executable}\n"
            "Install with: python -m pip install pypdf"
        ) from e

    reader = PdfReader(str(path))
    parts: list[str] = []
    for page in reader.pages:
        if cancel_event and cancel_event.is_set():
            raise CancelledError("Cancelled by user.")
        parts.append(page.extract_text() or "")
    return "\n".join(parts)


def _extract_text_from_docx(path: Path) -> str:
    try:
        import docx  # python-docx
    except ModuleNotFoundError as e:
        import sys
        raise ModuleNotFoundError(
            "Missing dependency 'python-docx' for DOCX extraction.\n"
            f"Python executable: {sys.executable}\n"
            "Install with: python -m pip install python-docx"
        ) from e

    doc = docx.Document(str(path))
    return "\n".join(p.text for p in doc.paragraphs)


def _extract_text_from_rtf(path: Path) -> str:
    try:
        from striprtf.striprtf import rtf_to_text
    except ModuleNotFoundError as e:
        import sys
        raise ModuleNotFoundError(
            "Missing dependency 'striprtf' for RTF extraction.\n"
            f"Python executable: {sys.executable}\n"
            "Install with: python -m pip install striprtf"
        ) from e

    raw = _extract_text_from_plain_text(path)
    return rtf_to_text(raw)


def _extract_text_from_html(path: Path) -> str:
    """
    If BeautifulSoup is available, strip tags properly.
    Otherwise, fall back to raw text.
    """
    html = _extract_text_from_plain_text(path)
    try:
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(html, "html.parser")
        return soup.get_text(separator=" ")
    except ModuleNotFoundError:
        return html


def _extract_text_from_ipynb(path: Path) -> str:
    raw = _extract_text_from_plain_text(path)
    nb = json.loads(raw)

    parts: list[str] = []
    for cell in nb.get("cells", []):
        src = cell.get("source", [])
        if isinstance(src, list):
            parts.append("".join(src))
        elif isinstance(src, str):
            parts.append(src)
    return "\n".join(parts)


def _extract_text_any(path: Path, cancel_event: threading.Event | None = None) -> str:
    ext = path.suffix.lower()

    if ext == ".pdf":
        return _extract_text_from_pdf(path, cancel_event=cancel_event)
    if ext == ".docx":
        return _extract_text_from_docx(path)
    if ext == ".rtf":
        return _extract_text_from_rtf(path)
    if ext in {".html", ".htm", ".xhtml"}:
        return _extract_text_from_html(path)
    if ext == ".ipynb":
        return _extract_text_from_ipynb(path)
    if ext in PLAIN_TEXT_EXTS:
        return _extract_text_from_plain_text(path)

    raise ValueError(f"Unsupported file type: {ext}")


def _parse_words(words_raw: str) -> list[str]:
    tokens = [t.strip() for t in re.split(r"[,\s]+", words_raw) if t.strip()]
    seen = set()
    words: list[str] = []
    for t in tokens:
        key = t.casefold()
        if key not in seen:
            seen.add(key)
            words.append(t)
    return words


def _normalize_ext_list(exts_raw: str) -> set[str]:
    tokens = [t.strip().lower() for t in re.split(r"[,\s]+", exts_raw or "") if t.strip()]
    out: set[str] = set()
    for t in tokens:
        if not t.startswith("."):
            t = "." + t
        out.add(t)
    return out


def _list_target_files(input_folder: Path, recursive: bool, exts: set[str]) -> list[Path]:
    exts = {e.lower() if e.startswith(".") else f".{e.lower()}" for e in exts}
    patterns = [f"*{ext}" for ext in sorted(exts)]
    files: list[Path] = []

    if recursive:
        for pat in patterns:
            files.extend(input_folder.rglob(pat))
    else:
        for pat in patterns:
            files.extend(input_folder.glob(pat))

    unique = {}
    for f in files:
        unique[str(f.resolve())] = f
    return sorted(unique.values(), key=lambda p: p.name.lower())


def generate_word_report_for_files(
    input_folder,
    output_folder,
    words,
    case_sensitive: bool = False,
    recursive: bool = False,
    scan_exts: set[str] | None = None,
    progress_cb=None,
    cancel_event: threading.Event | None = None, 
) -> Path:
    input_folder = Path(input_folder)
    output_folder = Path(output_folder)

    if not input_folder.exists() or not input_folder.is_dir():
        raise ValueError(f"Input folder does not exist or is not a folder: {input_folder}")

    if isinstance(words, str):
        words = _parse_words(words)
    if not words:
        raise ValueError("No words provided. Please enter at least one word to check for.")

    if scan_exts is None:
        scan_exts = set(PLAIN_TEXT_EXTS) | set(CONTAINER_EXTS)

    supported = set(PLAIN_TEXT_EXTS) | set(CONTAINER_EXTS)
    scan_exts = {e.lower() if e.startswith(".") else f".{e.lower()}" for e in scan_exts}
    scan_exts = {e for e in scan_exts if e in supported}

    if not scan_exts:
        raise ValueError("No supported file formats selected. Please select at least one format to scan.")

    report_file = _pick_unique_report_path(output_folder, base_name="report", ext=".txt")

    target_files = _list_target_files(input_folder, recursive=recursive, exts=scan_exts)
    if not target_files:
        raise ValueError(f"No files found for selected formats in: {input_folder}")

    flags = 0 if case_sensitive else re.IGNORECASE
    patterns = {w: re.compile(rf"\b{re.escape(w)}\b", flags) for w in words}

    rows: list[dict] = []
    total_count = len(target_files)

    for idx, file_path in enumerate(target_files, start=1):
        # cancel check inbetween file scan
        if cancel_event and cancel_event.is_set():
            raise CancelledError("Cancelled by user.")

        try:
            text = _extract_text_any(file_path, cancel_event=cancel_event)
            counts_for_file = {w: len(patterns[w].findall(text)) for w in words}
            rows.append({"file": file_path.name, "error": "", "counts": counts_for_file})
        except CancelledError:
            raise
        except Exception as e:
            rows.append(
                {"file": file_path.name, "error": f"{type(e).__name__}: {e}", "counts": {w: 0 for w in words}}
            )

        if progress_cb:
            progress_cb(idx, total_count, file_path.name)

    file_col_width = max(len("File"), max(len(r["file"]) for r in rows))
    word_col_widths = {w: max(len(w), max(len(str(r["counts"][w])) for r in rows)) for w in words}
    total_col_width = max(len("Total"), len(str(max(sum(r["counts"].values()) for r in rows))))

    with open(report_file, "w", encoding="utf-8") as report:
        report.write("Word Count Report (Multi-format)\n")
        report.write(f"Input folder : {input_folder}\n")
        report.write(f"Output file  : {report_file}\n")
        report.write(f"Recursive scan: {recursive}\n")
        report.write(f"Case sensitive: {case_sensitive}\n")
        report.write(f"Selected formats: {', '.join(sorted(scan_exts))}\n")
        report.write("\n")

        header_cells = [f"{'File':<{file_col_width}}"]
        for w in words:
            header_cells.append(f"{w:>{word_col_widths[w]}}")
        header_cells.append(f"{'Total':>{total_col_width}}")
        header_line = " | ".join(header_cells)

        report.write(header_line + "\n")
        report.write("-" * len(header_line) + "\n")

        for r in rows:
            row_total = sum(r["counts"].values())
            line_cells = [f"{r['file']:<{file_col_width}}"]
            for w in words:
                line_cells.append(f"{r['counts'][w]:>{word_col_widths[w]}}")
            line_cells.append(f"{row_total:>{total_col_width}}")
            report.write(" | ".join(line_cells) + "\n")

            if r["error"]:
                report.write(f"{'':<{file_col_width}}   [Extraction error: {r['error']}]\n")

        report.write("-" * len(header_line) + "\n")
        totals_per_word = {w: sum(r["counts"][w] for r in rows) for w in words}
        grand_total = sum(totals_per_word.values())

        total_cells = [f"{'TOTAL':<{file_col_width}}"]
        for w in words:
            total_cells.append(f"{totals_per_word[w]:>{word_col_widths[w]}}")
        total_cells.append(f"{grand_total:>{total_col_width}}")
        report.write(" | ".join(total_cells) + "\n")

    return report_file


# ----------------------------
# GUI helpers
# ----------------------------
def _browse_input_folder(var: tk.StringVar):
    folder = filedialog.askdirectory(title="Select folder to scan")
    if folder:
        var.set(folder)


def _browse_output_folder(var: tk.StringVar):
    folder = filedialog.askdirectory(title="Select folder to save the report")
    if folder:
        var.set(folder)


def _run_generate(
    root: tk.Tk,
    input_var: tk.StringVar,
    output_var: tk.StringVar,
    words_var: tk.StringVar,
    case_var: tk.BooleanVar,
    recursive_var: tk.BooleanVar,
    fmt_pdf_var: tk.BooleanVar,
    fmt_docx_var: tk.BooleanVar,
    fmt_rtf_var: tk.BooleanVar,
    fmt_txt_var: tk.BooleanVar,
    fmt_md_var: tk.BooleanVar,
    fmt_csv_var: tk.BooleanVar,
    fmt_json_var: tk.BooleanVar,
    fmt_html_var: tk.BooleanVar,
    fmt_ipynb_var: tk.BooleanVar,
    fmt_xml_var: tk.BooleanVar,
    fmt_yaml_var: tk.BooleanVar,
    fmt_ini_cfg_var: tk.BooleanVar,
    fmt_log_var: tk.BooleanVar,
    fmt_sql_var: tk.BooleanVar,
    fmt_tex_var: tk.BooleanVar,
    fmt_rst_var: tk.BooleanVar,
    fmt_properties_var: tk.BooleanVar,
    fmt_toml_var: tk.BooleanVar,
    fmt_env_var: tk.BooleanVar,
    fmt_ndjson_var: tk.BooleanVar,
    fmt_xhtml_var: tk.BooleanVar,
    fmt_srt_var: tk.BooleanVar,
    other_exts_var: tk.StringVar,
    status_var: tk.StringVar,
    progress_var: tk.DoubleVar,
    progress_label_var: tk.StringVar,
    generate_btn: tk.Button,
    cancel_btn: tk.Button, 
):
    input_folder = input_var.get().strip()
    output_folder = output_var.get().strip()
    words_raw = words_var.get().strip()

    if not input_folder:
        messagebox.showerror("Missing input folder", "Please select the folder to scan.")
        return
    if not output_folder:
        messagebox.showerror("Missing output folder", "Please select the folder to save the report.")
        return
    if not words_raw:
        messagebox.showerror("Missing words", "Please enter at least one word to check for.")
        return

    words = _parse_words(words_raw)
    if not words:
        messagebox.showerror("Missing words", "Please enter at least one valid word to check for.")
        return

    selected_exts: set[str] = set()
    if fmt_pdf_var.get():
        selected_exts.add(".pdf")
    if fmt_docx_var.get():
        selected_exts.add(".docx")
    if fmt_rtf_var.get():
        selected_exts.add(".rtf")
    if fmt_txt_var.get():
        selected_exts.add(".txt")
    if fmt_md_var.get():
        selected_exts.add(".md")
    if fmt_csv_var.get():
        selected_exts.update({".csv", ".tsv"})
    if fmt_json_var.get():
        selected_exts.add(".json")
    if fmt_html_var.get():
        selected_exts.update({".html", ".htm"})
    if fmt_ipynb_var.get():
        selected_exts.add(".ipynb")

    if fmt_xml_var.get():
        selected_exts.add(".xml")
    if fmt_yaml_var.get():
        selected_exts.update({".yaml", ".yml"})
    if fmt_ini_cfg_var.get():
        selected_exts.update({".ini", ".cfg"})
    if fmt_log_var.get():
        selected_exts.add(".log")
    if fmt_sql_var.get():
        selected_exts.add(".sql")
    if fmt_tex_var.get():
        selected_exts.add(".tex")
    if fmt_rst_var.get():
        selected_exts.add(".rst")
    if fmt_properties_var.get():
        selected_exts.add(".properties")
    if fmt_toml_var.get():
        selected_exts.add(".toml")
    if fmt_env_var.get():
        selected_exts.add(".env")
    if fmt_ndjson_var.get():
        selected_exts.update({".ndjson", ".jsonl"})
    if fmt_xhtml_var.get():
        selected_exts.add(".xhtml")
    if fmt_srt_var.get():
        selected_exts.add(".srt")

    selected_exts |= _normalize_ext_list(other_exts_var.get())

    supported = set(PLAIN_TEXT_EXTS) | set(CONTAINER_EXTS)
    selected_supported = {e.lower() if e.startswith(".") else f".{e.lower()}" for e in selected_exts}
    selected_supported = {e for e in selected_supported if e in supported}
    if not selected_supported:
        messagebox.showerror("Missing formats", "Please select at least one supported file format.")
        return


    cancel_event = threading.Event()


    generate_btn.config(state="disabled")
    cancel_btn.config(state="normal")


    progress_var.set(0)
    progress_label_var.set("Starting...")
    status_var.set("Working...")

    q: queue.Queue = queue.Queue()

    def progress_cb(done_count: int, total_count: int, filename: str):
        q.put(("progress", done_count, total_count, filename))

    def on_cancel():
        #feedback
        cancel_btn.config(state="disabled")
        status_var.set("Cancelling… (stops after the current file)")
        progress_label_var.set("Cancelling…")
        cancel_event.set()

    cancel_btn.config(command=on_cancel)

    def worker():
        try:
            report_path = generate_word_report_for_files(
                input_folder=input_folder,
                output_folder=output_folder,
                words=words,
                case_sensitive=bool(case_var.get()),
                recursive=bool(recursive_var.get()),
                scan_exts=selected_exts,
                progress_cb=progress_cb,
                cancel_event=cancel_event, 
            )
            q.put(("done", str(report_path)))
        except CancelledError:
            q.put(("cancelled",))
        except Exception as e:
            q.put(("error", str(e)))

    threading.Thread(target=worker, daemon=True).start()

    def poll_queue():
        try:
            while True:
                msg = q.get_nowait()

                if msg[0] == "progress":
                    _, done_count, total_count, filename = msg
                    percent = (done_count / total_count) * 100 if total_count else 0
                    progress_var.set(percent)
                    progress_label_var.set(f"{done_count}/{total_count}  {filename}")

                elif msg[0] == "done":
                    _, report_path = msg
                    progress_var.set(100)
                    progress_label_var.set("Completed.")
                    status_var.set(f"Report generated at: {report_path}")
                    generate_btn.config(state="normal")
                    cancel_btn.config(state="disabled")
                    messagebox.showinfo("Done", f"Report generated at:\n{report_path}")
                    return

                elif msg[0] == "cancelled":
                    progress_var.set(0)
                    progress_label_var.set("Cancelled.")
                    status_var.set("Cancelled. No report generated.")
                    generate_btn.config(state="normal")
                    cancel_btn.config(state="disabled")
                    return

                elif msg[0] == "error":
                    _, err = msg
                    progress_label_var.set("Failed.")
                    status_var.set("Error.")
                    generate_btn.config(state="normal")
                    cancel_btn.config(state="disabled")
                    messagebox.showerror("Error", err)
                    return

        except queue.Empty:
            pass

        root.after(100, poll_queue)

    poll_queue()


##### GUI ###### (mostly AI generated, with some manual tweaks)
class _ToolTip:
    def __init__(self, widget, text: str):
        self.widget = widget
        self.text = text
        self._tip = None
        widget.bind("<Enter>", self._show)
        widget.bind("<Leave>", self._hide)

    def _show(self, _event=None):
        if self._tip or not self.text:
            return
        x = self.widget.winfo_rootx() + 12
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 8
        self._tip = tk.Toplevel(self.widget)
        self._tip.wm_overrideredirect(True)
        self._tip.wm_geometry(f"+{x}+{y}")
        label = ttk.Label(self._tip, text=self.text, padding=(10, 6), style="Tip.TLabel")
        label.pack()

    def _hide(self, _event=None):
        if self._tip:
            self._tip.destroy()
            self._tip = None


def main():
    root = tk.Tk()
    root.title("Word Count Report (Multi-format)")
    root.minsize(760, 720)
    root.resizable(True, True)


    style = ttk.Style(root)
    for theme in ("clam", "vista", "xpnative"):
        try:
            style.theme_use(theme)
            break
        except tk.TclError:
            pass

    default_font = ("Segoe UI", 10)
    root.option_add("*Font", default_font)

    style.configure("Title.TLabel", font=("Segoe UI", 14, "bold"))
    style.configure("Subtitle.TLabel", foreground="#555555")
    style.configure("Tip.TLabel", background="#111111", foreground="#ffffff")
    style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"))
    style.configure("Danger.TButton", foreground="#b00020")

    # ---- Variables
    input_var = tk.StringVar()
    output_var = tk.StringVar()
    words_var = tk.StringVar()

    case_var = tk.BooleanVar(value=False)
    recursive_var = tk.BooleanVar(value=False)

    fmt_pdf_var = tk.BooleanVar(value=True)
    fmt_docx_var = tk.BooleanVar(value=False)
    fmt_rtf_var = tk.BooleanVar(value=False)
    fmt_txt_var = tk.BooleanVar(value=True)
    fmt_md_var = tk.BooleanVar(value=True)
    fmt_csv_var = tk.BooleanVar(value=False)
    fmt_json_var = tk.BooleanVar(value=False)
    fmt_html_var = tk.BooleanVar(value=False)
    fmt_ipynb_var = tk.BooleanVar(value=False)

    fmt_xml_var = tk.BooleanVar(value=False)
    fmt_yaml_var = tk.BooleanVar(value=False)
    fmt_ini_cfg_var = tk.BooleanVar(value=False)
    fmt_log_var = tk.BooleanVar(value=False)
    fmt_sql_var = tk.BooleanVar(value=False)
    fmt_tex_var = tk.BooleanVar(value=False)
    fmt_rst_var = tk.BooleanVar(value=False)
    fmt_properties_var = tk.BooleanVar(value=False)
    fmt_toml_var = tk.BooleanVar(value=False)
    fmt_env_var = tk.BooleanVar(value=False)
    fmt_ndjson_var = tk.BooleanVar(value=False)
    fmt_xhtml_var = tk.BooleanVar(value=False)
    fmt_srt_var = tk.BooleanVar(value=False)

    other_exts_var = tk.StringVar(value="")
    status_var = tk.StringVar(value="Ready.")
    progress_var = tk.DoubleVar(value=0.0)
    progress_label_var = tk.StringVar(value="Idle.")

    PADX, PADY = 12, 10

    # ---- Layout root grid
    root.columnconfigure(0, weight=1)
    root.rowconfigure(1, weight=1)

    header = ttk.Frame(root, padding=(PADX, PADY))
    header.grid(row=0, column=0, sticky="ew")
    header.columnconfigure(0, weight=1)

    ttk.Label(header, text="Multi-format Word Counter", style="Title.TLabel").grid(row=0, column=0, sticky="w")
    ttk.Label(
        header,
        text="Scan selected file types, count your words, export a table report (no overwrites).",
        style="Subtitle.TLabel",
    ).grid(row=1, column=0, sticky="w", pady=(4, 0))

    ttk.Separator(root).grid(row=2, column=0, sticky="ew")

    notebook = ttk.Notebook(root)
    notebook.grid(row=1, column=0, sticky="nsew", padx=PADX, pady=(PADY, 0))

    # ---- Tabs
    tab_folders = ttk.Frame(notebook, padding=(PADX, PADY))
    tab_formats = ttk.Frame(notebook, padding=(PADX, PADY))
    tab_words = ttk.Frame(notebook, padding=(PADX, PADY))
    tab_run = ttk.Frame(notebook, padding=(PADX, PADY))

    notebook.add(tab_folders, text="Folders")
    notebook.add(tab_formats, text="Formats")
    notebook.add(tab_words, text="Words")
    notebook.add(tab_run, text="Run")

    # -----------------
    # Folders tab
    # -----------------
    lf_folders = ttk.Labelframe(tab_folders, text="Where to scan and where to save", padding=(PADX, PADY))
    lf_folders.grid(row=0, column=0, sticky="nsew")
    tab_folders.columnconfigure(0, weight=1)

    lf_folders.columnconfigure(0, weight=1)

    ttk.Label(lf_folders, text="Scan folder").grid(row=0, column=0, sticky="w")
    in_row = ttk.Frame(lf_folders)
    in_row.grid(row=1, column=0, sticky="ew", pady=(6, 12))
    in_row.columnconfigure(0, weight=1)

    input_entry = ttk.Entry(in_row, textvariable=input_var)
    input_entry.grid(row=0, column=0, sticky="ew")
    btn_browse_in = ttk.Button(in_row, text="Browse…", command=lambda: _browse_input_folder(input_var))
    btn_browse_in.grid(row=0, column=1, padx=(10, 0))

    ttk.Label(lf_folders, text="Output folder").grid(row=2, column=0, sticky="w")
    out_row = ttk.Frame(lf_folders)
    out_row.grid(row=3, column=0, sticky="ew", pady=(6, 0))
    out_row.columnconfigure(0, weight=1)

    output_entry = ttk.Entry(out_row, textvariable=output_var)
    output_entry.grid(row=0, column=0, sticky="ew")
    btn_browse_out = ttk.Button(out_row, text="Browse…", command=lambda: _browse_output_folder(output_var))
    btn_browse_out.grid(row=0, column=1, padx=(10, 0))

    _ToolTip(btn_browse_in, "Select the folder that contains the files to scan.")
    _ToolTip(btn_browse_out, "Select the folder where the report.txt will be created.")

    # -----------------
    # Formats tab
    # -----------------
    tab_formats.columnconfigure(0, weight=1)

    lf_formats = ttk.Labelframe(tab_formats, text="File formats to include", padding=(PADX, PADY))
    lf_formats.grid(row=0, column=0, sticky="nsew")
    lf_formats.columnconfigure((0, 1, 2), weight=1)

    def _set_formats(value: bool):
        for v in (
            fmt_pdf_var, fmt_docx_var, fmt_rtf_var, fmt_txt_var, fmt_md_var, fmt_csv_var, fmt_json_var,
            fmt_html_var, fmt_ipynb_var, fmt_xml_var, fmt_yaml_var, fmt_ini_cfg_var, fmt_log_var, fmt_sql_var,
            fmt_tex_var, fmt_rst_var, fmt_properties_var, fmt_toml_var, fmt_env_var, fmt_ndjson_var,
            fmt_xhtml_var, fmt_srt_var,
        ):
            v.set(value)

    fmt_btns = ttk.Frame(lf_formats)
    fmt_btns.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 10))
    ttk.Button(fmt_btns, text="Select all", command=lambda: _set_formats(True)).grid(row=0, column=0, padx=(0, 8))
    ttk.Button(fmt_btns, text="Clear", command=lambda: _set_formats(False)).grid(row=0, column=1)

    # Three columns of checkboxes
    col1 = ttk.Frame(lf_formats)
    col2 = ttk.Frame(lf_formats)
    col3 = ttk.Frame(lf_formats)
    col1.grid(row=1, column=0, sticky="nw")
    col2.grid(row=1, column=1, sticky="nw", padx=(18, 0))
    col3.grid(row=1, column=2, sticky="nw", padx=(18, 0))

    # Containers + web
    ttk.Checkbutton(col1, text="PDF (.pdf)", variable=fmt_pdf_var).grid(row=0, column=0, sticky="w", pady=2)
    ttk.Checkbutton(col1, text="DOCX (.docx)", variable=fmt_docx_var).grid(row=1, column=0, sticky="w", pady=2)
    ttk.Checkbutton(col1, text="RTF (.rtf)", variable=fmt_rtf_var).grid(row=2, column=0, sticky="w", pady=2)
    ttk.Checkbutton(col1, text="HTML/HTM", variable=fmt_html_var).grid(row=3, column=0, sticky="w", pady=2)
    ttk.Checkbutton(col1, text="XHTML (.xhtml)", variable=fmt_xhtml_var).grid(row=4, column=0, sticky="w", pady=2)

    # Common text / docs
    ttk.Checkbutton(col2, text="TXT (.txt)", variable=fmt_txt_var).grid(row=0, column=0, sticky="w", pady=2)
    ttk.Checkbutton(col2, text="MD (.md)", variable=fmt_md_var).grid(row=1, column=0, sticky="w", pady=2)
    ttk.Checkbutton(col2, text="RST (.rst)", variable=fmt_rst_var).grid(row=2, column=0, sticky="w", pady=2)
    ttk.Checkbutton(col2, text="LOG (.log)", variable=fmt_log_var).grid(row=3, column=0, sticky="w", pady=2)
    ttk.Checkbutton(col2, text="SRT (.srt)", variable=fmt_srt_var).grid(row=4, column=0, sticky="w", pady=2)
    ttk.Checkbutton(col2, text="IPYNB (.ipynb)", variable=fmt_ipynb_var).grid(row=5, column=0, sticky="w", pady=2)

    # Structured/config/dev
    ttk.Checkbutton(col3, text="CSV/TSV", variable=fmt_csv_var).grid(row=0, column=0, sticky="w", pady=2)
    ttk.Checkbutton(col3, text="JSON (.json)", variable=fmt_json_var).grid(row=1, column=0, sticky="w", pady=2)
    ttk.Checkbutton(col3, text="NDJSON/JSONL", variable=fmt_ndjson_var).grid(row=2, column=0, sticky="w", pady=2)
    ttk.Checkbutton(col3, text="XML (.xml)", variable=fmt_xml_var).grid(row=3, column=0, sticky="w", pady=2)
    ttk.Checkbutton(col3, text="YAML/YML", variable=fmt_yaml_var).grid(row=4, column=0, sticky="w", pady=2)
    ttk.Checkbutton(col3, text="TOML (.toml)", variable=fmt_toml_var).grid(row=5, column=0, sticky="w", pady=2)
    ttk.Checkbutton(col3, text="INI/CFG", variable=fmt_ini_cfg_var).grid(row=6, column=0, sticky="w", pady=2)
    ttk.Checkbutton(col3, text="ENV (.env)", variable=fmt_env_var).grid(row=7, column=0, sticky="w", pady=2)
    ttk.Checkbutton(col3, text="PROPERTIES", variable=fmt_properties_var).grid(row=8, column=0, sticky="w", pady=2)
    ttk.Checkbutton(col3, text="SQL (.sql)", variable=fmt_sql_var).grid(row=9, column=0, sticky="w", pady=2)
    ttk.Checkbutton(col3, text="TeX (.tex)", variable=fmt_tex_var).grid(row=10, column=0, sticky="w", pady=2)

    ttk.Separator(lf_formats).grid(row=2, column=0, columnspan=3, sticky="ew", pady=(12, 10))

    other_row = ttk.Frame(lf_formats)
    other_row.grid(row=3, column=0, columnspan=3, sticky="ew")
    other_row.columnconfigure(1, weight=1)

    ttk.Label(other_row, text="Other extensions").grid(row=0, column=0, sticky="w", padx=(0, 10))
    other_entry = ttk.Entry(other_row, textvariable=other_exts_var)
    other_entry.grid(row=0, column=1, sticky="ew")
    ttk.Label(lf_formats, text="Example: .ini .cfg .toml (space or comma separated)", style="Subtitle.TLabel").grid(
        row=4, column=0, columnspan=3, sticky="w", pady=(6, 0)
    )

    # -----------------
    # Words tab
    # -----------------
    tab_words.columnconfigure(0, weight=1)

    lf_words = ttk.Labelframe(tab_words, text="Words to count", padding=(PADX, PADY))
    lf_words.grid(row=0, column=0, sticky="nsew")
    lf_words.columnconfigure(0, weight=1)

    ttk.Label(lf_words, text="Words (comma/space separated)").grid(row=0, column=0, sticky="w")
    words_entry = ttk.Entry(lf_words, textvariable=words_var)
    words_entry.grid(row=1, column=0, sticky="ew", pady=(6, 0))
    ttk.Label(lf_words, text="Example: and, or, hello", style="Subtitle.TLabel").grid(row=2, column=0, sticky="w", pady=(6, 0))

    opts = ttk.Frame(lf_words)
    opts.grid(row=3, column=0, sticky="w", pady=(12, 0))
    ttk.Checkbutton(opts, text="Case sensitive", variable=case_var).grid(row=0, column=0, sticky="w", padx=(0, 12))
    ttk.Checkbutton(opts, text="Include subfolders (recursive)", variable=recursive_var).grid(row=0, column=1, sticky="w")

    # -----------------
    # Run tab
    # -----------------
    tab_run.columnconfigure(0, weight=1)

    lf_run = ttk.Labelframe(tab_run, text="Progress", padding=(PADX, PADY))
    lf_run.grid(row=0, column=0, sticky="nsew")
    lf_run.columnconfigure(0, weight=1)

    progress = ttk.Progressbar(lf_run, variable=progress_var, maximum=100)
    progress.grid(row=0, column=0, sticky="ew")
    ttk.Label(lf_run, textvariable=progress_label_var, style="Subtitle.TLabel").grid(row=1, column=0, sticky="w", pady=(6, 0))
    ttk.Separator(lf_run).grid(row=2, column=0, sticky="ew", pady=(12, 10))

    btn_row = ttk.Frame(lf_run)
    btn_row.grid(row=3, column=0, sticky="w")

    generate_btn = ttk.Button(btn_row, text="Generate report", style="Accent.TButton")
    generate_btn.grid(row=0, column=0, padx=(0, 10))

    cancel_btn = ttk.Button(btn_row, text="Cancel", state="disabled", style="Danger.TButton")
    cancel_btn.grid(row=0, column=1)

    ttk.Label(lf_run, textvariable=status_var).grid(row=4, column=0, sticky="w", pady=(14, 0))

    # Wire up Generate (Cancel command is set inside _run_generate per-run)
    generate_btn.config(
        command=lambda: _run_generate(
            root,
            input_var,
            output_var,
            words_var,
            case_var,
            recursive_var,
            fmt_pdf_var,
            fmt_docx_var,
            fmt_rtf_var,
            fmt_txt_var,
            fmt_md_var,
            fmt_csv_var,
            fmt_json_var,
            fmt_html_var,
            fmt_ipynb_var,
            fmt_xml_var,
            fmt_yaml_var,
            fmt_ini_cfg_var,
            fmt_log_var,
            fmt_sql_var,
            fmt_tex_var,
            fmt_rst_var,
            fmt_properties_var,
            fmt_toml_var,
            fmt_env_var,
            fmt_ndjson_var,
            fmt_xhtml_var,
            fmt_srt_var,
            other_exts_var,
            status_var,
            progress_var,
            progress_label_var,
            generate_btn,
            cancel_btn,
        )
    )


    notebook.select(tab_folders)
    input_entry.focus_set()

    root.mainloop()


if __name__ == "__main__":
    main()