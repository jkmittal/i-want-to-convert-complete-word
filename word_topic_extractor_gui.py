from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from docx import Document

from word_topic_extractor import content_table_entries, extract_topics


HONEYWELL_RED = "#d71920"
INK = "#1f2933"
MUTED = "#6b7280"
PAGE_BG = "#f3f5f7"
PANEL_BG = "#ffffff"
BORDER = "#d9dee5"


class WordTopicExtractorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Honeywell Document Topic Extractor")
        self.geometry("1120x720")
        self.minsize(980, 620)
        self.configure(bg=PAGE_BG)

        self.input_path = None
        self.entries = []
        self.filtered_entry_indexes = []
        self.selected_entry_indexes = set()

        self._configure_styles()
        self._build_ui()

    def _configure_styles(self):
        self.style = ttk.Style(self)
        self.style.theme_use("clam")

        self.style.configure("App.TFrame", background=PAGE_BG)
        self.style.configure("Panel.TFrame", background=PANEL_BG, relief="flat")
        self.style.configure("Header.TFrame", background=HONEYWELL_RED)
        self.style.configure("Title.TLabel", background=HONEYWELL_RED, foreground="white", font=("Segoe UI", 23, "bold"))
        self.style.configure("HeaderMeta.TLabel", background=HONEYWELL_RED, foreground="#ffe8e8", font=("Segoe UI", 10))
        self.style.configure("Section.TLabel", background=PANEL_BG, foreground=INK, font=("Segoe UI", 13, "bold"))
        self.style.configure("Body.TLabel", background=PANEL_BG, foreground=MUTED, font=("Segoe UI", 10))
        self.style.configure("Metric.TLabel", background=PANEL_BG, foreground=INK, font=("Segoe UI", 18, "bold"))
        self.style.configure("MetricName.TLabel", background=PANEL_BG, foreground=MUTED, font=("Segoe UI", 9))
        self.style.configure("Path.TEntry", fieldbackground="#f8fafc")
        self.style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"), padding=(16, 9))
        self.style.configure("Action.TButton", font=("Segoe UI", 10), padding=(12, 8))
        self.style.configure("Treeview", rowheight=31, font=("Segoe UI", 10), background="white", fieldbackground="white")
        self.style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), background="#eef1f5", foreground=INK)
        self.style.map("Treeview", background=[("selected", "#fde8e9")], foreground=[("selected", INK)])

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        self._build_header()
        self._build_workspace()
        self._build_footer()

    def _build_header(self):
        header = ttk.Frame(self, style="Header.TFrame", padding=(24, 18, 24, 18))
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(1, weight=1)

        logo = ttk.Label(header, text="HONEYWELL", style="Title.TLabel")
        logo.grid(row=0, column=0, sticky="w")

        details = ttk.Frame(header, style="Header.TFrame")
        details.grid(row=0, column=1, sticky="e")
        ttk.Label(details, text="Document Intelligence Console", style="HeaderMeta.TLabel").grid(row=0, column=0, sticky="e")
        ttk.Label(details, text="Topic extraction from Word contents table", style="HeaderMeta.TLabel").grid(row=1, column=0, sticky="e")

    def _build_workspace(self):
        workspace = ttk.Frame(self, style="App.TFrame", padding=(22, 20, 22, 10))
        workspace.grid(row=1, column=0, sticky="nsew")
        workspace.columnconfigure(0, weight=3)
        workspace.columnconfigure(1, weight=1)
        workspace.rowconfigure(0, weight=1)

        left_panel = ttk.Frame(workspace, style="Panel.TFrame", padding=(18, 18, 18, 18))
        left_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 14))
        left_panel.columnconfigure(0, weight=1)
        left_panel.rowconfigure(4, weight=1)

        right_panel = ttk.Frame(workspace, style="Panel.TFrame", padding=(18, 18, 18, 18))
        right_panel.grid(row=0, column=1, sticky="nsew")
        right_panel.columnconfigure(0, weight=1)

        self._build_file_picker(left_panel)
        self._build_search_and_table(left_panel)
        self._build_control_panel(right_panel)

    def _build_file_picker(self, parent):
        ttk.Label(parent, text="1. Upload Word File", style="Section.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            parent,
            text="The app reads the contents table, page numbers, and matching heading sections.",
            style="Body.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(3, 12))

        file_row = ttk.Frame(parent, style="Panel.TFrame")
        file_row.grid(row=2, column=0, sticky="ew", pady=(0, 14))
        file_row.columnconfigure(0, weight=1)

        self.file_var = tk.StringVar(value="No file selected")
        ttk.Entry(file_row, textvariable=self.file_var, state="readonly", style="Path.TEntry").grid(
            row=0, column=0, sticky="ew", padx=(0, 10)
        )
        ttk.Button(file_row, text="Upload Word File", style="Primary.TButton", command=self.browse_file).grid(row=0, column=1)

    def _build_search_and_table(self, parent):
        heading_row = ttk.Frame(parent, style="Panel.TFrame")
        heading_row.grid(row=3, column=0, sticky="ew", pady=(0, 10))
        heading_row.columnconfigure(0, weight=1)

        ttk.Label(heading_row, text="2. Search and Select Topics", style="Section.TLabel").grid(row=0, column=0, sticky="w")
        self.status_var = tk.StringVar(value="Waiting for a Word file.")
        ttk.Label(heading_row, textvariable=self.status_var, style="Body.TLabel").grid(row=0, column=1, sticky="e")

        table_area = ttk.Frame(parent, style="Panel.TFrame")
        table_area.grid(row=4, column=0, sticky="nsew")
        table_area.columnconfigure(0, weight=1)
        table_area.rowconfigure(1, weight=1)

        search_row = ttk.Frame(table_area, style="Panel.TFrame")
        search_row.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        search_row.columnconfigure(1, weight=1)

        ttk.Label(search_row, text="Search", style="Body.TLabel").grid(row=0, column=0, sticky="w")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self.apply_search_filter)
        ttk.Entry(search_row, textvariable=self.search_var).grid(row=0, column=1, sticky="ew", padx=(10, 10))
        ttk.Button(search_row, text="Clear", style="Action.TButton", command=self.clear_search).grid(row=0, column=2)

        columns = ("selected", "topic", "page", "source")
        self.topic_tree = ttk.Treeview(table_area, columns=columns, show="headings", selectmode="extended")
        self.topic_tree.heading("selected", text="Selected")
        self.topic_tree.heading("topic", text="Topic")
        self.topic_tree.heading("page", text="Page")
        self.topic_tree.heading("source", text="Read From")
        self.topic_tree.column("selected", width=80, minwidth=70, anchor="center", stretch=False)
        self.topic_tree.column("topic", width=520, minwidth=260, anchor="w")
        self.topic_tree.column("page", width=80, minwidth=60, anchor="center", stretch=False)
        self.topic_tree.column("source", width=140, minwidth=110, anchor="center", stretch=False)
        self.topic_tree.grid(row=1, column=0, sticky="nsew")
        self.topic_tree.bind("<Double-1>", self.toggle_focused_topic)

        scrollbar = ttk.Scrollbar(table_area, orient="vertical", command=self.topic_tree.yview)
        scrollbar.grid(row=1, column=1, sticky="ns")
        self.topic_tree.configure(yscrollcommand=scrollbar.set)

        action_row = ttk.Frame(table_area, style="Panel.TFrame")
        action_row.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(10, 0))

        ttk.Button(action_row, text="Add Highlighted", style="Action.TButton", command=self.add_highlighted_topics).grid(
            row=0, column=0, padx=(0, 8)
        )
        ttk.Button(action_row, text="Remove Highlighted", style="Action.TButton", command=self.remove_highlighted_topics).grid(
            row=0, column=1, padx=(0, 8)
        )
        ttk.Button(action_row, text="Select Visible", style="Action.TButton", command=self.select_visible_topics).grid(
            row=0, column=2, padx=(0, 8)
        )
        ttk.Button(action_row, text="Clear Selected", style="Action.TButton", command=self.clear_selected_topics).grid(row=0, column=3)

    def _build_control_panel(self, parent):
        ttk.Label(parent, text="Selection Summary", style="Section.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(parent, text="Create a focused Word file from selected topics.", style="Body.TLabel").grid(
            row=1, column=0, sticky="w", pady=(3, 16)
        )

        metrics = ttk.Frame(parent, style="Panel.TFrame")
        metrics.grid(row=2, column=0, sticky="ew", pady=(0, 18))
        metrics.columnconfigure((0, 1), weight=1)

        self.total_topics_var = tk.StringVar(value="0")
        self.selected_topics_var = tk.StringVar(value="0")
        ttk.Label(metrics, textvariable=self.total_topics_var, style="Metric.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(metrics, text="Topics found", style="MetricName.TLabel").grid(row=1, column=0, sticky="w")
        ttk.Label(metrics, textvariable=self.selected_topics_var, style="Metric.TLabel").grid(row=0, column=1, sticky="w")
        ttk.Label(metrics, text="Topics selected", style="MetricName.TLabel").grid(row=1, column=1, sticky="w")

        ttk.Label(parent, text="3. Output File", style="Section.TLabel").grid(row=3, column=0, sticky="w")
        self.output_var = tk.StringVar(value=str(Path("outputs") / "selected_topics.docx"))
        ttk.Entry(parent, textvariable=self.output_var).grid(row=4, column=0, sticky="ew", pady=(8, 8))
        ttk.Button(parent, text="Choose Save Location", style="Action.TButton", command=self.choose_output).grid(
            row=5, column=0, sticky="ew", pady=(0, 18)
        )

        ttk.Button(parent, text="Create New Word File", style="Primary.TButton", command=self.create_word_file).grid(
            row=6, column=0, sticky="ew"
        )

        note = (
            "Workflow\n"
            "1. Upload a .docx file.\n"
            "2. Search the contents table.\n"
            "3. Highlight rows and add them.\n"
            "4. Export all selected topics."
        )
        ttk.Label(parent, text=note, style="Body.TLabel", justify="left").grid(row=7, column=0, sticky="w", pady=(18, 0))

    def _build_footer(self):
        footer = ttk.Frame(self, style="App.TFrame", padding=(22, 0, 22, 16))
        footer.grid(row=2, column=0, sticky="ew")
        footer.columnconfigure(0, weight=1)

        self.footer_var = tk.StringVar(value="Honeywell-styled desktop interface | Word Topic Extractor")
        ttk.Label(footer, textvariable=self.footer_var, background=PAGE_BG, foreground=MUTED, font=("Segoe UI", 9)).grid(
            row=0, column=0, sticky="w"
        )

    def browse_file(self):
        path = filedialog.askopenfilename(
            title="Select Word file",
            filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
        )
        if not path:
            return

        self.input_path = Path(path)
        self.file_var.set(str(self.input_path))
        self.load_topics()

    def load_topics(self):
        self.topic_tree.delete(*self.topic_tree.get_children())
        self.selected_entry_indexes.clear()

        try:
            document = Document(self.input_path)
            self.entries = content_table_entries(document)
        except Exception as error:
            messagebox.showerror("Could not read Word file", str(error))
            self.entries = []
            self.refresh_metrics()
            return

        if not self.entries:
            self.status_var.set("No contents table or heading topics were found.")
            self.refresh_metrics()
            return

        self.search_var.set("")
        self.apply_search_filter()
        self.refresh_metrics()

    def clear_search(self):
        self.search_var.set("")

    def apply_search_filter(self, *_):
        self.topic_tree.delete(*self.topic_tree.get_children())
        query = " ".join(self.search_var.get().casefold().split())
        self.filtered_entry_indexes = []

        for index, entry in enumerate(self.entries):
            searchable = f"{entry['title']} {entry.get('page', '')}".casefold()
            if query and query not in searchable:
                continue

            self.filtered_entry_indexes.append(index)
            self.insert_topic_row(index, entry)

        total = len(self.entries)
        shown = len(self.filtered_entry_indexes)
        if query:
            self.status_var.set(f"Showing {shown} of {total} topics.")
        elif total:
            source_label = "contents table" if self.entries[0]["source"] == "toc" else "headings"
            self.status_var.set(f"Found {total} topics from {source_label}.")
        else:
            self.status_var.set("Waiting for a Word file.")

        self.refresh_metrics()

    def insert_topic_row(self, index, entry):
        mark = "Yes" if index in self.selected_entry_indexes else ""
        indent = "    " * max(entry["level"] - 1, 0)
        source = "Contents table" if entry["source"] == "toc" else "Heading"
        self.topic_tree.insert(
            "",
            "end",
            iid=str(index),
            values=(mark, f"{indent}{entry['title']}", entry["page"] or "-", source),
        )

    def selected_tree_indexes(self):
        return [int(item_id) for item_id in self.topic_tree.selection()]

    def add_highlighted_topics(self):
        highlighted = self.selected_tree_indexes()
        if not highlighted:
            messagebox.showinfo("Select rows", "Highlight one or more rows first, then click Add Highlighted.")
            return

        self.selected_entry_indexes.update(highlighted)
        self.apply_search_filter()

    def remove_highlighted_topics(self):
        for index in self.selected_tree_indexes():
            self.selected_entry_indexes.discard(index)
        self.apply_search_filter()

    def select_visible_topics(self):
        self.selected_entry_indexes.update(self.filtered_entry_indexes)
        self.apply_search_filter()

    def clear_selected_topics(self):
        self.selected_entry_indexes.clear()
        self.apply_search_filter()

    def toggle_focused_topic(self, _event=None):
        focused = self.topic_tree.focus()
        if not focused:
            return

        index = int(focused)
        if index in self.selected_entry_indexes:
            self.selected_entry_indexes.remove(index)
        else:
            self.selected_entry_indexes.add(index)
        self.apply_search_filter()

    def refresh_metrics(self):
        self.total_topics_var.set(str(len(self.entries)))
        self.selected_topics_var.set(str(len(self.selected_entry_indexes)))

    def choose_output(self):
        path = filedialog.asksaveasfilename(
            title="Save selected topics as",
            defaultextension=".docx",
            filetypes=[("Word documents", "*.docx")],
            initialfile=Path(self.output_var.get()).name,
        )
        if path:
            self.output_var.set(path)

    def create_word_file(self):
        if not self.input_path:
            messagebox.showwarning("Select a file", "Please upload a Word file first.")
            return

        if not self.selected_entry_indexes:
            messagebox.showwarning("Select topics", "Please add one or more topics before creating the Word file.")
            return

        selected_entries = [self.entries[index] for index in sorted(self.selected_entry_indexes)]
        selected_titles = [entry["title"] for entry in selected_entries]
        output_path = Path(self.output_var.get())
        if output_path.suffix.lower() != ".docx":
            output_path = output_path.with_suffix(".docx")

        try:
            result = extract_topics(self.input_path, selected_titles, output_path)
        except Exception as error:
            messagebox.showerror("Could not create Word file", str(error))
            return

        messagebox.showinfo(
            "Word file created",
            f"Created a new Word file with {len(selected_entries)} selected topic(s):\n\n{result}",
        )


def main():
    app = WordTopicExtractorApp()
    app.mainloop()


if __name__ == "__main__":
    main()
