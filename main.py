import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import glob
import sys
import re
from docx import Document
from docxcompose.composer import Composer

# --- KONFIGURATION ---
MANDATORY_FILES = [
    "1.0.0 - Lage der Baustelle, Lageplanskizze.docx",
    "2.1.0 - Arbeitszeit.docx",
    "2.2.0 - Dauer der Gleissperrungen, gesperrte Gleise, Weichen.docx",
    "3.0.0 - Geschwindigkeitseinschränkungen und andere Besonderheiten für Zugfahrten.docx",
    "4.0.0 - Zuständige Berechtigte.docx",
    "4.1.0 - Fahrdienstleiter Weichenwärter Zugleiter BözM.docx",
    "4.2.0 - Technischer Berechtigter - UV-Berechtigter.docx",
    "5.0.0 - Betriebliche Regelungen.docx",
    "5.1.0 - Regelungen für die Sicherung des Bahnbetriebes.docx",
    "5.1.1 - Grundsatz.docx",
    "5.1.2 - Fernmündliche Aufträge und Meldungen.docx",
    "5.1.3 - Beginn der Arbeiten.docx",
    "5.2.0 - Regelungen für die Durchführung des Bahnbetriebes.docx",
    "5.2.15 - Außergewöhnliche Sendungen, Fahrzeuge, Züge.docx",
    "5.3.0 - Regelungen für das gesperrte Gleis - Baugleis - unterbrochene Arbeitszeit - Ortsstellbereiche.docx",
    "5.3.1 - Infrastrukturparameter.docx",
    "5.4.0 - Regelungen für den Einsatz von Schienenfahrzeugen, Maschinen und Geräten und deren besonderen Einsatzbedingungen.docx",
    "5.4.1 - Grundsätze.docx",
    "6.0.0 - Sicherung der Beschäftigten.docx",
    "6.1.0 - Arbeiten im Gleisbereich.docx",
    "6.2.0 - Arbeiten an oder in der Nähe von aktiven Teilen der Oberleitungsanlage.docx",
    "7.0.0 - Verantwortliche.docx",
    "8.0.0 - Angaben zur Bautechnologie-Bauablauf-Baustellenlogistik.docx",
    "9.0.0 - Sonstige Angaben.docx",
    "9.1.0 - Anlagen - Zugestimmt - Verteiler.docx"
]

CATEGORIES = {
    "Oberleitung": ["2.3.", "4.3.0", "5.3.20"],
    "Baugleis": ["5.1.11", "5.3.14", "5.3.15", "5.3.16", "5.3.17", "5.3.18", "5.3.21"],
    "BÜ": ["5.1.22", "5.1.23", "5.1.24", "5.1.25", "5.1.26", "5.1.27", "5.1.28", "5.3.11"],
    "Lfst (Pkt. 3)": ["3.1.", "3.2."],
    "VorGWB": ["5.1.20", "5.1.21"],
    "Alle": ["0.", "1.", "2.", "3.", "4.", "5.", "6.", "7.", "8.", "9."]
}


COLUMN_LAYOUT = {
    "0": 0,
    "1": 0,
    "2": 0,
    "3": 0,
    "4": 0,
    "5.0": 1,
    "5.1": 1,
    "5.2": 2,
    "5.3": 3,
    "5.4": 4,
    "6": 4,
    "7": 4,
    "8": 4,
    "9": 4,
    "10": 4,
    "Unsortiert": 5
}
NUM_MAIN_COLUMNS = 6
version = "0.2a"

# ---------------------

def natural_sort_key(s):
    filename = os.path.basename(s)
    return [int(c) if c.isdigit() else c.lower() for c in re.split('([0-9]+)', filename)]


class WordMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Betra Builder v" + version)
        self.root.geometry("1450x700")

        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))

        try:
            icon_path = os.path.join(base_path, "icon.ico")
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
            else:
                raise FileNotFoundError
        except (FileNotFoundError, tk.TclError):
            try:
                png_path = os.path.join(base_path, "icon.png")
                if os.path.exists(png_path):
                    png_icon = tk.PhotoImage(file=png_path)
                    self.root.iconphoto(False, png_icon)
            except Exception as e:
                print(f"Icon konnte nicht geladen werden: {e}")

        self.modules_dir = os.path.join(base_path, "modules")
        self.output_dir = os.path.join(base_path, "output")

        self.checkbox_items = []
        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        info_label = ttk.Label(main_frame, text=f"Module aus: '{self.modules_dir}'", font=("-default-", 9, "italic"))
        info_label.pack(anchor="w")

        list_frame = ttk.Frame(main_frame, padding=(0, 10, 0, 0))
        list_frame.pack(fill=tk.BOTH, expand=True)

        x_scrollbar = ttk.Scrollbar(list_frame, orient="horizontal")
        x_scrollbar.pack(side="bottom", fill="x")
        y_scrollbar = ttk.Scrollbar(list_frame, orient="vertical")
        y_scrollbar.pack(side="right", fill="y")

        self.canvas = tk.Canvas(list_frame)
        self.canvas.pack(side="left", fill="both", expand=True)

        self.canvas.configure(yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set)
        x_scrollbar.configure(command=self.canvas.xview)
        y_scrollbar.configure(command=self.canvas.yview)

        self.scrollable_frame = ttk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )

        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind("<Button-4>", self._on_mousewheel)
        self.canvas.bind("<Button-5>", self._on_mousewheel)
        self.scrollable_frame.bind("<MouseWheel>", self._on_mousewheel)
        self.scrollable_frame.bind("<Button-4>", self._on_mousewheel)
        self.scrollable_frame.bind("<Button-5>", self._on_mousewheel)

        category_frame = ttk.Frame(main_frame)
        category_frame.pack(fill=tk.X, pady=(10, 5))

        category_label = ttk.Label(category_frame, text="Kategorien (An/Aus):")
        category_label.pack(fill=tk.X, pady=(0, 4))

        btn_container = ttk.Frame(category_frame)
        btn_container.pack(fill=tk.X)

        for text, prefixes in CATEGORIES.items():
            btn = ttk.Button(btn_container,
                             text=text,
                             command=lambda p=prefixes: self.toggle_category(p))
            btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2, pady=2)

        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))

        self.help_button = ttk.Button(button_frame, text="Anleitung", command=self.show_help)
        self.help_button.pack(side=tk.LEFT)

        self.contact_button = ttk.Button(button_frame, text="Kontakt", command=self.show_contact)
        self.contact_button.pack(side=tk.LEFT, padx=5)

        self.reset_button = ttk.Button(button_frame, text="Auswahl zurücksetzen", command=self.reset_selection)
        self.reset_button.pack(side=tk.LEFT, padx=(5, 0))  # Kleiner Abstand

        self.start_button = ttk.Button(button_frame, text="Ausgewählte Dateien zusammenfügen", command=self.start_merge)
        self.start_button.pack(side=tk.RIGHT)
        self.start_button["state"] = "disabled"

        self.load_files()

    def _on_mousewheel(self, event):
        if sys.platform == "linux":
            if event.num == 4:
                self.canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                self.canvas.yview_scroll(1, "units")
        else:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _get_layout_key(self, filename):
        parts = filename.split('.')
        if not parts:
            return "Unsortiert"

        main_chapter = parts[0]

        if main_chapter == "5":
            if len(parts) > 1:
                key = f"{parts[0]}.{parts[1]}"
                if key in COLUMN_LAYOUT:
                    return key

        if main_chapter in COLUMN_LAYOUT:
            return main_chapter

        return "Unsortiert"

    def load_files(self):
        if not os.path.isdir(self.modules_dir):
            messagebox.showerror("Fehler", f"Der Ordner '{self.modules_dir}' wurde nicht gefunden.")
            self.root.quit()
            return

        for item in self.checkbox_items:
            item["widget"].destroy()
        self.checkbox_items.clear()

        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        search_path = os.path.join(self.modules_dir, "*.docx")
        file_paths = glob.glob(search_path)
        file_paths.sort(key=natural_sort_key)

        if not file_paths:
            messagebox.showinfo("Keine Dateien", "Keine .docx Dateien im 'modules'-Ordner gefunden.")
            return

        wrap_length_pixels = 220
        main_columns = []
        for i in range(NUM_MAIN_COLUMNS):
            main_col_frame = ttk.Frame(self.scrollable_frame)
            main_col_frame.grid(row=0, column=i, sticky="nw", padx=5)
            main_col_frame.bind("<MouseWheel>", self._on_mousewheel)
            main_columns.append(main_col_frame)

        group_frames = {}

        for file_path in file_paths:
            filename = os.path.basename(file_path)
            layout_key = self._get_layout_key(filename)

            if layout_key not in group_frames:
                main_col_index = COLUMN_LAYOUT.get(layout_key, NUM_MAIN_COLUMNS - 1)
                parent_frame = main_columns[main_col_index]

                group_frame = ttk.Frame(parent_frame, padding=5, borderwidth=1, relief="sunken")
                group_frame.pack(side=tk.TOP, fill="x", anchor="n", pady=5)

                title = f"Punkt {layout_key}" if layout_key != "Unsortiert" else "Unsortiert"
                title_label = ttk.Label(group_frame, text=title, font=("-default-", 10, "bold"),
                                        wraplength=wrap_length_pixels, justify=tk.LEFT)
                title_label.pack(pady=(0, 5), anchor="w")

                group_frames[layout_key] = group_frame

                group_frame.bind("<MouseWheel>", self._on_mousewheel)
                group_frame.bind("<Button-4>", self._on_mousewheel)
                group_frame.bind("<Button-5>", self._on_mousewheel)
                title_label.bind("<MouseWheel>", self._on_mousewheel)
                title_label.bind("<Button-4>", self._on_mousewheel)
                title_label.bind("<Button-5>", self._on_mousewheel)
            else:
                group_frame = group_frames[layout_key]

            is_mandatory = filename in MANDATORY_FILES
            var = tk.BooleanVar(value=is_mandatory)
            cb_state = "disabled" if is_mandatory else "normal"

            item_frame = ttk.Frame(group_frame)
            item_frame.pack(fill='x', anchor="w", pady=1)

            cb = ttk.Checkbutton(item_frame, variable=var, state=cb_state)
            cb.pack(side=tk.LEFT, anchor="n", padx=(0, 5))

            display_name = filename
            if filename.endswith(".docx"):
                display_name = filename[:-5]

            label = ttk.Label(item_frame, text=display_name, wraplength=wrap_length_pixels, justify=tk.LEFT)
            label.pack(side=tk.LEFT, fill='x', expand=True)

            def on_label_click(event, cb_widget=cb, cb_var=var):
                if cb_widget.cget("state") == "normal":
                    cb_var.set(not cb_var.get())

            label.bind("<Button-1>", on_label_click)

            item_frame.bind("<MouseWheel>", self._on_mousewheel)
            item_frame.bind("<Button-4>", self._on_mousewheel)
            item_frame.bind("<Button-5>", self._on_mousewheel)
            cb.bind("<MouseWheel>", self._on_mousewheel)
            cb.bind("<Button-4>", self._on_mousewheel)
            cb.bind("<Button-5>", self._on_mousewheel)
            label.bind("<MouseWheel>", self._on_mousewheel)
            label.bind("<Button-4>", self._on_mousewheel)
            label.bind("<Button-5>", self._on_mousewheel)

            self.checkbox_items.append({
                "var": var,
                "path": file_path,
                "filename": filename,
                "widget": cb,
                "is_mandatory": is_mandatory
            })

        self.start_button["state"] = "normal"

    def reset_selection(self):
        for item in self.checkbox_items:
            if not item["is_mandatory"]:
                item["var"].set(False)
            else:
                item["var"].set(True)

    def toggle_category(self, prefixes):
        items_in_category = []
        for item in self.checkbox_items:
            if item["is_mandatory"] or item["widget"].cget("state") == "disabled":
                continue

            for prefix in prefixes:
                if item["filename"].startswith(prefix):
                    items_in_category.append(item)
                    break

        if not items_in_category:
            return

        is_anything_deselected = any(not item["var"].get() for item in items_in_category)
        new_state = is_anything_deselected

        for item in items_in_category:
            item["var"].set(new_state)

    def show_help(self):
        anleitung_text = (
            "Anleitung Betra Builder\n\n"
            "1. Pflichtmodule sind bereits ausgewählt und können nicht abgewählt werden.\n\n"
            "2. Wählen Sie optionale Module aus, indem Sie die Haken setzen (Klick auf den Haken oder den Text).\n\n"
            "3. Nutzen Sie die 'Kategorien'-Buttons, um gängige Modul-Gruppen schnell an- oder abzuwählen.\n\n"
            "4. Mit 'Auswahl zurücksetzen' werden alle optionalen Module abgewählt.\n\n"
            "5. Klicken Sie auf 'Ausgewählte Dateien zusammenfügen', wählen Sie einen Speicherort und die Zieldatei wird erstellt."
        )
        messagebox.showinfo("Anleitung", anleitung_text)

    def show_contact(self):
        kontakt_text = (
            "Kontakt & Support\n\n"
            "Bei Fragen, Problemen, Ideen oder Vorschläge mit dem Betra Builder:\n\n"
            "Name: Dennis Heinze, I.IA-W-N-HA-B\n"
            "E-Mail: dennis.heinze@deutschebahn.com\n"
            "Telefon (dienstlich): 0152 33114237\n"
            "Version: " + version + "\n\n"
        )
        messagebox.showinfo("Kontakt", kontakt_text)

    def start_merge(self):
        selected_files = []
        for item in self.checkbox_items:
            if item["var"].get():
                selected_files.append(item["path"])

        if not selected_files:
            messagebox.showwarning("Keine Auswahl", "Es ist keine Datei ausgewählt.")
            return

        try:
            os.makedirs(self.output_dir, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Fehler", f"Konnte den Output-Ordner nicht erstellen:\n{e}")
            return

        save_path = filedialog.asksaveasfilename(
            initialdir=self.output_dir,
            initialfile="Betra F33 XXXX-26.docx",
            defaultextension=".docx",
            filetypes=[("Word-Dokumente", "*.docx"), ("Alle Dateien", "*.*")],
            title="Zieldatei speichern unter..."
        )

        if not save_path:
            return

        try:
            self.start_button.config(text="Arbeite...", state="disabled")
            self.root.update_idletasks()

            self.merge_documents(selected_files, save_path)

            messagebox.showinfo("Erfolg", f"Dateien erfolgreich zusammengefügt!\nGespeichert als: {save_path}")
        except Exception as e:
            messagebox.showerror("Fehler", f"Ein Fehler ist aufgetreten:\n{e}\n\n"
                                           f"Hinweis: Stellen Sie sicher, dass die Zieldatei (falls sie existiert) geschlossen ist.")
        finally:
            self.start_button.config(text="Ausgewählte Dateien zusammenfügen", state="normal")

    def merge_documents(self, file_paths, save_path):
        if not file_paths:
            return

        if not os.path.exists(file_paths[0]):
            raise FileNotFoundError(f"Die Basis-Datei konnte nicht gefunden werden: {file_paths[0]}")

        master_doc = Document(file_paths[0])
        composer = Composer(master_doc)

        if len(file_paths) > 1:
            for file_path in file_paths[1:]:
                if not os.path.exists(file_path):
                    print(f"Warnung: Datei übersprungen (nicht gefunden): {file_path}")
                    continue
                try:
                    doc_to_append = Document(file_path)
                    composer.append(doc_to_append)
                except Exception as e_inner:
                    print(f"Fehler beim Anhängen von {file_path}: {e_inner}")
                    pass

        composer.save(save_path)


if __name__ == "__main__":
    root = tk.Tk()
    app = WordMergerApp(root)
    root.mainloop()