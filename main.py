import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import glob
import sys
from docx import Document
from docxcompose.composer import Composer

# --- KONFIGURATION ---
# Trage hier die Dateinamen ein, die immer aktiv und gesperrt sein sollen.
# (Exakte Dateinamen aus dem 'modules'-Ordner)
MANDATORY_FILES = [
    "0.0.0 - Deckblatt.docx",
    "1.0.0 - Lage der Baustelle, Lageplanskizze.docx",
    "2.1.0 - Arbeitszeit.docx",
    "2.2.0 - Dauer der Gleissperrungen, gesperrte Gleise, Weichen.docx",
    "3.0.0 - Geschwindigkeitseinschränkungen und andere Besonderheiten für Zugfahrten.docx",
    "4.0.0 - Zuständige Berechtigte.docx",
    "4.1.0 - Fahrdienstleiter-Weichenwärter-Zugleiter-BözM.docx",
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

# Definieren Sie hier Ihre Kategorien als "Button-Text": ["Präfix1", "Präfix2", ...]
CATEGORIES = {
    "Oberleitung": ["2.3.1", "2.3.2", "2.3.3", "2.3.4", "2.3.5", "2.3.6", "4.3.0", "5.3.20"],
    "Baugleis": ["3.0.", "3.1.", "3.2.", "5.1.11", "5.3.14", "5.3.15", "5.3.16", "5.3.17", "5.3.18", "5.3.21"],
    "BÜ": ["5.1.22", "5.1.23", "5.1.24", "5.1.25", "5.1.26", "5.1.27", "5.1.28", "5.3.11"],
    "Lfst (Pkt. 3)": ["3.1.", "3.2."],
    "VorGWB": ["5.1.20", "5.1.21"],
    "UntArb": ["5.3.4", "5.3.6"]
}


# ---------------------


class WordMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("BetraBuilder - Word Merger (Spalten-Layout)")
        self.root.geometry("1000x600")  # Breiter/Höher für Spalten

        # --- Pfade ---
        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))

        self.modules_dir = os.path.join(base_path, "modules")
        self.output_dir = os.path.join(base_path, "output")

        self.checkbox_items = []

        # --- Haupt-Frame ---
        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        info_label = ttk.Label(main_frame, text=f"Module aus: '{self.modules_dir}'", font=("-default-", 9, "italic"))
        info_label.pack(anchor="w")

        # --- Checkbox-Liste (mit X- und Y-Scrollbars) ---
        list_frame = ttk.Frame(main_frame, padding=(0, 10, 0, 0))
        list_frame.pack(fill=tk.BOTH, expand=True)

        # X Scrollbar (unten)
        x_scrollbar = ttk.Scrollbar(list_frame, orient="horizontal")
        x_scrollbar.pack(side="bottom", fill="x")

        # Y Scrollbar (rechts)
        y_scrollbar = ttk.Scrollbar(list_frame, orient="vertical")
        y_scrollbar.pack(side="right", fill="y")

        # Canvas (füllt den Rest)
        self.canvas = tk.Canvas(list_frame)
        self.canvas.pack(side="left", fill="both", expand=True)

        # Scrollbars mit Canvas verknüpfen
        self.canvas.configure(yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set)
        x_scrollbar.configure(command=self.canvas.xview)
        y_scrollbar.configure(command=self.canvas.yview)

        # Das Frame, das gescrollt wird
        self.scrollable_frame = ttk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")

        # Scroll-Region an die Größe des Frames anpassen
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )

        # --- Kategorie-Buttons ---
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
            # Buttons einfach nebeneinander packen
            btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2, pady=2)

        # --- Haupt-Buttons (Reset/Start) ---
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))

        self.reset_button = ttk.Button(button_frame, text="Auswahl zurücksetzen", command=self.reset_selection)
        self.reset_button.pack(side=tk.LEFT)

        self.start_button = ttk.Button(button_frame, text="Ausgewählte Dateien zusammenfügen", command=self.start_merge)
        self.start_button.pack(side=tk.RIGHT)
        self.start_button["state"] = "disabled"

        self.load_files()

    def load_files(self):
        if not os.path.isdir(self.modules_dir):
            messagebox.showerror("Fehler", f"Der Ordner '{self.modules_dir}' wurde nicht gefunden.")
            self.root.quit()
            return

        # Alte Einträge löschen
        for item in self.checkbox_items:
            item["widget"].destroy()
        self.checkbox_items.clear()

        # Kinder des scrollbaren Frames löschen (alte Spalten)
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        search_path = os.path.join(self.modules_dir, "*.docx")
        # WICHTIG: Sortierte Liste aller Dateien (definiert die Merge-Reihenfolge)
        file_paths = glob.glob(search_path)
        file_paths.sort()

        if not file_paths:
            messagebox.showinfo("Keine Dateien", "Keine .docx Dateien im 'modules'-Ordner gefunden.")
            return

        # --- Spalten-Logik ---
        column_frames = {}  # Speichert die Spalten-Frames (z.B. "0": frame0, "1": frame1)

        for file_path in file_paths:
            filename = os.path.basename(file_path)

            # Schlüssel für die Spalte extrahieren (z.B. "0", "1", "2", "5")
            try:
                chapter_key = filename.split('.')[0]
            except IndexError:
                chapter_key = "Unsortiert"  # Fallback

            # Prüfen, ob für dieses Kapitel schon ein Frame existiert
            if chapter_key not in column_frames:
                # Neue Spalte (Frame) erstellen
                col_frame = ttk.Frame(self.scrollable_frame, padding=5, borderwidth=1, relief="sunken")
                col_frame.pack(side=tk.LEFT, fill=tk.Y, anchor="n", padx=5, pady=5)

                # Titel für die Spalte
                title = f"Kap. {chapter_key}.x" if chapter_key.isdigit() else chapter_key
                ttk.Label(col_frame, text=title, font=("-default-", 10, "bold")).pack(pady=(0, 5))

                # Frame im Dictionary speichern
                column_frames[chapter_key] = col_frame
            else:
                # Existierendes Frame holen
                col_frame = column_frames[chapter_key]

            # --- Checkbox erstellen (Logik von vorher) ---
            is_mandatory = filename in MANDATORY_FILES
            var = tk.BooleanVar(value=is_mandatory)
            cb_state = "disabled" if is_mandatory else "normal"

            cb = ttk.Checkbutton(col_frame, text=filename, variable=var, state=cb_state)
            cb.pack(anchor="w", padx=10, pady=2)

            # Wichtig: Wir speichern die Items in der sortierten Reihenfolge (file_paths)
            # NICHT in der Spaltenreihenfolge.
            self.checkbox_items.append({
                "var": var,
                "path": file_path,
                "filename": filename,
                "widget": cb,
                "is_mandatory": is_mandatory
            })

        self.start_button["state"] = "normal"

    def reset_selection(self):
        """Setzt alle Checkboxen auf den Startzustand zurück."""
        for item in self.checkbox_items:
            if not item["is_mandatory"]:
                item["var"].set(False)
            else:
                item["var"].set(True)

    def toggle_category(self, prefixes):
        """
        Schaltet alle (nicht-obligatorischen) Dateien um,
        die mit einem der Präfixe beginnen.
        """
        items_in_category = []
        for item in self.checkbox_items:
            if item["is_mandatory"]:
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

    def start_merge(self):
        # Sammelt die Dateien in der *sortierten Listenreihenfolge*
        selected_files = []
        for item in self.checkbox_items:
            if item["var"].get():
                selected_files.append(item["path"])

        if not selected_files:
            messagebox.showwarning("Keine Auswahl", "Es ist keine Datei ausgewählt.")
            return

        # Output-Ordner und Speicherpfad
        try:
            os.makedirs(self.output_dir, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Fehler", f"Konnte den Output-Ordner nicht erstellen:\n{e}")
            return

        save_path = filedialog.asksaveasfilename(
            initialdir=self.output_dir,
            initialfile="Betra_Zusammenstellung.docx",
            defaultextension=".docx",
            filetypes=[("Word-Dokumente", "*.docx"), ("Alle Dateien", "*.*")],
            title="Zieldatei speichern unter..."
        )

        if not save_path:
            return

        # Merge-Prozess
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
        """
        Fügt Word-Dokumente mit docxcompose zusammen (ohne Seitenumbrüche).
        """
        if not file_paths:
            return

        master_doc = Document(file_paths[0])
        composer = Composer(master_doc)

        if len(file_paths) > 1:
            for file_path in file_paths[1:]:
                doc_to_append = Document(file_path)
                composer.append(doc_to_append)

        composer.save(save_path)


# --- Hauptprogramm starten ---
if __name__ == "__main__":
    root = tk.Tk()
    app = WordMergerApp(root)
    root.mainloop()