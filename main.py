import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import os
import glob
import sys
import re
import configparser
from datetime import datetime  # Neu für AEL-Datum
import openpyxl  # Neu für Excel-Export (muss mit 'pip install openpyxl' installiert werden)
from docx import Document
from docxcompose.composer import Composer

# --- CONFIGURATION ---
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

NUM_PRESETS = 5

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
    "Unsorted": 5
}
NUM_MAIN_COLUMNS = 6
APP_VERSION = "1.0a"  # Version erhöht


# ---------------------

def natural_sort_key(s):
    """Sorts strings 'naturally' (e.g., 1, 2, 10) instead of alphabetically (1, 10, 2)."""
    filename = os.path.basename(s)
    return [int(c) if c.isdigit() else c.lower() for c in re.split('([0-9]+)', filename)]


class FileNameDialog(simpledialog.Dialog):
    """Custom dialog to ask for doc type, serial number, and AEL option."""

    def __init__(self, parent, title):
        self.result = None
        super().__init__(parent, title)

    def body(self, frame):
        type_frame = ttk.Frame(frame)
        ttk.Label(type_frame, text="Art:").pack(side=tk.LEFT, padx=5)

        self.doc_type_var = tk.StringVar(value="Betra")
        rb1 = ttk.Radiobutton(type_frame, text="Betra", variable=self.doc_type_var, value="Betra")
        rb1.pack(side=tk.LEFT, padx=5)
        rb2 = ttk.Radiobutton(type_frame, text="BA", variable=self.doc_type_var, value="BA")
        rb2.pack(side=tk.LEFT, padx=5)
        type_frame.pack(pady=5)

        num_frame = ttk.Frame(frame)
        ttk.Label(num_frame, text="Laufende Nummer (YYYY):").pack(side=tk.LEFT, padx=5)

        self.entry_var = tk.StringVar()
        self.entry_widget = ttk.Entry(num_frame, textvariable=self.entry_var, width=10)
        self.entry_widget.pack(side=tk.LEFT)
        num_frame.pack(pady=5)

        # --- NEUE AEL-CHECKBOX ---
        ael_frame = ttk.Frame(frame)
        self.ael_var = tk.BooleanVar(value=False)
        ael_check = ttk.Checkbutton(ael_frame, text="AEL-Verrechnung", variable=self.ael_var)
        ael_check.pack(side=tk.LEFT, padx=5, pady=5)
        ael_frame.pack()
        # --- ENDE NEU ---

        return self.entry_widget

    def validate(self):
        serial_num = self.entry_var.get().strip()
        if not serial_num:
            messagebox.showwarning("Eingabe fehlt", "Bitte eine laufende Nummer eingeben.", parent=self)
            return 0
        return 1

    def apply(self):
        # Result ist jetzt ein 3-Tuple
        self.result = (
            self.doc_type_var.get(),
            self.entry_var.get().strip(),
            self.ael_var.get()
        )


class InitialConfigDialog(simpledialog.Dialog):
    """Custom dialog for the first-time setup (RB, Network). Year is fixed."""

    def __init__(self, parent, title, network_data):
        self.network_data = network_data
        self.result = None
        super().__init__(parent, title)

    def body(self, frame):
        self.rb_var = tk.StringVar()
        self.network_var = tk.StringVar()

        # 1. Regionalbereich (RB)
        rb_frame = ttk.Frame(frame)
        ttk.Label(rb_frame, text="Regionalbereich (RB):").pack(side=tk.LEFT, padx=5, pady=5)
        self.rb_combo = ttk.Combobox(rb_frame, textvariable=self.rb_var, state="readonly", width=30)
        self.rb_combo['values'] = sorted(list(self.network_data.keys()))
        self.rb_combo.pack(side=tk.LEFT, padx=5, pady=5)
        rb_frame.pack()

        # 2. Network
        network_frame = ttk.Frame(frame)
        ttk.Label(network_frame, text="Netz auswählen:").pack(side=tk.LEFT, padx=5, pady=5)
        self.network_combo = ttk.Combobox(network_frame, textvariable=self.network_var, state="disabled", width=30)
        self.network_combo.pack(side=tk.LEFT, padx=5, pady=5)
        network_frame.pack()

        # 3. Year
        year_info_label = ttk.Label(frame, text="Das Jahr ist fest auf 2026 eingestellt (Module 2026).",
                                    font=("-default-", 9, "italic"))
        year_info_label.pack(pady=(10, 0))

        # Bind event
        self.rb_combo.bind("<<ComboboxSelected>>", self._on_rb_selected)

        return self.rb_combo  # Initial focus

    def _on_rb_selected(self, event=None):
        """Called when the RB combobox selection changes."""
        selected_rb = self.rb_var.get()
        networks = self.network_data.get(selected_rb, {})

        network_display_list = []
        for code, name in networks.items():
            network_display_list.append(f"{code} - {name}")

        if network_display_list:
            self.network_combo['values'] = sorted(network_display_list)
            self.network_combo.set(network_display_list[0])
            self.network_combo.config(state="readonly")
        else:
            self.network_combo.set("")
            self.network_combo.config(state="disabled")

    def validate(self):
        if not self.rb_var.get():
            messagebox.showwarning("Eingabe fehlt", "Bitte einen Regionalbereich auswählen.", parent=self)
            return 0
        if not self.network_var.get():
            messagebox.showwarning("Eingabe fehlt", "Bitte ein Netz auswählen.", parent=self)
            return 0
        return 1

    def apply(self):
        """Parses the result when OK is clicked."""
        try:
            full_network_string = self.network_var.get()
            parts = full_network_string.split(" - ", 1)
            code_full = parts[0].strip()
            name = parts[1].strip()

            self.result = (code_full, name)
        except Exception as e:
            print(f"Dialog apply error: {e}")
            messagebox.showerror("Fehler", "Auswahl konnte nicht verarbeitet werden.", parent=self)


class WordMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Betra Builder v" + APP_VERSION)
        self.root.geometry("1410x700")

        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))

        self.load_icon(base_path)

        # Define application paths
        self.modules_dir = os.path.join(base_path, "modules")
        self.output_dir = os.path.join(base_path, "output")
        self.configs_dir = os.path.join(base_path, "configs")
        self.config_file_path = os.path.join(self.configs_dir, "config.ini")
        self.presets_file_path = os.path.join(self.configs_dir, "presets.ini")
        self.network_data_file_path = os.path.join(self.configs_dir, "BetraNetzziffern.txt")

        # Config Parsers
        self.preset_config = configparser.ConfigParser()
        self.presets = {}
        self.config = configparser.ConfigParser()
        self.settings = {}

        # Data variables
        self.network_data = {}
        self.cover_pages = []
        self.selected_cover_page = tk.StringVar()
        self.checkbox_items = []

        # Load Configs
        self.load_or_create_network_data()  # Must be first
        self.load_or_create_config()
        self.load_or_create_presets()

        # Build UI
        self.create_main_widgets()

        # Load module files into UI
        self.load_files()

    def load_icon(self, base_path):
        """Try to load .ico, fallback to .png."""
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
                print(f"Could not load icon: {e}")

    def create_main_widgets(self):
        """Creates all widgets for the main application window."""
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Top Info Bar
        top_info_frame = ttk.Frame(main_frame)
        top_info_frame.pack(fill=tk.X, anchor="n", pady=(0, 5))

        info_label = ttk.Label(top_info_frame, text=f"Module aus: '{self.modules_dir}'",
                               font=("-default-", 9, "italic"))
        info_label.pack(side=tk.LEFT, anchor="w")

        year_short = self.settings.get('year', '??')
        year_display = f"20{year_short}" if year_short.isdigit() else "??"

        config_text = f"Aktuelle Konfiguration: {self.settings.get('regional_code_full', '??')} ({self.settings.get('network_name', '???')}), Jahr: {year_display}"

        self.config_label = ttk.Label(top_info_frame, text=config_text, font=("-default-", 9, "italic"))
        self.config_label.pack(side=tk.RIGHT, anchor="e")

        # Cover Page Selector
        cover_page_frame = ttk.Frame(main_frame)
        cover_page_frame.pack(fill=tk.X, anchor="n", pady=(0, 10))

        cover_label = ttk.Label(cover_page_frame, text="Deckblatt auswählen:", font=("-default-", 10, "bold"))
        cover_label.pack(side=tk.LEFT, anchor="w")

        self.cover_page_combo = ttk.Combobox(cover_page_frame, textvariable=self.selected_cover_page, state="readonly",
                                             width=60)
        self.cover_page_combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))

        # Scrollable Checkbox Area
        list_frame = ttk.Frame(main_frame, padding=(0, 10, 0, 0))
        list_frame.pack(fill=tk.BOTH, expand=True)

        x_scrollbar = ttk.Scrollbar(list_frame, orient="horizontal")
        x_scrollbar.pack(side="bottom", fill="x")
        y_scrollbar = ttk.Scrollbar(list_frame, orient="vertical")
        y_scrollbar.pack(side="right", fill="y")

        self.canvas = tk.Canvas(list_frame)
        self.canvas.pack(side=tk.LEFT, fill="both", expand=True)

        self.canvas.configure(yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set)
        x_scrollbar.configure(command=self.canvas.xview)
        y_scrollbar.configure(command=self.canvas.yview)

        self.scrollable_frame = ttk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")

        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind("<Button-4>", self._on_mousewheel)
        self.canvas.bind("<Button-5>", self._on_mousewheel)
        self.scrollable_frame.bind("<MouseWheel>", self._on_mousewheel)
        self.scrollable_frame.bind("<Button-4>", self._on_mousewheel)
        self.scrollable_frame.bind("<Button-5>", self._on_mousewheel)

        # Preset Buttons
        category_frame = ttk.Frame(main_frame)
        category_frame.pack(fill=tk.X, pady=(10, 5))

        category_label = ttk.Label(category_frame, text="Presets (An/Aus):")
        category_label.pack(fill=tk.X, pady=(0, 4))

        self.preset_btn_container = ttk.Frame(category_frame)
        self.preset_btn_container.pack(fill=tk.X)

        self.create_preset_buttons()

        # Bottom Button Bar
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))

        self.help_button = ttk.Button(button_frame, text="Anleitung", command=self.show_help)
        self.help_button.pack(side=tk.LEFT)

        self.contact_button = ttk.Button(button_frame, text="Kontakt", command=self.show_contact)
        self.contact_button.pack(side=tk.LEFT, padx=5)

        self.reset_button = ttk.Button(button_frame, text="Auswahl zurücksetzen", command=self.reset_selection)
        self.reset_button.pack(side=tk.LEFT, padx=(5, 0))

        self.edit_presets_button = ttk.Button(button_frame, text="Presets bearbeiten", command=self.open_preset_editor)
        self.edit_presets_button.pack(side=tk.LEFT, padx=(5, 0))

        self.start_button = ttk.Button(button_frame, text="Ausgewählte Dateien zusammenfügen", command=self.start_merge)
        self.start_button.pack(side=tk.RIGHT)
        self.start_button["state"] = "disabled"

    def load_or_create_network_data(self):
        """Loads BetraNetzziffern.txt, or creates it if it doesn't exist."""
        os.makedirs(self.configs_dir, exist_ok=True)
        if not os.path.exists(self.network_data_file_path):
            print(f"Datei '{self.network_data_file_path}' nicht gefunden, wird erstellt...")
            try:
                default_content = (
                    "RB Ost\n"
                    "S10, S-Bahn-Berlin\nF12, Netz Berlin\nF16, Netz Cottbus\nF18, Netz Schwerin\nF19, Netz Neustrelitz\n\n"
                    "RB Nord\n"
                    "F21, Netz Bremen\nF22, Netz Hamburg\nF23, Netz Hannover\nF24, Netz Kiel\nF25, Netz Osnabrück\n\n"
                    "RB West\n"
                    "F31, Netz Duisburg\nF32, Netz Düsseldorf\nF33, Netz Hagen\nF34, Netz Hamm\nF35, Netz Köln\n\n"
                    "RB Südost\n"
                    "F41, Netz Dresden\nF42, Netz Erfurt\nF43, Netz Halle(Saale)\nF44, Netz Leipzig\nF45, Netz Magdeburg\nF46, Netz Zwickau\n\n"
                    "RB Mitte\n"
                    "F51, Netz Frankfurt(Main)\nF52, Netz Kassel\nF53, Netz Koblenz\nF54, Netz Mainz\n\n"
                    "RB Südwest\n"
                    "F61, Netz Freiburg\nF62, Netz Karlsruhe\nF63, Netz Saarbrücken\nF64, Netz Stuttgart\nF65, Netz Ulm\n\n"
                    "RB Süd\n"
                    "F71, Netz Augsburg\nF72, Netz München\nF73, Netz Nürnberg\nF74, Netz Regensburg\nF75, Netz Würzburg\n"
                )
                with open(self.network_data_file_path, 'w', encoding='utf-8') as f:
                    f.write(default_content)
            except Exception as e:
                messagebox.showerror("Kritischer Fehler",
                                     f"Konnte '{self.network_data_file_path}' nicht erstellen: {e}")
                self.root.quit()
                return

        # Now, parse the file
        try:
            with open(self.network_data_file_path, 'r', encoding='utf-8') as f:
                current_rb = None
                for line in f:
                    line = line.strip()
                    if not line:
                        continue
                    if "," not in line:  # This is an RB header
                        current_rb = line
                        if current_rb not in self.network_data:
                            self.network_data[current_rb] = {}
                    else:  # This is a network line
                        if current_rb is None:
                            continue  # Skip lines before the first RB header
                        parts = line.split(",", 1)
                        if len(parts) == 2:
                            code = parts[0].strip()
                            name = parts[1].strip()
                            self.network_data[current_rb][code] = name
        except Exception as e:
            messagebox.showerror("Kritischer Fehler", f"Konnte '{self.network_data_file_path}' nicht lesen: {e}")
            self.root.quit()

    def load_or_create_config(self):
        """Loads config.ini or triggers first-time setup."""
        os.makedirs(self.configs_dir, exist_ok=True)
        try:
            if not os.path.exists(self.config_file_path):
                raise FileNotFoundError("Config file not found.")

            self.config.read(self.config_file_path)

            if 'SETTINGS' not in self.config or \
                    'RegionalCodeFull' not in self.config['SETTINGS'] or \
                    'NetworkName' not in self.config['SETTINGS'] or \
                    'Year' not in self.config['SETTINGS']:
                raise ValueError("Config file is incomplete.")

            self.settings['regional_code_full'] = self.config['SETTINGS']['RegionalCodeFull']
            self.settings['network_name'] = self.config['SETTINGS']['NetworkName']
            self.settings['year'] = self.config['SETTINGS']['Year']

            if self.settings['year'] != "26":
                print("Alte Jahr-Einstellung gefunden. Erzwinge '26' für Modul-Kompatibilität.")
                self.settings['year'] = "26"
                self.config['SETTINGS']['Year'] = "26"
                with open(self.config_file_path, 'w') as configfile:
                    self.config.write(configfile)

            if not self.settings['regional_code_full'] or not self.settings['network_name']:
                raise ValueError("Config values are empty.")

        except Exception as e:
            print(f"Configuration error: {e}. Starting first-time setup...")
            self.settings = {'regional_code_full': '??', 'network_name': '???', 'year': '26'}
            self.root.after_idle(self.ask_for_initial_config)

    def load_or_create_presets(self):
        """Loads presets.ini, or creates defaults."""
        os.makedirs(self.configs_dir, exist_ok=True)
        try:
            if not os.path.exists(self.presets_file_path):
                raise FileNotFoundError("Presets file not found.")

            self.preset_config.read(self.presets_file_path)

            for i in range(1, NUM_PRESETS + 1):
                section = f'PRESET_{i}'
                if section not in self.preset_config:
                    raise ValueError(f"Preset section {section} missing.")

                name = self.preset_config[section]['Name']
                modules = self.preset_config[section]['Modules']
                self.presets[section] = {'Name': name, 'Modules': modules}

            if len(self.presets) < NUM_PRESETS:
                raise ValueError("Not all presets were found.")

        except Exception as e:
            print(f"Preset config error: {e}. Creating default presets.")
            self.create_default_presets()

    def create_default_presets(self):
        """Creates and saves default presets."""
        default_presets_data = {
            "Oberleitung": ["2.3.", "4.3.0", "5.3.20"],
            "Baugleis": ["5.1.11", "5.3.14", "5.3.15", "5.3.16", "5.3.17", "5.3.18", "5.3.21"],
            "BÜ": ["5.1.22", "5.1.23", "5.1.24", "5.1.25", "5.1.26", "5.1.27", "5.1.28", "5.3.11"],
            "Lfst (Pkt. 3)": ["3.1.", "3.2."],
            "VorGWB": ["5.1.20", "5.1.21"],
        }

        self.presets.clear()
        self.preset_config = configparser.ConfigParser()

        i = 1
        for name, modules_list in default_presets_data.items():
            if i > NUM_PRESETS:
                break

            section = f'PRESET_{i}'
            modules_str = ', '.join(modules_list)

            self.preset_config[section] = {'Name': name, 'Modules': modules_str}
            self.presets[section] = {'Name': name, 'Modules': modules_str}
            i += 1

        while i <= NUM_PRESETS:
            section = f'PRESET_{i}'
            name = f"Preset {i}"
            modules_str = ""
            self.preset_config[section] = {'Name': name, 'Modules': modules_str}
            self.presets[section] = {'Name': name, 'Modules': modules_str}
            i += 1

        try:
            with open(self.presets_file_path, 'w') as f:
                self.preset_config.write(f)
        except Exception as e:
            print(f"Could not save default presets: {e}")

    def create_preset_buttons(self):
        """Clears and rebuilds the preset buttons from self.presets."""
        for widget in self.preset_btn_container.winfo_children():
            widget.destroy()

        for i in range(1, NUM_PRESETS + 1):
            section = f'PRESET_{i}'
            name = self.presets[section]['Name']
            modules_str = self.presets[section]['Modules']

            prefixes = [p.strip() for p in modules_str.split(',') if p.strip()]

            btn = ttk.Button(self.preset_btn_container,
                             text=name,
                             command=lambda p=prefixes: self.toggle_category(p))
            btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2, pady=2)

        alle_prefixes = ["0.", "1.", "2.", "3.", "4.", "5.", "6.", "7.", "8.", "9."]
        alle_btn = ttk.Button(self.preset_btn_container,
                              text="Alle",
                              command=lambda p=alle_prefixes: self.toggle_category(p))
        alle_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2, pady=2)

    def open_preset_editor(self):
        """Opens a new Toplevel window to edit the presets."""
        self.editor_window = tk.Toplevel(self.root)
        self.editor_window.title("Preset-Editor")
        self.editor_window.transient(self.root)
        self.editor_window.grab_set()
        self.editor_window.resizable(False, False)

        main_frame = ttk.Frame(self.editor_window, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        self.preset_name_vars = []
        self.preset_modules_vars = []

        notebook = ttk.Notebook(main_frame)
        notebook.pack(pady=10, padx=10, fill="x", expand=True)

        for i in range(1, NUM_PRESETS + 1):
            section = f'PRESET_{i}'
            preset_data = self.presets[section]

            tab_frame = ttk.Frame(notebook, padding="10")
            notebook.add(tab_frame, text=f"Preset {i}")

            name_var = tk.StringVar(value=preset_data['Name'])
            self.preset_name_vars.append(name_var)

            ttk.Label(tab_frame, text="Button-Name:").pack(anchor="w")
            ttk.Entry(tab_frame, textvariable=name_var, width=50).pack(fill="x", anchor="w", pady=(0, 10))

            modules_var = tk.StringVar(value=preset_data['Modules'])
            self.preset_modules_vars.append(modules_var)

            ttk.Label(tab_frame, text="Modul-Präfixe (durch Komma getrennt):").pack(anchor="w")
            ttk.Entry(tab_frame, textvariable=modules_var, width=50).pack(fill="x", anchor="w")

        editor_btn_frame = ttk.Frame(main_frame)
        editor_btn_frame.pack(fill="x", pady=(10, 0))

        cancel_btn = ttk.Button(editor_btn_frame, text="Abbrechen", command=self.editor_window.destroy)
        cancel_btn.pack(side=tk.RIGHT, padx=5)

        save_btn = ttk.Button(editor_btn_frame, text="Speichern", command=self.save_presets)
        save_btn.pack(side=tk.RIGHT)

    def save_presets(self):
        """Saves the edited presets to file and memory."""
        try:
            for i in range(1, NUM_PRESETS + 1):
                section = f'PRESET_{i}'
                name = self.preset_name_vars[i - 1].get()
                modules = self.preset_modules_vars[i - 1].get()

                if not name.strip():
                    messagebox.showerror("Fehler", f"Der Name für Preset {i} darf nicht leer sein.",
                                         parent=self.editor_window)
                    return

                self.preset_config[section]['Name'] = name
                self.preset_config[section]['Modules'] = modules
                self.presets[section] = {'Name': name, 'Modules': modules}

            with open(self.presets_file_path, 'w') as f:
                self.preset_config.write(f)

            self.create_preset_buttons()
            self.editor_window.destroy()
            messagebox.showinfo("Gespeichert", "Presets erfolgreich aktualisiert.", parent=self.root)

        except Exception as e:
            messagebox.showerror("Fehler", f"Presets konnten nicht gespeichert werden:\n{e}", parent=self.editor_window)

    def ask_for_initial_config(self):
        """Runs the new custom dialog for first-time setup."""
        if not self.network_data:
            messagebox.showerror("Kritischer Fehler", "Netzwerkdaten sind nicht geladen. Konfiguration nicht möglich.")
            self.root.quit()
            return

        messagebox.showinfo("Erstkonfiguration",
                            "Willkommen! Bitte gib dein Standard-Netz ein.",
                            parent=self.root)

        dialog = InitialConfigDialog(self.root, "Erstkonfiguration", self.network_data)

        if not dialog.result:
            messagebox.showerror("Abbruch", "Ohne Konfiguration kann das Programm nicht starten.")
            self.root.quit()
            return

        code_full, name = dialog.result
        year_short = "26"  # Hardcoded to match modules

        # Save values to config file
        self.config['SETTINGS'] = {
            'RegionalCodeFull': code_full,
            'NetworkName': name,
            'Year': year_short
        }
        with open(self.config_file_path, 'w') as configfile:
            self.config.write(configfile)

        # Save values to working memory
        self.settings['regional_code_full'] = code_full
        self.settings['network_name'] = name
        self.settings['year'] = year_short
        messagebox.showinfo("Konfiguration", "Einstellungen erfolgreich gespeichert.", parent=self.root)

        # Update UI label
        if hasattr(self, 'config_label'):
            year_short = self.settings.get('year', '??')
            year_display = f"20{year_short}" if year_short.isdigit() else "??"
            config_text = f"Aktuelle Konfiguration: {self.settings.get('regional_code_full', '??')} ({self.settings.get('network_name', '???')}), Jahr: {year_display}"
            self.config_label.config(text=config_text)

    def _on_mousewheel(self, event):
        """Cross-platform mousewheel scrolling."""
        if sys.platform == "linux":
            if event.num == 4:
                self.canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                self.canvas.yview_scroll(1, "units")
        else:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _on_label_click(self, checkbox_widget, check_var):
        """Toggles the associated checkbox when its label is clicked."""
        if checkbox_widget.instate(['!disabled']):
            check_var.set(not check_var.get())

    def _get_layout_key(self, filename):
        """Determines the layout group (and column) for a file."""
        parts = filename.split('.')
        if not parts:
            return "Unsorted"

        main_chapter = parts[0]

        if main_chapter == "5":
            if len(parts) > 1:
                key = f"{parts[0]}.{parts[1]}"
                if key in COLUMN_LAYOUT:
                    return key

        if main_chapter in COLUMN_LAYOUT:
            return main_chapter

        return "Unsorted"

    def load_files(self):
        """
        Loads all .docx files, separating them into cover pages (for combobox)
        and modules (for checkboxes).
        """
        if not os.path.isdir(self.modules_dir):
            messagebox.showerror("Fehler", f"Der Ordner '{self.modules_dir}' wurde nicht gefunden.")
            self.root.quit()
            return

        # --- Clear existing UI elements ---
        for item in self.checkbox_items:
            item["checkbox"].master.destroy()
        self.checkbox_items.clear()

        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        self.cover_pages.clear()
        self.cover_page_combo['values'] = []
        self.selected_cover_page.set("")
        # --- End Clear ---

        search_path = os.path.join(self.modules_dir, "*.docx")
        all_file_paths = glob.glob(search_path)

        if not all_file_paths:
            messagebox.showinfo("Keine Dateien", f"Keine .docx Dateien im Ordner '{self.modules_dir}' gefunden.")
            return

        cover_page_files = []
        module_files = []

        for file_path in all_file_paths:
            filename = os.path.basename(file_path)
            if filename.startswith("0."):
                cover_page_files.append(file_path)
            else:
                module_files.append(file_path)

        # 1. Populate Cover Page ComboBox
        cover_page_names = []
        cover_page_files.sort(key=natural_sort_key)

        for file_path in cover_page_files:
            filename = os.path.basename(file_path)
            display_name = os.path.splitext(filename)[0]
            self.cover_pages.append({'name': display_name, 'path': file_path})
            cover_page_names.append(display_name)

        self.cover_page_combo['values'] = cover_page_names
        if cover_page_names:
            self.selected_cover_page.set(cover_page_names[0])
        else:
            messagebox.showwarning("Deckblatt fehlt",
                                   f"Keine Deckblatt-Dateien (beginnend mit '0.') im Ordner '{self.modules_dir}' gefunden.\n Zusammenfügen ist nicht möglich.")
            self.start_button["state"] = "disabled"

        # 2. Populate Module Checkboxes
        module_files.sort(key=natural_sort_key)

        if not module_files and cover_page_files:
            messagebox.showinfo("Keine Module",
                                f"Keine Modul-Dateien (außer Deckblättern) im Ordner '{self.modules_dir}' gefunden.")

        wrap_length_pixels = 220
        main_columns = []
        for i in range(NUM_MAIN_COLUMNS):
            main_col_frame = ttk.Frame(self.scrollable_frame)
            main_col_frame.grid(row=0, column=i, sticky="nw", padx=5)
            main_col_frame.bind("<MouseWheel>", self._on_mousewheel)
            main_columns.append(main_col_frame)

        group_frames = {}

        for file_path in module_files:
            filename = os.path.basename(file_path)
            layout_key = self._get_layout_key(filename)

            if layout_key not in group_frames:
                main_col_index = COLUMN_LAYOUT.get(layout_key, NUM_MAIN_COLUMNS - 1)
                parent_frame = main_columns[main_col_index]

                group_frame = ttk.Frame(parent_frame, padding=5, borderwidth=1, relief="sunken")
                group_frame.pack(side=tk.TOP, fill="x", anchor="n", pady=5)

                title_text = f"Punkt {layout_key}" if layout_key != "Unsorted" else "Unsortiert"
                title_label = ttk.Label(group_frame, text=title_text, font=("-default-", 10, "bold"),
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
            check_var = tk.BooleanVar(value=is_mandatory)
            cb_state = "disabled" if is_mandatory else "normal"

            item_frame = ttk.Frame(group_frame)
            item_frame.pack(fill='x', anchor="w", pady=1)

            checkbox = ttk.Checkbutton(item_frame, variable=check_var, state=cb_state)
            checkbox.pack(side=tk.LEFT, anchor="n", padx=(0, 5))

            display_name = os.path.splitext(filename)[0]

            label = ttk.Label(item_frame, text=display_name, wraplength=wrap_length_pixels, justify=tk.LEFT)
            label.pack(side=tk.LEFT, fill='x', expand=True)

            label.bind("<Button-1>", lambda e, w=checkbox, v=check_var: self._on_label_click(w, v))

            item_frame.bind("<MouseWheel>", self._on_mousewheel)
            item_frame.bind("<Button-4>", self._on_mousewheel)
            item_frame.bind("<Button-5>", self._on_mousewheel)
            checkbox.bind("<MouseWheel>", self._on_mousewheel)
            checkbox.bind("<Button-4>", self._on_mousewheel)
            checkbox.bind("<Button-5>", self._on_mousewheel)
            label.bind("<MouseWheel>", self._on_mousewheel)
            label.bind("<Button-4>", self._on_mousewheel)
            label.bind("<Button-5>", self._on_mousewheel)

            self.checkbox_items.append({
                "check_var": check_var,
                "path": file_path,
                "filename": filename,
                "checkbox": checkbox,
                "is_mandatory": is_mandatory
            })

        if cover_page_files:
            self.start_button["state"] = "normal"

    def reset_selection(self):
        """Resets all optional checkboxes to False."""
        for item in self.checkbox_items:
            if not item["is_mandatory"]:
                item["check_var"].set(False)
            else:
                item["check_var"].set(True)

    def toggle_category(self, prefixes):
        """Toggles non-mandatory checkboxes matching the prefixes."""
        items_in_category = []
        for item in self.checkbox_items:
            if item["is_mandatory"] or item["checkbox"].cget("state") == "disabled":
                continue

            for prefix in prefixes:
                if item["filename"].startswith(prefix):
                    items_in_category.append(item)
                    break

        if not items_in_category:
            return

        is_anything_deselected = any(not item["check_var"].get() for item in items_in_category)
        new_state = is_anything_deselected

        for item in items_in_category:
            item["check_var"].set(new_state)

    def show_help(self):
        """Displays the help/instructions messagebox."""
        help_text = (
            "Anleitung Betra Builder (v1.0a)\n\n"  # Version angepasst
            "1. Wählen Sie oben das gewünschte Deckblatt aus der Liste aus.\n\n"
            "2. Pflichtmodule sind bereits ausgewählt und können nicht abgewählt werden.\n\n"
            "3. Wählen Sie optionale Module aus, indem Sie die Haken setzen (Klick auf den Haken oder den Text).\n\n"
            "4. Nutzen Sie die 'Presets'-Buttons, um gängige Modul-Gruppen schnell an- oder abzuwählen.\n\n"
            "5. Mit 'Auswahl zurücksetzen' werden alle optionalen Module abgewählt.\n\n"
            "6. Klicken Sie auf 'Ausgewählte Dateien zusammenfügen', geben Sie die laufende Nummer an.\n"
            "   -> Die Datei wird im 'output'-Ordner in einem eigenen Unterordner gespeichert.\n\n"
            "7. AEL-Verrechnung: Setzen Sie den Haken, um nach dem Speichern eine Projektnummer\n"
            "   für die Excel-Datei 'AEL-Verrechnung.xlsx' einzugeben.\n\n"
            "--- \n"
            "Eigene Presets:\n"
            "Mit 'Presets bearbeiten' können Sie die 5 Preset-Buttons an Ihre Bedürfnisse anpassen.\n\n"
            "Konfiguration (Netz/Jahr):\n"
            "Das Jahr ist auf 2026 festgelegt. Um Ihr Netz zu ändern, löschen Sie die Datei 'config.ini' im Ordner 'configs' und starten Sie das Programm neu."
        )
        messagebox.showinfo("Anleitung", help_text)

    def show_contact(self):
        """Displays the contact/support messagebox."""
        contact_text = (
                "Kontakt & Support\n\n"
                "Bei Fragen, Problemen, Ideen oder Vorschläge mit dem Betra Builder:\n\n"
                "Name: Dennis Heinze, I.IA-W-N-HA-B\n"
                "E-Mail: dennis.heinze@deutschebahn.com\n"
                "Telefon (dienstlich): 0152 33114237\n"
                "Version: " + APP_VERSION + "\n\n"
        )
        messagebox.showinfo("Kontakt", contact_text)

    def start_merge(self):
        """Gathers selected files and triggers the document merge process."""

        # 1. Get Cover Page
        selected_cover_name = self.selected_cover_page.get()
        if not selected_cover_name:
            messagebox.showwarning("Deckblatt fehlt",
                                   "Bitte ein Deckblatt aus der Liste auswählen, bevor Sie fortfahren.")
            return

        cover_path = ""
        for cover in self.cover_pages:
            if cover['name'] == selected_cover_name:
                cover_path = cover['path']
                break

        if not cover_path or not os.path.exists(cover_path):
            messagebox.showerror("Fehler", f"Die Deckblatt-Datei '{selected_cover_name}' konnte nicht gefunden werden.")
            return

        selected_files_for_merge = [cover_path]

        # 2. Get Selected Modules
        for item in self.checkbox_items:
            if item["check_var"].get():
                selected_files_for_merge.append(item["path"])

        # 3. Validation
        if len(selected_files_for_merge) == 1:
            if not messagebox.askyesno("Warnung",
                                       "Es sind keine Module ausgewählt.\n\nMöchten Sie nur das Deckblatt unter dem neuen Namen speichern?",
                                       parent=self.root):
                return

        # 4. Ensure *base* output directory exists
        try:
            os.makedirs(self.output_dir, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Fehler", f"Konnte den Output-Ordner nicht erstellen:\n{e}")
            return

        # 5. Get Filename and AEL option
        dialog = FileNameDialog(self.root, "Dateiname festlegen")
        if not dialog.result:
            return

        # (MODIFIZIERT) Entpackt jetzt 3 Werte
        doc_type, serial_num, ael_checked = dialog.result

        base_name = f"{doc_type} {self.settings['regional_code_full']} {serial_num}-{self.settings['year']}"
        new_folder_path = os.path.join(self.output_dir, base_name)
        file_name_with_ext = f"{base_name}.docx"
        save_path = os.path.join(new_folder_path, file_name_with_ext)

        # 6. Create subfolder
        try:
            os.makedirs(new_folder_path, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Fehler", f"Konnte den Unterordner nicht erstellen:\n{new_folder_path}\n\nFehler: {e}")
            return

        # 7. Check existence
        if os.path.exists(save_path):
            relative_file_name = os.path.join(base_name, file_name_with_ext)
            if not messagebox.askyesno("Warnung",
                                       f"Die Datei:\n{relative_file_name}\n\nexistiert bereits.\nSoll sie überschrieben werden?",
                                       parent=self.root):
                return

        # 8. Run the merge process
        try:
            self.start_button.config(text="Arbeite...", state="disabled")
            self.root.update_idletasks()

            self.merge_documents(selected_files_for_merge, save_path)

            messagebox.showinfo("Erfolg", f"Dateien erfolgreich zusammengefügt!\nGespeichert als: {save_path}")

            # --- (NEUE AEL-LOGIK) ---
            if ael_checked:
                project_num = simpledialog.askstring("AEL-Verrechnung",
                                                     "Bitte Projektnummer eingeben:",
                                                     parent=self.root)

                # Fährt nur fort, wenn der Nutzer nicht "Abbrechen" klickt und etwas eingibt
                if project_num:
                    today_date = datetime.now().strftime("%d.%m.%Y")
                    # Wir verwenden base_name, z.B. "Betra F33 1234-26"
                    self.update_ael_excel(project_num, today_date, base_name)
            # --- (ENDE NEUE AEL-LOGIK) ---

        except Exception as e:
            messagebox.showerror("Fehler", f"Ein Fehler ist aufgetreten:\n{e}\n\n"
                                           f"Hinweis: Stellen Sie sicher, dass die Zieldatei (falls sie existiert) geschlossen ist.")
        finally:
            self.start_button.config(text="Ausgewählte Dateien zusammenfügen", state="normal")

    # --- NEUE METHODE ---
    def update_ael_excel(self, project_num, today_date, file_base_name):
        """Erstellt oder aktualisiert die AEL-Verrechnung.xlsx im output-Ordner."""

        excel_path = os.path.join(self.output_dir, "AEL-Verrechnung.xlsx")
        headers = ["Projektnummer", "Datum", "Name"]
        new_row = [project_num, today_date, file_base_name]

        try:
            if not os.path.exists(excel_path):
                # Datei existiert nicht -> Neu erstellen mit Kopfzeile
                wb = openpyxl.Workbook()
                sheet = wb.active
                sheet.title = "Verrechnung"
                sheet.append(headers)
                sheet.append(new_row)
            else:
                # Datei existiert -> Laden und Zeile anhängen
                wb = openpyxl.load_workbook(excel_path)
                sheet = wb.active

                # (Optional) Prüfen, ob die Kopfzeile korrekt ist
                if sheet["A1"].value != headers[0]:
                    # Wenn die Datei komisch formatiert ist, einfach anhängen
                    print("Warnung: AEL-Excel-Kopfzeile weicht ab, Zeile wird trotzdem angehängt.")

                sheet.append(new_row)

            # Spaltenbreite automatisch anpassen (optional, aber nützlich)
            for col in sheet.columns:
                max_length = 0
                column = col[0].column_letter  # Spaltenbuchstabe
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = (max_length + 2)
                sheet.column_dimensions[column].width = adjusted_width

            wb.save(excel_path)
            messagebox.showinfo("AEL-Verrechnung",
                                f"Excel-Datei '{excel_path}' erfolgreich aktualisiert.",
                                parent=self.root)

        except PermissionError:
            messagebox.showerror("Fehler (Excel)",
                                 f"Speichern fehlgeschlagen!\nDie Datei '{excel_path}' ist eventuell geöffnet.\n\n"
                                 "Bitte schließen Sie die Datei und tragen Sie die Zeile manuell ein:\n"
                                 f"Projekt: {project_num}\nDatum: {today_date}\nName: {file_base_name}",
                                 parent=self.root)
        except Exception as e:
            messagebox.showerror("Fehler (Excel)",
                                 f"Ein unerwarteter Fehler beim Speichern der Excel-Datei ist aufgetreten:\n{e}",
                                 parent=self.root)

    def merge_documents(self, file_paths, save_path):
        """
        Merges a list of .docx files into a single document.
        The first file (file_paths[0]) is the base document.
        """
        if not file_paths:
            return

        if not os.path.exists(file_paths[0]):
            raise FileNotFoundError(f"Die Basis-Datei (Deckblatt) konnte nicht gefunden werden: {file_paths[0]}")

        master_doc = Document(file_paths[0])
        composer = Composer(master_doc)

        if len(file_paths) > 1:
            for file_path in file_paths[1:]:
                if not os.path.exists(file_path):
                    print(f"Warning: Skipping file (not found): {file_path}")
                    continue
                try:
                    doc_to_append = Document(file_path)
                    composer.append(doc_to_append)
                except Exception as inner_exception:
                    print(f"Error appending {file_path}: {inner_exception}")
                    pass

        composer.save(save_path)


if __name__ == "__main__":
    root = tk.Tk()
    app = WordMergerApp(root)
    root.mainloop()