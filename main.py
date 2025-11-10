import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import os
import glob
import sys
import re
import configparser
from docx import Document
from docxcompose.composer import Composer

# --- CONFIGURATION ---
# These filenames are data and must match the files in the 'modules' directory.
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

# --- (MODIFIZIERT) CATEGORIES wurde entfernt, wird jetzt aus presets.ini geladen ---
NUM_PRESETS = 5  # --- NEU ---

# Defines which chapter key maps to which UI column index.
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
APP_VERSION = "0.4a"  # Updated version


# ---------------------

def natural_sort_key(s):
    """Sorts strings 'naturally' (e.g., 1, 2, 10) instead of alphabetically (1, 10, 2)."""
    filename = os.path.basename(s)
    return [int(c) if c.isdigit() else c.lower() for c in re.split('([0-9]+)', filename)]


# --- CUSTOM DIALOG FOR FILENAME INPUT ---
class FileNameDialog(simpledialog.Dialog):
    """
    A custom dialog to ask the user for the document type ("Betra" or "BA")
    and the serial number (YYYY).
    """

    def __init__(self, parent, title):
        self.result = None
        super().__init__(parent, title)

    def body(self, frame):
        # 1. Document Type (Radiobuttons)
        type_frame = ttk.Frame(frame)
        ttk.Label(type_frame, text="Art:").pack(side=tk.LEFT, padx=5)

        self.doc_type_var = tk.StringVar(value="Betra")
        rb1 = ttk.Radiobutton(type_frame, text="Betra", variable=self.doc_type_var, value="Betra")
        rb1.pack(side=tk.LEFT, padx=5)
        rb2 = ttk.Radiobutton(type_frame, text="BA", variable=self.doc_type_var, value="BA")
        rb2.pack(side=tk.LEFT, padx=5)
        type_frame.pack(pady=5)

        # 2. Serial Number (Entry)
        num_frame = ttk.Frame(frame)
        ttk.Label(num_frame, text="Laufende Nummer (YYYY):").pack(side=tk.LEFT, padx=5)

        self.entry_var = tk.StringVar()
        self.entry_widget = ttk.Entry(num_frame, textvariable=self.entry_var, width=10)
        self.entry_widget.pack(side=tk.LEFT)
        num_frame.pack(pady=5)

        return self.entry_widget  # Set focus to the entry widget

    def validate(self):
        serial_num = self.entry_var.get().strip()
        if not serial_num:
            messagebox.showwarning("Eingabe fehlt", "Bitte eine laufende Nummer eingeben.", parent=self)
            return 0  # Validation failed

        return 1  # Validation successful

    def apply(self):
        # This is called when "OK" is pressed. Store the result.
        self.result = (self.doc_type_var.get(), self.entry_var.get().strip())


# -----------------------------------------------


class WordMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Betra Builder v" + APP_VERSION)
        self.root.geometry("1410x700")

        # Determine base path for bundled (PyInstaller) or script execution
        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))

        # Try to load application icon
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

        # Define application paths
        self.modules_dir = os.path.join(base_path, "modules")
        self.output_dir = os.path.join(base_path, "output")
        self.configs_dir = os.path.join(base_path, "configs")
        self.config_file_path = os.path.join(self.configs_dir, "config.ini")

        # --- (MODIFIZIERT) Config-Pfade erweitert ---
        self.presets_file_path = os.path.join(self.configs_dir, "presets.ini")
        self.preset_config = configparser.ConfigParser()
        self.presets = {}
        # --- ENDE MODIFIKATION ---

        # --- CONFIGURATION LOGIC ---
        self.config = configparser.ConfigParser()
        self.settings = {}
        self.load_or_create_config()
        self.load_or_create_presets()  # --- NEU ---
        # --- END CONFIGURATION LOGIC ---

        self.checkbox_items = []  # Stores dicts of {var, path, filename, widget, is_mandatory}
        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- Top Info Bar (NEW FRAME) ---
        top_info_frame = ttk.Frame(main_frame)
        top_info_frame.pack(fill=tk.X, anchor="n", pady=(0, 5))  # Pack at the top, fill horizontally

        # 1. Module Path Label (packed left)
        info_label = ttk.Label(top_info_frame, text=f"Module aus: '{self.modules_dir}'",
                               font=("-default-", 9, "italic"))
        info_label.pack(side=tk.LEFT, anchor="w")

        # 2. Config Settings Label (packed right)
        config_text = f"Aktuelle Konfiguration: Code: {self.settings.get('regional_code', '??')}, Jahr: {self.settings.get('year', '??')}"
        self.config_label = ttk.Label(top_info_frame, text=config_text, font=("-default-", 9, "italic"))
        self.config_label.pack(side=tk.RIGHT, anchor="e")
        # --- End Top Info Bar ---

        # --- SCROLLABLE AREA SETUP ---
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

        # Bind mousewheel events for scrolling
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind("<Button-4>", self._on_mousewheel)  # Linux scroll up
        self.canvas.bind("<Button-5>", self._on_mousewheel)  # Linux scroll down
        self.scrollable_frame.bind("<MouseWheel>", self._on_mousewheel)
        self.scrollable_frame.bind("<Button-4>", self._on_mousewheel)
        self.scrollable_frame.bind("<Button-5>", self._on_mousewheel)
        # --- END SCROLLABLE AREA SETUP ---

        # --- (MODIFIZIERT) CATEGORY BUTTONS ---
        category_frame = ttk.Frame(main_frame)
        category_frame.pack(fill=tk.X, pady=(10, 5))

        category_label = ttk.Label(category_frame, text="Presets (An/Aus):")
        category_label.pack(fill=tk.X, pady=(0, 4))

        self.preset_btn_container = ttk.Frame(category_frame)
        self.preset_btn_container.pack(fill=tk.X)

        self.create_preset_buttons()  # --- NEU ---
        # --- END CATEGORY BUTTONS ---

        # --- (MODIFIZIERT) BOTTOM BUTTON BAR ---
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))

        self.help_button = ttk.Button(button_frame, text="Anleitung", command=self.show_help)
        self.help_button.pack(side=tk.LEFT)

        self.contact_button = ttk.Button(button_frame, text="Kontakt", command=self.show_contact)
        self.contact_button.pack(side=tk.LEFT, padx=5)

        self.reset_button = ttk.Button(button_frame, text="Auswahl zurücksetzen", command=self.reset_selection)
        self.reset_button.pack(side=tk.LEFT, padx=(5, 0))

        # --- NEUER BUTTON ---
        self.edit_presets_button = ttk.Button(button_frame, text="Presets bearbeiten", command=self.open_preset_editor)
        self.edit_presets_button.pack(side=tk.LEFT, padx=(5, 0))
        # --- ENDE NEU ---

        self.start_button = ttk.Button(button_frame, text="Ausgewählte Dateien zusammenfügen", command=self.start_merge)
        self.start_button.pack(side=tk.RIGHT)
        self.start_button["state"] = "disabled"
        # --- END BOTTOM BUTTON BAR ---

        self.load_files()

    def load_or_create_config(self):
        """
        Loads the configuration from 'configs/config.ini'.
        If the file or required settings are missing, triggers the first-time setup.
        """
        os.makedirs(self.configs_dir, exist_ok=True)
        try:
            if not os.path.exists(self.config_file_path):
                raise FileNotFoundError("Config file not found.")

            self.config.read(self.config_file_path)

            if 'SETTINGS' not in self.config or \
                    'RegionalCode' not in self.config['SETTINGS'] or \
                    'Year' not in self.config['SETTINGS']:
                raise ValueError("Config file is incomplete.")

            self.settings['regional_code'] = self.config['SETTINGS']['RegionalCode']
            self.settings['year'] = self.config['SETTINGS']['Year']

            if not self.settings['regional_code'] or not self.settings['year']:
                raise ValueError("Config values are empty.")

        except Exception as e:
            # If anything fails (file missing, section missing, keys missing, values empty):
            # Run the first-time setup.
            print(f"Configuration error: {e}. Starting first-time setup...")
            # We must set defaults here in case ask_for_initial_config is cancelled
            self.settings = {'regional_code': '??', 'year': '??'}
            self.root.after_idle(self.ask_for_initial_config)  # Call after main UI is built

    # --- NEUE METHODE ---
    def load_or_create_presets(self):
        """
        Loads presets from 'configs/presets.ini'.
        If the file is missing or corrupt, creates default presets.
        """
        os.makedirs(self.configs_dir, exist_ok=True)
        try:
            if not os.path.exists(self.presets_file_path):
                raise FileNotFoundError("Presets file not found.")

            self.preset_config.read(self.presets_file_path)

            for i in range(1, NUM_PRESETS + 1):
                section = f'PRESET_{i}'
                if section not in self.preset_config:
                    raise ValueError(f"Preset-Sektion {section} fehlt.")

                name = self.preset_config[section]['Name']
                modules = self.preset_config[section]['Modules']
                self.presets[section] = {'Name': name, 'Modules': modules}

            if len(self.presets) < NUM_PRESETS:
                raise ValueError("Nicht alle Presets wurden gefunden.")

        except Exception as e:
            print(f"Preset-Konfigurationsfehler: {e}. Erstelle Standard-Presets.")
            self.create_default_presets()

    # --- NEUE METHODE ---
    def create_default_presets(self):
        """Creates and saves default presets based on the old CATEGORIES."""
        # Defaults based on old CATEGORIES dict
        default_presets_data = {
            "Oberleitung": ["2.3.", "4.3.0", "5.3.20"],
            "Baugleis": ["5.1.11", "5.3.14", "5.3.15", "5.3.16", "5.3.17", "5.3.18", "5.3.21"],
            "BÜ": ["5.1.22", "5.1.23", "5.1.24", "5.1.25", "5.1.26", "5.1.27", "5.1.28", "5.3.11"],
            "Lfst (Pkt. 3)": ["3.1.", "3.2."],
            "VorGWB": ["5.1.20", "5.1.21"],
        }

        self.presets.clear()
        self.preset_config = configparser.ConfigParser()  # Clear old config

        i = 1
        for name, modules_list in default_presets_data.items():
            if i > NUM_PRESETS:
                break

            section = f'PRESET_{i}'
            modules_str = ', '.join(modules_list)

            self.preset_config[section] = {'Name': name, 'Modules': modules_str}
            self.presets[section] = {'Name': name, 'Modules': modules_str}
            i += 1

        # Fill any remaining preset slots if NUM_PRESETS > 5
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
            print(f"Konnte Standard-Presets nicht speichern: {e}")

    # --- NEUE METHODE ---
    def create_preset_buttons(self):
        """Clears and rebuilds the preset buttons from self.presets."""
        # Clear existing buttons
        for widget in self.preset_btn_container.winfo_children():
            widget.destroy()

        # Add dynamic preset buttons
        for i in range(1, NUM_PRESETS + 1):
            section = f'PRESET_{i}'
            name = self.presets[section]['Name']
            modules_str = self.presets[section]['Modules']

            # Get prefixes from comma-separated string
            prefixes = [p.strip() for p in modules_str.split(',') if p.strip()]

            btn = ttk.Button(self.preset_btn_container,
                             text=name,
                             command=lambda p=prefixes: self.toggle_category(p))
            btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2, pady=2)

        # Add the static "Alle" button (hardcoded)
        alle_prefixes = ["0.", "1.", "2.", "3.", "4.", "5.", "6.", "7.", "8.", "9."]
        alle_btn = ttk.Button(self.preset_btn_container,
                              text="Alle",
                              command=lambda p=alle_prefixes: self.toggle_category(p))
        alle_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2, pady=2)

    # --- NEUE METHODE ---
    def open_preset_editor(self):
        """Opens a new Toplevel window to edit the presets."""

        # Create the Toplevel window
        self.editor_window = tk.Toplevel(self.root)
        self.editor_window.title("Preset-Editor")
        self.editor_window.transient(self.root)  # Keep on top
        self.editor_window.grab_set()  # Modal behavior
        self.editor_window.resizable(False, False)

        main_frame = ttk.Frame(self.editor_window, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Create StringVars to hold the editor's content
        self.preset_name_vars = []
        self.preset_modules_vars = []

        # Use a Notebook (Tabs) for clarity
        notebook = ttk.Notebook(main_frame)
        notebook.pack(pady=10, padx=10, fill="x", expand=True)

        for i in range(1, NUM_PRESETS + 1):
            section = f'PRESET_{i}'
            preset_data = self.presets[section]

            tab_frame = ttk.Frame(notebook, padding="10")
            notebook.add(tab_frame, text=f"Preset {i}")

            # Name Entry
            name_var = tk.StringVar(value=preset_data['Name'])
            self.preset_name_vars.append(name_var)

            ttk.Label(tab_frame, text="Button-Name:").pack(anchor="w")
            ttk.Entry(tab_frame, textvariable=name_var, width=50).pack(fill="x", anchor="w", pady=(0, 10))

            # Modules Entry
            modules_var = tk.StringVar(value=preset_data['Modules'])
            self.preset_modules_vars.append(modules_var)

            ttk.Label(tab_frame, text="Modul-Präfixe (durch Komma getrennt):").pack(anchor="w")
            ttk.Entry(tab_frame, textvariable=modules_var, width=50).pack(fill="x", anchor="w")

        # Save/Cancel buttons for the editor window
        editor_btn_frame = ttk.Frame(main_frame)
        editor_btn_frame.pack(fill="x", pady=(10, 0))

        cancel_btn = ttk.Button(editor_btn_frame, text="Abbrechen", command=self.editor_window.destroy)
        cancel_btn.pack(side=tk.RIGHT, padx=5)

        save_btn = ttk.Button(editor_btn_frame, text="Speichern", command=self.save_presets)
        save_btn.pack(side=tk.RIGHT)

    # --- NEUE METHODE ---
    def save_presets(self):
        """Saves the edited presets from the editor window to file and memory."""
        try:
            for i in range(1, NUM_PRESETS + 1):
                section = f'PRESET_{i}'
                name = self.preset_name_vars[i - 1].get()
                modules = self.preset_modules_vars[i - 1].get()

                # Validate: Name should not be empty
                if not name.strip():
                    messagebox.showerror("Fehler", f"Der Name für Preset {i} darf nicht leer sein.",
                                         parent=self.editor_window)
                    return

                # Update config object
                self.preset_config[section]['Name'] = name
                self.preset_config[section]['Modules'] = modules
                # Update in-memory presets
                self.presets[section] = {'Name': name, 'Modules': modules}

            # Save to file
            with open(self.presets_file_path, 'w') as f:
                self.preset_config.write(f)

            # Refresh the main UI buttons
            self.create_preset_buttons()

            # Close editor
            self.editor_window.destroy()
            messagebox.showinfo("Gespeichert", "Presets erfolgreich aktualisiert.", parent=self.root)

        except Exception as e:
            messagebox.showerror("Fehler", f"Presets konnten nicht gespeichert werden:\n{e}", parent=self.editor_window)

    def ask_for_initial_config(self):
        """
        Prompts the user for initial settings (Regional Code, Year)
        and saves them to config.ini.
        """
        messagebox.showinfo("Erstkonfiguration",
                            "Willkommen! Bitte gib deine Standardwerte ein.",
                            parent=self.root)

        regional_code = simpledialog.askstring("Erstkonfiguration",
                                               "Bitte regionalen Code eingeben (z.B. 33):",
                                               parent=self.root)
        # Abort if user cancels
        if not regional_code:
            messagebox.showerror("Abbruch", "Ohne regionalen Code kann das Programm nicht starten.")
            self.root.quit()
            return

        year = simpledialog.askstring("Erstkonfiguration",
                                      "Bitte Jahr eingeben (z.B. 26 für 2026):",
                                      parent=self.root)
        # Abort if user cancels
        if not year:
            messagebox.showerror("Abbruch", "Ohne Jahr kann das Programm nicht starten.")
            self.root.quit()
            return

        # Save values to config file
        self.config['SETTINGS'] = {
            'RegionalCode': regional_code,
            'Year': year
        }
        with open(self.config_file_path, 'w') as configfile:
            self.config.write(configfile)

        # Save values to working memory
        self.settings['regional_code'] = regional_code
        self.settings['year'] = year
        messagebox.showinfo("Konfiguration", "Einstellungen erfolgreich gespeichert.", parent=self.root)

        # --- (NEW) Update UI label now that we have the values ---
        if hasattr(self, 'config_label'):
            config_text = f"Aktuelle Konfiguration: Code: {self.settings.get('regional_code', '??')}, Jahr: {self.settings.get('year', '??')}"
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
        """
        Determines the layout group (and column) for a file
        based on its chapter number.
        """
        parts = filename.split('.')
        if not parts:
            return "Unsorted"

        main_chapter = parts[0]

        # Special handling for chapter 5 sub-sections (5.0, 5.1, etc.)
        if main_chapter == "5":
            if len(parts) > 1:
                key = f"{parts[0]}.{parts[1]}"
                if key in COLUMN_LAYOUT:
                    return key

        if main_chapter in COLUMN_LAYOUT:
            return main_chapter

        return "Unsorted"  # Fallback for unclassified files

    def load_files(self):
        """
        Loads all .docx files from the 'modules' directory,
        sorts them, and displays them in the scrollable UI.
        """
        if not os.path.isdir(self.modules_dir):
            messagebox.showerror("Fehler", f"Der Ordner '{self.modules_dir}' wurde nicht gefunden.")
            self.root.quit()
            return

        # Clear existing UI elements
        for item in self.checkbox_items:
            item["checkbox"].destroy()
        self.checkbox_items.clear()

        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        # Find, sort, and process files
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

        group_frames = {}  # Cache for group frames

        for file_path in file_paths:
            filename = os.path.basename(file_path)
            layout_key = self._get_layout_key(filename)

            # Create a new group frame if it doesn't exist
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

                # Bind mousewheel to group frames and titles as well
                group_frame.bind("<MouseWheel>", self._on_mousewheel)
                group_frame.bind("<Button-4>", self._on_mousewheel)
                group_frame.bind("<Button-5>", self._on_mousewheel)
                title_label.bind("<MouseWheel>", self._on_mousewheel)
                title_label.bind("<Button-4>", self._on_mousewheel)
                title_label.bind("<Button-5>", self._on_mousewheel)
            else:
                group_frame = group_frames[layout_key]

            # Create checkbox item
            is_mandatory = filename in MANDATORY_FILES
            check_var = tk.BooleanVar(value=is_mandatory)
            cb_state = "disabled" if is_mandatory else "normal"

            item_frame = ttk.Frame(group_frame)
            item_frame.pack(fill='x', anchor="w", pady=1)

            checkbox = ttk.Checkbutton(item_frame, variable=check_var, state=cb_state)
            checkbox.pack(side=tk.LEFT, anchor="n", padx=(0, 5))

            display_name = os.path.splitext(filename)[0]  # Remove .docx

            label = ttk.Label(item_frame, text=display_name, wraplength=wrap_length_pixels, justify=tk.LEFT)
            label.pack(side=tk.LEFT, fill='x', expand=True)

            # Bind click event to the label
            label.bind("<Button-1>", lambda e, w=checkbox, v=check_var: self._on_label_click(w, v))

            # Bind mousewheel to all elements to ensure scrolling works
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

        self.start_button["state"] = "normal"  # Enable merge button after loading

    def reset_selection(self):
        """Resets all optional checkboxes to False."""
        for item in self.checkbox_items:
            if not item["is_mandatory"]:
                item["check_var"].set(False)
            else:
                item["check_var"].set(True)  # Ensure mandatory ones stay checked

    def toggle_category(self, prefixes):
        """
        Toggles all non-mandatory checkboxes whose filename starts
        with one of the given prefixes.
        """
        items_in_category = []
        for item in self.checkbox_items:
            if item["is_mandatory"] or item["checkbox"].cget("state") == "disabled":
                continue

            for prefix in prefixes:
                if item["filename"].startswith(prefix):
                    items_in_category.append(item)
                    break  # Go to next item once matched

        if not items_in_category:
            return

        # Toggle logic: If anything is deselected, select all.
        # If all are selected, deselect all.
        is_anything_deselected = any(not item["check_var"].get() for item in items_in_category)
        new_state = is_anything_deselected  # True (select all) or False (deselect all)

        for item in items_in_category:
            item["check_var"].set(new_state)

    def show_help(self):
        """Displays the help/instructions messagebox."""
        # --- (MODIFIZIERT) Anleitungstext angepasst ---
        help_text = (
            "Anleitung Betra Builder\n\n"
            "1. Pflichtmodule sind bereits ausgewählt und können nicht abgewählt werden.\n\n"
            "2. Wählen Sie optionale Module aus, indem Sie die Haken setzen (Klick auf den Haken oder den Text).\n\n"
            "3. Nutzen Sie die 'Presets'-Buttons, um gängige Modul-Gruppen schnell an- oder abzuwählen.\n\n"
            "4. Mit 'Auswahl zurücksetzen' werden alle optionalen Module abgewählt.\n\n"
            "5. Klicken Sie auf 'Ausgewählte Dateien zusammenfügen', geben Sie die laufende Nummer an und die Zieldatei wird erstellt.\n\n"
            "--- \n"
            "Eigene Presets:\n"
            "Mit 'Presets bearbeiten' können Sie die 5 Preset-Buttons an Ihre Bedürfnisse anpassen.\n\n"
            "Konfiguration (Code/Jahr):\n"
            "Um den regionalen Code oder das Jahr (oben rechts) zu ändern, löschen Sie die Datei 'config.ini' im Ordner 'configs' und starten Sie das Programm neu."
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
        """
        Gathers selected files, asks for file name details,
        and triggers the document merge process.
        """
        selected_files = []
        for item in self.checkbox_items:
            if item["check_var"].get():
                selected_files.append(item["path"])

        if not selected_files:
            messagebox.showwarning("Keine Auswahl", "Es ist keine Datei ausgewählt.")
            return

        # Ensure output directory exists
        try:
            os.makedirs(self.output_dir, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Fehler", f"Konnte den Output-Ordner nicht erstellen:\n{e}")
            return

        # --- NEW SAVING LOGIC ---

        # 1. Open custom dialog for serial number and type
        dialog = FileNameDialog(self.root, "Dateiname festlegen")

        # 2. Check if user clicked "OK" (dialog.result) or "Cancel" (None)
        if not dialog.result:
            return  # User cancelled

        doc_type, serial_num = dialog.result

        # 3. Assemble filename from config and dialog results
        file_name = f"{doc_type} F{self.settings['regional_code']} {serial_num}-{self.settings['year']}.docx"
        save_path = os.path.join(self.output_dir, file_name)

        # 4. Check if file already exists
        if os.path.exists(save_path):
            if not messagebox.askyesno("Warnung",
                                       f"Die Datei:\n{file_name}\n\nexistiert bereits im Ordner 'output'.\nSoll sie überschrieben werden?",
                                       parent=self.root):
                return  # User chose not to overwrite

        # --- END NEW SAVING LOGIC ---

        # Run the merge process
        try:
            self.start_button.config(text="Arbeite...", state="disabled")
            self.root.update_idletasks()

            self.merge_documents(selected_files, save_path)

            messagebox.showinfo("Erfolg", f"Dateien erfolgreich zusammengefügt!\nGespeichert als: {save_path}")
        except Exception as e:
            messagebox.showerror("Fehler", f"Ein Fehler ist aufgetreten:\n{e}\n\n"
                                           f"Hinweis: Stellen Sie sicher, dass die Zieldatei (falls sie existiert) geschlossen ist.")
        finally:
            # Restore button state
            self.start_button.config(text="Ausgewählte Dateien zusammenfügen", state="normal")

    def merge_documents(self, file_paths, save_path):
        """
        Merges a list of .docx files into a single document at 'save_path'.
        The first file in the list is used as the base document.
        """
        if not file_paths:
            return

        if not os.path.exists(file_paths[0]):
            raise FileNotFoundError(f"Die Basis-Datei konnte nicht gefunden werden: {file_paths[0]}")

        # Use the first selected file as the master
        master_doc = Document(file_paths[0])
        composer = Composer(master_doc)

        # Append all subsequent files
        if len(file_paths) > 1:
            for file_path in file_paths[1:]:
                if not os.path.exists(file_path):
                    print(f"Warning: Skipping file (not found): {file_path}")
                    continue
                try:
                    doc_to_append = Document(file_path)
                    composer.append(doc_to_append)
                except Exception as inner_exception:
                    # Log error but try to continue
                    print(f"Error appending {file_path}: {inner_exception}")
                    pass

        composer.save(save_path)


if __name__ == "__main__":
    root = tk.Tk()
    app = WordMergerApp(root)
    root.mainloop()