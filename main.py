import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import os
import glob
import sys
import re
import subprocess
import configparser
from datetime import datetime
import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement, ns
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
    "4.2.3 - Abweichungen vom geplanten Bauablauf.docx",
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
    "9.1.0 - Anlagen - Zugestimmt - Verteiler.docx"
]

NUM_PRESETS = 5
APP_VERSION = "2.8"

# --- UPDATE PFAD KONFIGURATION ---
SHAREPOINT_RELATIVE_PATH = r"Deutsche Bahn\BuS Hagen - Dokumente\Betra\BetraTool (in Arbeit)\_Update"
UPDATE_FILE_NAME = "version.txt"

COLUMN_LAYOUT = {
    "0": 0, "1": 0, "2": 0,       # Spalte 1
    "3": 1, "4": 1,               # Spalte 2
    "5.0": 2, "5.1": 2,           # Spalte 3
    "5.2": 3,                     # Spalte 4
    "5.3": 4,                     # Spalte 5
    "5.4": 5, "6": 5,             # Spalte 6
    "7": 6, "8": 6, "9": 6, "10": 6, "Unsorted": 6 # Spalte 7
}
NUM_MAIN_COLUMNS = 7


def natural_sort_key(s):
    """Sorts strings 'naturally' handling embedded numbers."""
    filename = os.path.basename(s)
    return [int(c) if c.isdigit() else c.lower() for c in re.split('([0-9]+)', filename)]


def sanitize_filename(text):
    """Removes illegal characters for filenames."""
    if not text:
        return ""
    return re.sub(r'[\\/*?:"<>|]', "", text).strip()


class ToolTip(object):
    """Creates a tooltip for a given widget."""
    active_instance = None 

    def __init__(self, widget, text='widget info'):
        self.waittime = 500
        self.wraplength = 400
        self.widget = widget
        self.text = text
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)
        self.widget.bind("<ButtonPress>", self.leave)
        self.id = None
        self.tw = None

    def enter(self, event=None):
        self.schedule()

    def leave(self, event=None):
        self.unschedule()
        self.hidetip()

    def schedule(self):
        self.unschedule()
        self.id = self.widget.after(self.waittime, self.showtip)

    def unschedule(self):
        id = self.id
        self.id = None
        if id:
            try:
                self.widget.after_cancel(id)
            except Exception:
                pass

    def showtip(self):
        if ToolTip.active_instance and ToolTip.active_instance != self:
            ToolTip.active_instance.hidetip()
            
        if self.tw:
            self.hidetip()
            
        try:
            x = self.widget.winfo_rootx() + 20
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5
            
            self.tw = tk.Toplevel(self.widget)
            self.tw.wm_overrideredirect(True) 
            self.tw.wm_geometry("+%d+%d" % (x, y))
            self.tw.lift() 
            
            label = tk.Label(self.tw, text=self.text, justify='left',
                           background="#ffffe0", relief='solid', borderwidth=1,
                           wraplength=self.wraplength,
                           font=("tahoma", "9", "normal"), padx=5, pady=2)
            label.pack(ipadx=1)
            
            ToolTip.active_instance = self
            
        except Exception as e:
            pass

    def hidetip(self):
        if ToolTip.active_instance == self:
            ToolTip.active_instance = None
            
        tw = self.tw
        self.tw = None
        if tw:
            try:
                tw.destroy()
            except:
                pass


class FileNameDialog(simpledialog.Dialog):
    """Dialog to ask for document type, serial number, location, desc, AEL and Kompensation."""
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
        type_frame.pack(pady=5, fill="x")

        num_frame = ttk.Frame(frame)
        ttk.Label(num_frame, text="Laufende Nummer (YYYY):").pack(side=tk.LEFT, padx=5)
        self.entry_var = tk.StringVar()
        self.entry_widget = ttk.Entry(num_frame, textvariable=self.entry_var, width=10)
        self.entry_widget.pack(side=tk.LEFT)
        num_frame.pack(pady=5, fill="x")

        loc_frame = ttk.Frame(frame)
        ttk.Label(loc_frame, text="Ort / Betriebsstelle:").pack(side=tk.LEFT, padx=5)
        self.location_var = tk.StringVar()
        ttk.Entry(loc_frame, textvariable=self.location_var, width=30).pack(side=tk.LEFT)
        loc_frame.pack(pady=5, fill="x")

        desc_frame = ttk.Frame(frame)
        ttk.Label(desc_frame, text="Maßnahme (Kurz):").pack(side=tk.LEFT, padx=5)
        self.desc_var = tk.StringVar()
        ttk.Entry(desc_frame, textvariable=self.desc_var, width=30).pack(side=tk.LEFT)
        desc_frame.pack(pady=5, fill="x")

        opts_frame = ttk.Frame(frame)
        opts_frame.pack(fill="x", pady=10)

        self.ael_var = tk.BooleanVar(value=False)
        ael_check = ttk.Checkbutton(opts_frame, text="AEL-Verrechnung", variable=self.ael_var)
        ael_check.pack(side=tk.LEFT, padx=5)

        self.komp_var = tk.BooleanVar(value=False)
        komp_check = ttk.Checkbutton(opts_frame, text="Kompensationsmaßnahmen", variable=self.komp_var)
        komp_check.pack(side=tk.LEFT, padx=5)

        return self.entry_widget

    def validate(self):
        serial_num = self.entry_var.get().strip()
        if not serial_num:
            messagebox.showwarning("Eingabe fehlt", "Bitte eine laufende Nummer eingeben.", parent=self)
            return 0
        return 1

    def apply(self):
        self.result = (
            self.doc_type_var.get(),
            self.entry_var.get().strip(),
            self.location_var.get().strip(),
            self.desc_var.get().strip(),
            self.ael_var.get(),
            self.komp_var.get()
        )


class InitialConfigDialog(simpledialog.Dialog):
    def __init__(self, parent, title, network_data):
        self.network_data = network_data
        self.result = None
        super().__init__(parent, title)

    def body(self, frame):
        self.rb_var = tk.StringVar()
        self.network_var = tk.StringVar()
        self.user_name_var = tk.StringVar()

        rb_frame = ttk.Frame(frame)
        ttk.Label(rb_frame, text="Regionalbereich (RB):").pack(side=tk.LEFT, padx=5, pady=5)
        self.rb_combo = ttk.Combobox(rb_frame, textvariable=self.rb_var, state="readonly", width=30)
        self.rb_combo['values'] = sorted(list(self.network_data.keys()))
        self.rb_combo.pack(side=tk.LEFT, padx=5, pady=5)
        rb_frame.pack()
        
        network_frame = ttk.Frame(frame)
        ttk.Label(network_frame, text="Netz auswählen:").pack(side=tk.LEFT, padx=5, pady=5)
        self.network_combo = ttk.Combobox(network_frame, textvariable=self.network_var, state="disabled", width=30)
        self.network_combo.pack(side=tk.LEFT, padx=5, pady=5)
        network_frame.pack()

        name_frame = ttk.Frame(frame)
        ttk.Label(name_frame, text="Ihr Name (für AEL-Export):").pack(side=tk.LEFT, padx=5, pady=5)
        self.name_entry = ttk.Entry(name_frame, textvariable=self.user_name_var, width=32)
        self.name_entry.pack(side=tk.LEFT, padx=5, pady=5)
        name_frame.pack()

        year_info_label = ttk.Label(frame, text="Das Jahr ist fest auf 2026 eingestellt (Module 2026).", font=("-default-", 9, "italic"))
        year_info_label.pack(pady=(10, 0))

        self.rb_combo.bind("<<ComboboxSelected>>", self._on_rb_selected)
        
        return self.rb_combo

    def _on_rb_selected(self, event=None):
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
        if not self.user_name_var.get().strip():
            messagebox.showwarning("Eingabe fehlt", "Bitte Ihren Namen für den AEL-Export eingeben.", parent=self)
            return 0
        return 1

    def apply(self):
        try:
            full_network_string = self.network_var.get()
            parts = full_network_string.split(" - ", 1)
            code_full = parts[0].strip()
            name = parts[1].strip()
            user_name = self.user_name_var.get().strip()
            
            self.result = (code_full, name, user_name)
        except Exception as e:
            print(f"Dialog apply error: {e}")
            messagebox.showerror("Fehler", "Auswahl konnte nicht verarbeitet werden.", parent=self)


class AelDetailsDialog(simpledialog.Dialog):
    def __init__(self, parent, title):
        self.result = None
        super().__init__(parent, title)

    def body(self, frame):
        proj_frame = ttk.Frame(frame)
        ttk.Label(proj_frame, text="Projektnummer/Planungs-AIB:").pack(side=tk.LEFT, padx=5, pady=5)
        self.proj_var = tk.StringVar()
        self.proj_entry = ttk.Entry(proj_frame, textvariable=self.proj_var, width=20)
        self.proj_entry.pack(side=tk.LEFT, padx=5, pady=5)
        proj_frame.pack()

        kurz_frame = ttk.Frame(frame)
        ttk.Label(kurz_frame, text="Kurztext (optional):").pack(side=tk.LEFT, padx=5, pady=5)
        self.kurz_var = tk.StringVar()
        kurz_entry = ttk.Entry(kurz_frame, textvariable=self.kurz_var, width=20)
        kurz_entry.pack(side=tk.LEFT, padx=5, pady=5)
        kurz_frame.pack()
        
        sonst_frame = ttk.Frame(frame)
        ttk.Label(sonst_frame, text="Sonstiges (optional):").pack(side=tk.LEFT, padx=5, pady=5)
        self.sonstiges_var = tk.StringVar()
        sonst_entry = ttk.Entry(sonst_frame, textvariable=self.sonstiges_var, width=20)
        sonst_entry.pack(side=tk.LEFT, padx=5, pady=5)
        sonst_frame.pack()

        dritte_frame = ttk.Frame(frame)
        self.dritte_var = tk.BooleanVar(value=False)
        dritte_check = ttk.Checkbutton(dritte_frame, text="Leistung für Dritte?", variable=self.dritte_var)
        dritte_check.pack(side=tk.LEFT, padx=5, pady=10)
        dritte_frame.pack()

        return self.proj_entry

    def validate(self):
        if not self.proj_var.get().strip():
            messagebox.showwarning("Eingabe fehlt", "Bitte eine Projektnummer eingeben.", parent=self)
            return 0
        return 1

    def apply(self):
        self.result = (
            self.proj_var.get().strip(),
            self.kurz_var.get().strip(),
            self.dritte_var.get(),
            self.sonstiges_var.get().strip()
        )


class KompensationsDialog(simpledialog.Dialog):
    def __init__(self, parent, title):
        self.result = None
        super().__init__(parent, title)

    def body(self, frame):
        # 1. In Kraft ab
        ttk.Label(frame, text="In Kraft ab:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.in_kraft_datum = tk.StringVar()
        self.in_kraft_datum_entry = ttk.Entry(frame, textvariable=self.in_kraft_datum, width=15)
        self.in_kraft_datum_entry.grid(row=0, column=1, padx=5, pady=2)
        ttk.Label(frame, text="(Datum)").grid(row=0, column=2, sticky="w")
        
        self.in_kraft_uhrzeit = tk.StringVar()
        ttk.Entry(frame, textvariable=self.in_kraft_uhrzeit, width=10).grid(row=0, column=3, padx=5, pady=2)
        ttk.Label(frame, text="(Uhrzeit)").grid(row=0, column=4, sticky="w")

        # 2. Außer Kraft ab
        ttk.Label(frame, text="Außer Kraft ab:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.aus_kraft_datum = tk.StringVar()
        ttk.Entry(frame, textvariable=self.aus_kraft_datum, width=15).grid(row=1, column=1, padx=5, pady=2)
        ttk.Label(frame, text="(Datum)").grid(row=1, column=2, sticky="w")
        
        self.aus_kraft_uhrzeit = tk.StringVar()
        ttk.Entry(frame, textvariable=self.aus_kraft_uhrzeit, width=10).grid(row=1, column=3, padx=5, pady=2)
        ttk.Label(frame, text="(Uhrzeit)").grid(row=1, column=4, sticky="w")

        # 3. Inhalt
        ttk.Label(frame, text="Inhalt (Maßnahme):").grid(row=2, column=0, sticky="w", padx=5, pady=2)
        self.inhalt = tk.StringVar()
        ttk.Entry(frame, textvariable=self.inhalt, width=40).grid(row=2, column=1, columnspan=4, sticky="w", padx=5, pady=2)

        # 4. Antragsnummer
        ttk.Label(frame, text="Antragsnummer (6-stellig):").grid(row=3, column=0, sticky="w", padx=5, pady=2)
        self.antrag_nr = tk.StringVar()
        ttk.Entry(frame, textvariable=self.antrag_nr, width=15).grid(row=3, column=1, padx=5, pady=2)

        # 5. Arbeitszeiten
        ttk.Label(frame, text="Arbeitszeiten:").grid(row=4, column=0, sticky="w", padx=5, pady=2)
        self.arbeitszeit = tk.StringVar()
        ttk.Entry(frame, textvariable=self.arbeitszeit, width=40).grid(row=4, column=1, columnspan=4, sticky="w", padx=5, pady=2)

        # 6. Folgen
        ttk.Label(frame, text="Folgen bei Nichtzulassen:").grid(row=5, column=0, sticky="w", padx=5, pady=2)
        self.folgen = tk.StringVar()
        ttk.Entry(frame, textvariable=self.folgen, width=40).grid(row=5, column=1, columnspan=4, sticky="w", padx=5, pady=2)

        return self.in_kraft_datum_entry

    def validate(self):
        if not self.antrag_nr.get():
             messagebox.showwarning("Eingabe fehlt", "Bitte Antragsnummer eingeben.", parent=self)
             return 0
        return 1

    def apply(self):
        self.result = {
            "InKraftAbDatum": self.in_kraft_datum.get(),
            "InKraftAbUhrzeit": self.in_kraft_uhrzeit.get(),
            "AußerKraftAbDatum": self.aus_kraft_datum.get(),
            "AußerKraftAbUhrzeit": self.aus_kraft_uhrzeit.get(),
            "Inhalt": self.inhalt.get(),
            "Antragsnummer": self.antrag_nr.get(),
            "Arbeitszeit": self.arbeitszeit.get(),
            "Folgen": self.folgen.get()
        }


class WordMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("BetraTool v" + APP_VERSION)
        
        try:
            self.root.state('zoomed')
        except:
            self.root.geometry("1410x700")

        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))

        self.load_icon(base_path)

        self.modules_dir = os.path.join(base_path, "modules")
        self.output_dir = os.path.join(base_path, "output")
        self.configs_dir = os.path.join(base_path, "configs")
        self.config_file_path = os.path.join(self.configs_dir, "config.ini")
        self.presets_file_path = os.path.join(self.configs_dir, "presets.ini")
        self.module_infos_file_path = os.path.join(self.configs_dir, "ModulInfos.ini")
        self.network_data_file_path = os.path.join(self.configs_dir, "BetraNetzziffern.txt")
        
        self.preset_config = configparser.ConfigParser()
        self.module_info_config = configparser.ConfigParser()
        self.presets = {}
        self.module_infos = {} 
        self.config = configparser.ConfigParser()
        self.settings = {}
        
        self.network_data = {} 
        self.cover_pages = [] 
        self.selected_cover_page = tk.StringVar()
        self.checkbox_items = []
        self.module_files = [] 

        self.load_or_create_network_data()
        self.load_or_create_config()
        self.load_or_create_presets() 
        self.load_or_create_module_infos()
        
        self.create_main_widgets()
        self.load_files()

    def load_icon(self, base_path):
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

    def check_for_updates(self):
        """Prüft im OneDrive-Ordner nach einer neueren Version."""
        try:
            user_profile = os.environ.get('USERPROFILE')
            update_folder = os.path.join(user_profile, SHAREPOINT_RELATIVE_PATH)
            update_path = os.path.join(update_folder, UPDATE_FILE_NAME)
            
            if not os.path.exists(update_path):
                return

            with open(update_path, 'r') as f:
                server_version = f.read().strip()

            try:
                local_clean = re.sub(r'[a-zA-Z]', '', APP_VERSION)
                server_clean = re.sub(r'[a-zA-Z]', '', server_version)
                
                if float(server_clean) > float(local_clean):
                    response = messagebox.askyesno(
                        "Update verfügbar",
                        f"Eine neue Version ({server_version}) ist verfügbar!\n"
                        f"Sie nutzen Version {APP_VERSION}.\n\n"
                        "Möchten Sie den Update-Ordner öffnen?"
                    )
                    if response:
                        os.startfile(update_folder)
            except Exception:
                pass 

        except Exception as e:
            print(f"Update Check Error: {e}")

    def create_main_widgets(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        top_info_frame = ttk.Frame(main_frame)
        top_info_frame.pack(fill=tk.X, anchor="n", pady=(0, 5)) 

        info_label = ttk.Label(top_info_frame, text=f"Module aus: '{self.modules_dir}'", font=("-default-", 9, "italic"))
        info_label.pack(side=tk.LEFT, anchor="w")

        year_short = self.settings.get('year', '??')
        year_display = f"20{year_short}" if year_short.isdigit() else "??"
        
        config_text = f"Aktuelle Konfiguration: {self.settings.get('regional_code_full', '??')} ({self.settings.get('network_name', '???')}), Jahr: {year_display}, Bearbeiter: {self.settings.get('user_name', '???')}"
        
        self.config_label = ttk.Label(top_info_frame, text=config_text, font=("-default-", 9, "italic"))
        self.config_label.pack(side=tk.RIGHT, anchor="e")

        cover_page_frame = ttk.Frame(main_frame)
        cover_page_frame.pack(fill=tk.X, anchor="n", pady=(0, 10))
        
        cover_label = ttk.Label(cover_page_frame, text="Deckblatt auswählen:", font=("-default-", 10, "bold"))
        cover_label.pack(side=tk.LEFT, anchor="w")
        
        self.cover_page_combo = ttk.Combobox(cover_page_frame, textvariable=self.selected_cover_page, state="readonly", width=60)
        self.cover_page_combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5,0))

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

        category_frame = ttk.Frame(main_frame)
        category_frame.pack(fill=tk.X, pady=(10, 5))

        category_label = ttk.Label(category_frame, text="Presets (An/Aus):")
        category_label.pack(fill=tk.X, pady=(0, 4))

        self.preset_btn_container = ttk.Frame(category_frame)
        self.preset_btn_container.pack(fill=tk.X)
        
        self.create_preset_buttons() 

        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))

        self.help_button = ttk.Button(button_frame, text="Anleitung", command=self.show_help)
        self.help_button.pack(side=tk.LEFT)

        self.contact_button = ttk.Button(button_frame, text="Kontakt", command=self.show_contact)
        self.contact_button.pack(side=tk.LEFT, padx=5)

        self.open_output_button = ttk.Button(button_frame, text="Output öffnen", command=self.open_output_folder)
        self.open_output_button.pack(side=tk.LEFT, padx=5)

        self.map_button = ttk.Button(button_frame, text="Netzkarte", command=self.open_netzkarte)
        self.map_button.pack(side=tk.LEFT, padx=5)

        self.reset_button = ttk.Button(button_frame, text="Auswahl zurücksetzen", command=self.reset_selection)
        self.reset_button.pack(side=tk.LEFT, padx=(5, 0))
        
        self.edit_presets_button = ttk.Button(button_frame, text="Presets bearbeiten", command=self.open_preset_editor)
        self.edit_presets_button.pack(side=tk.LEFT, padx=(5, 0))

        self.start_button = ttk.Button(button_frame, text="Ausgewählte Dateien zusammenfügen", command=self.start_merge)
        self.start_button.pack(side=tk.RIGHT)
        self.start_button["state"] = "disabled"
        
        self.root.after(2000, self.check_for_updates)

    def open_output_folder(self):
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir, exist_ok=True)
        
        try:
            os.startfile(self.output_dir)
        except AttributeError:
            import subprocess
            if sys.platform == 'darwin':
                subprocess.call(['open', self.output_dir])
            else:
                subprocess.call(['xdg-open', self.output_dir])
        except Exception as e:
            messagebox.showerror("Fehler", f"Konnte Ordner nicht öffnen:\n{e}")

    def open_netzkarte(self):
        """Startet die externe Netzkarte.exe Anwendung."""
        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))

        exe_path = os.path.join(base_path, "netzkarte.exe")

        if os.path.exists(exe_path):
            try:
                # Popen startet das Programm unabhängig (non-blocking)
                subprocess.Popen([exe_path])
            except Exception as e:
                messagebox.showerror("Fehler", f"Konnte Netzkarte nicht starten:\n{e}")
        else:
            messagebox.showerror("Fehler", f"Die Datei 'netzkarte.exe' wurde nicht gefunden.\nPfad: {exe_path}")

    def load_or_create_network_data(self):
        os.makedirs(self.configs_dir, exist_ok=True)
        if not os.path.exists(self.network_data_file_path):
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
                messagebox.showerror("Kritischer Fehler", f"Konnte '{self.network_data_file_path}' nicht erstellen: {e}")
                self.root.quit()
                return

        try:
            with open(self.network_data_file_path, 'r', encoding='utf-8') as f:
                current_rb = None
                for line in f:
                    line = line.strip()
                    if not line:
                        continue
                    if "," not in line:
                        current_rb = line
                        if current_rb not in self.network_data:
                            self.network_data[current_rb] = {}
                    else:
                        if current_rb is None:
                            continue 
                        parts = line.split(",", 1)
                        if len(parts) == 2:
                            code = parts[0].strip()
                            name = parts[1].strip()
                            self.network_data[current_rb][code] = name
        except Exception as e:
            messagebox.showerror("Kritischer Fehler", f"Konnte '{self.network_data_file_path}' nicht lesen: {e}")
            self.root.quit()

    def load_or_create_config(self):
        os.makedirs(self.configs_dir, exist_ok=True)
        try:
            if not os.path.exists(self.config_file_path):
                raise FileNotFoundError("Config file not found.")
            
            self.config.read(self.config_file_path, encoding='utf-8')
            
            if 'SETTINGS' not in self.config or \
               'RegionalCodeFull' not in self.config['SETTINGS'] or \
               'NetworkName' not in self.config['SETTINGS'] or \
               'Year' not in self.config['SETTINGS'] or \
               'UserName' not in self.config['SETTINGS']:
                raise ValueError("Config file is incomplete.")

            self.settings['regional_code_full'] = self.config['SETTINGS']['RegionalCodeFull']
            self.settings['network_name'] = self.config['SETTINGS']['NetworkName']
            self.settings['year'] = self.config['SETTINGS']['Year']
            self.settings['user_name'] = self.config['SETTINGS']['UserName']
            
            if self.settings['year'] != "26":
                self.settings['year'] = "26"
                self.config['SETTINGS']['Year'] = "26"
                with open(self.config_file_path, 'w', encoding='utf-8') as configfile:
                    self.config.write(configfile)

            if not self.settings['regional_code_full'] or not self.settings['network_name'] or not self.settings['user_name']:
                raise ValueError("Config values are empty.")

        except Exception as e:
            self.settings = {'regional_code_full': '??', 'network_name': '???', 'year': '26', 'user_name': '???'} 
            self.root.after_idle(self.ask_for_initial_config)

    def load_or_create_presets(self):
        os.makedirs(self.configs_dir, exist_ok=True)
        try:
            if not os.path.exists(self.presets_file_path):
                raise FileNotFoundError("Presets file not found.")
            
            self.preset_config.read(self.presets_file_path, encoding='utf-8')
            
            for i in range(1, NUM_PRESETS + 1):
                section = f'PRESET_{i}'
                if section not in self.preset_config:
                    raise ValueError(f"Preset section {section} missing.")
                
                name = self.preset_config[section]['Name']
                if 'Bausteine' in self.preset_config[section]:
                    modules = self.preset_config[section]['Bausteine']
                    self.preset_config[section]['Modules'] = modules
                    del self.preset_config[section]['Bausteine']
                    with open(self.presets_file_path, 'w', encoding='utf-8') as f:
                        self.preset_config.write(f)
                else:
                    modules = self.preset_config[section]['Modules']
                    
                self.presets[section] = {'Name': name, 'Modules': modules}

            if len(self.presets) < NUM_PRESETS:
                raise ValueError("Not all presets were found.")

        except Exception as e:
            self.create_default_presets()

    def create_default_presets(self):
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
            with open(self.presets_file_path, 'w', encoding='utf-8') as f:
                self.preset_config.write(f)
        except Exception as e:
            print(f"Could not save default presets: {e}")

    def load_or_create_module_infos(self):
        os.makedirs(self.configs_dir, exist_ok=True)
        self.module_info_config.read(self.module_infos_file_path, encoding='utf-8')
        
        if not self.module_info_config.has_section('INFOS'):
            self.module_info_config['INFOS'] = {}
            
            if not os.path.exists(self.module_infos_file_path):
                search_path = os.path.join(self.modules_dir, "*.docx")
                all_files = glob.glob(search_path)
                for f in all_files:
                    fname = os.path.basename(f)
                    self.module_info_config['INFOS'][fname] = "Hier Info-Text in ModulInfos.ini eintragen" 
                
                try:
                    with open(self.module_infos_file_path, 'w', encoding='utf-8') as f:
                        self.module_info_config.write(f)
                except Exception as e:
                    print(f"Could not create ModulInfos.ini: {e}")

        if self.module_info_config.has_section('INFOS'):
            for key, value in self.module_info_config['INFOS'].items():
                if value.strip():
                    self.module_infos[key.lower()] = value

    def create_preset_buttons(self):
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
        self.editor_window = tk.Toplevel(self.root)
        self.editor_window.title("Preset-Editor")
        self.editor_window.transient(self.root)
        self.editor_window.grab_set()
        self.editor_window.resizable(False, False)
        self.editor_window.geometry("600x500")

        notebook = ttk.Notebook(self.editor_window)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.preset_editor_data = {}

        for i in range(1, NUM_PRESETS + 1):
            section = f'PRESET_{i}'
            preset_data = self.presets[section]
            
            tab_frame = ttk.Frame(notebook, padding=10)
            notebook.add(tab_frame, text=f"Preset {i}")
            
            name_frame = ttk.Frame(tab_frame)
            name_frame.pack(fill=tk.X, pady=(0, 10))
            ttk.Label(name_frame, text="Button-Name:").pack(side=tk.LEFT)
            
            name_var = tk.StringVar(value=preset_data['Name'])
            ttk.Entry(name_frame, textvariable=name_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))
            
            list_label = ttk.Label(tab_frame, text="Bausteine auswählen:")
            list_label.pack(anchor="w", pady=(0, 5))

            list_container = ttk.Frame(tab_frame, borderwidth=1, relief="sunken")
            list_container.pack(fill=tk.BOTH, expand=True)
            
            canvas = tk.Canvas(list_container)
            scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=canvas.yview)
            scrollable_frame = ttk.Frame(canvas)
            
            scrollable_frame.bind(
                "<Configure>",
                lambda e, c=canvas: c.configure(scrollregion=c.bbox("all"))
            )
            
            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)
            
            canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

            def _on_mousewheel_editor(event, c=canvas):
                if sys.platform == "linux":
                    if event.num == 4: c.yview_scroll(-1, "units")
                    elif event.num == 5: c.yview_scroll(1, "units")
                else:
                    c.yview_scroll(int(-1 * (event.delta / 120)), "units")
            
            canvas.bind("<MouseWheel>", _on_mousewheel_editor)
            scrollable_frame.bind("<MouseWheel>", _on_mousewheel_editor)

            current_modules_str = preset_data['Modules']
            current_prefixes = [p.strip() for p in current_modules_str.split(',') if p.strip()]
            
            checkboxes = []
            
            for file_path in self.module_files:
                filename = os.path.basename(file_path)
                display_name = os.path.splitext(filename)[0]
                
                is_checked = False
                for prefix in current_prefixes:
                    # Match full filename start to handle partial matches like "5.1" vs "5.1.1"
                    if filename.startswith(prefix):
                        is_checked = True
                        break
                
                var = tk.BooleanVar(value=is_checked)
                cb = ttk.Checkbutton(scrollable_frame, text=display_name, variable=var)
                cb.pack(anchor="w", padx=5, pady=1)
                cb.bind("<MouseWheel>", _on_mousewheel_editor)

                checkboxes.append({'path': file_path, 'var': var, 'filename': filename})

            self.preset_editor_data[i] = {
                'name_var': name_var,
                'checks': checkboxes
            }

        btn_frame = ttk.Frame(self.editor_window, padding=(0, 10, 0, 10))
        btn_frame.pack(fill=tk.X, padx=10)
        
        ttk.Button(btn_frame, text="Speichern", command=self.save_presets_graphical).pack(side=tk.RIGHT)
        ttk.Button(btn_frame, text="Abbrechen", command=self.editor_window.destroy).pack(side=tk.RIGHT, padx=5)

    def save_presets_graphical(self):
        try:
            for i in range(1, NUM_PRESETS + 1):
                section = f'PRESET_{i}'
                data = self.preset_editor_data[i]
                
                name = data['name_var'].get().strip()
                if not name:
                    messagebox.showerror("Fehler", f"Der Name für Preset {i} darf nicht leer sein.", parent=self.editor_window)
                    return

                selected_prefixes = []
                for item in data['checks']:
                    if item['var'].get():
                        # Use the starting number as the ID/prefix
                        parts = item['filename'].split(' - ')
                        if parts:
                            prefix = parts[0]
                            selected_prefixes.append(prefix)
                        else:
                            selected_prefixes.append(item['filename']) 

                modules_str = ", ".join(selected_prefixes)
                
                self.preset_config[section]['Name'] = name
                self.preset_config[section]['Modules'] = modules_str
                self.presets[section] = {'Name': name, 'Modules': modules_str}

            with open(self.presets_file_path, 'w', encoding='utf-8') as f:
                self.preset_config.write(f)

            self.create_preset_buttons()
            self.editor_window.destroy()
            messagebox.showinfo("Gespeichert", "Presets erfolgreich aktualisiert.", parent=self.root)

        except Exception as e:
            messagebox.showerror("Fehler", f"Presets konnten nicht gespeichert werden:\n{e}", parent=self.editor_window)

    def save_presets(self):
        try:
            for i in range(1, NUM_PRESETS + 1):
                section = f'PRESET_{i}'
                name = self.preset_name_vars[i-1].get()
                modules = self.preset_modules_vars[i-1].get()

                if not name.strip():
                    messagebox.showerror("Fehler", f"Der Name für Preset {i} darf nicht leer sein.", parent=self.editor_window)
                    return
                
                self.preset_config[section]['Name'] = name
                self.preset_config[section]['Modules'] = modules
                self.presets[section] = {'Name': name, 'Modules': modules}

            with open(self.presets_file_path, 'w', encoding='utf-8') as f:
                self.preset_config.write(f)

            self.create_preset_buttons()
            self.editor_window.destroy()
            messagebox.showinfo("Gespeichert", "Presets erfolgreich aktualisiert.", parent=self.root)

        except Exception as e:
            messagebox.showerror("Fehler", f"Presets konnten nicht gespeichert werden:\n{e}", parent=self.editor_window)
            
            
    def ask_for_initial_config(self):
        if not self.network_data:
             messagebox.showerror("Kritischer Fehler", "Netzwerkdaten sind nicht geladen. Konfiguration nicht möglich.")
             self.root.quit()
             return
             
        messagebox.showinfo("Erstkonfiguration", 
                            "Willkommen! Bitte gib dein Standard-Netz und deinen Namen ein.", 
                            parent=self.root)
        
        dialog = InitialConfigDialog(self.root, "Erstkonfiguration", self.network_data)
        
        if not dialog.result:
            messagebox.showerror("Abbruch", "Ohne Konfiguration kann das Programm nicht starten.")
            self.root.quit()
            return
            
        code_full, name, user_name = dialog.result
        year_short = "26"

        self.config['SETTINGS'] = {
            'RegionalCodeFull': code_full,
            'NetworkName': name,
            'Year': year_short,
            'UserName': user_name
        }
        with open(self.config_file_path, 'w', encoding='utf-8') as configfile:
            self.config.write(configfile)

        self.settings['regional_code_full'] = code_full
        self.settings['network_name'] = name
        self.settings['year'] = year_short
        self.settings['user_name'] = user_name
        messagebox.showinfo("Konfiguration", "Einstellungen erfolgreich gespeichert.", parent=self.root)
        
        if hasattr(self, 'config_label'):
            year_short = self.settings.get('year', '??')
            year_display = f"20{year_short}" if year_short.isdigit() else "??"
            config_text = f"Aktuelle Konfiguration: {self.settings.get('regional_code_full', '??')} ({self.settings.get('network_name', '???')}), Jahr: {year_display}, Bearbeiter: {self.settings.get('user_name', '???')}"
            self.config_label.config(text=config_text)

    def _on_mousewheel(self, event):
        if sys.platform == "linux":
            if event.num == 4:
                self.canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                self.canvas.yview_scroll(1, "units")
        else:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _on_label_click(self, checkbox_widget, check_var):
        if checkbox_widget.instate(['!disabled']):
            check_var.set(not check_var.get())

    def _get_layout_key(self, filename):
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
        if not os.path.isdir(self.modules_dir):
            messagebox.showerror("Fehler", f"Der Ordner '{self.modules_dir}' wurde nicht gefunden.")
            self.root.quit()
            return

        for item in self.checkbox_items:
            item["checkbox"].master.destroy() 
        self.checkbox_items.clear()
        
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
            
        self.cover_pages.clear()
        self.cover_page_combo['values'] = []
        self.selected_cover_page.set("")

        search_path = os.path.join(self.modules_dir, "*.docx")
        all_file_paths = glob.glob(search_path)
        
        if not all_file_paths:
            messagebox.showinfo("Keine Dateien", f"Keine .docx Dateien im Ordner '{self.modules_dir}' gefunden.")
            return
            
        cover_page_files = []
        self.module_files = []

        for file_path in all_file_paths:
            filename = os.path.basename(file_path)
            if filename.startswith("0."):
                cover_page_files.append(file_path)
            else:
                self.module_files.append(file_path)
        
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

        self.module_files.sort(key=natural_sort_key)

        if not self.module_files and cover_page_files:
            messagebox.showinfo("Keine Module", f"Keine Modul-Dateien (außer Deckblättern) im Ordner '{self.modules_dir}' gefunden.")
        
        wrap_length_pixels = 218
        main_columns = []
        for i in range(NUM_MAIN_COLUMNS):
            main_col_frame = ttk.Frame(self.scrollable_frame)
            main_col_frame.grid(row=0, column=i, sticky="nw", padx=5)
            main_col_frame.bind("<MouseWheel>", self._on_mousewheel)
            main_columns.append(main_col_frame)

        group_frames = {}

        for file_path in self.module_files:
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

            info_text = self.module_infos.get(filename.lower())
            if info_text:
                ToolTip(item_frame, info_text)
                ToolTip(checkbox, info_text)
                ToolTip(label, info_text)

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
        for item in self.checkbox_items:
            if not item["is_mandatory"]:
                item["check_var"].set(False)
            else:
                item["check_var"].set(True)

    def toggle_category(self, prefixes):
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
        help_text = (
            f"Anleitung BetraTool (v{APP_VERSION})\n\n"
            "1. Wählen Sie oben das gewünschte Deckblatt aus der Liste aus.\n\n"
            "2. Pflicht-Module sind bereits ausgewählt und können nicht abgewählt werden.\n\n"
            "3. Wählen Sie optionale Module aus, indem Sie die Haken setzen.\n\n"
            "4. Nutzen Sie die 'Presets'-Buttons, um gängige Modul-Gruppen schnell an- oder abzuwählen.\n\n"
            "5. Mit 'Auswahl zurücksetzen' werden alle optionalen Module abgewählt.\n\n"
            "6. Klicken Sie auf 'Ausgewählte Dateien zusammenfügen', geben Sie die laufende Nummer an.\n"
            "   -> Die Datei wird im 'output'-Ordner in einem eigenen Unterordner gespeichert.\n\n"
            "7. AEL-Verrechnung: Setzen Sie den Haken, um nach dem Speichern Details\n"
            "   (Projekt-Nr., Kurztext, etc.) für die Excel-Liste 'AEL-Verrechnung.xlsx' einzugeben.\n\n"
            "8. Kompensationsmaßnahmen: Setzen Sie den Haken, um ein zusätzliches Dokument\n"
            "   für Kompensationsmaßnahmen zu erstellen und auszufüllen.\n\n"
            "--- \n"
            "Eigene Presets:\n"
            "Mit 'Presets bearbeiten' können Sie die 5 Preset-Buttons an Ihre Bedürfnisse anpassen.\n\n"
            "Konfiguration (Netz/Jahr/Name):\n"
            "Um Ihre Konfiguration zu ändern, löschen Sie die Datei 'config.ini' \n"
            "im Ordner 'configs' und starten Sie das Programm neu."
        )
        messagebox.showinfo("Anleitung", help_text)

    def show_contact(self):
        contact_text = (
            "Kontakt & Support\n\n"
            "Bei Fragen, Problemen, Ideen oder Vorschläge mit dem BetraTool:\n\n"
            "Name: Dennis Heinze, I.IA-W-N-HA-B\n"
            "E-Mail: dennis.heinze@deutschebahn.com\n"
            "Telefon (dienstlich): 0152 33114237\n"
            "Version: " + APP_VERSION + "\n\n"
        )
        messagebox.showinfo("Kontakt", contact_text)

    def fill_kompensations_doc(self, template_path, output_path, data):
        """
        Fills the content controls (SDT) in the Kompensationsmaßnahmen docx.
        """
        if not os.path.exists(template_path):
            messagebox.showerror("Fehler", f"Vorlage nicht gefunden:\n{template_path}")
            return

        # Tags that should NOT be bold
        tags_no_bold = ['Antragsnummer', 'Datum', 'Arbeitszeit']

        try:
            doc = Document(template_path)
            
            # Helper to find sdt tags. python-docx doesn't support sdt natively well.
            # We iterate over the element tree.
            for element in doc.element.body.iter():
                if element.tag.endswith('sdt'):
                    sdtPr = element.find(ns.qn('w:sdtPr'))
                    if sdtPr is not None:
                        tag_element = sdtPr.find(ns.qn('w:tag'))
                        if tag_element is not None:
                            tag_val = tag_element.get(ns.qn('w:val'))
                            
                            if tag_val in data:
                                # Found a matching tag. Replace content.
                                sdtContent = element.find(ns.qn('w:sdtContent'))
                                if sdtContent is not None:
                                    # Clear existing content
                                    sdtContent.clear()
                                    
                                    # Create new run with text
                                    r = OxmlElement('w:r')
                                    
                                    # Run Properties (rPr) container
                                    rPr = OxmlElement('w:rPr')

                                    # 1. Set Font (Global for all fields)
                                    rFonts = OxmlElement('w:rFonts')
                                    rFonts.set(ns.qn('w:ascii'), 'DB Neo Office')
                                    rFonts.set(ns.qn('w:hAnsi'), 'DB Neo Office')
                                    rFonts.set(ns.qn('w:cs'), 'DB Neo Office')
                                    rPr.append(rFonts)
                                    
                                    # 2. Apply Bold if not in exclusion list
                                    if tag_val not in tags_no_bold:
                                        b = OxmlElement('w:b')
                                        rPr.append(b)

                                    # 3. Special case: LfdNr needs font size 16
                                    if tag_val == "LfdNr":
                                        sz = OxmlElement('w:sz')
                                        sz.set(ns.qn('w:val'), '32') # 16pt * 2
                                        rPr.append(sz)
                                        sz_cs = OxmlElement('w:szCs')
                                        sz_cs.set(ns.qn('w:val'), '32')
                                        rPr.append(sz_cs)
                                    
                                    if len(rPr) > 0:
                                        r.append(rPr)

                                    t = OxmlElement('w:t')
                                    t.text = str(data[tag_val])
                                    r.append(t)

                                    # Check context: Inline (inside p) or Block?
                                    parent = element.getparent()
                                    if parent.tag.endswith('}p'): 
                                        # Inline SDT: append run
                                        sdtContent.append(r)
                                    else:
                                        # Block SDT: wrap in paragraph
                                        p = OxmlElement('w:p')
                                        
                                        # Paragraph Properties (pPr)
                                        pPr = OxmlElement('w:pPr')
                                        
                                        # Special case: Inhalt -> Center
                                        if tag_val == "Inhalt":
                                            jc = OxmlElement('w:jc')
                                            jc.set(ns.qn('w:val'), 'center')
                                            pPr.append(jc)
                                        
                                        # Special case: Datum -> Right
                                        elif tag_val == "Datum":
                                            jc = OxmlElement('w:jc')
                                            jc.set(ns.qn('w:val'), 'right')
                                            pPr.append(jc)
                                            
                                        if len(pPr) > 0:
                                            p.append(pPr)
                                            
                                        p.append(r)
                                        sdtContent.append(p)

            doc.save(output_path)

        except Exception as e:
            messagebox.showerror("Fehler (Kompensation)", f"Fehler beim Erstellen des Dokuments:\n{e}")


    def start_merge(self):
        check_modules = ["3.0.1", "5.3.22", "5.3.26", "5.3.27"]
        found_any_special = False

        for item in self.checkbox_items:
            if item["check_var"].get():
                for prefix in check_modules:
                    if item["filename"].startswith(prefix):
                        found_any_special = True
                        break
            if found_any_special:
                break

        if not found_any_special:
            msg = (
                "Hinweis: Es wurde keiner der folgenden Module ausgewählt:\n\n"
                "• 3.0.1 - Punkt 3 entfällt komplett\n"
                "• 5.3.22 - Meldung über die Befahrbarkeit der Gleise\n"
                "• 5.3.26 - Aufheben der Sperrung\n"
                "• 5.3.27 - Aufheben der UV-Sperrung\n\n"
                "Möchten Sie trotzdem fortfahren?"
            )
            if not messagebox.askyesno("Modul-Hinweis", msg, icon="warning"):
                return

        selected_cover_name = self.selected_cover_page.get()
        if not selected_cover_name:
            messagebox.showwarning("Deckblatt fehlt", "Bitte ein Deckblatt aus der Liste auswählen, bevor Sie fortfahren.")
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
        
        for item in self.checkbox_items:
            if item["check_var"].get():
                selected_files_for_merge.append(item["path"])

        if len(selected_files_for_merge) == 1:
            if not messagebox.askyesno("Warnung", 
                                       "Es sind keine Module ausgewählt.\n\nMöchten Sie nur das Deckblatt unter dem neuen Namen speichern?", 
                                       parent=self.root):
                return

        try:
            os.makedirs(self.output_dir, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Fehler", f"Konnte den Output-Ordner nicht erstellen:\n{e}")
            return

        dialog = FileNameDialog(self.root, "Dateiname festlegen")
        if not dialog.result:
            return

        # 6 values returned from dialog
        doc_type, serial_num, location, description, ael_checked, komp_checked = dialog.result
        
        # Sanitize input strings for filename safety
        loc_clean = sanitize_filename(location)
        desc_clean = sanitize_filename(description)
        
        # Construct Base Name
        # Format: Betra F33 1001-26 [Ort] [Maßnahme]
        base_name = f"{doc_type} {self.settings['regional_code_full']} {serial_num}-{self.settings['year']}"
        short_name = base_name
            
        if loc_clean:
            base_name += f" {loc_clean}"
        if desc_clean:
            base_name += f" {desc_clean}"

        new_folder_path = os.path.join(self.output_dir, base_name)
        # For the file itself, we usually keep the same name as the folder or base name
        file_name_with_ext = f"{short_name}.docx"
        save_path = os.path.join(new_folder_path, file_name_with_ext)

        try:
            os.makedirs(new_folder_path, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Fehler", f"Konnte den Unterordner nicht erstellen:\n{new_folder_path}\n\nFehler: {e}")
            return
            
        if os.path.exists(save_path):
            relative_file_name = os.path.join(base_name, file_name_with_ext)
            if not messagebox.askyesno("Warnung", 
                                       f"Die Datei:\n{relative_file_name}\n\nexistiert bereits.\nSoll sie überschrieben werden?", 
                                       parent=self.root):
                return
        
        try:
            self.start_button.config(text="Arbeite...", state="disabled")
            self.root.update_idletasks()

            self.merge_documents(selected_files_for_merge, save_path)
            self.add_footer_to_doc(save_path, short_name)

            messagebox.showinfo("Erfolg", f"Dateien erfolgreich zusammengefügt!\nGespeichert als: {save_path}")
            
            # --- Kompensationsmaßnahmen Workflow ---
            if komp_checked:
                komp_dialog = KompensationsDialog(self.root, "Kompensationsmaßnahmen erfassen")
                if komp_dialog.result:
                    komp_data = komp_dialog.result
                    # Add automatic fields
                    komp_data["Datum"] = datetime.now().strftime("%d.%m.%Y")
                    komp_data["LfdNr"] = serial_num
                    
                    template_path = os.path.join(self.configs_dir, "Kompensationsmaßnahmen.docx")
                    
                    # Neuer Dateiname: Kompensationsmaßnahme <RB> <LfdNr>-<Jahr>.docx
                    # z.B. Kompensationsmaßnahme F33 1234-26.docx
                    komp_filename = f"Kompensationsmaßnahme {self.settings['regional_code_full']} {serial_num}-{self.settings['year']}.docx"
                    komp_output_path = os.path.join(new_folder_path, komp_filename)
                    
                    self.fill_kompensations_doc(template_path, komp_output_path, komp_data)
                    messagebox.showinfo("Info", f"Kompensations-Dokument erstellt:\n{komp_output_path}")

            # --- AEL Workflow ---
            if ael_checked:
                ael_dialog = AelDetailsDialog(self.root, "AEL-Verrechnungsdetails")
                
                if ael_dialog.result: 
                    project_num, kurztext, leistung_dritte, sonstiges = ael_dialog.result
                    today_date = datetime.now().strftime("%d.%m.%Y")
                    user_name = self.settings.get('user_name', 'UNBEKANNT')
                    
                    self.update_ael_excel(
                        project_num=project_num,
                        kurztext=kurztext,
                        leistung_dritte=leistung_dritte,
                        user_name=user_name,
                        today_date=today_date,
                        betra_name=short_name,
                        sonstiges=sonstiges
                    )

        except Exception as e:
            messagebox.showerror("Fehler", f"Ein Fehler ist aufgetreten:\n{e}\n\n"
                                           f"Hinweis: Stellen Sie sicher, dass die Zieldatei (falls sie existiert) geschlossen ist.")
        finally:
            self.start_button.config(text="Ausgewählte Dateien zusammenfügen", state="normal")

    def update_ael_excel(self, project_num, kurztext, leistung_dritte, user_name, today_date, betra_name, sonstiges):
        excel_path = os.path.join(self.output_dir, "AEL-Verrechnung.xlsx")
        
        headers = [
            "Auftragsart", "Eckstarttermin", "Eckendtermin", "AAR-ProjektNr", "Kurztext", 
            "Verantw.ArbPl.", "Auftrag", "AAR-Auftr.-Nr.", "(Buchungsdatum)\nDatum", "Name", 
            "Arbeitsplatz (A0BBK, A0BETRA oder A0SIPLA)", "(LArt)\nFAA\n(immer 1065)", 
            "Werk\n(immer 16ES)", "Einheit\n(immer MIN)", "(Ist-Arbeit)\nMenge", 
            "(Rückmeldetext)\nTätigkeitsbezeichnung/Betra-Nr./SiPla-Nr.", "Bemerkung / Frage"
        ]
        
        new_row_data = [""] * len(headers)
        
        new_row_data[4] = kurztext
        new_row_data[7] = project_num
        new_row_data[8] = today_date
        new_row_data[9] = user_name
        new_row_data[10] = "A0BETRA"
        new_row_data[11] = "1065"
        new_row_data[12] = "16ES"
        new_row_data[13] = "MIN"
        new_row_data[15] = betra_name
        new_row_data[16] = sonstiges
        
        fill_yellow_header = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
        fill_red_header = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")
        fill_yellow_row = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
        
        header_font = Font(name='DB Neo Office Head', size=11, bold=True)
        data_font = Font(name='Db Neo Office', size=11, bold=False)
        
        red_header_indices = [8, 9, 10, 14, 15] 
        
        thin_border_side = Side(border_style="thin", color="000000")
        full_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
        
        header_alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')

        try:
            if not os.path.exists(excel_path):
                wb = openpyxl.Workbook()
                sheet = wb.active
                sheet.title = "AEL-Aufträge"
                sheet.append(headers)
                
                for col_idx, cell in enumerate(sheet[1], 1):
                    if (col_idx - 1) in red_header_indices:
                        cell.fill = fill_red_header
                    else:
                        cell.fill = fill_yellow_header
                    cell.border = full_border
                    cell.alignment = header_alignment
                    cell.font = header_font
                
                sheet.append(new_row_data)
                new_row_index = sheet.max_row
                for cell in sheet[new_row_index]:
                    cell.border = full_border
                    cell.font = data_font

            else:
                wb = openpyxl.load_workbook(excel_path)
                sheet = wb.active
                sheet.append(new_row_data)
                
                new_row_index = sheet.max_row
                for cell in sheet[new_row_index]:
                    cell.border = full_border
                    cell.font = data_font
            
            if leistung_dritte:
                new_row_index = sheet.max_row 
                for cell in sheet[new_row_index]:
                    cell.fill = fill_yellow_row
                    cell.border = full_border 
                    cell.font = data_font

            for col in sheet.columns:
                max_length = 0
                column_letter = col[0].column_letter
                for cell in col:
                    try:
                        cell_value = str(cell.value)
                        lines = cell_value.split('\n')
                        cell_max_line = max(len(line) for line in lines)
                        
                        if cell_max_line > max_length:
                            max_length = cell_max_line
                    except:
                        pass
                adjusted_width = max(10, max_length + 2) 
                sheet.column_dimensions[column_letter].width = adjusted_width
            
            wb.save(excel_path)
            messagebox.showinfo("AEL-Verrechnung", 
                                f"Excel-Datei '{excel_path}' erfolgreich aktualisiert.", 
                                parent=self.root)
            
        except PermissionError:
             error_details = (
                f"Projekt-Nr.: {project_num}\n"
                f"Kurztext: {kurztext}\n"
                f"Sonstiges: {sonstiges}\n"
                f"Datum: {today_date}\n"
                f"Name: {user_name}\n"
                f"Betra-Bez.: {betra_name}\n"
                f"Leistung Dritte: {'Ja' if leistung_dritte else 'Nein'}\n\n"
                f"Statisch: A0BETRA, 1065, 16ES, MIN"
             )
             messagebox.showerror("Fehler (Excel)", 
                                f"Speichern fehlgeschlagen!\nDie Datei '{excel_path}' ist eventuell geöffnet.\n\n"
                                "Bitte schließen Sie die Datei und tragen Sie die Zeile manuell ein:\n\n"
                                f"{error_details}", 
                                parent=self.root)
        except Exception as e:
            messagebox.showerror("Fehler (Excel)", 
                                f"Ein unerwarteter Fehler beim Speichern der Excel-Datei ist aufgetreten:\n{e}", 
                                parent=self.root)

    def merge_documents(self, file_paths, save_path):
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
    
    def add_footer_to_doc(self, file_path, footer_text):
        """Adds a footer with filename (left) and page number (right)."""
        try:
            doc = Document(file_path)
            
            for section in doc.sections:
                footer = section.footer
                footer.is_linked_to_previous = False
                
                for paragraph in footer.paragraphs:
                    p_element = paragraph._element
                    p_element.getparent().remove(p_element)
                
                paragraph = footer.add_paragraph()
                
                page_width = section.page_width or Cm(21)
                left_margin = section.left_margin or Cm(2.5)
                right_margin = section.right_margin or Cm(2.5)
                tab_pos = page_width - left_margin - right_margin
                
                tab_stops = paragraph.paragraph_format.tab_stops
                tab_stops.add_tab_stop(tab_pos, WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.SPACES)
                
                run = paragraph.add_run(footer_text)
                run = paragraph.add_run("\t")
                run = paragraph.add_run("Seite ")
                self._add_field(run, "PAGE")
                run = paragraph.add_run(" von ")
                self._add_field(run, "NUMPAGES")
                
            doc.save(file_path)
            
        except Exception as e:
            print(f"Error adding footer: {e}")
            
    def _add_field(self, run, field_code):
        """Helper to insert a Word field code."""
        fldChar1 = OxmlElement('w:fldChar')
        fldChar1.set(ns.qn('w:fldCharType'), 'begin')
        
        instrText = OxmlElement('w:instrText')
        instrText.set(ns.qn('xml:space'), 'preserve')
        instrText.text = field_code
        
        fldChar2 = OxmlElement('w:fldChar')
        fldChar2.set(ns.qn('w:fldCharType'), 'end')
        
        run._r.append(fldChar1)
        run._r.append(instrText)
        run._r.append(fldChar2)


if __name__ == "__main__":
    root = tk.Tk()
    app = WordMergerApp(root)
    root.mainloop()