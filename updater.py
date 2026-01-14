import os
import sys
import shutil
import subprocess
import time
import tkinter as tk
import re  # NEU: Für Version-Parsing
from tkinter import messagebox, ttk

# --- KONFIGURATION ---
MAIN_APP_EXE = "BetraTool.exe"

# LISTE der möglichen Pfade relativ zum User-Profil
SHAREPOINT_RELATIVE_PATHS = [
    r"Deutsche Bahn\BuS Hagen - Dokumente\Betra\BetraTool (in Arbeit)\_Update",
    r"Deutsche Bahn\BuS Hagen - Betra\BetraTool (in Arbeit)\_Update"
]

# NEU: Version als String, nicht als Float!
UPDATER_INTERNAL_VERSION = "1.2" 
UPDATER_VERSION_FILE = "updater_version.txt"

SERVER_EXE_NAME = "BetraTool.exe"
VERSION_FILE_NAME = "version.txt"

# Ordner die synchronisiert werden sollen
FOLDERS_TO_SYNC = ["modules", "configs", "output"]

# MAPPING: Welcher Ordner hat welche geschützten Dateien?
PROTECTED_FILES_MAP = {
    "configs": ["config.ini", "presets.ini"],
    "output":  ["AEL-Verrechnung.xlsx"]
}

def get_version_key(version_str):
    """
    Wandelt einen Versions-String (z.B. '2.5.1' oder '2.5a') in eine 
    vergleichbare Liste um.
    Beispiel: '2.5.1' -> [2, 5, 1]
    Beispiel: '2.5a'  -> [2, 5, 'a']
    Dies löst das Problem, dass float('2.5.1') abstürzt.
    """
    if not version_str:
        return []
    
    # 1. String bereinigen
    s = str(version_str).strip()
    
    # 2. Splitten an Punkten und Wechseln zwischen Zahl/Text
    # Das Regex splittet bei Zahlen: '1.2a' -> ['1', '.', '2', 'a']
    parts = re.split(r'(\d+)', s)
    
    # 3. Leere Strings entfernen und Zahlen in echte Integers wandeln
    key = []
    for part in parts:
        if part == '' or part == '.':
            continue
        if part.isdigit():
            key.append(int(part))
        else:
            # Textteile (z.B. 'a', 'beta') lowercase machen für Vergleich
            key.append(part.lower())
    return key

def get_valid_server_path():
    """Prüft alle Pfade und gibt den ersten existierenden zurück."""
    user_profile = os.environ.get('USERPROFILE')
    
    for relative_path in SHAREPOINT_RELATIVE_PATHS:
        full_path = os.path.join(user_profile, relative_path)
        if os.path.exists(full_path):
            return full_path
            
    return None

def fill_missing_files(src_folder, dst_folder, protected=[]):
    """Kopiert fehlende Dateien von src nach dst."""
    if not os.path.exists(src_folder):
        return

    if not os.path.exists(dst_folder):
        os.makedirs(dst_folder)

    for filename in os.listdir(src_folder):
        src_file = os.path.join(src_folder, filename)
        dst_file = os.path.join(dst_folder, filename)

        if os.path.isfile(src_file):
            if filename in protected:
                if not os.path.exists(dst_file):
                    try:
                        shutil.copy2(src_file, dst_file)
                        print(f"Protected file missing, creating default: {filename}")
                    except: pass
                continue

            if not os.path.exists(dst_file):
                try:
                    shutil.copy2(src_file, dst_file)
                    print(f"Added missing: {filename}")
                except Exception as e:
                    print(f"Error copying {filename}: {e}")

def check_and_update():
    """Prüft Version und führt Update durch."""
    server_dir = get_valid_server_path()
    
    if not server_dir:
        # Falls Server nicht gefunden, starten wir einfach lokal
        launch_main_app()
        return

    # === CHECK AUF UPDATER-UPDATE ===
    server_updater_ver_file = os.path.join(server_dir, UPDATER_VERSION_FILE)
    
    if os.path.exists(server_updater_ver_file):
        try:
            with open(server_updater_ver_file, 'r') as f:
                server_updater_ver_str = f.read().strip()
            
            # Vergleich mit neuer Logik
            if get_version_key(server_updater_ver_str) > get_version_key(UPDATER_INTERNAL_VERSION):
                msg = (
                    "Es gibt eine neue Version dieses Update-Programms!\n\n"
                    f"Ihre Version: {UPDATER_INTERNAL_VERSION}\n"
                    f"Neue Version: {server_updater_ver_str}\n\n"
                    "Bitte kopieren Sie die neue 'Start BetraTool.exe' "
                    "manuell aus dem Sharepoint."
                )
                messagebox.showwarning("Updater veraltet", msg)
                os.startfile(server_dir)
                sys.exit() 
        except Exception as e:
            print(f"Fehler beim Updater-Version-Check: {e}")

    # === CHECK HAUPTPROGRAMM VERSION ===
    server_version_file = os.path.join(server_dir, VERSION_FILE_NAME)
    server_exe_file = os.path.join(server_dir, SERVER_EXE_NAME)
    local_version_file = "version.txt"
    
    if not os.path.exists(server_version_file) or not os.path.exists(server_exe_file):
        # Falls Server-Dateien fehlen, warnen aber lokal starten
        messagebox.showwarning(
            "Update-Fehler", 
            f"Der Update-Ordner wurde gefunden, aber 'version.txt' fehlt.\nPfad: {server_dir}"
        )
        launch_main_app()
        return

    try:
        # Server Version lesen (als String)
        with open(server_version_file, 'r') as f:
            server_version_str = f.read().strip()
            
        # Lokale Version lesen (als String)
        local_version_str = "0.0"
        if os.path.exists(local_version_file):
            with open(local_version_file, 'r') as f:
                content = f.read().strip()
                if content:
                    local_version_str = content
        
        # VERGLEICH: Nutzen der neuen Helper-Funktion
        if get_version_key(server_version_str) > get_version_key(local_version_str):
            perform_update_process(server_dir, server_exe_file, server_version_str)
        else:
            launch_main_app()
            
    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Versions-Check:\n{e}")
        launch_main_app()

def perform_update_process(server_root, server_exe_path, new_version_str):
    """Führt den Kopiervorgang durch."""
    
    lbl_status.config(text=f"Installiere Update auf Version {new_version_str}...")
    progress['value'] = 0
    root.update()

    try:
        # 1. Hauptprogramm aktualisieren
        lbl_status.config(text="Aktualisiere Programmdatei...")
        shutil.copy2(server_exe_path, MAIN_APP_EXE)
        
        # 2. Netzkarte.exe
        server_map_path = os.path.join(server_root, "netzkarte.exe")
        if os.path.exists(server_map_path):
            lbl_status.config(text="Aktualisiere Netzkarte...")
            shutil.copy2(server_map_path, "netzkarte.exe")

        # 3. Netzdaten.xlsx
        server_data_file = os.path.join(server_root, "configs", "Netzdaten.xlsx")
        local_data_file = os.path.join("configs", "Netzdaten.xlsx")
        
        if not os.path.exists("configs"):
            os.makedirs("configs")

        if os.path.exists(server_data_file):
            lbl_status.config(text="Aktualisiere Netzdaten...")
            try:
                shutil.copy2(server_data_file, local_data_file)
            except Exception as e:
                print(f"Konnte Netzdaten nicht kopieren: {e}")

        progress['value'] = 40
        root.update()
        
        # 4. Fehlende Dateien ergänzen
        step = 60 / len(FOLDERS_TO_SYNC)
        current_progress = 40

        for folder in FOLDERS_TO_SYNC:
            lbl_status.config(text=f"Prüfe Ordner: {folder}...")
            
            src = os.path.join(server_root, folder)
            dst = folder
            prot = PROTECTED_FILES_MAP.get(folder, [])
            
            fill_missing_files(src, dst, protected=prot)
            
            current_progress += step
            progress['value'] = current_progress
            root.update()

        # 5. Lokale Version schreiben (als String)
        with open("version.txt", "w") as f:
            f.write(str(new_version_str))
            
        lbl_status.config(text="Update erfolgreich!")
        progress['value'] = 100
        root.update()
        time.sleep(1)
        
        launch_main_app()
        
    except Exception as e:
        messagebox.showerror("Update Fehler", f"Konnte Update nicht installieren:\n{e}")
        launch_main_app()

def launch_main_app():
    """Startet das Hauptprogramm."""
    if os.path.exists(MAIN_APP_EXE):
        subprocess.Popen([MAIN_APP_EXE])
    else:
        messagebox.showerror("Fehler", f"Die Datei '{MAIN_APP_EXE}' wurde nicht gefunden.")
    
    root.destroy()
    sys.exit()

# --- GUI ---
root = tk.Tk()
root.title("BetraTool Launcher v" + UPDATER_INTERNAL_VERSION)
root.geometry("400x150")
root.eval('tk::PlaceWindow . center')

lbl_status = tk.Label(root, text="Suche nach Updates...", pady=10)
lbl_status.pack()

progress = ttk.Progressbar(root, orient="horizontal", length=350, mode="determinate")
progress.pack(pady=5)

root.after(500, check_and_update)
root.mainloop()