import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog
import os
import subprocess
import csv
from PyPDF2 import PdfReader
import sys
import platform
import configparser
import re
import logging

# Benutzerverzeichnis ermitteln
user_directory = os.path.expanduser("~")
log_directory = os.path.join(user_directory, 'ADAC_Logs')
os.makedirs(log_directory, exist_ok=True)

# Logging einrichten
log_file = os.path.join(log_directory, 'converter.log')
logging.basicConfig(
    filename=log_file,
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Aktuelles Arbeitsverzeichnis ermitteln
current_directory = os.path.dirname(os.path.abspath(__file__))

def read_settings():
    config = configparser.ConfigParser()
    try:
        config.read(os.path.join(current_directory, 'settings.ini'))
        extensions = config['ClubExtensions']['extensions'].split(', ')
        return extensions
    except Exception as e:
        logging.error(f"Fehler beim Lesen der Einstellungen: {e}")
        messagebox.showerror("Fehler", "Fehler beim Lesen der Einstellungen.")
        return []

def get_csv_encoding():
    system = platform.system()
    if system == "Windows":
        return "cp1252"
    elif system == "Darwin":  # macOS
        return "macroman"
    else:
        return "utf-8"

if getattr(sys, 'frozen', False):
    import pyi_splash

def install_and_import(package):
    try:
        __import__(package)
    except ImportError:
        logging.warning(f"{package} ist nicht installiert. Installiere {package}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        __import__(package)

# Installiere PyPDF2 falls nicht vorhanden
install_and_import("PyPDF2")

def open_file_dialog(type_label, entry_widget):
    try:
        initial_dir = current_directory
        if type_label == "PDF":
            filename = filedialog.askopenfilename(
                initialdir=initial_dir,
                filetypes=[
                    ("PDF files", "*.pdf")
                ])
        elif type_label == "CSV":
            filename = filedialog.asksaveasfilename(
                initialdir=initial_dir,
                defaultextension=".csv", filetypes=[
                    ("CSV files", "*.csv")
                ]
            )
        elif type_label == "Word":
            filename = filedialog.askopenfilename(
                initialdir=initial_dir,
                filetypes=[
                    ("Word Templates (*.dotx)", "*.dotx"),
                    ("Word Documents (*.docx)", "*.docx"),
                    ("All Files", "*"),
                ]
            )
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, filename)
    except Exception as e:
        logging.error(f"Fehler im Dateidialog: {e}")
        messagebox.showerror("Fehler", f"Fehler im Dateidialog: {e}")

def remove_extension_and_following_string(file_name, extensions):
    try:
        for ext in extensions:
            if file_name.endswith(ext):
                file_name = re.sub(rf'{ext}.*?(?=\d)', ext, file_name)
                break
        return file_name
    except Exception as e:
        logging.error(f"Fehler beim Entfernen der Erweiterung: {e}")
        messagebox.showerror("Fehler", f"Fehler beim Entfernen der Erweiterung: {e}")
        return file_name

def execute_script(pdf_path, csv_path, word_path):
    extensions = read_settings()
    
    pdf_path = remove_extension_and_following_string(pdf_path, extensions)
    csv_path = remove_extension_and_following_string(csv_path, extensions)
    word_path = remove_extension_and_following_string(word_path, extensions)
    
    try:
        pdf_text_lines, klasse = extract_data_from_pdf(pdf_path)
        processed_data = process_text(pdf_text_lines, klasse, extensions)
        logging.info("Datenverarbeitung erfolgreich abgeschlossen.")
    except Exception as e:
        logging.error(f"Fehler bei der Datenverarbeitung: {e}")
        messagebox.showerror("Fehler", f"Fehler bei der Datenverarbeitung: {e}")
        return
    
    try:
        write_data_to_csv(processed_data, csv_path)
        logging.info("Daten wurden erfolgreich in CSV exportiert.")
        messagebox.showinfo("Erfolg", "Daten wurden erfolgreich in CSV exportiert.")
    except PermissionError:
        retry = messagebox.askretrycancel(
            "Berechtigungsfehler",
            "Keine Berechtigung zum Speichern der Datei. Möchten Sie es erneut versuchen?",
        )
        if retry:
            csv_path = filedialog.asksaveasfilename(
                initialdir=current_directory,
                defaultextension=".csv", filetypes=[("CSV files", "*.csv")]
            )
            execute_script(pdf_path, csv_path, word_path)
        else:
            logging.error("Speichern der Datei abgebrochen aufgrund von Berechtigungsfehlern.")
    
    if word_path:
        try:
            if sys.platform == "win32":
                os.startfile(word_path)
            elif sys.platform == "darwin":
                subprocess.call(["open", word_path])
            else:
                subprocess.call(["xdg-open", word_path])
        except Exception as e:
            logging.error(f"Fehler beim Öffnen der Datei: {e}")
            messagebox.showerror(f"Fehler beim Öffnen der Datei: {e}")

    sys.exit()

def close_application():
    logging.info("Anwendung wird beendet.")
    root.quit()

def extract_data_from_pdf(pdf_path):
    try:
        with open(pdf_path, "rb") as file:
            reader = PdfReader(file)
            text = []
            klasse = ""
            for page in reader.pages:
                page_text = page.extract_text()
                text.extend(page_text.split("\n"))
                if "Klasse: K" in page_text and not klasse:
                    klasse_line = [
                        line for line in page_text.split("\n") if "Klasse: K" in line
                    ][0]
                    klasse = klasse_line.split("Klasse: K")[1].strip().split()[0]
        return text, klasse
    except Exception as e:
        logging.error(f"Fehler beim Extrahieren von Daten aus der PDF-Datei: {e}")
        messagebox.showerror("Fehler", f"Fehler beim Extrahieren von Daten aus der PDF-Datei: {e}")
        return [], ""

def prompt_user_for_missing_data(line):
    missing_data = {}

    def on_confirm():
        missing_data['place'] = place_entry.get()
        missing_data['fahrername'] = fahrername_entry.get()
        missing_data['club'] = club_entry.get()

        if not missing_data['place'] or not missing_data['fahrername'] or not missing_data['club']:
            messagebox.showerror("Fehler", "Alle Felder müssen ausgefüllt werden.")
            return
        prompt_window.destroy()

    prompt_window = tk.Toplevel(root)
    prompt_window.title("Fehlende Daten eingeben")

    tk.Label(prompt_window, text="Fehlende Daten in Zeile gefunden. Bitte ergänzen:").pack(pady=10)
    
    tk.Label(prompt_window, text=f"Zeile: {line}").pack()

    tk.Label(prompt_window, text="Platz:").pack()
    place_entry = tk.Entry(prompt_window, width=50)
    place_entry.pack()

    tk.Label(prompt_window, text="Fahrername:").pack()
    fahrername_entry = tk.Entry(prompt_window, width=50)
    fahrername_entry.pack()

    tk.Label(prompt_window, text="Club:").pack()
    club_entry = tk.Entry(prompt_window, width=50)
    club_entry.pack()

    tk.Button(prompt_window, text="Bestätigen", command=on_confirm).pack(pady=10)
    
    prompt_window.grab_set()
    root.wait_window(prompt_window)
    
    return missing_data['place'], missing_data['fahrername'], missing_data['club']

def process_text(text_lines, klasse, extensions):
    data = []
    try:
        for line in text_lines:
            parts = line.split()
            if len(parts) < 5:
                continue  # Zeile überspringen, wenn sie zu kurz ist

            if parts[0].isdigit():
                try:
                    place = parts[0]
                    startnummer = parts[1]
                    license_index = next((i for i, part in enumerate(parts) if re.match(r'\d{6}', part)), len(parts))
                    if license_index == len(parts):
                        logging.warning(f"Keine Lizenznummer in Zeile gefunden: {line}")
                        place, fahrername, club_name = prompt_user_for_missing_data(line)
                    else:
                        license_number = parts[license_index]
                        # Bestimme den Clubnamen anhand der Extension
                        club_name_parts = []
                        for i in range(2, license_index):
                            if any(parts[i].startswith(ext) for ext in extensions):
                                club_name_parts = parts[i:license_index]
                                break

                        club_name = ' '.join(club_name_parts)
                        fahrername = ' '.join(parts[2:license_index - len(club_name_parts)])  # Fahrernamen extrahieren, ohne Startnummer und Clubnamenerweiterung

                    data.append([place, fahrername, club_name, klasse])
                except IndexError as e:
                    logging.error(f"Fehler beim Verarbeiten der Zeile: {line} - {e}")
                    continue
        return data
    except Exception as e:
        logging.error(f"Fehler beim Verarbeiten des Textes: {e}")
        messagebox.showerror("Fehler", f"Fehler beim Verarbeiten des Textes: {e}")
        return data

def write_data_to_csv(data, csv_path):
    encoding = get_csv_encoding()
    try:
        with open(csv_path, mode='w', newline='', encoding=encoding) as file:
            writer = csv.writer(file, delimiter=';')
            writer.writerow(['Platz', 'Fahrername', 'Club', 'Klasse'])
            for record in data:
                record_encoded = [
                    item.encode(encoding, errors='replace').decode(encoding) if isinstance(item, str) else item
                    for item in record
                ]
                writer.writerow(record_encoded)
    except Exception as e:
        logging.error(f"Fehler beim Schreiben der CSV-Datei: {e}")
        messagebox.showerror("Fehler", f"Fehler beim Schreiben der CSV-Datei: {e}")

class App:
    def __init__(self, root):
        root.title("ADAC-Ergebnislisten Converter")
        width = 410
        height = 200
        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()
        alignstr = "%dx%d+%d+%d" % (
            width,
            height,
            (screenwidth - width) / 2,
            (screenheight - height) / 2,
        )
        root.geometry(alignstr)
        root.resizable(width=False, height=False)

        self.pdf_entry = tk.Entry(root, width=40)
        self.pdf_entry.place(x=150, y=15)
        tk.Button(
            root,
            text="ADAC-Ergebnisliste",
            command=lambda: open_file_dialog("PDF", self.pdf_entry),
        ).place(x=10, y=10)

        self.csv_entry = tk.Entry(root, width=40)
        self.csv_entry.place(x=150, y=55)
        tk.Button(
            root,
            text="CSV-Exportpfad",
            command=lambda: open_file_dialog("CSV", self.csv_entry),
        ).place(x=10, y=50)

        self.word_entry = tk.Entry(root, width=40)
        self.word_entry.place(x=150, y=95)
        tk.Button(
            root,
            text="Word-Vorlage",
            command=lambda: open_file_dialog("Word", self.word_entry),
        ).place(x=10, y=90)

        tk.Button(root, text="Ausführen", command=self.start_command).place(x=150, y=130)
        tk.Button(root, text="Beenden", command=close_application).place(x=150, y=160)

    def start_command(self):
        execute_script(
            self.pdf_entry.get(), self.csv_entry.get(), self.word_entry.get()
        )

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.lift()  # Fenster in den Vordergrund bringen
    root.attributes('-topmost', True)  # Setze das Fenster ganz nach oben
    root.after_idle(root.attributes, '-topmost', False)  # Setze das Fenster wieder normal, wenn das Fenster idle ist
    if getattr(sys, 'frozen', False):
        pyi_splash.close()
    logging.info("Anwendung gestartet.")
    root.mainloop()
