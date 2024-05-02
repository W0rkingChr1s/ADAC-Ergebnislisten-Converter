import tkinter as tk
from tkinter import filedialog, messagebox
import os
import subprocess
import csv
from PyPDF2 import PdfReader
import sys

if getattr(sys, 'frozen', False):
    import pyi_splash

def install_and_import(package):
    try:
        __import__(package)
    except ImportError:
        print(f"{package} ist nicht installiert. Installiere {package}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        __import__(package)


# Installiere PyPDF2 falls nicht vorhanden
install_and_import("PyPDF2")


def open_file_dialog(type_label, entry_widget):
    if type_label == "PDF":
        filename = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    elif type_label == "CSV":
        filename = filedialog.asksaveasfilename(
            defaultextension=".csv", filetypes=[("CSV files", "*.csv")]
        )
    elif type_label == "Word":
        filename = filedialog.askopenfilename(
            filetypes=[
                ("Word Documents (*.docx)", "*.docx"),
                ("Word Templates (*.dotx)", "*.dotx"),
                ("All Files", "*"),
            ]
        )
    entry_widget.delete(0, tk.END)
    entry_widget.insert(0, filename)


def execute_script(pdf_path, csv_path, word_path):
    print("Datenverarbeitung starten...")
    pdf_text_lines, klasse = extract_data_from_pdf(pdf_path)
    processed_data = process_text(pdf_text_lines, klasse)
    write_data_to_csv(processed_data, csv_path)
    print("Daten wurden erfolgreich in CSV exportiert.")

    messagebox.showinfo("Erfolg", "Daten wurden erfolgreich in CSV exportiert.")

    if word_path:
        try:
            if sys.platform == "win32":
                os.startfile(word_path)
            elif sys.platform == "darwin":
                subprocess.call(["open", word_path])  # Verwende 'open' auf macOS
            else:
                subprocess.call(
                    ["xdg-open", word_path]
                )  # Verwende 'xdg-open' auf Linux
        except Exception as e:
            print(f"Fehler beim Öffnen der Datei: {e}")


def close_application():
    root.quit()


def extract_data_from_pdf(pdf_path):
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


def process_text(text_lines, klasse):
    data = []
    for line in text_lines:
        parts = line.split()
        if parts and parts[0].isdigit() and len(parts) > 3:
            place = parts[0]
            nachname = parts[2]
            vorname = parts[3]
            data.append([place, nachname, vorname, klasse])
    return data


def write_data_to_csv(data, csv_path):
    with open(csv_path, mode='w', newline='', encoding='cp437') as file:
        writer = csv.writer(file, delimiter=';')
        writer.writerow(['Platz', 'Nachname', 'Vorname', 'Klasse'])
        for record in data:
            # Konvertiere alle Datensätze in cp437, um Sonderzeichen korrekt zu behandeln
            record_cp437 = [
                item.encode('cp437', errors='replace').decode('cp437') if isinstance(item, str) else item
                for item in record
            ]
            writer.writerow(record_cp437)


class App:
    def __init__(self, root):
        # setting title
        root.title("ADAC-Ergebnislisten Converter")
        # setting window size
        width = 400
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

        self.pdf_entry = tk.Entry(root, width=50)
        self.pdf_entry.place(x=240, y=60)
        tk.Button(
            root,
            text="ADAC-Ergebnisliste",
            command=lambda: open_file_dialog("PDF", self.pdf_entry),
        ).place(x=130, y=10)

        self.csv_entry = tk.Entry(root, width=50)
        self.csv_entry.place(x=240, y=150)
        tk.Button(
            root,
            text="CSV-Exportpfad",
            command=lambda: open_file_dialog("CSV", self.csv_entry),
        ).place(x=140, y=40)

        self.word_entry = tk.Entry(root, width=50)
        self.word_entry.place(x=240, y=240)
        tk.Button(
            root,
            text="Word-Vorlage",
            command=lambda: open_file_dialog("Word", self.word_entry),
        ).place(x=150, y=80)

        tk.Button(root, text="Ausführen", command=self.start_command).place(x=160, y=110)
        tk.Button(root, text="Beenden", command=close_application).place(x=165, y=160)

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
    root.mainloop()
