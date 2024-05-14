# ADAC-Ergebnislisten Converter 🏁

Dies ist ein Python-Programm zur Extraktion von Daten aus ADAC-Ergebnislisten im PDF-Format und zum Exportieren dieser Daten in eine CSV-Datei. Es bietet auch die Möglichkeit, eine Word-Vorlage auszuwählen, um die exportierten Daten zu formatieren.

## Einleitung 🚀

Dieses Programm bietet eine grafische Benutzeroberfläche (GUI) zur einfachen Verwendung. Es ist in Python geschrieben und verwendet das Tkinter-Modul für die GUI sowie andere Module wie PyPDF2 für die PDF-Verarbeitung und csv für das Schreiben von CSV-Dateien.

## Funktionen

- **PDF-zu-CSV-Konvertierung:** Extrahieren Sie Daten aus ADAC-Ergebnislisten im PDF-Format und speichern Sie sie in einer CSV-Datei.
- **Word-Vorlagenintegration:** Möglichkeit zur Auswahl einer Word-Vorlage zur Formatierung der exportierten Daten.
- **Fehlerbehandlung:** Umfassende Fehlerbehandlung und Protokollierung zur Fehlerbehebung.
- **Plattformübergreifende Kompatibilität:** Funktioniert unter Windows, macOS und Linux.

## Voraussetzungen

- Python 3.x
- Erforderliche Python-Pakete: tkinter, PyPDF2, configparser

## Installation

1. **Repository klonen:**

    ```bash
    git clone https://github.com/your-username/adac-ergebnislisten-converter.git
    cd adac-ergebnislisten-converter
    ```

2. **Erforderliche Pakete installieren:**

    ```bash
    pip install -r requirements.txt
    ```

3. **Kontrollien Sie die settings.ini-Datei:**

    Hier sollten alle Clubnamenserweiterungen der Vereine hinterlegt sein

    ```ini
    [ClubExtensions]
    extensions = MC, MSC, AC, Rapid, AMC, MSF, RT
    ```

## Verwendung 🛠️

### Ausführen als Python-Skript:

1. **Hauptskript ausführen:**

    ```bash
    python main.py
    ```

2. **GUI verwenden:**

- Wählen Sie die PDF-Datei mit der ADAC-Ergebnisliste aus.
- Wählen Sie den Zielpfad für die CSV-Datei aus.
- Optional: Wählen Sie eine Word-Vorlagendatei aus.
- Klicken Sie auf "Ausführen", um den Konvertierungsvorgang zu starten.

### Ausführen als eigenständige ausführbare Datei:

1. Laden Sie die ausführbare Datei herunter.

2. Führen Sie die ausführbare Datei aus:

- Wählen Sie die PDF-Datei mit der ADAC-Ergebnisliste aus.
- Wählen Sie den Zielpfad für die CSV-Datei aus.
- Optional: Wählen Sie eine Word-Vorlagendatei aus.
- Klicken Sie auf "Ausführen", um den Konvertierungsvorgang zu starten.

## Module 📦

- **tkinter:** Wird für die Erstellung der grafischen Benutzeroberfläche verwendet.
- **PyPDF2:** Wird für das Extrahieren von Text aus PDF-Dateien verwendet.
- **csv:** Wird für das Schreiben von Daten in CSV-Dateien verwendet.
- **configparser:** Wird für das Lesen von Konfigurationseinstellungen verwendet.
- **logging:** Wird für die Protokollierung von Fehlern und Informationen verwendet.

## Lizenz ©️

Dieses Programm steht unter der [MIT-Lizenz](LICENSE).

## Protokollierung

Eine Protokolldatei *(converter.log)* wird im Verzeichnis des Skripts erstellt, um Fehlermeldungen und andere Informationen zur Fehlerbehebung zu erfassen.

## Mitwirken

Fühlen Sie sich frei, dieses Repository zu forken, Änderungen vorzunehmen und Pull-Requests einzureichen. Bei größeren Änderungen öffnen Sie bitte zuerst ein Issue, um zu diskutieren, was Sie ändern möchten.

------------------------------------------------------

Wenn Sie auf Probleme stoßen oder Fragen haben, öffnen Sie bitte ein Issue im GitHub-Repository.

Viel Erfolg beim Konvertieren! 🏁