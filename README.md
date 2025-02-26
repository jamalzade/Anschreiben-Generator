# Anschreiben Generator

## 📌 Beschreibung
Ein einfaches Python-Tool zur Generierung von individuellen Bewerbungsanschreiben auf Deutsch.  
Es extrahiert Stellenanzeigen aus Webseiten und generiert basierend auf den Anforderungen ein professionelles Bewerbungsschreiben.

## 🚀 Funktionen
- **Web-Scraping** von Stellenanzeigen aus Indeed, LinkedIn und StepStone
- **Generierung** von personalisierten Bewerbungsschreiben mit KI-Unterstützung
- **GUI mit PyQt6** für einfache Bedienung
- **Speichern** der Anschreiben als `.txt`, `.docx` oder `.pdf`
- **Automatische Datensicherung** für wiederholte Bewerbungen

## ⚠️ WICHTIG: OpenAI API-Token erforderlich
Dieses Projekt verwendet die OpenAI-API für die Generierung der Anschreiben.  
**Du musst deinen eigenen API-Token einfügen, bevor du das Programm nutzen kannst.**  

**🔧 So ersetzt du den API-Token:**
1. Öffne die Datei `n.py` in einem Texteditor.
2. Suche nach folgender Zeile:
   ```python
   openai.api_key = "DEIN_OPENAI_API_KEY"

3. Ersetze "DEIN_OPENAI_API_KEY" mit deinem persönlichen API-Token von OpenAI.
4. Speichere die Datei und starte das Programm erneut.

🛠 Installation
✅ Voraussetzungen
Python 3.8 oder höher
Google Chrome
WebDriver für Chrome (wird automatisch installiert)
📌 Installation von Abhängigkeiten
Führe den folgenden Befehl aus, um alle erforderlichen Bibliotheken zu installieren:
pip install -r requirements.txt

Falls requirements.txt nicht vorhanden ist, installiere die Abhängigkeiten manuell: sh Copy Edit
pip install selenium PyQt6 docx pdfkit webdriver-manager beautifulsoup4 openai

🎯 Benutzung
Das Programm kann direkt aus der Konsole oder über eine GUI gestartet werden:
python Anschreiben-Generator.py
Falls Web-Scraping fehlschlägt, stelle sicher, dass Chrome aktuell ist.


