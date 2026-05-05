# exportPowerpointSlides
exports Powerpoint slides to png and text information to txt
# PowerPoint Extractor (Windows)

Ein Python-Tool mit GUI zum automatischen Verarbeiten von PowerPoint-Dateien:

- Rekursive Suche nach `.pptx`, `.pptm`, `.ppt`
- Export aller Folien als PNG (über Microsoft PowerPoint)
- Extraktion aller Texte und Notizen

---

## Features

### Dateiverarbeitung
- Durchsucht einen Ordner **rekursiv**
- Unterstützt:
  - `.pptx`
  - `.pptm`
  - `.ppt`

---

### Strukturierte Ausgabe
Für jede PowerPoint wird automatisch ein Ordner erzeugt:

```text
<JJJJMMDDHHMMSS>_<Dateiname>
```

Beispiel:

```text
20260504120030_MeinePraesentation
```

---

### Folienexport
- Jede Folie wird als PNG exportiert
- Dateiname = Folientitel (bereinigt)
- Beispiel:

```text
Einleitung.png
Architektur.png
```

---

### 📝 Textextraktion

#### 1. Alle Folientexte

Datei:

```text
Alle_Folientexte.txt
```

Inhalt:

```text
===== Folie 1: Titel =====
Text...

===== Folie 2: Titel =====
Text...
```

---

#### 2. Alle Notizen / Kommentare

Datei:

```text
Alle_Notizen.txt
```

Nur wenn vorhanden.

---

###  GUI
- Auswahl von:
  - Source Folder
  - Zielordner
- Start-Button
- Live-Log-Anzeige

---

###  Technik
- Nutzung von **Microsoft PowerPoint (COM)**
- Kein LibreOffice notwendig
- Hohe Qualität beim PNG-Export

---

##  Voraussetzungen

###  System
- Windows
- Microsoft PowerPoint installiert (64-bit empfohlen)

---

###  Python
- Python 3.10+ (empfohlen: 3.11)

---

###  Python-Pakete

Installation:

```powershell
& "C:\Program Files\Python311\python.exe" -m pip install python-pptx pywin32
```

---

##  Installation

### 1. Script speichern

```text
exporter.py
```

---

### 2. pywin32 installieren

```powershell
& "C:\Program Files\Python311\python.exe" -m pip install pywin32
```

---

### 3. Test (optional)

```powershell
& "C:\Program Files\Python311\python.exe" -c "import pythoncom; import win32com.client; print('OK')"
```

---

##  Starten

```powershell
& "C:\Program Files\Python311\python.exe" exporter.py
```

---

##  Ablauf

1. Source Folder auswählen  
2. Zielordner auswählen  
3. „Start“ klicken  

Das Tool verarbeitet automatisch alle PowerPoints.

---

##  Beispiel Ausgabe

```text
out/
└── 20260504120030_MeinePraesentation/
    ├── Einleitung.png
    ├── Architektur.png
    ├── Fazit.png
    ├── Alle_Folientexte.txt
    └── Alle_Notizen.txt
```

---

## Bekannte Einschränkungen

- PowerPoint muss installiert sein
- PowerPoint darf keine Dialoge anzeigen (z. B. „Datei reparieren“)
- Sehr alte `.ppt` Dateien können Probleme machen
- COM funktioniert nur unter Windows

---

## Troubleshooting

### ModuleNotFoundError: pythoncom

```powershell
& "C:\Program Files\Python311\python.exe" -m pip install pywin32
```

---

### Keine PNGs erzeugt

- PowerPoint installiert?
- Datei lässt sich manuell öffnen?
- Schreibrechte im Zielordner?

---

### [WinError 2] Datei nicht gefunden

Wird durch Duplikat-Handling der PNG-Dateien verhindert

---

### Script startet nicht in PowerShell

```powershell
& "Pfad\zu\python.exe" script.py
```

---


## Lizenz

Frei nutzbar für private und interne Zwecke.
