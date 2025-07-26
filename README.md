# AutoBANF

**Automatische SAP BANF-Erstellung mit Excel-Import**

AutoBANF ist ein Python-Tool zur automatisierten Erstellung von Bestellanforderungen (BANF) in SAP-Systemen. Das Tool liest Excel-Dateien ein und füllt automatisch SAP-Formulare aus.

## 🚀 Schnellstart

### 1. Repository klonen
```bash
git clone [REPOSITORY-URL]
cd AutoBANF
```

### 2. Setup ausführen
```cmd
cd autoBANF
setup.bat
```

Das Setup-Script:
- Erstellt automatisch ein Python Virtual Environment
- Installiert alle benötigten Pakete
- Richtet den Playwright Browser ein
- Erstellt die requirements.txt

### 3. Programm starten
```cmd
python autoBANF.py testfiles\mouser.xlsx
```

## 📋 Systemvoraussetzungen

- **Python 3.8 oder höher** (von [python.org](https://python.org))
- **Windows** (getestet auf Windows 10/11)
- **Internet-Verbindung** für SAP-Zugang

## 📊 Excel-Format

Das Tool erwartet Excel-Dateien mit folgenden Spalten:

| Spalte | Beschreibung | Beispiel |
|--------|--------------|----------|
| Kategorie | Haupt->Unterkategorie | `Bedarf Labore und Werkstätten->Laborutensilien, Werkzeuge und Kleinteile` |
| Artikelbeschreibung | Artikel-Name | `1N4007` |
| Steuerkennzeichen | Steuer-Code | `Keine Steuer` |
| Preisart | Preis-Typ | `Netto-Preis` |
| Preis_je_Mengeneinheit | Stückpreis | `1,23` |
| Waehrung | Währung | `EUR` |
| Rabatttyp | Rabatt-Art | `Prozentsatz` |
| Rabattwert | Rabatt-Wert | `0` |
| Laufzeit | Gültigkeitszeitraum | `01.08.2025 - 31.12.2025` |
| Bestellmenge | Anzahl | `100` |
| Einheit | Mengeneinheit | `ST Stück` |
| Lange_Artikelbeschreibung | Detailbeschreibung | `Detaillierte Produktbeschreibung` |
| Angebotsreferenz | Angebots-ID | `Q158DC2` |
| Angebotsdatum | Angebotsdatum | `26.07.2025` |
| Kontierungsobjekttyp | Kontierung-Typ | `Projekt` |
| Kontierungsobjekt | Kontierung-Objekt | `9/000002002` |

## 🔧 Funktionen

### ✅ Was das Tool kann:
- **Excel-Import**: Beliebig viele Artikel aus Excel-Dateien
- **Automatische SAP-Navigation**: Vollautomatische Anmeldung und Navigation
- **Formular-Ausfüllung**: Alle 15 SAP-Felder werden automatisch ausgefüllt
- **Kategorieauswahl**: Robuste Auswahl von Haupt- und Unterkategorien
- **Deutsche Zahlenformatierung**: Automatische Komma-Dezimaltrennung
- **Sichere Anmeldedaten**: Verschlüsselte Speicherung der SAP-Credentials
- **Monitoring**: Screenshots für jeden Verarbeitungsschritt

### 🎯 Besondere Features:
- **Robuste Kategorieauswahl**: Funktioniert auch bei dynamischen Element-IDs
- **Multi-Artikel-Support**: Verarbeitet automatisch alle Artikel in einer Session
- **Fehlerbehandlung**: Detaillierte Logs und Fallback-Mechanismen
- **Performance-Optimiert**: Reduzierte Wartezeiten für schnelle Verarbeitung

## 🔐 Sicherheit

### Anmeldedaten
- Beim ersten Start werden SAP-Benutzername und Passwort abgefragt
- Sichere Verschlüsselung mit Fernet (AES 128)
- Speicherung in `.credentials/` (versteckter Ordner)
- Keine Klartext-Speicherung

### Ordnerstruktur
```
autoBANF/
├── autoBANF.py              # Hauptprogramm
├── setup.bat               # Setup-Script
├── testfiles/              # Excel-Vorlagen
│   └── mouser.xlsx         # Beispiel-Datei
├── .credentials/           # Verschlüsselte Anmeldedaten
├── .temp/                 # Screenshots & Logs
└── venv/                  # Python-Umgebung (nach Setup)
```

## 📝 Verwendung

### Erste Ausführung:
1. `setup.bat` ausführen
2. `python autoBANF.py testfiles\mouser.xlsx` starten
3. SAP-Anmeldedaten eingeben (werden sicher gespeichert)
4. Automatische Verarbeitung startet

### Weitere Ausführungen:
```cmd
python autoBANF.py [IHRE-EXCEL-DATEI.xlsx]
```

### Beispiele:
```cmd
# Mit Beispiel-Datei
python autoBANF.py testfiles\mouser.xlsx

# Mit eigener Excel-Datei
python autoBANF.py "C:\Meine Bestellungen\elektronik.xlsx"

# Mit relativer Datei
python autoBANF.py meine_artikel.xlsx
```

## 📸 Monitoring

Das Tool erstellt automatisch Screenshots in `.temp/`:
- `excel_import_artikel_1.png` - Nach Artikel 1
- `excel_import_artikel_2.png` - Nach Artikel 2
- `excel_import_complete.png` - Finale Übersicht
- `excel_import_error.png` - Bei Fehlern

## 🐛 Troubleshooting

### Häufige Probleme:

**Python nicht gefunden**
```
Lösung: Python von python.org installieren
```

**Playwright-Fehler**
```
Lösung: setup.bat erneut ausführen
```

**SAP-Anmeldung fehlgeschlagen**
```
Lösung: Anmeldedaten löschen und neu eingeben
- .credentials/ Ordner löschen
- Programm neu starten
```

**Excel-Datei nicht gefunden**
```
Lösung: Vollständigen Pfad verwenden oder Datei ins autoBANF/ Verzeichnis kopieren
```

## 🔄 Updates

```cmd
git pull
setup.bat
```

## 📞 Support

Bei Problemen:
1. Screenshots aus `.temp/` prüfen
2. Log-Ausgaben beachten
3. Git Issues erstellen

---

**🎯 AutoBANF - Effiziente SAP-Automatisierung mit Excel-Integration**