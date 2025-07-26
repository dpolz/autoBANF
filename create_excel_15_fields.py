#!/usr/bin/env python3
"""
create_excel_15_fields.py - Erstellt Excel-Template mit allen 15 Feldern
"""

import pandas as pd

def create_complete_excel_template():
    """
    Erstellt Excel-Template mit allen 15 Feldern und Testdaten
    """
    print("=== EXCEL-TEMPLATE MIT ALLEN 15 FELDERN ===")
    
    # Template-Daten mit mehreren Testartikeln
    data = {
        # Feld 0: Kategorie (Hauptkategorie->Unterkategorie)
        'Kategorie': [
            'Bedarf Labore und Werkstätten->Laborutensilien, Werkzeuge und Kleinteile',
            'Bedarf Labore und Werkstätten->Laborutensilien, Werkzeuge und Kleinteile',
            'Bedarf Labore und Werkstätten->Laborutensilien, Werkzeuge und Kleinteile',
            'Bedarf Labore und Werkstätten->Laborutensilien, Werkzeuge und Kleinteile',
            'Bedarf Labore und Werkstätten->Laborutensilien, Werkzeuge und Kleinteile'
        ],
        
        # Feld 1: Artikelbeschreibung (Text)
        'Artikelbeschreibung': [
            '1N4007',
            'Widerstand 220 Ohm',
            'LED rot 5mm',
            'Kondensator 100µF',
            'Schraubendreher Set'
        ],
        
        # Feld 2: Steuerkennzeichen (Dropdown)
        'Steuerkennzeichen': [
            'Keine Steuer',
            'Keine Steuer', 
            'kein separater Steuerabzug',
            'Keine Steuer',
            'Voller Steuersatz (19 %)'
        ],
        
        # Feld 3: Preisart (Dropdown)
        'Preisart': [
            'Brutto-Preis',
            'Netto-Preis',
            'Brutto-Preis',
            'Netto-Preis',
            'Brutto-Preis'
        ],
        
        # Feld 4: Preis je Mengeneinheit (Numeric)
        'Preis_je_Mengeneinheit': [
            '1,23',
            '0,15',
            '2,50',
            '8,99',
            '25,00'
        ],
        
        # Feld 5: Währung (Dropdown)
        'Waehrung': [
            'USD',
            'EUR',
            'EUR',
            'USD',
            'EUR'
        ],
        
        # Feld 6: Rabatttyp (Dropdown)
        'Rabatttyp': [
            'Prozentsatz',
            'Absoluter Wert',
            'Prozentsatz',
            'Absoluter Wert',
            'Prozentsatz'
        ],
        
        # Feld 7: Rabattwert (Numeric)
        'Rabattwert': [
            '10',
            '0',
            '5',
            '1',
            '15'
        ],
        
        # Feld 8: Laufzeit (Text mit Datum)
        'Laufzeit': [
            '10.07.2025 - 23.08.2025',
            '01.08.2025 - 31.12.2025',
            '15.08.2025 - 30.09.2025',
            '01.09.2025 - 31.10.2025',
            '01.11.2025 - 31.12.2025'
        ],
        
        # Feld 9: Bestellmenge (Numeric)
        'Bestellmenge': [
            '10',
            '100',
            '25',
            '5',
            '1'
        ],
        
        # Feld 10: Einheit (ComboBox)
        'Einheit': [
            'G g',
            'ST Stück',
            'ST Stück',
            'ST Stück',
            'ST Stück'
        ],
        
        # Feld 11: Lange Artikelbeschreibung (Text)
        'Lange_Artikelbeschreibung': [
            'Gleichrichterdiode 1N4007, 1000V, 1A',
            'Kohleschichtwiderstand 220 Ohm, 1/4W',
            'Rote LED 5mm, 20mA, 2V',
            'Elektrolytkondensator 100µF, 25V',
            'Schraubendreher Set, 6-teilig, Kreuz und Schlitz'
        ],
        
        # Feld 12: Angebotsreferenz (Text)
        'Angebotsreferenz': [
            'DIODE-2025-001',
            'WIDERSTAND-2025',
            'LED-ROT-2025',
            'KONDENSATOR-100UF',
            'WERKZEUG-SET-001'
        ],
        
        # Feld 13: Angebotsdatum (Text mit Datum)
        'Angebotsdatum': [
            '01.07.2025',
            '26.07.2025',
            '15.08.2025',
            '01.09.2025',
            '15.10.2025'
        ],
        
        # Feld 14: Kontierungsobjekttyp (Dropdown)
        'Kontierungsobjekttyp': [
            'Projekt',
            'Projekt',
            'Kostenstelle',
            'Projekt',
            'Innenauftrag'
        ],
        
        # Feld 15: Kontierungsobjekt (Text)
        'Kontierungsobjekt': [
            '9/000002002',
            '9/000002002',
            '8/000001001',
            '9/000002003',
            '7/000003001'
        ]
    }
    
    # DataFrame erstellen
    df = pd.DataFrame(data)
    
    # Excel-Datei erstellen
    excel_filename = 'AutoBANF_15_Felder_Template.xlsx'
    
    with pd.ExcelWriter(excel_filename, engine='openpyxl') as writer:
        # Haupttabelle mit Daten
        df.to_excel(writer, sheet_name='Artikel_Import', index=False)
        
        # Zusätzliches Sheet mit Feldbeschreibungen
        field_descriptions = pd.DataFrame({
            'Feld_Nr': [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15],
            'Feldname': [
                'Kategorie',
                'Artikelbeschreibung', 
                'Steuerkennzeichen',
                'Preisart',
                'Preis_je_Mengeneinheit',
                'Waehrung',
                'Rabatttyp', 
                'Rabattwert',
                'Laufzeit',
                'Bestellmenge',
                'Einheit',
                'Lange_Artikelbeschreibung',
                'Angebotsreferenz',
                'Angebotsdatum',
                'Kontierungsobjekttyp',
                'Kontierungsobjekt'
            ],
            'Typ': [
                'Kategorie-Dialog',
                'Text',
                'Dropdown',
                'Dropdown', 
                'Numeric',
                'Dropdown',
                'Dropdown',
                'Numeric',
                'Text/Datum',
                'Numeric',
                'ComboBox',
                'Text',
                'Text',
                'Text/Datum',
                'Dropdown',
                'Text'
            ],
            'Beispielwerte': [
                'Bedarf Labore und Werkstätten->Laborutensilien, Werkzeuge und Kleinteile',
                '1N4007, Widerstand 220 Ohm, LED rot 5mm',
                'Keine Steuer, kein separater Steuerabzug, Voller Steuersatz (19 %)',
                'Brutto-Preis, Netto-Preis',
                '1,23 (mit Komma als Dezimaltrennzeichen)',
                'EUR, USD',
                'Prozentsatz, Absoluter Wert',
                '10, 0, 5 (numerisch)',
                '10.07.2025 - 23.08.2025',
                '10, 100, 25 (numerisch)',
                'G g, ST Stück, KG kg, L l, H Stunde, LE LeistEinh., PAU Pauschal',
                'Detaillierte Artikelbeschreibung',
                'DIODE-2025-001, WIDERSTAND-2025',
                '01.07.2025, 26.07.2025',
                'Projekt, Kostenstelle, Innenauftrag',
                '9/000002002, 8/000001001'
            ]
        })
        
        field_descriptions.to_excel(writer, sheet_name='Feldbeschreibungen', index=False)
        
        # Verfügbare Dropdown-Werte
        dropdown_values = pd.DataFrame({
            'Steuerkennzeichen': [
                'EU-Ausland (19% ohne VST-abzug)',
                'Keine Steuer',
                'kein separater Steuerabzug', 
                'Drittland (ohne VSt-Abz)',
                'Voller Steuersatz (19 %)',
                'Gemäßigter Steuersatz (7%)',
                '', '', '', '', '', ''
            ],
            'Preisart': [
                'Netto-Preis',
                'Brutto-Preis',
                '', '', '', '', '', '', '', '', '', ''
            ],
            'Waehrung': [
                'EUR',
                'USD',
                '', '', '', '', '', '', '', '', '', ''
            ],
            'Rabatttyp': [
                'Absoluter Wert',
                'Prozentsatz',
                '', '', '', '', '', '', '', '', '', ''
            ],
            'Einheit': [
                'G g',
                'H Stunde',
                'KG kg',
                'L l',
                'LE LeistEinh.',
                'ST Stück',
                'PAU Pauschal',
                '', '', '', '', ''
            ],
            'Kontierungsobjekttyp': [
                'Projekt',
                'Kostenstelle',
                'Innenauftrag',
                '', '', '', '', '', '', '', '', ''
            ]
        })
        
        dropdown_values.to_excel(writer, sheet_name='Dropdown_Werte', index=False)
    
    print(f">> Excel-Template erstellt: {excel_filename}")
    print(f">> Anzahl Artikel: {len(df)}")
    print(f">> Anzahl Felder: {len(df.columns)}")
    print()
    print(">> SHEETS:")
    print("- Artikel_Import: Hauptdaten zum Ausfuellen")
    print("- Feldbeschreibungen: Erklaerung aller Felder")
    print("- Dropdown_Werte: Verfuegbare Dropdown-Optionen")
    print()
    print(">> VERWENDUNG:")
    print("1. Oeffne die Excel-Datei")
    print("2. Bearbeite die Daten im 'Artikel_Import' Sheet")
    print("3. Speichere die Datei")
    print("4. Verwende sie mit dem AutoBANF-Import-Skript")
    
    return excel_filename

if __name__ == "__main__":
    create_complete_excel_template()