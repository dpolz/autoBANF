#!/usr/bin/env python3
"""
excel_import_with_params.py - AutoBANF Excel-Import mit Dateinamen-Parameter
"""

import sys
import time
import os
import pandas as pd
from lib.autobanf_base import (
    navigate_to_artikel_page, 
    close_browser_safely,
    SecureCredentials
)

# Import der Funktionen aus complete_form_fill.py
from lib.complete_form_fill import (
    fill_text_field,
    fill_numeric_field,
    select_dropdown_option,
    select_combobox_option
)

def convert_to_german_number(value):
    """
    Konvertiert einen numerischen Wert zur deutschen Notation (Komma als Dezimaltrenner)
    """
    if pd.isna(value):
        return ""
    
    # Konvertiere zu String falls es ein numerischer Wert ist
    str_value = str(value).strip()
    
    # Wenn bereits ein Komma enthalten ist, lasse es so
    if ',' in str_value:
        return str_value
    
    # Wenn es ein Punkt als Dezimaltrenner gibt, ersetze durch Komma
    if '.' in str_value:
        return str_value.replace('.', ',')
    
    # Für ganze Zahlen: prüfe ob es als Dezimalzahl behandelt werden sollte
    try:
        num_value = float(str_value)
        # Wenn es eine ganze Zahl ist, gebe sie ohne Dezimalstellen zurück
        if num_value == int(num_value):
            return str(int(num_value))
        else:
            # Konvertiere zu deutscher Notation mit Komma
            return str(num_value).replace('.', ',')
    except ValueError:
        # Falls keine Zahl, gebe ursprünglichen Wert zurück
        return str_value

def select_category_robust(page, main_category, subcategory, artikel_nr):
    """
    Robuste Kategorieauswahl die verschiedene Zustände des Dialogs behandelt
    """
    try:
        print(f"   Robuste Kategorie-Auswahl für Artikel {artikel_nr}: '{main_category}' -> '{subcategory}'")
        
        # Schritt 1: Prüfen ob Kategorie-Dialog bereits offen ist
        dialog_already_open = page.evaluate('''() => {
            const dialogs = document.querySelectorAll('[role="dialog"], .sapMDialog');
            return Array.from(dialogs).some(d => d.offsetHeight > 0);
        }''')
        
        if not dialog_already_open:
            # Dialog muss geöffnet werden
            print(f"   Dialog ist geschlossen - öffne für Artikel {artikel_nr}")
            kategorie_button = page.locator("text=Kategorie auswählen")
            if kategorie_button.count() > 0:
                kategorie_button.click()
                time.sleep(1.5)
                print("   Dialog geöffnet")
            else:
                print("   WARNUNG: Kategorie-Button nicht gefunden")
                return False
        else:
            print(f"   Dialog bereits offen für Artikel {artikel_nr}")
        
        # Schritt 2: Warten bis Dialog vollständig geladen ist
        time.sleep(1)
        
        # Schritt 3: Hauptkategorie aufklappen (falls nötig)
        print(f"   Klappe Hauptkategorie '{main_category}' auf...")
        expand_success = page.evaluate(f'''() => {{
            const allElements = Array.from(document.querySelectorAll('*'));
            const categoryElements = allElements.filter(el => {{
                const text = el.textContent?.trim();
                return text === '{main_category}' || 
                       (text && text.includes('{main_category}') && text.length < 100);
            }});
            
            for (let categoryEl of categoryElements) {{
                const parent = categoryEl.closest('[role="treeitem"], li, .sapMLIB');
                if (parent) {{
                    const expandSelectors = [
                        'button:first-child',
                        'span:first-child',
                        '[class*="expand"]:first-child',
                        '[class*="arrow"]:first-child',
                        '[class*="toggle"]:first-child'
                    ];
                    
                    for (let selector of expandSelectors) {{
                        const expandBtn = parent.querySelector(selector);
                        if (expandBtn && expandBtn !== categoryEl) {{
                            try {{
                                expandBtn.click();
                                console.log("Hauptkategorie aufgeklappt:", '{main_category}');
                                return true;
                            }} catch (e) {{
                                console.log("Expand-Klick fehlgeschlagen:", e);
                            }}
                        }}
                    }}
                }}
            }}
            return false;
        }}''')
        
        if expand_success:
            print(f"   Hauptkategorie '{main_category}' aufgeklappt")
            time.sleep(0.5)
        else:
            print(f"   Hauptkategorie '{main_category}' schon aufgeklappt oder nicht gefunden")
        
        # Schritt 4: Unterkategorie auswählen - mit mehreren Versuchen
        print(f"   Wähle Unterkategorie '{subcategory}' aus...")
        
        # Versuch 1: Mit Playwright Locator
        try:
            subcategory_locator = page.locator(f"text={subcategory}").first
            if subcategory_locator.count() > 0:
                # Prüfe ob Element sichtbar ist
                is_visible = subcategory_locator.is_visible()
                print(f"   Unterkategorie gefunden, sichtbar: {is_visible}")
                
                if is_visible:
                    subcategory_locator.click(timeout=10000)
                    print(f"   Unterkategorie '{subcategory}' erfolgreich geklickt!")
                    time.sleep(0.5)
                    return True
                else:
                    print(f"   Unterkategorie '{subcategory}' nicht sichtbar - verwende JavaScript")
        except Exception as e:
            print(f"   Playwright-Klick fehlgeschlagen: {e}")
        
        # Versuch 2: Mit JavaScript - spezifisch für SAP CategorySelPopover
        print(f"   Verwende JavaScript für Unterkategorie '{subcategory}' (SAP-spezifisch)...")
        js_success = page.evaluate(f'''() => {{
            // Suche speziell nach SAP Category Tree Items
            console.log("=== SAP Kategorie-Suche für '{subcategory}' ===");
            
            // 1. Suche nach allen CategorySelPopover Tree Items
            const categoryTreeItems = document.querySelectorAll('[id*="CategorySelPopover--idCategoryTree"][role="treeitem"]');
            console.log("CategorySelPopover Tree Items gefunden:", categoryTreeItems.length);
            
            for (let item of categoryTreeItems) {{
                const contentDiv = item.querySelector('.sapMLIBContent');
                const itemText = contentDiv ? contentDiv.textContent?.trim() : '';
                
                console.log("Item ID:", item.id, "Text:", itemText);
                
                if (itemText === '{subcategory}') {{
                    const rect = item.getBoundingClientRect();
                    const isVisible = rect.width > 0 && rect.height > 0;
                    const isInPopover = item.closest('[id*="CategorySelPopover"]');
                    
                    console.log("=== MATCH GEFUNDEN ===");
                    console.log("Item ID:", item.id);
                    console.log("Sichtbar:", isVisible);
                    console.log("In Popover:", !!isInPopover);
                    console.log("Position:", rect);
                    console.log("Classes:", item.className);
                    
                    if (isVisible && isInPopover) {{
                        // Klick-Strategien in optimierter Reihenfolge
                        
                        // 1. Focus + Enter (SAP-typisch)
                        try {{
                            item.focus();
                            item.dispatchEvent(new KeyboardEvent('keydown', {{
                                key: 'Enter',
                                code: 'Enter',
                                which: 13,
                                keyCode: 13,
                                bubbles: true
                            }}));
                            console.log("Focus + Enter verwendet");
                            return true;
                        }} catch (e1) {{
                            console.log("Focus + Enter fehlgeschlagen:", e1);
                        }}
                        
                        // 2. Auf Content Div klicken
                        try {{
                            contentDiv.click();
                            console.log("Content Div geklickt");
                            return true;
                        }} catch (e2) {{
                            console.log("Content Div Klick fehlgeschlagen:", e2);
                        }}
                        
                        // 3. MouseEvent auf Tree Item
                        try {{
                            const clickEvent = new MouseEvent('click', {{
                                bubbles: true,
                                cancelable: true,
                                view: window,
                                clientX: rect.left + rect.width/2,
                                clientY: rect.top + rect.height/2
                            }});
                            item.dispatchEvent(clickEvent);
                            console.log("MouseEvent mit Koordinaten verwendet");
                            return true;
                        }} catch (e3) {{
                            console.log("MouseEvent fehlgeschlagen:", e3);
                        }}
                        
                        // 4. Direkter Tree Item Klick
                        try {{
                            item.click();
                            console.log("Tree Item direkt geklickt");
                            return true;
                        }} catch (e4) {{
                            console.log("Tree Item Klick fehlgeschlagen:", e4);
                        }}
                    }} else {{
                        console.log("Element nicht klickbar - nicht sichtbar oder nicht in Popover");
                    }}
                }}
            }}
            
            console.log("=== KEINE KLICKBARE UNTERKATEGORIE GEFUNDEN ===");
            return false;
        }}''')
        
        if js_success:
            print(f"   Unterkategorie '{subcategory}' mit SAP-JavaScript erfolgreich geklickt!")
            time.sleep(0.5)
            return True
        else:
            print(f"   FEHLER: Unterkategorie '{subcategory}' konnte nicht geklickt werden")
            # Dialog schließen
            page.keyboard.press("Escape")
            time.sleep(0.5)
            return False
            
    except Exception as e:
        print(f"   FEHLER bei robuster Kategorieauswahl: {e}")
        try:
            page.keyboard.press("Escape")
        except:
            pass
        return False


def read_excel_file(filename):
    """Liest Excel-Datei und gibt DataFrame zurück"""
    try:
        print(f">> Excel-Datei einlesen: {filename}")
        df = pd.read_excel(filename, sheet_name='Artikel_Import')
        print(f">> {len(df)} Artikel gefunden")
        print(f">> Spalten: {list(df.columns)}")
        return df
    except FileNotFoundError:
        print(f"FEHLER: Datei '{filename}' nicht gefunden!")
        return None
    except Exception as e:
        print(f"FEHLER beim Lesen der Excel-Datei: {e}")
        return None

def fill_article_form(page, row_data, artikel_nr):
    """Füllt das Formular für einen Artikel aus"""
    print(f"\n{'='*60}")
    print(f"ARTIKEL {artikel_nr} - FORMULAR AUSFÜLLEN")
    print(f"{'='*60}")
    
    success_count = 0
    total_fields = 16
    
    try:
        # 0. KATEGORIE AUSWÄHLEN (robuste Logik für alle Artikel)
        if pd.notna(row_data.get('Kategorie')):
            print(f"\n0. KATEGORIE auswählen (Artikel {artikel_nr})...")
            kategorie = str(row_data['Kategorie']).strip()
            if '->' in kategorie:
                main_cat, sub_cat = kategorie.split('->', 1)
                # Robuste Kategorieauswahl mit mehreren Versuchen
                if select_category_robust(page, main_cat.strip(), sub_cat.strip(), artikel_nr):
                    success_count += 1
                time.sleep(1.5)
            
        # 1. ARTIKELBESCHREIBUNG
        if pd.notna(row_data.get('Artikelbeschreibung')):
            print(f"\n1. ARTIKELBESCHREIBUNG ausfüllen (Artikel {artikel_nr})...")
            if fill_text_field(page, "Artikelbeschreibung", str(row_data['Artikelbeschreibung'])):
                success_count += 1
                
        # 2. STEUERKENNZEICHEN
        if pd.notna(row_data.get('Steuerkennzeichen')):
            print(f"\n2. STEUERKENNZEICHEN auswählen (Artikel {artikel_nr})...")
            if select_dropdown_option(page, "Steuerkennzeichen", str(row_data['Steuerkennzeichen'])):
                success_count += 1
                
        # 3. PREISART
        if pd.notna(row_data.get('Preisart')):
            print(f"\n3. PREISART auswählen (Artikel {artikel_nr})...")
            if select_dropdown_option(page, "Preisart", str(row_data['Preisart'])):
                success_count += 1
                
        # 4. PREIS JE MENGENEINHEIT
        if pd.notna(row_data.get('Preis_je_Mengeneinheit')):
            print(f"\n4. PREIS JE MENGENEINHEIT eingeben (Artikel {artikel_nr})...")
            german_price = convert_to_german_number(row_data['Preis_je_Mengeneinheit'])
            if fill_numeric_field(page, "Preis je Mengeneinheit", german_price):
                success_count += 1
                
        # 5. WÄHRUNG
        if pd.notna(row_data.get('Waehrung')):
            print(f"\n5. WÄHRUNG auswählen (Artikel {artikel_nr})...")
            if select_dropdown_option(page, "Waehrung", str(row_data['Waehrung'])):
                success_count += 1
                
        # 6. RABATTTYP
        if pd.notna(row_data.get('Rabatttyp')):
            print(f"\n6. RABATTTYP auswählen (Artikel {artikel_nr})...")
            if select_dropdown_option(page, "Rabatttyp", str(row_data['Rabatttyp'])):
                success_count += 1
                
        # 7. RABATTWERT
        if pd.notna(row_data.get('Rabattwert')):
            print(f"\n7. RABATTWERT eingeben (Artikel {artikel_nr})...")
            german_rabatt = convert_to_german_number(row_data['Rabattwert'])
            if fill_numeric_field(page, "Rabattwert", german_rabatt):
                success_count += 1
                
        # 8. LAUFZEIT
        if pd.notna(row_data.get('Laufzeit')):
            print(f"\n8. LAUFZEIT eingeben (Artikel {artikel_nr})...")
            if fill_text_field(page, "Laufzeit", str(row_data['Laufzeit'])):
                success_count += 1
                
        # 9. BESTELLMENGE
        if pd.notna(row_data.get('Bestellmenge')):
            print(f"\n9. BESTELLMENGE eingeben (Artikel {artikel_nr})...")
            german_menge = convert_to_german_number(row_data['Bestellmenge'])
            if fill_numeric_field(page, "Bestellmenge", german_menge):
                success_count += 1
                
        # 10. EINHEIT
        if pd.notna(row_data.get('Einheit')):
            print(f"\n10. EINHEIT auswählen (Artikel {artikel_nr})...")
            if select_combobox_option(page, str(row_data['Einheit'])):
                success_count += 1
                
        # 11. LANGE ARTIKELBESCHREIBUNG
        if pd.notna(row_data.get('Lange_Artikelbeschreibung')):
            print(f"\n11. LANGE ARTIKELBESCHREIBUNG eingeben (Artikel {artikel_nr})...")
            if fill_text_field(page, "Lange Artikelbeschreibung", str(row_data['Lange_Artikelbeschreibung'])):
                success_count += 1
                
        # 12. ANGEBOTSREFERENZ
        if pd.notna(row_data.get('Angebotsreferenz')):
            print(f"\n12. ANGEBOTSREFERENZ eingeben (Artikel {artikel_nr})...")
            if fill_text_field(page, "Angebotsreferenz", str(row_data['Angebotsreferenz'])):
                success_count += 1
                
        # 13. ANGEBOTSDATUM
        if pd.notna(row_data.get('Angebotsdatum')):
            print(f"\n13. ANGEBOTSDATUM eingeben (Artikel {artikel_nr})...")
            if fill_text_field(page, "Angebotsdatum", str(row_data['Angebotsdatum'])):
                success_count += 1
                
        # 14. KONTIERUNGSOBJEKTTYP
        if pd.notna(row_data.get('Kontierungsobjekttyp')):
            print(f"\n14. KONTIERUNGSOBJEKTTYP auswählen (Artikel {artikel_nr})...")
            if select_dropdown_option(page, "Kontierungsobjekttyp", str(row_data['Kontierungsobjekttyp'])):
                success_count += 1
                
        # 15. KONTIERUNGSOBJEKT
        if pd.notna(row_data.get('Kontierungsobjekt')):
            print(f"\n15. KONTIERUNGSOBJEKT eingeben (Artikel {artikel_nr})...")
            if fill_text_field(page, "Kontierungsobjekt", str(row_data['Kontierungsobjekt'])):
                success_count += 1
                
        # 16. BEARBEITUNG ABSCHLIESSEN
        print(f"\n16. BEARBEITUNG ABSCHLIESSEN (Artikel {artikel_nr})...")
        selectors_to_try = [
            "[id*='idButtonTransfer']",
            "[id$='--idButtonTransfer']",
            "button:has-text('Bearbeitung abschließen')",
            "text=Bearbeitung abschließen",
            "[id*='BDI-content']:has-text('Bearbeitung abschließen')"
        ]
        
        button_found = False
        for selector in selectors_to_try:
            try:
                bearbeitung_button = page.locator(selector)
                if bearbeitung_button.count() > 0:
                    bearbeitung_button.first.click()
                    print(f"   OK 'Bearbeitung abschließen' geklickt")
                    time.sleep(1.5)
                    button_found = True
                    success_count += 1
                    break
            except:
                continue
        
        if not button_found:
            print("   FEHLER: 'Bearbeitung abschließen' Button nicht gefunden")
            
    except Exception as e:
        print(f"FEHLER bei Artikel {artikel_nr}: {e}")
    
    print(f"\nArtikel {artikel_nr} abgeschlossen: {success_count}/{total_fields} Felder erfolgreich")
    return success_count, total_fields

def add_new_article_position(page, artikel_nr):
    """Fügt eine neue Artikelposition hinzu"""
    if artikel_nr == 1:
        return True  # Erste Position ist bereits da
        
    print(f"\n{'='*60}")
    print(f"NEUE POSITION FÜR ARTIKEL {artikel_nr} HINZUFÜGEN")
    print(f"{'='*60}")
    
    try:
        # "Neue Position anlegen" klicken
        print(f"\n>> 'Neue Position anlegen' für Artikel {artikel_nr} klicken...")
        neue_position_button = page.locator("text=Neue Position anlegen")
        if neue_position_button.count() > 0:
            neue_position_button.click()
            time.sleep(0.5)
            print("   OK 'Neue Position anlegen' geklickt")
        else:
            print("   FEHLER: 'Neue Position anlegen' Button nicht gefunden")
            return False
            
        # "Freitext" auswählen
        print(f"\n>> 'Freitext' für Artikel {artikel_nr} auswählen...")
        freitext_button = page.locator("text=Freitext").first
        freitext_button.click()
        time.sleep(1.5)
        print("   OK 'Freitext' ausgewählt")
        
        print(f"\n>> Artikel-Eingabe-Seite für Artikel {artikel_nr} erreicht!")
        time.sleep(1)
        return True
        
    except Exception as e:
        print(f"FEHLER beim Hinzufügen neuer Position für Artikel {artikel_nr}: {e}")
        return False

def excel_import_test(excel_filename):
    """Hauptfunktion für Excel-Import-Test"""
    print("=== AUTOBANF EXCEL-IMPORT TEST ===")
    print(f"Excel-Datei: {excel_filename}")
    print()
    
    # Excel-Datei einlesen
    df = read_excel_file(excel_filename)
    if df is None:
        return
        
    print(f"\n>> {len(df)} Artikel werden importiert:")
    for i, row in df.iterrows():
        artikel_name = row.get('Artikelbeschreibung', f'Artikel {i+1}')
        print(f"   {i+1}. {artikel_name}")
    print()
    
    # Anmelden
    cred_manager = SecureCredentials()
    username, password = cred_manager.get_credentials_interactive()
    
    if not username or not password:
        print("ABBRUCH: Keine gültigen Anmeldedaten erhalten")
        return
    
    # Zur Artikel-Seite navigieren
    success, browser, page = navigate_to_artikel_page(username, password, create_screenshots=False)
    if not success:
        return
        
    try:
        time.sleep(1)
        
        # Temp-Ordner für Screenshots erstellen
        temp_dir = ".temp"
        if not os.path.exists(temp_dir):
            os.makedirs(temp_dir)
            print(f">> Temp-Ordner '{temp_dir}' erstellt")
        
        total_success = 0
        total_fields = 0
        
        # Alle Artikel durchgehen
        for index, row in df.iterrows():
            artikel_nr = index + 1
            
            # Neue Position hinzufügen (außer beim ersten Artikel)
            if not add_new_article_position(page, artikel_nr):
                print(f"ABBRUCH: Konnte keine neue Position für Artikel {artikel_nr} hinzufügen")
                break
                
            # Artikel-Formular ausfüllen
            success_count, field_count = fill_article_form(page, row, artikel_nr)
            total_success += success_count
            total_fields += field_count
            
            # Screenshot nach jedem Artikel
            screenshot_path = os.path.join(temp_dir, f"excel_import_artikel_{artikel_nr}.png")
            page.screenshot(path=screenshot_path, full_page=True)
            print(f"Screenshot: {screenshot_path}")
            
        # Final Screenshot
        final_screenshot_path = os.path.join(temp_dir, "excel_import_complete.png")
        page.screenshot(path=final_screenshot_path, full_page=True)
        print(f"\nFinal Screenshot: {final_screenshot_path}")
        
        print(f"\n{'='*60}")
        print("EXCEL-IMPORT ABGESCHLOSSEN!")
        print(f"Artikel verarbeitet: {len(df)}")
        print(f"Gesamterfolg: {total_success}/{total_fields} Felder ({total_success/total_fields*100:.1f}%)")
        print(f"{'='*60}")
        
    except Exception as e:
        print(f"FEHLER: {e}")
        error_screenshot_path = os.path.join(temp_dir, "excel_import_error.png")
        page.screenshot(path=error_screenshot_path, full_page=True)
        
    finally:
        close_browser_safely(browser)

def main():
    """Hauptprogramm mit Parameterverarbeitung"""
    if len(sys.argv) != 2:
        print("VERWENDUNG: python excel_import_with_params.py <excel_filename>")
        print("BEISPIEL:   python excel_import_with_params.py Test.xlsx")
        return
        
    excel_filename = sys.argv[1]
    excel_import_test(excel_filename)

if __name__ == "__main__":
    main()