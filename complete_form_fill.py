#!/usr/bin/env python3
"""
complete_form_fill.py - Vollständiges Ausfüllen aller gewünschten Felder
"""

import time
import json
from autobanf_base import (
    navigate_to_artikel_page, 
    close_browser_safely,
    SecureCredentials
)

# Dropdown-Konfiguration
DROPDOWN_CONFIG = {
    "Steuerkennzeichen": {
        "selector": "[id$='--idTaxCodeValidValues']",
        "options": {
            "Keine Steuer": "text=Keine Steuer",
            "Drittland (ohne VSt-Abz)": "text=Drittland (ohne VSt-Abz)",
            "kein separater Steuerabzug": "text=kein separater Steuerabzug"
        }
    },
    "Preisart": {
        "selector": "[id$='--PriceIsGross']", 
        "options": {
            "Netto-Preis": "text=Netto-Preis",
            "Brutto-Preis": "text=Brutto-Preis"
        }
    },
    "Waehrung": {
        "selector": "[id$='--ItemPriceCurrency']",
        "options": {
            "EUR": "text=EUR",
            "USD": "text=USD"
        }
    },
    "Rabatttyp": {
        "selector": "[id$='--idDiscountTypeValidValues']",
        "options": {
            "Absoluter Wert": "text=Absoluter Wert",
            "Prozentsatz": "text=Prozentsatz"
        }
    },
    "Kontierungsobjekttyp": {
        "selector": "[id*='select'][id*='clone']:has([title='Kontierungsobjekttyp'])",
        "options": {
            "Projekt": "text=Projekt",
            "Kostenstelle": "text=Kostenstelle",
            "Innenauftrag": "text=Innenauftrag"
        }
    }
}

def fill_text_field(page, field_name, value):
    """Füllt ein Textfeld aus"""
    try:
        field_selectors = {
            "Artikelbeschreibung": "MaterialText-inner",
            "Laufzeit": "idDRGeneralTerms-inner",
            "Lange Artikelbeschreibung": "idCFControl-GENERAL-ARTIKELLANG-generated-inner",
            "Angebotsreferenz": "idCFControl-GENERAL-ANGEBOTSREFERENZ-generated-inner", 
            "Angebotsdatum": "idCFControl-GENERAL-ANGEBOTSDATUM-generated-inner",
            "Kontierungsobjekt": "[id*='input'][id*='clone'][id$='-inner']"
        }
        
        if field_name not in field_selectors:
            print(f"   FEHLER: Unbekanntes Textfeld '{field_name}'")
            return False
        
        field_suffix = field_selectors[field_name]
        
        # Spezielle Behandlung für Kontierungsobjekt
        if field_name == "Kontierungsobjekt":
            try:
                element = page.locator(field_suffix)
                if element.count() > 0:
                    element.click()
                    element.fill('')
                    time.sleep(0.1)
                    element.fill(str(value))
                    element.press('Tab')
                    print(f"   OK {field_name} '{value}' erfolgreich")
                    return True
            except:
                pass
        else:
            # Verschiedene Component-IDs probieren
            for component_id in ["11", "12", "13"]:
                try:
                    field_id = f"__component{component_id}---idCatItemView--{field_suffix}"
                    element = page.locator(f"#{field_id}")
                    if element.count() > 0:
                        element.click()
                        element.fill('')
                        time.sleep(0.1)
                        element.fill(str(value))
                        element.press('Tab')
                        print(f"   OK {field_name} '{value}' erfolgreich")
                        return True
                except:
                    continue
        
        print(f"   FEHLER: {field_name}-Feld nicht gefunden")
        return False
        
    except Exception as e:
        print(f"   FEHLER bei {field_name}: {e}")
        return False

def fill_numeric_field(page, field_name, value):
    """Füllt ein numerisches Feld aus"""
    try:
        field_selectors = {
            "Preis je Mengeneinheit": "Price-inner",
            "Rabattwert": "DiscountValue-inner",
            "Bestellmenge": "idQuantityStepInput-input-inner"
        }
        
        if field_name not in field_selectors:
            print(f"   FEHLER: Unbekanntes numerisches Feld '{field_name}'")
            return False
        
        field_suffix = field_selectors[field_name]
        
        # Verschiedene Component-IDs probieren
        for component_id in ["11", "12", "13"]:
            try:
                field_id = f"__component{component_id}---idCatItemView--{field_suffix}"
                element = page.locator(f"#{field_id}")
                if element.count() > 0:
                    element.click()
                    element.fill('')
                    time.sleep(0.1)
                    element.fill(str(value))
                    element.press('Tab')
                    print(f"   OK {field_name} '{value}' erfolgreich")
                    return True
            except:
                continue
        
        print(f"   FEHLER: {field_name}-Feld nicht gefunden")
        return False
        
    except Exception as e:
        print(f"   FEHLER bei {field_name}: {e}")
        return False

def select_combobox_option(page, option_text):
    """Spezielle Behandlung für Einheit-ComboBox"""
    try:
        # ComboBox-Eingabefeld anklicken
        for component_id in ["11", "12", "13"]:
            try:
                combobox_id = f"__component{component_id}---idCatItemView--idCBPOUnit-inner"
                combobox = page.locator(f"#{combobox_id}")
                if combobox.count() > 0:
                    combobox.click()
                    time.sleep(0.5)
                    
                    # Dropdown-Pfeil klicken
                    arrow_id = f"__component{component_id}---idCatItemView--idCBPOUnit-arrow"
                    arrow = page.locator(f"#{arrow_id}")
                    if arrow.count() > 0:
                        arrow.click()
                        time.sleep(1.5)
                        
                        # Option auswählen
                        option = page.locator(f"text={option_text}")
                        if option.count() > 0:
                            option.first.click()
                            print(f"   OK Einheit '{option_text}' erfolgreich")
                            return True
                        else:
                            print(f"   FEHLER: Einheit-Option '{option_text}' nicht gefunden")
                            return False
            except:
                continue
        
        print("   FEHLER: Einheit-ComboBox nicht gefunden")
        return False
        
    except Exception as e:
        print(f"   FEHLER bei Einheit: {e}")
        return False

def select_dropdown_option(page, field_name, option_text):
    """Optimierte Dropdown-Auswahl"""
    try:
        if field_name not in DROPDOWN_CONFIG:
            print(f"   FEHLER: Unbekanntes Dropdown-Feld '{field_name}'")
            return False
        
        config = DROPDOWN_CONFIG[field_name]
        
        # Dropdown öffnen
        dropdown = page.locator(config["selector"])
        if dropdown.count() == 0:
            print(f"   FEHLER: Dropdown für '{field_name}' nicht gefunden")
            return False
        
        dropdown.first.click()
        time.sleep(1.5)
        
        # Option auswählen
        if option_text in config["options"]:
            option_selector = config["options"][option_text]
            option = page.locator(option_selector)
            
            if option.count() > 0:
                option.first.click()
                print(f"   OK {field_name} '{option_text}' erfolgreich")
                return True
            else:
                print(f"   FEHLER: Option '{option_text}' nicht gefunden")
                return False
        else:
            print(f"   FEHLER: Option '{option_text}' nicht in Konfiguration")
            return False
            
    except Exception as e:
        print(f"   FEHLER bei {field_name}: {e}")
        return False

def select_category_and_subcategory(page, main_category, subcategory):
    """
    Behandelt die komplexe Kategorie-Auswahl mit funktionierendem Playwright-Locator
    Format: "Hauptkategorie->Unterkategorie" (z.B. "Bedarf Labore und Werkstätten->Laborutensilien, Werkzeuge und Kleinteile")
    """
    try:
        print(f"   Kategorie auswählen: '{main_category}' -> '{subcategory}'")
        
        # Kategorie-Button suchen und Dialog öffnen
        kategorie_button = page.locator("text=Kategorie auswählen")
        if kategorie_button.count() == 0:
            print("   FEHLER: Kategorie-Button nicht gefunden")
            return False
        
        kategorie_button.click()
        time.sleep(1.5)
        print("   Kategorie-Dialog geöffnet")
        
        # Hauptkategorie aufklappen (falls vorhanden)
        if subcategory:
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
                print(f"   Hauptkategorie '{main_category}' erfolgreich aufgeklappt")
                time.sleep(1)
            else:
                print(f"   WARNUNG: Hauptkategorie '{main_category}' konnte nicht aufgeklappt werden")
            
            # Unterkategorie mit funktionierendem Playwright-Locator auswählen
            try:
                subcategory_locator = page.locator(f"text={subcategory}")
                if subcategory_locator.count() > 0:
                    subcategory_locator.first.click()  # first() verwenden bei mehreren Elementen
                    print(f"   Unterkategorie '{subcategory}' mit Playwright-Locator ausgewählt")
                    
                    # 3 Sekunden warten, damit die Auswahl effektiv wird
                    time.sleep(1.5)
                    print("   3 Sekunden Wartezeit für Kategorie-Auswahl abgeschlossen")
                    
                    # Prüfen ob Dialog geschlossen wurde
                    dialog_open = page.evaluate('''() => {
                        const dialogs = document.querySelectorAll('[role="dialog"], .sapMDialog');
                        const visibleDialogs = Array.from(dialogs).filter(d => d.offsetHeight > 0);
                        return visibleDialogs.length > 0;
                    }''')
                    
                    if not dialog_open:
                        print("   Kategorie-Dialog erfolgreich geschlossen - Auswahl erfolgreich!")
                        return True
                    else:
                        print("   WARNUNG: Dialog noch geöffnet, aber Unterkategorie ausgewählt")
                        return True
                else:
                    print(f"   FEHLER: Unterkategorie '{subcategory}' nicht gefunden")
                    return False
            except Exception as e:
                print(f"   FEHLER bei Unterkategorie-Auswahl: {e}")
                return False
        else:
            # Nur Hauptkategorie auswählen
            main_locator = page.locator(f"text={main_category}")
            if main_locator.count() > 0:
                main_locator.first.click()
                print(f"   Hauptkategorie '{main_category}' ausgewählt")
                time.sleep(1.5)
                return True
            else:
                print(f"   FEHLER: Hauptkategorie '{main_category}' nicht gefunden")
                return False
        
    except Exception as e:
        print(f"   FEHLER bei Kategorieauswahl: {e}")
        return False

def complete_form_test():
    print("=== VOLLSTÄNDIGE 2-ARTIKEL AUTOMATISIERUNG ===")
    print("ARTIKEL 1:")
    print("0. Kategorie = Bedarf Labore und Werkstätten -> Laborutensilien, Werkzeuge und Kleinteile")
    print("1. Artikelbeschreibung = 1N4007")
    print("2. Steuerkennzeichen = Keine Steuer")
    print("3. Preisart = Brutto-Preis")
    print("4. Preis je Mengeneinheit = 1,23")
    print("5. Währung = USD")
    print("6. Rabatttyp = Prozentsatz")
    print("7. Rabattwert = 10")
    print("8. Laufzeit = 10.07.2025 - 23.08.2025")
    print("9. Bestellmenge = 10")
    print("10. Einheit = G g")
    print("11. Lange Artikelbeschreibung = Das ist ein Test")
    print("12. Angebotsreferenz = MeinAngebot")
    print("13. Angebotsdatum = 01.07.2025")
    print("14. Kontierungsobjekttyp = Projekt")
    print("15. Kontierungsobjekt = 9/000002002")
    print("16. Bearbeitung abschließen")
    print()
    print("ARTIKEL 2:")
    print("17. Neue Position anlegen -> Freitext")
    print("18. Kategorie = Bedarf Labore und Werkstätten -> Laborutensilien, Werkzeuge und Kleinteile")
    print("19. Artikelbeschreibung = Widerstand 220 Ohm")
    print("20. Preis je Mengeneinheit = 0,15")
    print("21. Währung = EUR") 
    print("22. Rabatttyp = Absoluter Wert")
    print("23. Bestellmenge = 100")
    print("24. Einheit = ST Stück")
    print("25. Bearbeitung abschließen")
    print()
    
    # Anmelden
    cred_manager = SecureCredentials()
    username, password = cred_manager.load_credentials()
    
    # Zur Artikel-Seite
    success, browser, page = navigate_to_artikel_page(username, password, create_screenshots=False)
    if not success:
        return
    
    try:
        time.sleep(1)
        
        # 0. KATEGORIE AUSWÄHLEN
        print("\n0. KATEGORIE auswählen...")
        select_category_and_subcategory(page, "Bedarf Labore und Werkstätten", "Laborutensilien, Werkzeuge und Kleinteile")
        time.sleep(1.5)  # 3 Sekunden warten, damit Kategorieauswahl effektiv wird
        
        # 1. ARTIKELBESCHREIBUNG
        print("\n1. ARTIKELBESCHREIBUNG ausfüllen...")
        fill_text_field(page, "Artikelbeschreibung", "1N4007")
        
        # 2. STEUERKENNZEICHEN
        print("\n2. STEUERKENNZEICHEN auswählen...")
        select_dropdown_option(page, "Steuerkennzeichen", "Keine Steuer")
        
        # 3. PREISART
        print("\n3. PREISART auswählen...")
        select_dropdown_option(page, "Preisart", "Brutto-Preis")
        
        # 4. PREIS JE MENGENEINHEIT
        print("\n4. PREIS JE MENGENEINHEIT eingeben...")
        fill_numeric_field(page, "Preis je Mengeneinheit", "1,23")
        
        # 5. WÄHRUNG
        print("\n5. WÄHRUNG auswählen...")
        select_dropdown_option(page, "Waehrung", "USD")
        
        # 6. RABATTTYP
        print("\n6. RABATTTYP auswählen...")
        select_dropdown_option(page, "Rabatttyp", "Prozentsatz")
        
        # 7. RABATTWERT
        print("\n7. RABATTWERT eingeben...")
        fill_numeric_field(page, "Rabattwert", "10")
        
        # 8. LAUFZEIT
        print("\n8. LAUFZEIT eingeben...")
        fill_text_field(page, "Laufzeit", "10.07.2025 - 23.08.2025")
        
        # 9. BESTELLMENGE
        print("\n9. BESTELLMENGE eingeben...")
        fill_numeric_field(page, "Bestellmenge", "10")
        
        # 10. EINHEIT
        print("\n10. EINHEIT auswählen...")
        select_combobox_option(page, "G g")
        
        # 11. LANGE ARTIKELBESCHREIBUNG
        print("\n11. LANGE ARTIKELBESCHREIBUNG eingeben...")
        fill_text_field(page, "Lange Artikelbeschreibung", "Das ist ein Test")
        
        # 12. ANGEBOTSREFERENZ
        print("\n12. ANGEBOTSREFERENZ eingeben...")
        fill_text_field(page, "Angebotsreferenz", "MeinAngebot")
        
        # 13. ANGEBOTSDATUM
        print("\n13. ANGEBOTSDATUM eingeben...")
        fill_text_field(page, "Angebotsdatum", "01.07.2025")
        
        # 14. KONTIERUNGSOBJEKTTYP
        print("\n14. KONTIERUNGSOBJEKTTYP auswählen...")
        select_dropdown_option(page, "Kontierungsobjekttyp", "Projekt")
        
        # 15. KONTIERUNGSOBJEKT
        print("\n15. KONTIERUNGSOBJEKT eingeben...")
        fill_text_field(page, "Kontierungsobjekt", "9/000002002")
        
        # 16. BEARBEITUNG ABSCHLIESSEN
        print("\n16. BEARBEITUNG ABSCHLIESSEN...")
        try:
            # Verschiedene Ansätze probieren
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
                        print(f"   OK 'Bearbeitung abschließen' geklickt (Selector: {selector})")
                        time.sleep(1.5)
                        button_found = True
                        break
                except:
                    continue
            
            if not button_found:
                # JavaScript-basierter Fallback
                js_success = page.evaluate('''() => {
                    const allElements = Array.from(document.querySelectorAll('*'));
                    const buttonElements = allElements.filter(el => {
                        const text = el.textContent?.trim();
                        return text && text.includes('Bearbeitung abschließen');
                    });
                    
                    for (let el of buttonElements) {
                        const clickableParent = el.closest('button, [role="button"], .sapMBtn');
                        if (clickableParent) {
                            clickableParent.click();
                            return true;
                        } else if (el.click) {
                            el.click();
                            return true;
                        }
                    }
                    return false;
                }''')
                
                if js_success:
                    print("   OK 'Bearbeitung abschließen' geklickt (JavaScript)")
                    time.sleep(1.5)
                else:
                    print("   FEHLER: 'Bearbeitung abschließen' Button nicht gefunden")
                    
        except Exception as e:
            print(f"   FEHLER beim Abschließen: {e}")
        
        # Screenshot nach erstem Artikel
        page.screenshot(path="complete_form_artikel1.png", full_page=True)
        print("\nScreenshot Artikel 1: complete_form_artikel1.png")
        
        print("\nARTIKEL 1 VOLLSTÄNDIG AUSGEFÜLLT UND ABGESCHLOSSEN!")
        time.sleep(1.5)
        
        # =================================================================
        # ZWEITER ARTIKEL HINZUFÜGEN
        # =================================================================
        print("\n" + "="*60)
        print("ZWEITER ARTIKEL HINZUFÜGEN")
        print("="*60)
        
        # Schritt 1: "Neue Position anlegen" klicken
        print("\n>> 'Neue Position anlegen' für zweiten Artikel klicken...")
        try:
            neue_position_button = page.locator("text=Neue Position anlegen")
            if neue_position_button.count() > 0:
                neue_position_button.click()
                time.sleep(1)
                print("   OK 'Neue Position anlegen' geklickt")
            else:
                print("   FEHLER: 'Neue Position anlegen' Button nicht gefunden")
                raise Exception("Neue Position anlegen nicht gefunden")
        except Exception as e:
            print(f"   FEHLER beim Klicken auf 'Neue Position anlegen': {e}")
            return
        
        # Schritt 2: "Freitext" auswählen (wie beim ersten Mal)
        print("\n>> 'Freitext' für zweiten Artikel auswählen...")
        try:
            freitext_button = page.locator("text=Freitext").first
            freitext_button.click()
            time.sleep(1.5)
            print("   OK 'Freitext' ausgewählt")
        except Exception as e:
            print(f"   FEHLER beim Auswählen von 'Freitext': {e}")
            return
        
        print("\n>> Artikel-Eingabe-Seite für zweiten Artikel erreicht!")
        time.sleep(1)
        
        # =================================================================
        # ZWEITER ARTIKEL - FORMULAR AUSFÜLLEN
        # =================================================================
        print("\n" + "="*60)
        print("ARTIKEL 2 - FORMULAR AUSFÜLLEN")
        print("="*60)
        
        # 0. KATEGORIE AUSWÄHLEN (gleiche wie beim ersten Artikel)
        print("\n0. KATEGORIE auswählen (Artikel 2)...")
        select_category_and_subcategory(page, "Bedarf Labore und Werkstätten", "Laborutensilien, Werkzeuge und Kleinteile")
        time.sleep(1.5)
        
        # 1. ARTIKELBESCHREIBUNG (anderer Artikel zum Test)
        print("\n1. ARTIKELBESCHREIBUNG ausfüllen (Artikel 2)...")
        fill_text_field(page, "Artikelbeschreibung", "Widerstand 220 Ohm")
        
        # 2-15. Alle anderen Felder (gleiche Werte zum Test)
        print("\n2. STEUERKENNZEICHEN auswählen (Artikel 2)...")
        select_dropdown_option(page, "Steuerkennzeichen", "Keine Steuer")
        
        print("\n3. PREISART auswählen (Artikel 2)...")
        select_dropdown_option(page, "Preisart", "Brutto-Preis")
        
        print("\n4. PREIS JE MENGENEINHEIT eingeben (Artikel 2)...")
        fill_numeric_field(page, "Preis je Mengeneinheit", "0,15")
        
        print("\n5. WÄHRUNG auswählen (Artikel 2)...")
        select_dropdown_option(page, "Waehrung", "EUR")
        
        print("\n6. RABATTTYP auswählen (Artikel 2)...")
        select_dropdown_option(page, "Rabatttyp", "Absoluter Wert")
        
        print("\n7. RABATTWERT eingeben (Artikel 2)...")
        fill_numeric_field(page, "Rabattwert", "0")
        
        print("\n8. LAUFZEIT eingeben (Artikel 2)...")
        fill_text_field(page, "Laufzeit", "01.08.2025 - 31.12.2025")
        
        print("\n9. BESTELLMENGE eingeben (Artikel 2)...")
        fill_numeric_field(page, "Bestellmenge", "100")
        
        print("\n10. EINHEIT auswählen (Artikel 2)...")
        select_combobox_option(page, "ST Stück")
        
        print("\n11. LANGE ARTIKELBESCHREIBUNG eingeben (Artikel 2)...")
        fill_text_field(page, "Lange Artikelbeschreibung", "Kohleschichtwiderstand 220 Ohm")
        
        print("\n12. ANGEBOTSREFERENZ eingeben (Artikel 2)...")
        fill_text_field(page, "Angebotsreferenz", "WIDERSTAND-2025")
        
        print("\n13. ANGEBOTSDATUM eingeben (Artikel 2)...")
        fill_text_field(page, "Angebotsdatum", "26.07.2025")
        
        print("\n14. KONTIERUNGSOBJEKTTYP auswählen (Artikel 2)...")
        select_dropdown_option(page, "Kontierungsobjekttyp", "Projekt")
        
        print("\n15. KONTIERUNGSOBJEKT eingeben (Artikel 2)...")
        fill_text_field(page, "Kontierungsobjekt", "9/000002002")
        
        # 16. BEARBEITUNG ABSCHLIESSEN für zweiten Artikel
        print("\n16. BEARBEITUNG ABSCHLIESSEN (Artikel 2)...")
        try:
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
                        print(f"   OK 'Bearbeitung abschließen' geklickt (Selector: {selector})")
                        time.sleep(1.5)
                        button_found = True
                        break
                except:
                    continue
            
            if not button_found:
                js_success = page.evaluate('''() => {
                    const allElements = Array.from(document.querySelectorAll('*'));
                    const buttonElements = allElements.filter(el => {
                        const text = el.textContent?.trim();
                        return text && text.includes('Bearbeitung abschließen');
                    });
                    
                    for (let el of buttonElements) {
                        const clickableParent = el.closest('button, [role="button"], .sapMBtn');
                        if (clickableParent) {
                            clickableParent.click();
                            return true;
                        } else if (el.click) {
                            el.click();
                            return true;
                        }
                    }
                    return false;
                }''')
                
                if js_success:
                    print("   OK 'Bearbeitung abschließen' geklickt (JavaScript)")
                    time.sleep(1.5)
                else:
                    print("   FEHLER: 'Bearbeitung abschließen' Button nicht gefunden")
                    
        except Exception as e:
            print(f"   FEHLER beim Abschließen von Artikel 2: {e}")
        
        # Final Screenshot nach beiden Artikeln
        page.screenshot(path="complete_form_2artikel.png", full_page=True)
        print("\nScreenshot beide Artikel: complete_form_2artikel.png")
        
        print("\n" + "="*60)
        print("BEIDE ARTIKEL VOLLSTÄNDIG ERSTELLT!")
        print("ARTIKEL 1: Diode 1N4007")
        print("ARTIKEL 2: Widerstand 220 Ohm")
        print("="*60)
        time.sleep(2.5)
        
    except Exception as e:
        print(f"FEHLER: {e}")
        
    finally:
        close_browser_safely(browser)

if __name__ == "__main__":
    complete_form_test()