"""
autobanf_base.py - Basis-Module für autoBANF
============================================

Diese Datei enthält alle Basis-Funktionen für die autoBANF-Automatisierung:
- Sichere Anmeldedaten-Verwaltung
- SAP-Navigation bis zur Artikel-Eingabe-Seite
- Utility-Funktionen für Playwright

Verwendung:
    from autobanf_base import navigate_to_artikel_page, get_credentials
"""

from playwright.sync_api import sync_playwright
import time
import json
import os
from pathlib import Path
from cryptography.fernet import Fernet
import getpass

class SecureCredentials:
    """
    Sichere Speicherung und Verwaltung von Anmeldedaten mit Verschlüsselung
    """
    
    def __init__(self, config_file="autobanf_config.enc"):
        # Credentials-Ordner erstellen falls nicht vorhanden
        self.cred_dir = Path("credentials")
        self.cred_dir.mkdir(exist_ok=True)
        
        self.config_file = self.cred_dir / config_file
        self.key_file = self.cred_dir / ".autobanf_key"
        
    def _get_or_create_key(self):
        """Erstellt oder lädt den Verschlüsselungsschlüssel"""
        if self.key_file.exists():
            with open(self.key_file, 'rb') as f:
                key = f.read()
        else:
            key = Fernet.generate_key()
            with open(self.key_file, 'wb') as f:
                f.write(key)
            try:
                os.chmod(self.key_file, 0o600)
            except:
                pass
        return key
    
    def save_credentials(self, username, password):
        """Speichert Anmeldedaten verschlüsselt"""
        try:
            key = self._get_or_create_key()
            fernet = Fernet(key)
            
            credentials = {
                "username": username,
                "password": password,
                "saved_at": time.time()
            }
            
            json_data = json.dumps(credentials).encode()
            encrypted_data = fernet.encrypt(json_data)
            
            with open(self.config_file, 'wb') as f:
                f.write(encrypted_data)
            
            try:
                os.chmod(self.config_file, 0o600)
            except:
                pass
                
            print(">> Anmeldedaten sicher gespeichert")
            return True
            
        except Exception as e:
            print(f">> Fehler beim Speichern: {e}")
            return False
    
    def load_credentials(self):
        """Lädt und entschlüsselt Anmeldedaten"""
        try:
            if not self.config_file.exists():
                return None, None
            
            key = self._get_or_create_key()
            fernet = Fernet(key)
            
            with open(self.config_file, 'rb') as f:
                encrypted_data = f.read()
            
            decrypted_data = fernet.decrypt(encrypted_data)
            credentials = json.loads(decrypted_data.decode())
            
            return credentials["username"], credentials["password"]
            
        except Exception as e:
            print(f">> Fehler beim Laden der Anmeldedaten: {e}")
            return None, None
    
    def get_credentials_interactive(self):
        """
        Lädt Credentials oder fragt interaktiv nach, falls nicht vorhanden
        """
        # Zuerst versuchen zu laden
        username, password = self.load_credentials()
        
        if username and password:
            print(f">> Anmeldedaten für Benutzer '{username}' geladen")
            return username, password
        
        # Falls nicht vorhanden, interaktiv nachfragen
        print("\n" + "="*60)
        print("ANMELDEDATEN EINGEBEN")
        print("="*60)
        print("Keine gespeicherten Anmeldedaten gefunden.")
        print("Bitte geben Sie Ihre SAP-Anmeldedaten ein:")
        print()
        
        username = input("Benutzername: ").strip()
        if not username:
            print("FEHLER: Benutzername darf nicht leer sein!")
            return None, None
        
        password = getpass.getpass("Passwort: ").strip()
        if not password:
            print("FEHLER: Passwort darf nicht leer sein!")
            return None, None
        
        # Speichern anbieten
        save_choice = input("\nAnmeldedaten für zukünftige Verwendung speichern? (j/n): ").strip().lower()
        if save_choice in ['j', 'ja', 'y', 'yes']:
            if self.save_credentials(username, password):
                print(">> Anmeldedaten sicher gespeichert")
            else:
                print(">> WARNUNG: Anmeldedaten konnten nicht gespeichert werden")
        
        return username, password
    
    def delete_credentials(self):
        """Löscht gespeicherte Anmeldedaten"""
        try:
            if self.config_file.exists():
                self.config_file.unlink()
            if self.key_file.exists():
                self.key_file.unlink()
            print(">> Anmeldedaten geloescht")
            return True
        except Exception as e:
            print(f">> Fehler beim Loeschen: {e}")
            return False

def get_credentials():
    """
    Lädt Anmeldedaten aus sicherem Speicher oder fragt sie ab
    
    Returns:
        tuple: (username, password, cred_manager) oder (None, None, None)
    """
    cred_manager = SecureCredentials()
    
    username, password = cred_manager.load_credentials()
    
    if username and password:
        print(f">> Gespeicherte Anmeldedaten gefunden fuer: {username}")
        use_saved = input("Gespeicherte Anmeldedaten verwenden? (j/n): ").strip().lower()
        
        if use_saved in ['j', 'ja', 'y', 'yes', '']:
            return username, password, cred_manager
    
    print("\n>> Anmeldedaten eingeben:")
    username = input("HSA-Benutzername: ").strip()
    if not username:
        return None, None, None
    
    password = getpass.getpass("HSA-Passwort: ")
    if not password:
        return None, None, None
    
    save_creds = input("\nAnmeldedaten sicher speichern? (j/n): ").strip().lower()
    if save_creds in ['j', 'ja', 'y', 'yes']:
        cred_manager.save_credentials(username, password)
    
    return username, password, cred_manager

def create_browser_page(headless=False, slow_mo=800, viewport_width=1600, viewport_height=1000):
    """
    Erstellt einen Browser und eine Seite mit Standardeinstellungen
    
    Args:
        headless (bool): Browser im Hintergrund ausführen
        slow_mo (int): Millisekunden zwischen Aktionen (reduziert von 1500 auf 800)
        viewport_width (int): Browser-Breite
        viewport_height (int): Browser-Höhe
    
    Returns:
        tuple: (browser, page)
    """
    playwright = sync_playwright().start()
    browser = playwright.chromium.launch(headless=headless, slow_mo=slow_mo)
    page = browser.new_page()
    page.set_viewport_size({"width": viewport_width, "height": viewport_height})
    return browser, page

def navigate_to_artikel_page(username, password, browser=None, page=None, create_screenshots=True):
    """
    Navigiert zur SAP Artikel-Eingabe-Seite
    
    Durchläuft alle Schritte:
    Login > GISA easyBANF > Neuer Artikel > Neue Position anlegen > Freitext
    
    Args:
        username (str): HSA-Benutzername
        password (str): HSA-Passwort  
        browser: Existing browser instance (optional)
        page: Existing page instance (optional)
        create_screenshots (bool): Screenshots erstellen
    
    Returns:
        tuple: (success, browser, page) - browser und page für weitere Verwendung
    """
    
    # Browser erstellen falls nicht übergeben
    own_browser = browser is None
    if own_browser:
        browser, page = create_browser_page()
    
    try:
        print("\n=== Navigation zur Artikel-Eingabe-Seite ===")
        
        # ========================================
        # SCHRITT 1: LOGIN
        # ========================================
        print(">> SCHRITT 1: Anmeldung bei SAP")
        page.goto("https://prod.sap.hsa.fms-bayern.de/", timeout=30000)
        page.wait_for_load_state("networkidle", timeout=15000)
        page.wait_for_selector("input[name='username']", timeout=10000)
        
        page.locator("input[name='username']").fill(username)
        page.locator("input[name='password']").fill(password)
        page.locator("button[type='submit']").click()
        
        time.sleep(3)
        page.wait_for_function(
            "() => !window.location.href.includes('loginuserpass.php')",
            timeout=30000
        )
        print(">> Login erfolgreich")
        
        if create_screenshots:
            page.screenshot(path="base_01_nach_login.png")
        
        # ========================================
        # SCHRITT 2: GISA easyBANF
        # ========================================
        print("\n>> SCHRITT 2: GISA easyBANF klicken")
        page.wait_for_load_state("networkidle", timeout=15000)
        time.sleep(2)  # Reduziert von 3 auf 2
        
        gisa_element = page.locator("text=GISA easyBANF")
        gisa_element.wait_for(state="visible", timeout=10000)
        gisa_element.click()
        print(">> GISA easyBANF geklickt")
        
        time.sleep(2)  # Reduziert von 4 auf 2
        page.wait_for_load_state("networkidle", timeout=15000)
        
        if create_screenshots:
            page.screenshot(path="base_02_warenkorb.png")
        
        # ========================================
        # SCHRITT 3: NEUER ARTIKEL
        # ========================================
        print("\n>> SCHRITT 3: 'Neuer Artikel' klicken")
        
        neuer_artikel_element = page.locator("text=Neuer Artikel")
        neuer_artikel_element.wait_for(state="visible", timeout=10000)
        neuer_artikel_element.click()
        print(">> 'Neuer Artikel' geklickt")
        
        time.sleep(2)  # Reduziert von 4 auf 2
        page.wait_for_load_state("networkidle", timeout=15000)
        
        if create_screenshots:
            page.screenshot(path="base_03_nach_neuer_artikel.png")
        
        # ========================================
        # SCHRITT 4: NEUE POSITION ANLEGEN
        # ========================================
        print("\n>> SCHRITT 4: 'Neue Position anlegen' suchen und klicken")
        
        position_selectors = [
            "text=Neue Position anlegen",
            "*:has-text('Neue Position anlegen')",
            "button:has-text('Neue Position anlegen')",
            "a:has-text('Neue Position anlegen')",
            "*[title*='Neue Position anlegen']"
        ]
        
        position_element = None
        for selector in position_selectors:
            try:
                elements = page.locator(selector)
                if elements.count() > 0:
                    position_element = elements.first
                    print(f">> 'Neue Position anlegen' gefunden mit: {selector}")
                    break
            except:
                continue
        
        if not position_element:
            print(">> 'Neue Position anlegen' nicht gefunden!")
            return False, browser, page
        
        position_element.click()
        print(">> 'Neue Position anlegen' geklickt")
        
        time.sleep(2)  # Reduziert von 4 auf 2
        page.wait_for_load_state("networkidle", timeout=15000)
        
        if create_screenshots:
            page.screenshot(path="base_04_nach_neue_position.png")
        
        # ========================================
        # SCHRITT 5: FREITEXT AUSWÄHLEN
        # ========================================
        print("\n>> SCHRITT 5: 'Freitext' auswaehlen")
        
        # Erst direkt nach Freitext-Elementen suchen
        freitext_selectors = [
            "text=Freitext",
            "*:has-text('Freitext')",
            "option:has-text('Freitext')",
            "button:has-text('Freitext')",
            "a:has-text('Freitext')"
        ]
        
        freitext_element = None
        for selector in freitext_selectors:
            try:
                elements = page.locator(selector)
                if elements.count() > 0:
                    freitext_element = elements.first
                    print(f">> 'Freitext' gefunden mit: {selector}")
                    break
            except:
                continue
        
        # Falls nicht direkt gefunden, in Dropdowns suchen
        if not freitext_element:
            print(">> Suche 'Freitext' in Dropdown-Menues...")
            
            selects = page.locator("select")
            select_count = selects.count()
            
            for i in range(select_count):
                try:
                    select_element = selects.nth(i)
                    options = select_element.locator("option")
                    option_count = options.count()
                    
                    for j in range(option_count):
                        try:
                            option_text = options.nth(j).inner_text().strip()
                            
                            if "freitext" in option_text.lower():
                                print(f">> FREITEXT GEFUNDEN in Dropdown! Waehle: '{option_text}'")
                                select_element.select_option(index=j)
                                freitext_element = "found_in_dropdown"
                                time.sleep(2)  # Reduziert von 3 auf 2
                                break
                        except:
                            continue
                    
                    if freitext_element:
                        break
                        
                except:
                    continue
        
        # Normales Element klicken falls gefunden
        if freitext_element and freitext_element != "found_in_dropdown":
            freitext_element.click()
            print(">> 'Freitext' geklickt")
            time.sleep(2)  # Reduziert von 3 auf 2
        elif freitext_element == "found_in_dropdown":
            print(">> 'Freitext' aus Dropdown ausgewaehlt")
        else:
            print(">> 'Freitext' nicht gefunden!")
            return False, browser, page
        
        # ========================================
        # SCHRITT 6: ARTIKEL-SEITE ERREICHT
        # ========================================
        print("\n>> SCHRITT 6: Artikel-Eingabe-Seite erreicht!")
        
        time.sleep(2)  # Reduziert von 4 auf 2
        page.wait_for_load_state("networkidle", timeout=15000)
        
        if create_screenshots:
            page.screenshot(path="base_05_artikelseite.png")
        
        # Seiteninformationen
        current_url = page.url
        current_title = page.title()
        print(f">> Aktuelle Seite: {current_title}")
        print(f">> URL: {current_url}")
        
        return True, browser, page
        
    except Exception as e:
        print(f">> Fehler bei der Navigation: {e}")
        
        if create_screenshots:
            try:
                page.screenshot(path="base_fehler.png")
                print(">> Fehler-Screenshot gespeichert")
            except:
                pass
        
        return False, browser, page

def analyze_form_fields(page, verbose=True):
    """
    Analysiert alle Formularfelder auf der aktuellen Seite
    
    Args:
        page: Playwright page object
        verbose (bool): Detaillierte Ausgabe
    
    Returns:
        list: Liste von Formularfeld-Informationen
    """
    if verbose:
        print("\n>> FORMULARFELD-ANALYSE")
    
    fields = []
    inputs = page.locator("input, select, textarea")
    input_count = inputs.count()
    
    if verbose:
        print(f">> Gefundene Eingabefelder: {input_count}")
    
    for i in range(input_count):
        try:
            field = inputs.nth(i)
            
            # Nur sichtbare und aktivierte Felder
            if field.is_visible() and field.is_enabled():
                field_info = {
                    'index': i,
                    'element': field,
                    'type': field.get_attribute("type") or "text",
                    'id': field.get_attribute("id") or "",
                    'name': field.get_attribute("name") or "",
                    'placeholder': field.get_attribute("placeholder") or "",
                    'class': field.get_attribute("class") or "",
                    'value': field.input_value() if field.get_attribute("type") != "password" else "",
                    'tag': field.evaluate("el => el.tagName")
                }
                
                # Versuche Label zu finden
                try:
                    if field_info['id']:
                        label_element = page.locator(f"label[for='{field_info['id']}']")
                        if label_element.count() > 0:
                            field_info['label'] = label_element.inner_text().strip()
                        else:
                            field_info['label'] = ""
                    else:
                        field_info['label'] = ""
                except:
                    field_info['label'] = ""
                
                fields.append(field_info)
                
                if verbose:
                    info_text = f"[{field_info['tag']}:{field_info['type']}]"
                    if field_info['label']:
                        info_text += f" Label:'{field_info['label']}'"
                    if field_info['id']:
                        info_text += f" ID:{field_info['id']}"
                    if field_info['name']:
                        info_text += f" Name:{field_info['name']}"
                    if field_info['placeholder']:
                        info_text += f" Placeholder:'{field_info['placeholder']}'"
                    
                    print(f"  {i+1}. {info_text}")
        except:
            continue
    
    return fields

def fill_field_by_criteria(page, field_info, value, verbose=True):
    """
    Füllt ein Formularfeld mit einem Wert aus
    
    Args:
        page: Playwright page object
        field_info (dict): Feld-Informationen von analyze_form_fields()
        value (str): Einzutragender Wert
        verbose (bool): Detaillierte Ausgabe
    
    Returns:
        bool: Erfolg der Operation
    """
    try:
        field = field_info['element']
        
        if field_info['tag'].upper() == 'SELECT':
            # Dropdown: Versuche verschiedene Methoden
            try:
                field.select_option(label=str(value))
                if verbose:
                    print(f">> Dropdown gefuellt (Label): {value}")
                return True
            except:
                try:
                    field.select_option(value=str(value))
                    if verbose:
                        print(f">> Dropdown gefuellt (Value): {value}")
                    return True
                except:
                    if verbose:
                        print(f">> Dropdown-Wert nicht verfuegbar: {value}")
                    return False
        else:
            # Text-Feld
            field.fill(str(value))
            if verbose:
                print(f">> Feld gefuellt: {value}")
            return True
            
    except Exception as e:
        if verbose:
            print(f">> Fehler beim Ausfuellen: {e}")
        return False

def close_browser_safely(browser):
    """
    Schließt Browser sicher
    
    Args:
        browser: Playwright browser object
    """
    try:
        browser.close()
        print("🔚 Browser geschlossen")
    except:
        print(">> Browser bereits geschlossen")

# Beispiel-Daten für Tests
EXAMPLE_FORM_DATA = {
    "Artikelbeschreibung": "Testartikel",
    "Steuerkennzeichen": "kein separater Steuerabzug", 
    "Preisart": "Netto-Preis",
    "Netto-Preis je Mengeneinheit": "1,23",
    "Währung": "EUR",
    "Rabattyp": "Absoluter Wert",
    "Rabattwert": "0",
    "Lange Artikelbeschreibung": "Das ist ein Automatisierungstest",
    "Angebotsreferenz": "ABC123456",
    "Angebotsdatum": "25.07.2025"
}

if __name__ == "__main__":
    print("autobanf_base.py - Basis-Module für autoBANF")
    print("Diese Datei ist zum Importieren gedacht, nicht zur direkten Ausführung.")
    print("\nVerfügbare Funktionen:")
    print("- get_credentials(): Anmeldedaten verwalten")
    print("- navigate_to_artikel_page(): Zur Artikel-Seite navigieren")
    print("- analyze_form_fields(): Formularfelder analysieren")
    print("- fill_field_by_criteria(): Einzelnes Feld ausfüllen")
    print("- create_browser_page(): Browser erstellen")
    print("- close_browser_safely(): Browser schließen")