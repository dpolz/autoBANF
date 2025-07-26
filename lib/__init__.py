"""
autoBANF Library Module
======================

Dieses Paket enthält die Kernfunktionen für die autoBANF SAP-Automatisierung:

- autobanf_base: Basis-Funktionen für SAP-Navigation und Credential-Management
- complete_form_fill: Spezialisierte Formular-Ausfüllfunktionen
"""

# Imports für einfache Verwendung
from .autobanf_base import (
    SecureCredentials,
    get_credentials,
    create_browser_page,
    navigate_to_artikel_page,
    analyze_form_fields,
    fill_field_by_criteria,
    close_browser_safely
)

from .complete_form_fill import (
    fill_text_field,
    fill_numeric_field,
    select_dropdown_option,
    select_combobox_option,
    select_category_and_subcategory
)

__version__ = "1.0.0"
__author__ = "autoBANF Project"