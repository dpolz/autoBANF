# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

AutoBANF is a Python automation tool for SAP procurement processes using Playwright web automation. The project automates article entry forms in SAP systems, including form filling, category selection, and data validation.

## Development Environment Setup

```bash
# Install dependencies
pip install -r requirements.txt

# Run the main application
python autoBANF_main.py

# Run the simplified final version
python autobanf_final.py

# Run specific modules
python autobanf_kategorien.py      # Category-focused automation
python autobanf_vollausfuellung.py # Complete form filling
python autobanf_katfokus.py        # Category focus functionality
```

## Architecture and Code Structure

### Core Modules

- **autobanf_base.py**: Foundation module containing:
  - `SecureCredentials` class for encrypted credential management
  - `navigate_to_artikel_page()` - Core navigation function
  - `analyze_form_fields()` - Form field detection and analysis
  - `fill_field_by_criteria()` - Intelligent form filling
  - Browser management utilities

- **autoBANF_main.py**: Main application with comprehensive form handling
- **autobanf_final.py**: Simplified, production-ready version focusing on reliability
- **autobanf_kategorien.py**: Specialized category selection automation
- **autobanf_vollausfuellung.py**: Complete form filling with advanced field mapping
- **autobanf_katfokus.py**: Category-focused automation with enhanced search

### Data Files

- **sap_kategorien_*.json**: Category data cache files for SAP integration
- **artikel_*_analyse.json**: Form field analysis results
- **autobanf_config.enc**: Encrypted credentials storage
- **.autobanf_key**: Encryption key file (auto-generated)

### Form Field Mapping Strategy

The project uses intelligent field mapping based on multiple criteria:
- Field labels, IDs, names, placeholders, and CSS classes
- Context-aware matching using search terms
- Prioritized mapping to avoid field conflicts
- Support for different field types (text, dropdown, date)

Example mapping pattern:
```python
field_mappings = [
    (['artikel', 'beschreibung'], 'Artikelbeschreibung'),
    (['steuer', 'kennzeichen'], 'Steuerkennzeichen'),
    (['preis', 'art'], 'Preisart'),
]
```

### Key Functions

- `get_credentials()`: Secure credential retrieval with fallback prompts
- `navigate_to_artikel_page()`: Complete SAP navigation workflow
- `analyze_form_fields()`: Dynamic form structure analysis
- `fill_field_by_criteria()`: Smart form filling with error handling
- `close_browser_safely()`: Proper browser cleanup

## Development Workflow

1. Use `autobanf_base.py` functions for core automation tasks
2. Test navigation with `navigate_to_artikel_page()` first
3. Use `analyze_form_fields()` to understand form structure before filling
4. Implement field mappings based on actual form analysis
5. Always use `close_browser_safely()` for proper cleanup

## Security Considerations

- Credentials are encrypted using Fernet symmetric encryption
- Config files have restricted permissions (0o600)
- No sensitive data in screenshots or logs
- Secure credential deletion functionality available

## Testing and Debugging

- Screenshots are automatically created at key steps
- Verbose mode available in analysis functions
- Error screenshots saved on failures
- Manual inspection periods built into workflows

## Common Patterns

- Always check field existence before interaction
- Use time.sleep() for UI stability after interactions
- Implement fallback selectors for UI elements
- Handle both success and failure cases gracefully
- Provide user feedback throughout automation process