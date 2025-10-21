# AI Assistance Instructions for Report Generator Backend

This document guides AI agents in understanding and working with this codebase effectively.

## Project Overview

This is a Python-based report generator system that creates customized Word documents from Excel data and templates. The project consists of:

- `app.py`: Main Streamlit web application for uploading files and generating reports
- `gerador_relatorios.py`: Core report generation logic and data processing

### Key Components

1. **Data Sources**:
   - Control Excel file (`.xlsx`) - Contains client information and contract details
   - Template Word file (`.docx`) - Base template for report generation
   - Tickets Excel file(s) (`.xlsx`) - Contains service tickets data

2. **Report Generation Flow**:
   - Load control data and ticket data from Excel files
   - Process each client's data separately
   - Generate Word reports using templates with placeholder substitution
   - Package multiple reports into a ZIP file

## Important Patterns

### Template Placeholders
Reports use a specific placeholder format in Word templates:
```
{{NOME_DO_CLIENTE}} - Client name
{{NUMERO_CONTRATO}} - Contract number
{{MES_REFERENCIA_MAIUSCULO}} - Reference month (uppercase)
{{tk_*}} - Ticket-related fields (numero, abertura, inicio, etc.)
```

### Data Processing
- All dates are expected to be in Brazilian format (dd/mm/yyyy)
- Locale settings attempt to use 'pt_BR.UTF-8' for proper month names
- Excel files must contain specific column names as referenced in the code

## Development Workflow

### Environment Setup
- Required Python packages: streamlit, pandas, python-docx, locale
- Locale setting for Brazilian Portuguese is important for date formatting

### Testing
- `gerador_relatorios.py` contains example test data in `DADOS_FALSOS_SERVICE_NOW`
- Use this for testing without real ServiceNow integration

### Common Operations
1. Adding new placeholder fields:
   - Add to `substituicoes_gerais` or `substituicoes_ticket` dictionaries
   - Update Word template with corresponding placeholders
   
2. Modifying report layout:
   - Edit Word template while maintaining placeholder syntax
   - Table structure in template is critical for ticket display

## Integration Points
- Excel files must match expected column names exactly
- Word templates must contain at least one table for ticket data
- File encodings should be UTF-8 compatible

## Troubleshooting
- Check Excel column names if KeyError occurs
- Verify locale settings if month names appear in English
- Ensure Word template contains required table structure