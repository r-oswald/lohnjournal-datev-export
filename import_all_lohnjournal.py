#!/usr/bin/env python3
"""Import all Lohnjournal PDFs from a folder into SQLite and Excel."""

import os
import re
import argparse
from pathlib import Path

import pandas as pd
import pdfplumber

from lohnjournal_parser import CoordinateLohnjournalParser, LohnjournalDatabase

MONTH_ORDER = {
    'Januar': 1, 'Februar': 2, 'März': 3, 'April': 4, 'Mai': 5, 'Juni': 6,
    'Juli': 7, 'August': 8, 'September': 9, 'Oktober': 10, 'November': 11, 'Dezember': 12
}

EXPORT_COLUMNS = [
    'pers_nr', 'name', 'steuerklasse', 'faktor', 'ki_freibetrag',
    'steuerbrutto', 'pausch_verst_bezuege', 'kv_brutto', 'rv_brutto', 'av_brutto', 'pv_brutto', 'gesamtbrutto',
    'lohnsteuer', 'pausch_lohnsteuer', 'kirchensteuer', 'pausch_kirchensteuer',
    'solidaritaetszuschlag', 'pausch_solidaritaetszuschlag', 'foerderbetrag',
    'kv_beitrag_an', 'rv_beitrag_an', 'av_beitrag_an', 'pv_beitrag_an',
    'kv_beitrag_ag', 'rv_beitrag_ag', 'av_beitrag_ag', 'pv_beitrag_ag',
    'umlage_1', 'umlage_2', 'umlage_insolvenz',
    'netto_bezuege', 'auszahlungsbetrag', 'st_tage', 'sv_tage'
]

SUM_COLUMNS = EXPORT_COLUMNS[5:]  # All numeric columns after ki_freibetrag


def extract_month_year(filename: str) -> tuple:
    """Extract month and year from filename like 'Januar_2025.pdf'."""
    if match := re.search(r'(\w+)_(\d{4})', filename):
        month_name, year = match.group(1), int(match.group(2))
        return month_name, year, MONTH_ORDER.get(month_name, 0)
    return None, None, 0


def find_password(pdf_path: str, passwords: list) -> str | None:
    """Find working password for PDF."""
    for pwd in passwords:
        try:
            with pdfplumber.open(pdf_path, password=pwd) as pdf:
                _ = pdf.pages[0].extract_text()
                return pwd
        except Exception:
            continue
    return None


def process_pdfs(pdf_folder: str, passwords: list) -> list:
    """Process all PDFs in folder, return list of results."""
    # Find and sort PDF files
    pdf_files = []
    for f in os.listdir(pdf_folder):
        if f.lower().endswith('.pdf'):
            month_name, year, month_num = extract_month_year(f)
            if month_name:
                pdf_files.append({
                    'filename': f,
                    'path': os.path.join(pdf_folder, f),
                    'month_name': month_name,
                    'year': year,
                    'month_num': month_num,
                    'sort_key': year * 100 + month_num
                })
    
    pdf_files.sort(key=lambda x: x['sort_key'])
    print(f"Found {len(pdf_files)} PDF files")
    
    results = []
    for pf in pdf_files:
        print(f"\nProcessing: {pf['month_name']} {pf['year']}")
        
        password = find_password(pf['path'], passwords)
        if not password and passwords[0]:  # Skip password check if no password needed
            print(f"  ERROR: Could not find password for {pf['filename']}")
            continue
        
        try:
            parser = CoordinateLohnjournalParser(pf['path'], password=password)
            employees = parser.parse()
            print(f"  Extracted: {len(employees)} employees")
            
            results.append({
                'month_name': pf['month_name'],
                'year': pf['year'],
                'month_num': pf['month_num'],
                'sort_key': pf['sort_key'],
                'table_name': re.sub(r'[^a-zA-Z0-9_]', '_', f"lohnjournal_{pf['month_name']}_{pf['year']}"),
                'employees': employees,
                'metadata': parser.metadata
            })
        except Exception as e:
            print(f"  ERROR: {e}")
    
    return results


def save_to_database(results: list, db_path: str):
    """Save all results to SQLite database."""
    if os.path.exists(db_path):
        os.remove(db_path)
    
    db = LohnjournalDatabase(db_path)
    for result in results:
        print(f"Saving {result['month_name']} {result['year']}...")
        db.create_table(result['table_name'])
        db.insert_employees(result['table_name'], result['employees'])
    db.close()
    print(f"Database saved: {db_path}")


def create_summary(results: list) -> tuple[pd.DataFrame, list]:
    """Create summary DataFrame aggregating all months."""
    employee_data = {}
    months = []
    
    for result in results:
        months.append(f"{result['month_name']} {result['year']}")
        
        for emp in result['employees']:
            key = emp.pers_nr
            if key not in employee_data:
                employee_data[key] = {'pers_nr': key, 'name': emp.name, 'months_count': 0}
                for col in SUM_COLUMNS:
                    employee_data[key][col] = 0.0
            
            if emp.name and len(emp.name) > len(employee_data[key].get('name', '')):
                employee_data[key]['name'] = emp.name
            
            employee_data[key]['months_count'] += 1
            
            for col in SUM_COLUMNS:
                val = getattr(emp, col, None)
                if val is not None:
                    employee_data[key][col] += val
    
    df = pd.DataFrame(list(employee_data.values())).sort_values('pers_nr')
    return df, months


def export_to_excel(results: list, excel_path: str):
    """Export all data to Excel with summary and monthly sheets."""
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        # Summary sheet
        summary_df, months = create_summary(results)
        
        header = pd.DataFrame([
            {'Info': 'ZUSAMMENFASSUNG', 'Value': ''},
            {'Info': 'Zeitraum:', 'Value': f"{months[0]} - {months[-1]}" if months else 'N/A'},
            {'Info': 'Anzahl Monate:', 'Value': len(months)},
            {'Info': '', 'Value': ''},
        ])
        header.to_excel(writer, sheet_name='Zusammenfassung', index=False)
        
        summary_cols = ['pers_nr', 'name', 'months_count'] + SUM_COLUMNS
        available_cols = [c for c in summary_cols if c in summary_df.columns]
        summary_df[available_cols].to_excel(writer, sheet_name='Zusammenfassung', index=False, startrow=6)
        print(f"  Created summary: {len(summary_df)} employees")
        
        # Monthly sheets
        for result in results:
            sheet_name = f"{result['month_name']}_{result['year']}"[:31]
            
            rows = [{col: getattr(emp, col, None) for col in EXPORT_COLUMNS} for emp in result['employees']]
            pd.DataFrame(rows).to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"  Created: {sheet_name} ({len(rows)} employees)")
    
    print(f"Excel exported: {excel_path}")


def main():
    parser = argparse.ArgumentParser(description='Import Lohnjournal PDFs to SQLite and Excel')
    parser.add_argument('--pdf-folder', '-p', default='./pdfs', help='Folder containing PDF files')
    parser.add_argument('--db', '-d', help='Database output path')
    parser.add_argument('--excel', '-e', help='Excel output path')
    parser.add_argument('--name', '-n', default='lohnjournal_complete', help='Base name for output files')
    parser.add_argument('--password', '-P', help='PDF password')
    
    args = parser.parse_args()
    
    base_dir = Path(__file__).parent
    db_path = args.db or base_dir / f'{args.name}.db'
    excel_path = args.excel or base_dir / f'{args.name}_export.xlsx'
    passwords = [args.password] if args.password else ['']
    
    print("=" * 50)
    print("LOHNJOURNAL IMPORT")
    print("=" * 50)
    
    results = process_pdfs(args.pdf_folder, passwords)
    
    if not results:
        print("No PDFs processed!")
        return
    
    print(f"\nProcessed {len(results)} months")
    
    save_to_database(results, str(db_path))
    export_to_excel(results, str(excel_path))
    
    total = sum(len(r['employees']) for r in results)
    print(f"\nComplete! {len(results)} months, {total} total records")
    print(f"Database: {db_path}")
    print(f"Excel: {excel_path}")


if __name__ == '__main__':
    main()
