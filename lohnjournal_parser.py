#!/usr/bin/env python3
"""DATEV Lohnjournal PDF Parser - Coordinate-based extraction."""

import re
import sqlite3
import argparse
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional, List, Dict, Any
from collections import defaultdict

import pdfplumber


@dataclass
class EmployeeRecord:
    """Employee payroll record from Lohnjournal."""
    pers_nr: str
    name: str = ""
    steuerklasse: Optional[str] = None
    faktor: Optional[str] = None
    ki_freibetrag: Optional[str] = None
    
    # Brutto
    steuerbrutto: Optional[float] = None
    pausch_verst_bezuege: Optional[float] = None
    kv_brutto: Optional[float] = None
    rv_brutto: Optional[float] = None
    av_brutto: Optional[float] = None
    pv_brutto: Optional[float] = None
    gesamtbrutto: Optional[float] = None
    
    # Taxes
    lohnsteuer: Optional[float] = None
    pausch_lohnsteuer: Optional[float] = None
    kirchensteuer: Optional[float] = None
    pausch_kirchensteuer: Optional[float] = None
    solidaritaetszuschlag: Optional[float] = None
    pausch_solidaritaetszuschlag: Optional[float] = None
    foerderbetrag: Optional[float] = None
    
    # Employee contributions (AN)
    kv_beitrag_an: Optional[float] = None
    rv_beitrag_an: Optional[float] = None
    av_beitrag_an: Optional[float] = None
    pv_beitrag_an: Optional[float] = None
    
    # Employer contributions (AG)
    kv_beitrag_ag: Optional[float] = None
    rv_beitrag_ag: Optional[float] = None
    av_beitrag_ag: Optional[float] = None
    pv_beitrag_ag: Optional[float] = None
    
    # Umlagen
    umlage_1: Optional[float] = None
    umlage_2: Optional[float] = None
    umlage_insolvenz: Optional[float] = None
    
    # Netto
    netto_bezuege: Optional[float] = None
    auszahlungsbetrag: Optional[float] = None
    
    # Metadata
    sv_tage: Optional[int] = None
    st_tage: Optional[int] = None
    sub_row_codes: List[str] = field(default_factory=list)
    raw_lines: List[str] = field(default_factory=list)


def parse_german_number(value: str) -> Optional[float]:
    """Parse DATEV number format (e.g., '2.43000' = 2430.00, '18041' = 180.41)."""
    if not value or value.strip() in ('', 'Z', 'E'):
        return None
    
    value = value.strip()
    is_negative = value.endswith('-')
    if is_negative:
        value = value[:-1]
    
    value = re.sub(r'[^\d]', '', value)
    if not value:
        return None
    
    try:
        # Last 2 digits are cents
        result = int(value) / 100 if len(value) > 2 else float(value) / 100
        return -result if is_negative else result
    except ValueError:
        return None


# Column X-coordinate ranges: (min_x, max_x)
COLUMNS = {
    # Main row (employee header)
    'main': {
        'steuerklasse': (60, 78),
        'faktor': (80, 115),
        'ki_freibetrag': (95, 135),
        'name': (140, 350),
        'kv_brutto': (455, 510),
        'rv_brutto': (515, 570),
        'av_brutto': (575, 630),
        'pv_brutto': (635, 690),
        'umlage_1': (700, 745),
        'gesamtbrutto': (765, 820),
    },
    # Tax row (codes 1-6)
    'tax': {
        'st_tage': (130, 155),
        'steuerbrutto': (170, 225),
        'lohnsteuer': (245, 310),
        'kirchensteuer': (310, 380),
        'solidaritaetszuschlag': (380, 445),
        'kv_beitrag_an': (465, 510),
        'rv_beitrag_an': (525, 570),
        'av_beitrag_an': (592, 625),
        'pv_beitrag_an': (652, 690),
        'umlage_2': (705, 745),
        'netto_bezuege': (775, 825),
    },
    # AG row (01111, 00110)
    'ag': {
        'sv_tage': (130, 155),
        'pausch_verst_bezuege': (185, 225),
        'pausch_lohnsteuer': (268, 305),
        'pausch_kirchensteuer': (345, 375),
        'pausch_solidaritaetszuschlag': (418, 450),
        'kv_beitrag_ag': (465, 510),
        'rv_beitrag_ag': (525, 570),
        'av_beitrag_ag': (590, 630),
        'pv_beitrag_ag': (650, 690),
        'umlage_insolvenz': (710, 745),
        'auszahlungsbetrag': (765, 825),
    },
    # Minijob row (26500, 26100)
    'minijob': {
        'sv_tage': (130, 155),
        'pausch_verst_bezuege': (180, 230),
        'pausch_lohnsteuer': (265, 310),
        'kv_beitrag_ag': (465, 510),
        'rv_beitrag_ag': (525, 570),
        'umlage_insolvenz': (705, 745),
        'auszahlungsbetrag': (775, 825),
    },
}

# Fields that need numeric parsing
NUMERIC_FIELDS = {
    'kv_brutto', 'rv_brutto', 'av_brutto', 'pv_brutto', 'gesamtbrutto',
    'steuerbrutto', 'lohnsteuer', 'kirchensteuer', 'solidaritaetszuschlag',
    'kv_beitrag_an', 'rv_beitrag_an', 'av_beitrag_an', 'pv_beitrag_an',
    'kv_beitrag_ag', 'rv_beitrag_ag', 'av_beitrag_ag', 'pv_beitrag_ag',
    'umlage_1', 'umlage_2', 'umlage_insolvenz', 'netto_bezuege', 'auszahlungsbetrag',
    'pausch_verst_bezuege', 'pausch_lohnsteuer', 'pausch_kirchensteuer',
    'pausch_solidaritaetszuschlag',
}

# Fields that are integers
INT_FIELDS = {'st_tage', 'sv_tage'}

AG_ROW_CODES = {'01111', '00110'}
MINIJOB_ROW_CODES = {'26500', '26100'}
TAX_ROW_CODES = {'1', '2', '3', '4', '5', '6'}


class CoordinateLohnjournalParser:
    """Parser using coordinate-based column detection."""
    
    def __init__(self, pdf_path: str, password: str | None = None):
        self.pdf_path = pdf_path
        self.password = password
        self.employees: List[EmployeeRecord] = []
        self.metadata: Dict[str, Any] = {}
        
    def parse(self) -> List[EmployeeRecord]:
        """Parse the PDF and return employee records."""
        with pdfplumber.open(self.pdf_path, password=self.password) as pdf:
            pages = [
                (i, p) for i, p in enumerate(pdf.pages)
                if 'Lohnjournal' in (p.extract_text() or '') and 'Form.-Nr.LOA313' in (p.extract_text() or '')
            ]
            
            print(f"Found {len(pages)} Lohnjournal pages")
            
            for page_num, page in pages:
                self._parse_page(page)
                
            if pages:
                self._extract_metadata(pages[0][1])
                
        return self.employees
    
    def _extract_metadata(self, page):
        """Extract document metadata from first page."""
        text = page.extract_text() or ''
        patterns = {
            'berater': r'Berater:\s*(\d+)',
            'mandant': r'Mandant:\s*(\d+)',
            'datum': r'Datum:\s*([\d.]+)',
            'monat': r'Lohnjournal\s+(\w+\s+\d{4})',
        }
        for key, pattern in patterns.items():
            if match := re.search(pattern, text):
                self.metadata[key] = match.group(1)
        print(f"Metadata: {self.metadata}")
    
    def _parse_page(self, page):
        """Parse a single page."""
        words = page.extract_words(keep_blank_chars=False, x_tolerance=3, y_tolerance=3)
        
        # Group words by Y coordinate
        rows = defaultdict(list)
        for word in words:
            y_key = round(word['top'] / 2) * 2
            rows[y_key].append(word)
        
        current_emp = None
        
        for y_pos, row_words in sorted(rows.items()):
            row_words.sort(key=lambda w: w['x0'])
            if not row_words or y_pos < 95:
                continue
            
            first = row_words[0]['text']
            
            # New employee row: 5-digit ID at left edge
            if (first.isdigit() and len(first) == 5 and 
                first not in AG_ROW_CODES and first not in MINIJOB_ROW_CODES and
                row_words[0]['x0'] < 35):
                
                if current_emp:
                    self.employees.append(current_emp)
                current_emp = EmployeeRecord(pers_nr=first)
                self._parse_row(current_emp, row_words, 'main')
                
            elif current_emp:
                # Sub-row for current employee
                current_emp.raw_lines.append(' '.join(w['text'] for w in row_words))
                
                if first in TAX_ROW_CODES:
                    current_emp.sub_row_codes.append(first)
                    self._parse_row(current_emp, row_words, 'tax', skip_x=130)
                elif first in AG_ROW_CODES:
                    current_emp.sub_row_codes.append(first)
                    self._parse_row(current_emp, row_words, 'ag', skip_x=65)
                elif first in MINIJOB_ROW_CODES:
                    current_emp.sub_row_codes.append(first)
                    self._parse_row(current_emp, row_words, 'minijob', skip_x=65)
        
        if current_emp:
            self.employees.append(current_emp)
    
    def _parse_row(self, emp: EmployeeRecord, words: List[dict], row_type: str, skip_x: float = 60):
        """Parse a row using column definitions."""
        if row_type == 'main':
            emp.raw_lines.append(' '.join(w['text'] for w in words))
        
        columns = COLUMNS[row_type]
        
        for word in words:
            x0 = word['x0']
            text = word['text']
            
            if x0 < skip_x or text in ('Z', 'E'):
                continue
            
            # Special handling for name field (accumulates text)
            if row_type == 'main' and 'name' in columns:
                name_range = columns['name']
                if name_range[0] <= x0 <= name_range[1] and not text.replace('.', '').replace(',', '').isdigit():
                    emp.name = f"{emp.name} {text}".strip() if emp.name else text
                    continue
            
            # Match word to column
            for field_name, (min_x, max_x) in columns.items():
                if field_name == 'name':
                    continue
                if min_x <= x0 <= max_x:
                    self._set_field(emp, field_name, text)
                    break
        
        # Clean name
        if row_type == 'main' and emp.name:
            emp.name = re.sub(r'\s*NB\s*$', '', emp.name).strip()
    
    def _set_field(self, emp: EmployeeRecord, field: str, text: str):
        """Set a field value with appropriate type conversion."""
        if field in NUMERIC_FIELDS:
            value = parse_german_number(text)
            if value is not None:
                setattr(emp, field, value)
        elif field in INT_FIELDS:
            if text.isdigit():
                setattr(emp, field, int(text))
        else:
            setattr(emp, field, text)


class LohnjournalDatabase:
    """SQLite database handler for Lohnjournal data."""
    
    FIELDS = [
        'pers_nr', 'name', 'steuerklasse', 'faktor', 'ki_freibetrag',
        'steuerbrutto', 'pausch_verst_bezuege', 'kv_brutto', 'rv_brutto', 
        'av_brutto', 'pv_brutto', 'gesamtbrutto',
        'lohnsteuer', 'pausch_lohnsteuer', 'kirchensteuer', 'pausch_kirchensteuer',
        'solidaritaetszuschlag', 'pausch_solidaritaetszuschlag', 'foerderbetrag',
        'kv_beitrag_an', 'rv_beitrag_an', 'av_beitrag_an', 'pv_beitrag_an',
        'kv_beitrag_ag', 'rv_beitrag_ag', 'av_beitrag_ag', 'pv_beitrag_ag',
        'umlage_1', 'umlage_2', 'umlage_insolvenz',
        'netto_bezuege', 'auszahlungsbetrag',
        'st_tage', 'sv_tage', 'sub_row_codes', 'raw_lines',
    ]
    
    def __init__(self, db_path: str):
        self.conn = sqlite3.connect(db_path)
        
    def create_table(self, table_name: str) -> str:
        """Create table for Lohnjournal data."""
        table_name = re.sub(r'[^a-zA-Z0-9_]', '_', table_name)
        
        self.conn.execute(f"""
            CREATE TABLE IF NOT EXISTS {table_name} (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                pers_nr TEXT NOT NULL, name TEXT, steuerklasse TEXT, faktor TEXT, ki_freibetrag TEXT,
                steuerbrutto REAL, pausch_verst_bezuege REAL, kv_brutto REAL, rv_brutto REAL,
                av_brutto REAL, pv_brutto REAL, gesamtbrutto REAL,
                lohnsteuer REAL, pausch_lohnsteuer REAL, kirchensteuer REAL, pausch_kirchensteuer REAL,
                solidaritaetszuschlag REAL, pausch_solidaritaetszuschlag REAL, foerderbetrag REAL,
                kv_beitrag_an REAL, rv_beitrag_an REAL, av_beitrag_an REAL, pv_beitrag_an REAL,
                kv_beitrag_ag REAL, rv_beitrag_ag REAL, av_beitrag_ag REAL, pv_beitrag_ag REAL,
                umlage_1 REAL, umlage_2 REAL, umlage_insolvenz REAL,
                netto_bezuege REAL, auszahlungsbetrag REAL,
                st_tage INTEGER, sv_tage INTEGER, sub_row_codes TEXT, raw_lines TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        self.conn.commit()
        return table_name
    
    def insert_employees(self, table_name: str, employees: List[EmployeeRecord]):
        """Insert employee records."""
        table_name = re.sub(r'[^a-zA-Z0-9_]', '_', table_name)
        placeholders = ', '.join(['?'] * len(self.FIELDS))
        columns = ', '.join(self.FIELDS)
        
        for emp in employees:
            values = []
            for f in self.FIELDS:
                val = getattr(emp, f, None)
                if f == 'sub_row_codes':
                    val = ','.join(val) if val else ''
                elif f == 'raw_lines':
                    val = '\n'.join(val) if val else ''
                values.append(val)
            self.conn.execute(f"INSERT INTO {table_name} ({columns}) VALUES ({placeholders})", values)
        
        self.conn.commit()
        
    def close(self):
        self.conn.close()


def main():
    """CLI entry point."""
    parser = argparse.ArgumentParser(description='Parse DATEV Lohnjournal PDF')
    parser.add_argument('pdf_path', help='Path to PDF file')
    parser.add_argument('--password', '-p', default='', help='PDF password')
    parser.add_argument('--output', '-o', default='lohnjournal.db', help='Output SQLite database')
    parser.add_argument('--table', '-t', help='Table name')
    parser.add_argument('--debug', '-d', action='store_true', help='Print debug info')
    
    args = parser.parse_args()
    
    print(f"Parsing {args.pdf_path}...")
    lj_parser = CoordinateLohnjournalParser(args.pdf_path, args.password)
    employees = lj_parser.parse()
    
    print(f"\nFound {len(employees)} employees")
    
    if args.debug:
        for emp in employees[:10]:
            print(f"\n{'='*60}")
            print(f"Pers.-Nr.: {emp.pers_nr}, Name: {emp.name}")
            print(f"Steuerbrutto: {emp.steuerbrutto}, Lohnsteuer: {emp.lohnsteuer}")
            print(f"Netto: {emp.netto_bezuege}, Auszahlung: {emp.auszahlungsbetrag}")
    
    table_name = args.table or f"lohnjournal_{lj_parser.metadata.get('monat', Path(args.pdf_path).stem).replace(' ', '_')}"
    
    print(f"\nSaving to {args.output}, table: {table_name}")
    db = LohnjournalDatabase(args.output)
    db.create_table(table_name)
    db.insert_employees(table_name, employees)
    db.close()
    
    print(f"\nSummary: {len(employees)} employees, "
          f"{sum(1 for e in employees if e.steuerbrutto)} with Steuerbrutto, "
          f"{sum(1 for e in employees if e.auszahlungsbetrag)} with Auszahlung")


if __name__ == '__main__':
    main()
