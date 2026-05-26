"""
Arctic Q1 Project — PDF text extraction script.

Usage:
1. Put this script in the Arctic_Q1_Project folder.
2. Ensure the folder structure is:
   01_Strategy_Documents/<Country>/<PDF files>
3. Run: python extract_arctic_pdfs.py

Outputs:
- extracted_texts/<Country>/<PDF as TXT>
- document_text_quality_report.csv
- extraction_summary.json
"""
from pathlib import Path
import csv
import json
import re
import fitz

SOURCE_DIR = Path('01_Strategy_Documents')
OUTPUT_ROOT = Path('extracted_texts')

EXPECTED = [
    ('USA','USA_2022_National_Arctic_Strategy.pdf','Yes'),
    ('USA','USA_2024_DoD_Arctic_Strategy.pdf','Yes'),
    ('Canada','Canada_2019_Arctic_Northern_Policy_Framework.pdf','Yes'),
    ('Canada','Canada_2024_Our_North_Strong_Free.pdf','Yes'),
    ('Russia','Russia_2020_Arctic_State_Policy_2035.pdf','Yes'),
    ('Russia','Russia_2020_Arctic_Zone_National_Security_2035.pdf','Yes'),
    ('Norway','Norway_2025_Norway_in_the_North.pdf','Yes'),
    ('Norway','Norway_2020_Arctic_Policy_Abstract.pdf','No'),
    ('Denmark_Greenland','Denmark_2011_Kingdom_Arctic_Strategy_2011_2020.pdf','Yes'),
    ('Denmark_Greenland','Greenland_2024_Foreign_Security_Defence_Policy.pdf','Yes'),
    ('Finland','Finland_2021_Strategy_for_Arctic_Policy.pdf','Yes'),
    ('Finland','Finland_2024_Defence_Report.pdf','Yes'),
    ('Sweden','Sweden_2020_Strategy_for_the_Arctic_Region.pdf','Yes'),
    ('Sweden','Sweden_2024_National_Security_Strategy.pdf','Yes'),
    ('Iceland','Iceland_2021_Arctic_Policy.pdf','Yes'),
    ('Iceland','Iceland_2023_National_Security_Policy.pdf','Yes'),
    ('Iceland','Iceland_2026_EU_Security_Defence_Partnership.pdf','No'),
]

def count_words(text: str) -> int:
    return len(re.findall(r"\b\w+\b", text))

def extract_one(country: str, fname: str, main: str) -> dict:
    path = SOURCE_DIR / country / fname
    row = {
        'country': country,
        'file_name': fname,
        'use_in_main_analysis': main,
        'file_found': 'Yes' if path.exists() else 'No',
        'page_count': None,
        'text_chars': 0,
        'word_count': 0,
        'avg_words_per_page': None,
        'extraction_status': '',
        'quality_flag': '',
        'output_text_file': '',
        'notes': ''
    }
    if not path.exists():
        row['extraction_status'] = 'Missing'
        row['quality_flag'] = 'MISSING'
        return row
    try:
        doc = fitz.open(str(path))
        pages = []
        for i, page in enumerate(doc, start=1):
            text = page.get_text('text') or ''
            pages.append(f"\n\n--- PAGE {i} ---\n{text}")
        full_text = ''.join(pages).strip()
        row['page_count'] = len(doc)
        row['text_chars'] = len(full_text)
        row['word_count'] = count_words(full_text)
        row['avg_words_per_page'] = round(row['word_count'] / len(doc), 1) if len(doc) else None
        if row['word_count'] < 500:
            row['quality_flag'] = 'LOW_TEXT_CHECK_MANUALLY'
        elif row['avg_words_per_page'] < 80:
            row['quality_flag'] = 'POSSIBLE_LOW_EXTRACTION'
        else:
            row['quality_flag'] = 'OK'
        row['extraction_status'] = 'Extracted'
        out_dir = OUTPUT_ROOT / country
        out_dir.mkdir(parents=True, exist_ok=True)
        out_path = out_dir / fname.replace('.pdf', '.txt')
        out_path.write_text(full_text, encoding='utf-8')
        row['output_text_file'] = str(out_path)
    except Exception as exc:
        row['extraction_status'] = 'Error'
        row['quality_flag'] = 'ERROR'
        row['notes'] = repr(exc)
    return row

def main():
    OUTPUT_ROOT.mkdir(exist_ok=True)
    rows = [extract_one(country, fname, main) for country, fname, main in EXPECTED]
    fieldnames = list(rows[0].keys())
    with open('document_text_quality_report.csv', 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)
    summary = {
        'expected_documents': len(EXPECTED),
        'main_documents': sum(1 for _, _, main in EXPECTED if main == 'Yes'),
        'supplementary_documents': sum(1 for _, _, main in EXPECTED if main == 'No'),
        'found_expected': sum(1 for r in rows if r['file_found'] == 'Yes'),
        'missing_expected': [r['file_name'] for r in rows if r['file_found'] != 'Yes'],
        'quality_flags': {flag: sum(1 for r in rows if r['quality_flag'] == flag) for flag in sorted(set(r['quality_flag'] for r in rows))},
    }
    Path('extraction_summary.json').write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding='utf-8')
    print(json.dumps(summary, indent=2, ensure_ascii=False))

if __name__ == '__main__':
    main()
