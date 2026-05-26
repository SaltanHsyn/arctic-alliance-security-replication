from artifact_tool import Workbook, SpreadsheetFile
import csv, json, os, statistics
from pathlib import Path

BASE = Path('/mnt/data')
index_csv = BASE/'arctic_dictionary_analysis'/'index_scores_by_document.csv'
category_csv = BASE/'arctic_dictionary_analysis'/'category_scores_long.csv'
manual_flags_csv = BASE/'arctic_dictionary_analysis'/'manual_review_flags.csv'
quality_csv = BASE/'arctic_text_extraction'/'document_text_quality_report.csv'
output_path = BASE/'Arctic_Analysis_Ready_Content_Scores_v2.xlsx'

# Helpers

def read_csv(path):
    with open(path, newline='', encoding='utf-8-sig') as f:
        return list(csv.DictReader(f))

def to_float(x):
    try:
        return float(x)
    except Exception:
        return None

def fmt_val(x):
    if isinstance(x, float):
        return round(x, 3)
    return x

rows = read_csv(index_csv)
cat_rows = read_csv(category_csv)
flags = read_csv(manual_flags_csv)
quality = read_csv(quality_csv)
quality_by_file = {r.get('file_name') or r.get('\ufeffcountry'): r for r in quality}

# Determine document layer and refined validity
def get_layer(r):
    if r['use_in_main_analysis'] != 'Yes':
        return 'Supplementary / Excluded'
    dt = r['document_type'].lower()
    fn = r['file_name']
    if fn in ['Canada_2024_Our_North_Strong_Free.pdf','Finland_2024_Defence_Report.pdf','Sweden_2024_National_Security_Strategy.pdf','Iceland_2023_National_Security_Policy.pdf']:
        return 'Security / Alliance Supplement'
    if 'defence arctic' in dt:
        return 'Arctic Defence Strategy'
    if 'arctic' in dt or 'high north' in dt or 'greenland' in fn.lower():
        return 'Arctic Strategy / Policy Main'
    return 'Security / Alliance Supplement'

def score_or_blank(r, col):
    return to_float(r[col])

refined_rows = []
for r in rows:
    layer = get_layer(r)
    include_core = 'Yes' if layer in ['Arctic Strategy / Policy Main','Arctic Defence Strategy'] else 'No'
    include_security = 'Yes' if layer in ['Arctic Defence Strategy','Security / Alliance Supplement'] else ('No' if layer == 'Supplementary / Excluded' else 'Partial')
    scfi_raw = score_or_blank(r, 'SCFI_Strategic_Competition_Framing')
    cgfi_raw = score_or_blank(r, 'CGFI_Cooperative_Governance_Framing')
    nafi_raw = score_or_blank(r, 'NAFI_NATO_Alliance_Framing')
    rtfi_raw = score_or_blank(r, 'RTFI_Russia_Threat_Framing')
    csli_raw = score_or_blank(r, 'CSLI_Climate_Security_Linkage')
    res_raw = score_or_blank(r, 'Resilience_Total_Defence')
    ci_raw = score_or_blank(r, 'Critical_Infrastructure_Military_Mobility')
    # Adjusted scores: keep SCFI/CGFI only for Arctic-specific layer; keep security scores for supplement layer
    adj_scfi = scfi_raw if include_core == 'Yes' else None
    adj_cgfi = cgfi_raw if include_core == 'Yes' else None
    adj_csli = csli_raw if include_core == 'Yes' else None
    adj_nafi = nafi_raw if layer != 'Supplementary / Excluded' else None
    adj_rtfi = rtfi_raw if layer != 'Supplementary / Excluded' else None
    adj_res = res_raw if layer != 'Supplementary / Excluded' else None
    adj_ci = ci_raw if layer != 'Supplementary / Excluded' else None
    if layer == 'Supplementary / Excluded':
        status = 'Exclude from main models'
    elif layer == 'Security / Alliance Supplement':
        status = 'Use only for NATO/Russia/security supplement variables'
    elif layer == 'Arctic Defence Strategy':
        status = 'Use in both Arctic strategy and security/alliance layers'
    else:
        status = 'Use in Arctic strategy framing layer'
    refined_rows.append({
        'country': r['country'],
        'file_name': r['file_name'],
        'document_title': r['document_title'],
        'year': int(float(r['year'])) if r['year'].replace('.','',1).isdigit() else r['year'],
        'document_type': r['document_type'],
        'war_period': r['war_period'],
        'nato_period': r['nato_period'],
        'document_layer': layer,
        'model_use_status': status,
        'word_count': int(float(r['word_count'])) if r['word_count'] else '',
        'SCFI_raw': scfi_raw,
        'SCFI_adjusted': adj_scfi,
        'NAFI_raw': nafi_raw,
        'NAFI_adjusted': adj_nafi,
        'RussiaThreat_raw': rtfi_raw,
        'RussiaThreat_adjusted': adj_rtfi,
        'CSLI_raw': csli_raw,
        'CSLI_adjusted': adj_csli,
        'CGFI_raw': cgfi_raw,
        'CGFI_adjusted': adj_cgfi,
        'Resilience_raw': res_raw,
        'Resilience_adjusted': adj_res,
        'CriticalInfra_raw': ci_raw,
        'CriticalInfra_adjusted': adj_ci,
    })

# Country-period summary from adjusted scores
summary_map = {}
for r in refined_rows:
    if r['document_layer'] == 'Supplementary / Excluded':
        continue
    key = (r['country'], r['war_period'], r['nato_period'])
    d = summary_map.setdefault(key, {k: [] for k in ['SCFI_adjusted','NAFI_adjusted','RussiaThreat_adjusted','CSLI_adjusted','CGFI_adjusted','Resilience_adjusted','CriticalInfra_adjusted']})
    for k in d:
        if r[k] is not None:
            d[k].append(r[k])

country_period_rows = []
for (country, war_period, nato_period), vals in sorted(summary_map.items()):
    row = [country, war_period, nato_period]
    for k in ['SCFI_adjusted','NAFI_adjusted','RussiaThreat_adjusted','CSLI_adjusted','CGFI_adjusted','Resilience_adjusted','CriticalInfra_adjusted']:
        row.append(round(sum(vals[k])/len(vals[k]),3) if vals[k] else '')
    country_period_rows.append(row)

# Country-level pre/post comparison: compute deltas where possible for selected security metrics
country_metrics = {}
for r in refined_rows:
    if r['document_layer'] == 'Supplementary / Excluded':
        continue
    c = country_metrics.setdefault(r['country'], {'Pre-war': [], 'Post-war': []})
    period = 'Post-war' if str(r['war_period']).startswith('Post') else 'Pre-war'
    c[period].append(r)

delta_rows = []
for c, periods in sorted(country_metrics.items()):
    def avg(rows, k):
        vals=[x[k] for x in rows if x[k] is not None]
        return round(sum(vals)/len(vals),3) if vals else ''
    pre_nafi=avg(periods['Pre-war'],'NAFI_adjusted')
    post_nafi=avg(periods['Post-war'],'NAFI_adjusted')
    pre_rt=avg(periods['Pre-war'],'RussiaThreat_adjusted')
    post_rt=avg(periods['Post-war'],'RussiaThreat_adjusted')
    pre_scfi=avg(periods['Pre-war'],'SCFI_adjusted')
    post_scfi=avg(periods['Post-war'],'SCFI_adjusted')
    def delta(a,b):
        return round(b-a,3) if isinstance(a,(int,float)) and isinstance(b,(int,float)) else ''
    delta_rows.append([c, pre_scfi, post_scfi, delta(pre_scfi,post_scfi), pre_nafi, post_nafi, delta(pre_nafi,post_nafi), pre_rt, post_rt, delta(pre_rt,post_rt)])

# Manual check rows
manual_rows = []
for r in flags:
    manual_rows.append([r['country'], r['file_name'], r['document_title'], r['year'], r['use_in_main_analysis'], r['word_count'], r['flags'], ''])
# Add Iceland confirmation based on quality report if not present
for r in quality:
    if r.get('file_name') == 'Iceland_2023_National_Security_Policy.pdf':
        manual_rows.append(['Iceland', r['file_name'], 'Text extraction quality confirmation', '2023', 'Yes', r.get('word_count',''), 'Quality report marks extraction_status=Extracted and quality_flag=OK; two-page document, short length is expected.', 'Confirmed usable for security/alliance supplement'])

# Build workbook
wb = Workbook.create()

header_fill = '#1F4E78'
sub_fill = '#D9EAF7'
warn_fill = '#FFF2CC'
light_green = '#E2F0D9'
light_red = '#F4CCCC'

# README
sh = wb.worksheets.add('README')
readme = [
    ['Arctic Q1 Project – Analysis-ready content scores v2'],
    ['Purpose', 'This workbook converts dictionary-based content analysis into a cleaner analysis-ready structure by separating Arctic-strategy framing from general security/alliance supplement documents.'],
    ['Core methodological decision', 'General defence/national-security documents are not treated as full Arctic strategies. They are used only for NATO, Russia-threat, resilience, deterrence, and alliance-security variables.'],
    ['Main Arctic framing layer', 'Arctic-specific strategies and High North/Arctic policy documents. SCFI, CGFI, and CSLI are valid here.'],
    ['Security/alliance supplement layer', 'Post-2022 or post-NATO defence/security documents. NAFI, RussiaThreat, Resilience, and CriticalInfrastructure are valid here.'],
    ['Excluded layer', 'Supplementary or outside-baseline documents. They are preserved for discussion but excluded from main modeling.'],
    ['Next step', 'Use Refined_Content_Scores_v2 for document-level interpretation and Country_Period_Panel_v2 for the first panel-ready coding table.'],
]
sh.get_range('A1:B7').values = readme
sh.get_range('A1:B1').merge()
sh.get_range('A1').format = {'font': {'bold': True, 'size': 15, 'color': '#FFFFFF'}, 'fill': header_fill, 'horizontal_alignment': 'center'}
sh.get_range('A2:A7').format = {'font': {'bold': True}, 'fill': sub_fill}
sh.get_range('A1:B7').format.wrap_text = True
sh.get_range('A:B').format.column_width = 35
sh.get_range('B:B').format.column_width = 90

# Decisions
sh = wb.worksheets.add('Method_Decisions')
decisions = [
    ['Decision', 'Operational rule', 'Why it matters'],
    ['Arctic-specific vs general security documents', 'Only Arctic/High North documents receive full SCFI/CGFI/CSLI treatment.', 'Prevents overestimating Arctic framing in general defence documents.'],
    ['Post-2022 war breakpoint', 'Documents dated 2022 onward are read as post-war unless otherwise noted.', 'Captures Russia-Ukraine war as a structural security break.'],
    ['Finland and Sweden NATO breakpoints', 'Finland 2023 onward and Sweden 2024 onward are post-NATO.', 'Captures alliance accession effects.'],
    ['Norway 2025 document', 'Keep as post-war High North strategy; not used for 2000–2024 strict baseline unless the period is extended.', 'Substantively strong but date exceeds baseline.'],
    ['Iceland 2023 document', 'Use for security/alliance supplement only.', 'It is a short national security policy, not an Arctic strategy.'],
    ['Iceland 2026 document', 'Exclude from main analysis; keep only for extended discussion.', 'Outside baseline period.'],
]
sh.get_range_by_indexes(0,0,len(decisions),len(decisions[0])).values = decisions
sh.get_range('A1:C1').format = {'fill': header_fill, 'font': {'bold': True, 'color': '#FFFFFF'}, 'horizontal_alignment': 'center'}
sh.get_range('A:C').format.wrap_text = True
sh.get_range('A:A').format.column_width = 34
sh.get_range('B:B').format.column_width = 55
sh.get_range('C:C').format.column_width = 62

# Refined content scores
sh = wb.worksheets.add('Refined_Content_Scores_v2')
ref_headers = list(refined_rows[0].keys())
ref_values = [ref_headers] + [[fmt_val(r[h]) if r[h] is not None else '' for h in ref_headers] for r in refined_rows]
sh.get_range_by_indexes(0,0,len(ref_values),len(ref_headers)).values = ref_values
sh.get_range_by_indexes(0,0,1,len(ref_headers)).format = {'fill': header_fill, 'font': {'bold': True, 'color': '#FFFFFF'}, 'horizontal_alignment': 'center'}
sh.freeze_panes.freeze_rows(1)
sh.get_range_by_indexes(0,0,len(ref_values),len(ref_headers)).format.wrap_text = True
# widths
for col, width in [('A:A',18),('B:B',42),('C:C',55),('D:D',10),('E:E',24),('F:F',22),('G:G',24),('H:H',28),('I:I',60)]:
    sh.get_range(col).format.column_width = width
for col_letter in list('JKLMNOPQRSTUVWX'):
    try:
        sh.get_range(f'{col_letter}:{col_letter}').format.column_width = 14
    except Exception:
        pass
sh.tables.add(sh.get_range_by_indexes(0,0,len(ref_values),len(ref_headers)), True, 'RefinedScoresTable')

# Country period panel
sh = wb.worksheets.add('Country_Period_Panel_v2')
panel_headers = ['country','war_period','nato_period','SCFI_adj_avg','NAFI_adj_avg','RussiaThreat_adj_avg','CSLI_adj_avg','CGFI_adj_avg','Resilience_adj_avg','CriticalInfra_adj_avg']
panel_values = [panel_headers] + country_period_rows
sh.get_range_by_indexes(0,0,len(panel_values),len(panel_headers)).values = panel_values
sh.get_range_by_indexes(0,0,1,len(panel_headers)).format = {'fill': header_fill, 'font': {'bold': True, 'color': '#FFFFFF'}, 'horizontal_alignment': 'center'}
sh.freeze_panes.freeze_rows(1)
sh.get_range('A:J').format.wrap_text = True
sh.get_range('A:C').format.column_width = 22
sh.get_range('D:J').format.column_width = 16
sh.tables.add(sh.get_range_by_indexes(0,0,len(panel_values),len(panel_headers)), True, 'CountryPeriodPanelTable')

# Pre post delta
sh = wb.worksheets.add('Pre_Post_Delta_Check')
delta_headers = ['country','Pre_SCFI','Post_SCFI','Delta_SCFI','Pre_NAFI','Post_NAFI','Delta_NAFI','Pre_RussiaThreat','Post_RussiaThreat','Delta_RussiaThreat']
delta_values = [delta_headers] + delta_rows
sh.get_range_by_indexes(0,0,len(delta_values),len(delta_headers)).values = delta_values
sh.get_range_by_indexes(0,0,1,len(delta_headers)).format = {'fill': header_fill, 'font': {'bold': True, 'color': '#FFFFFF'}, 'horizontal_alignment': 'center'}
sh.freeze_panes.freeze_rows(1)
sh.get_range('A:J').format.column_width = 15
sh.get_range('A:J').format.wrap_text = True
sh.tables.add(sh.get_range_by_indexes(0,0,len(delta_values),len(delta_headers)), True, 'DeltaCheckTable')

# Manual check
sh = wb.worksheets.add('Manual_Check_Log')
manual_headers = ['country','file_name','document_title','year','use_in_main_analysis','word_count','flag_or_issue','manual_decision']
manual_values = [manual_headers] + manual_rows
sh.get_range_by_indexes(0,0,len(manual_values),len(manual_headers)).values = manual_values
sh.get_range_by_indexes(0,0,1,len(manual_headers)).format = {'fill': header_fill, 'font': {'bold': True, 'color': '#FFFFFF'}, 'horizontal_alignment': 'center'}
sh.freeze_panes.freeze_rows(1)
sh.get_range('A:H').format.wrap_text = True
sh.get_range('A:A').format.column_width = 18
sh.get_range('B:B').format.column_width = 42
sh.get_range('C:C').format.column_width = 45
sh.get_range('G:H').format.column_width = 60
sh.tables.add(sh.get_range_by_indexes(0,0,len(manual_values),len(manual_headers)), True, 'ManualCheckTable')

# Category scores long, include adjusted layer info by merging map
layer_by_file = {r['file_name']: r['document_layer'] for r in refined_rows}
cat_headers = list(cat_rows[0].keys()) + ['document_layer']
cat_values = [cat_headers]
for r in cat_rows:
    cat_values.append([r.get(h,'') for h in cat_headers[:-1]] + [layer_by_file.get(r['file_name'],'')])
sh = wb.worksheets.add('Category_Scores_Long')
sh.get_range_by_indexes(0,0,len(cat_values),len(cat_headers)).values = cat_values
sh.get_range_by_indexes(0,0,1,len(cat_headers)).format = {'fill': header_fill, 'font': {'bold': True, 'color': '#FFFFFF'}, 'horizontal_alignment': 'center'}
sh.freeze_panes.freeze_rows(1)
sh.get_range_by_indexes(0,0,len(cat_values),len(cat_headers)).format.wrap_text = True
for col in ['A:A','B:B','C:C','G:G','H:H','K:K','R:R']:
    try: sh.get_range(col).format.column_width = 24
    except Exception: pass
sh.tables.add(sh.get_range_by_indexes(0,0,len(cat_values),len(cat_headers)), True, 'CategoryScoresTable')

# Next steps
sh = wb.worksheets.add('Next_Steps')
steps = [
    ['Step', 'Action', 'Output'],
    [1, 'Accept this v2 layer structure unless there is a theoretical objection.', 'Final document-layer map.'],
    [2, 'Run one manual evidence check for high-leverage categories: NATO, Russia threat, military/defence, cooperation.', 'Manual-score overrides where needed.'],
    [3, 'Create article-ready descriptive tables: document list, index summary, pre/post comparison.', 'Tables for Methods/Data section.'],
    [4, 'Prepare econometric panel skeleton by country-year using SCFI/NAFI/RTFI variables.', 'Arctic_Panel_Content_Scores.csv/xlsx.'],
    [5, 'Merge with external quantitative variables: military burden, sea ice decline, energy prices, NATO status, controls.', 'Full panel dataset for Python econometrics.'],
]
sh.get_range_by_indexes(0,0,len(steps),3).values = steps
sh.get_range('A1:C1').format = {'fill': header_fill, 'font': {'bold': True, 'color': '#FFFFFF'}, 'horizontal_alignment': 'center'}
sh.get_range('A:C').format.wrap_text = True
sh.get_range('A:A').format.column_width = 8
sh.get_range('B:B').format.column_width = 70
sh.get_range('C:C').format.column_width = 45

# Apply colors/conditional formatting based on model_use_status maybe simple cell fill values by rows? Use manual row styling.
# For simplified artifact_tool, leave table style sufficient.

# Verify sheet list and export
SpreadsheetFile.export_xlsx(wb).save(str(output_path))
print(output_path)
