import os, re, csv, json, zipfile, math, statistics
from pathlib import Path
from collections import defaultdict, Counter

BASE = Path('/mnt/data/arctic_text_extraction/extracted_texts')
OUT = Path('/mnt/data/arctic_dictionary_analysis')
OUT.mkdir(exist_ok=True)

# Master metadata (aligned with Arctic_Q1_Strategic_Documents_Master_FINAL_checked.xlsx)
docs = [
    ('USA','USA_2022_National_Arctic_Strategy.txt','USA_2022_National_Arctic_Strategy.pdf','National Strategy for the Arctic Region',2022,'Arctic Strategy','Yes','Post-war? no, pre-full-war baseline','Long-standing NATO member'),
    ('USA','USA_2024_DoD_Arctic_Strategy.txt','USA_2024_DoD_Arctic_Strategy.pdf','Department of Defense Arctic Strategy',2024,'Defence Arctic Strategy','Yes','Post-war','Long-standing NATO member'),
    ('Canada','Canada_2019_Arctic_Northern_Policy_Framework.txt','Canada_2019_Arctic_Northern_Policy_Framework.pdf','Arctic and Northern Policy Framework',2019,'Arctic Policy Framework','Yes','Pre-war','Long-standing NATO member'),
    ('Canada','Canada_2024_Our_North_Strong_Free.txt','Canada_2024_Our_North_Strong_Free.pdf','Our North, Strong and Free',2024,'Defence Policy','Yes','Post-war','Long-standing NATO member'),
    ('Russia','Russia_2020_Arctic_State_Policy_2035.txt','Russia_2020_Arctic_State_Policy_2035.pdf','Foundations of the Russian Federation State Policy in the Arctic to 2035',2020,'Arctic State Policy','Yes','Pre-war document / amended post-war','Non-NATO / Russia'),
    ('Russia','Russia_2020_Arctic_Zone_National_Security_2035.txt','Russia_2020_Arctic_Zone_National_Security_2035.pdf','Arctic Zone Development and National Security Strategy to 2035',2020,'Arctic Zone and National Security Strategy','Yes','Pre-war document / amended post-war','Non-NATO / Russia'),
    ('Norway','Norway_2025_Norway_in_the_North.txt','Norway_2025_Norway_in_the_North.pdf','Norway in the North: A High North Policy for a New Era',2025,'High North Strategy','Yes','Post-war','Long-standing NATO member'),
    ('Norway','Norway_2020_Arctic_Policy_Abstract.txt','Norway_2020_Arctic_Policy_Abstract.pdf','The Norwegian Government’s Arctic Policy Abstract',2020,'Arctic Policy Abstract','No','Pre-war','Long-standing NATO member'),
    ('Denmark_Greenland','Denmark_2011_Kingdom_Arctic_Strategy_2011_2020.txt','Denmark_2011_Kingdom_Arctic_Strategy_2011_2020.pdf','Kingdom of Denmark Strategy for the Arctic 2011–2020',2011,'Arctic Strategy','Yes','Pre-war','Long-standing NATO member'),
    ('Denmark_Greenland','Greenland_2024_Foreign_Security_Defence_Policy.txt','Greenland_2024_Foreign_Security_Defence_Policy.pdf','Greenland Foreign, Security and Defence Policy 2024–2033',2024,'Foreign, Security and Defence Policy','Yes','Post-war','Long-standing NATO member'),
    ('Finland','Finland_2021_Strategy_for_Arctic_Policy.txt','Finland_2021_Strategy_for_Arctic_Policy.pdf','Finland’s Strategy for Arctic Policy',2021,'Arctic Policy Strategy','Yes','Pre-war','Pre-NATO'),
    ('Finland','Finland_2024_Defence_Report.txt','Finland_2024_Defence_Report.pdf','Government Defence Report',2024,'Defence Report','Yes','Post-war','Post-NATO'),
    ('Sweden','Sweden_2020_Strategy_for_the_Arctic_Region.txt','Sweden_2020_Strategy_for_the_Arctic_Region.pdf','Sweden’s Strategy for the Arctic Region',2020,'Arctic Strategy','Yes','Pre-war','Pre-NATO'),
    ('Sweden','Sweden_2024_National_Security_Strategy.txt','Sweden_2024_National_Security_Strategy.pdf','National Security Strategy',2024,'National Security Strategy','Yes','Post-war','Post-NATO'),
    ('Iceland','Iceland_2021_Arctic_Policy.txt','Iceland_2021_Arctic_Policy.pdf','Iceland’s Policy on Matters Concerning the Arctic Region',2021,'Arctic Policy','Yes','Pre-war','Long-standing NATO member'),
    ('Iceland','Iceland_2023_National_Security_Policy.txt','Iceland_2023_National_Security_Policy.pdf','Parliamentary Resolution on a National Security Policy for Iceland',2023,'National Security Policy','Yes','Post-war','Long-standing NATO member'),
    ('Iceland','Iceland_2026_EU_Security_Defence_Partnership.txt','Iceland_2026_EU_Security_Defence_Partnership.pdf','EU-Iceland Security and Defence Partnership',2026,'Security and Defence Partnership','No','Post-war / outside baseline','Long-standing NATO member'),
]

# Dictionary terms: transparent, theory-derived first-pass list. Lowercase matching.
dictionaries = {
    'Security_Framing': ['security', 'secure', 'stability', 'threat', 'risk', 'strategic', 'national security', 'security policy', 'security interests', 'hostile', 'conflict', 'crisis', 'safe', 'safety'],
    'Military_Defence_Framing': ['military', 'defence', 'defense', 'armed forces', 'force posture', 'capability', 'capabilities', 'deterrence', 'defence capability', 'defense capability', 'exercise', 'exercises', 'training', 'operations', 'surveillance', 'submarine', 'air force', 'navy', 'coast guard'],
    'Sovereignty_Framing': ['sovereignty', 'sovereign rights', 'jurisdiction', 'territorial', 'territory', 'exclusive economic zone', 'continental shelf', 'maritime delimitation', 'border', 'borders', 'law of the sea', 'uncLOS', 'national interests', 'national interest'],
    'Energy_Resource_Access_Framing': ['energy', 'resources', 'natural resources', 'oil', 'gas', 'petroleum', 'minerals', 'critical raw materials', 'hydrocarbon', 'resource management', 'electricity', 'energy security', 'security of supply', 'supply'],
    'Shipping_Route_Framing': ['shipping', 'navigation', 'sea route', 'northern sea route', 'northeast passage', 'northwest passage', 'maritime traffic', 'transport', 'ports', 'sea lanes', 'freedom of navigation', 'mobility corridor'],
    'Great_Power_Rivalry_Framing': ['great power', 'strategic competition', 'competition', 'rivalry', 'china', 'russia', 'united states', 'us', 'usa', 'nato', 'western countries', 'non-arctic states', 'global competition'],
    'NATO_Alliance_Framing': ['nato', 'alliance', 'allied', 'allies', 'collective defence', 'collective defense', 'article 5', 'deterrence', 'interoperability', 'reinforcements', 'allied presence', 'euro-atlantic', 'transatlantic'],
    'Russia_Threat_Framing': ['russia', 'russian', 'ukraine', 'aggression', 'annexation', 'crimea', 'northern fleet', 'military build-up', 'military buildup', 'bastion', 'threat from russia', 'russian activity', 'war of aggression'],
    'Climate_Change_Framing': ['climate', 'climate change', 'warming', 'global warming', 'sea ice', 'ice cover', 'melting', 'permafrost', 'emissions', 'low-emission', 'paris agreement', 'adaptation', 'environment', 'biodiversity'],
    'Cooperation_Governance_Framing': ['cooperation', 'collaboration', 'governance', 'arctic council', 'international law', 'multilateral', 'dialogue', 'diplomacy', 'rules-based', 'partnership', 'partners', 'barents', 'un convention', 'UNCLOS'],
    'Indigenous_Human_Security_Framing': ['indigenous', 'sami', 'sámi', 'inuit', 'local communities', 'communities', 'livelihoods', 'culture', 'youth', 'living conditions', 'human security', 'welfare', 'people'],
    'Resilience_Total_Defence_Framing': ['resilience', 'preparedness', 'total defence', 'total defense', 'civilian preparedness', 'civil protection', 'emergency preparedness', 'crisis preparedness', 'security of supply', 'critical functions', 'societal functions'],
    'Critical_Infrastructure_Military_Mobility': ['critical infrastructure', 'infrastructure', 'military mobility', 'mobility', 'transport infrastructure', 'communications', 'digital infrastructure', 'surveillance', 'space infrastructure', 'undersea', 'ports', 'rail', 'corridor'],
}

# Compile regex patterns. Use loose phrase matching with optional hyphens/space for multiwords.
def term_pattern(term):
    term = term.lower()
    esc = re.escape(term)
    esc = esc.replace('\\ ', r'[\s\-]+')
    return re.compile(r'(?<![a-z])' + esc + r'(?![a-z])', re.IGNORECASE)
patterns = {cat: [(term, term_pattern(term)) for term in terms] for cat, terms in dictionaries.items()}

word_re = re.compile(r"[A-Za-zÀ-ÖØ-öø-ÿ][A-Za-zÀ-ÖØ-öø-ÿ'\-]*")

raw_rows = []
term_rows = []
for country, txtfile, pdffile, title, year, dtype, main, war, nato in docs:
    path = BASE / country / txtfile
    if not path.exists():
        # handle potential folder names in zip, but should not happen
        text = ''
        exists = False
    else:
        text = path.read_text(encoding='utf-8', errors='ignore')
        exists = True
    lower = text.lower()
    words = word_re.findall(text)
    word_count = len(words)
    result_base = {
        'country': country,
        'file_name': pdffile,
        'document_title': title,
        'year': year,
        'document_type': dtype,
        'use_in_main_analysis': main,
        'war_period': war,
        'nato_period': nato,
        'text_file_exists': 'Yes' if exists else 'No',
        'word_count': word_count,
    }
    for cat, pats in patterns.items():
        raw = 0
        term_counts = []
        for term, pat in pats:
            c = len(pat.findall(lower))
            if c:
                term_counts.append((term, c))
            raw += c
        norm = raw / word_count * 10000 if word_count else 0
        raw_rows.append({**result_base, 'category': cat, 'raw_hits': raw, 'hits_per_10k_words': round(norm, 3)})
        for term, c in sorted(term_counts, key=lambda x: (-x[1], x[0]))[:15]:
            term_rows.append({**result_base, 'category': cat, 'term': term, 'term_hits': c, 'term_hits_per_10k_words': round(c/word_count*10000, 3) if word_count else 0})

# Score 0-3 based on category-specific positive tertiles among main docs only.
positive_norms = defaultdict(list)
for row in raw_rows:
    if row['use_in_main_analysis'] == 'Yes' and row['hits_per_10k_words'] > 0:
        positive_norms[row['category']].append(row['hits_per_10k_words'])
thresholds = {}
for cat, vals in positive_norms.items():
    vals = sorted(vals)
    if not vals:
        thresholds[cat] = (0,0)
    else:
        # Use 33rd and 67th percentile positions
        def pct(p):
            if len(vals) == 1: return vals[0]
            k = (len(vals)-1) * p
            f = math.floor(k); c = math.ceil(k)
            if f == c: return vals[int(k)]
            return vals[f]*(c-k) + vals[c]*(k-f)
        thresholds[cat] = (pct(0.33), pct(0.67))

for row in raw_rows:
    norm = row['hits_per_10k_words']; cat = row['category']
    q1,q2 = thresholds.get(cat,(0,0))
    if norm <= 0:
        score = 0
    elif norm <= q1:
        score = 1
    elif norm <= q2:
        score = 2
    else:
        score = 3
    row['auto_score_0_3'] = score
    row['scoring_rule'] = '0 if no hits; otherwise category-specific tertiles among main-analysis documents.'

# Convert to wide scores by document
by_doc = {}
for row in raw_rows:
    key = row['file_name']
    if key not in by_doc:
        by_doc[key] = {k: row[k] for k in ['country','file_name','document_title','year','document_type','use_in_main_analysis','war_period','nato_period','word_count']}
    by_doc[key][row['category'] + '_hits10k'] = row['hits_per_10k_words']
    by_doc[key][row['category'] + '_score'] = row['auto_score_0_3']

def mean_scores(doc, cats):
    vals = [doc.get(c+'_score',0) for c in cats]
    return round(sum(vals)/len(vals), 3) if vals else 0

index_rows = []
SCFI_CATS = ['Security_Framing','Military_Defence_Framing','Sovereignty_Framing','Energy_Resource_Access_Framing','Great_Power_Rivalry_Framing']
CGFI_CATS = ['Cooperation_Governance_Framing','Indigenous_Human_Security_Framing','Climate_Change_Framing']
for fn, doc in by_doc.items():
    scfi = mean_scores(doc, SCFI_CATS)
    nafi = doc.get('NATO_Alliance_Framing_score',0)
    rtf = doc.get('Russia_Threat_Framing_score',0)
    climate = doc.get('Climate_Change_Framing_score',0)
    security = doc.get('Security_Framing_score',0)
    csli = round((climate * security) / 3, 3)  # 0-3 scale product linkage
    cgfi = mean_scores(doc, CGFI_CATS)
    resi = doc.get('Resilience_Total_Defence_Framing_score',0)
    infra = doc.get('Critical_Infrastructure_Military_Mobility_score',0)
    index_rows.append({**{k: doc[k] for k in ['country','file_name','document_title','year','document_type','use_in_main_analysis','war_period','nato_period','word_count']},
                       'SCFI_Strategic_Competition_Framing': scfi,
                       'NAFI_NATO_Alliance_Framing': nafi,
                       'RTFI_Russia_Threat_Framing': rtf,
                       'CSLI_Climate_Security_Linkage': csli,
                       'CGFI_Cooperative_Governance_Framing': cgfi,
                       'Resilience_Total_Defence': resi,
                       'Critical_Infrastructure_Military_Mobility': infra})

# Country-period summary for main docs only
summary_groups = defaultdict(list)
for row in index_rows:
    if row['use_in_main_analysis'] == 'Yes':
        summary_groups[(row['country'], row['war_period'], row['nato_period'])].append(row)
summary_rows = []
for (country, war, nato), rows in sorted(summary_groups.items()):
    def avg(field): return round(sum(r[field] for r in rows)/len(rows), 3)
    summary_rows.append({
        'country': country, 'war_period': war, 'nato_period': nato, 'n_documents': len(rows),
        'avg_SCFI': avg('SCFI_Strategic_Competition_Framing'),
        'avg_NAFI': avg('NAFI_NATO_Alliance_Framing'),
        'avg_Russia_Threat': avg('RTFI_Russia_Threat_Framing'),
        'avg_CSLI': avg('CSLI_Climate_Security_Linkage'),
        'avg_CGFI': avg('CGFI_Cooperative_Governance_Framing'),
        'avg_Resilience': avg('Resilience_Total_Defence'),
        'avg_Critical_Infrastructure': avg('Critical_Infrastructure_Military_Mobility')
    })

# Manual review flags
flag_rows = []
for row in index_rows:
    flags=[]
    if row['word_count'] < 1000: flags.append('Very short extracted text; manual PDF check recommended.')
    if row['use_in_main_analysis']=='No': flags.append('Supplementary document; exclude from main index unless explicitly needed.')
    if row['year'] > 2024: flags.append('Outside baseline 2000-2024 period.')
    if 'translation' in row['document_type'].lower(): flags.append('Translation document; compare with original if possible.')
    if flags:
        flag_rows.append({k: row[k] for k in ['country','file_name','document_title','year','use_in_main_analysis','word_count']} | {'flags':' | '.join(flags)})

# Write CSVs
files_to_write = {
    'category_scores_long.csv': raw_rows,
    'index_scores_by_document.csv': index_rows,
    'country_period_summary.csv': summary_rows,
    'dictionary_terms.csv': [{'category':cat,'term':term} for cat, terms in dictionaries.items() for term in terms],
    'top_terms_by_document.csv': term_rows,
    'manual_review_flags.csv': flag_rows,
}
for fname, rows in files_to_write.items():
    path = OUT/fname
    if rows:
        headers = list(rows[0].keys())
    else:
        headers = []
    with path.open('w', newline='', encoding='utf-8') as f:
        if headers:
            writer = csv.DictWriter(f, fieldnames=headers)
            writer.writeheader(); writer.writerows(rows)

# JSON metadata
(OUT/'dictionary_analysis_summary.json').write_text(json.dumps({
    'n_documents_total': len(docs),
    'n_main_documents': sum(1 for d in docs if d[6]=='Yes'),
    'n_categories': len(dictionaries),
    'scoring': '0=no hits; 1/2/3=category-specific tertiles of hits per 10k words among main-analysis documents.',
    'indices': {
        'SCFI': SCFI_CATS,
        'NAFI': ['NATO_Alliance_Framing'],
        'RTFI': ['Russia_Threat_Framing'],
        'CSLI': 'Climate_Change_Framing_score * Security_Framing_score / 3',
        'CGFI': CGFI_CATS,
    }
}, indent=2), encoding='utf-8')

print('Created CSVs in', OUT)
print('Index rows', len(index_rows), 'Category rows', len(raw_rows), 'Term rows', len(term_rows))
