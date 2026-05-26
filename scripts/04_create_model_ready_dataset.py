import math, os

from pathlib import Path
ROOT=Path(__file__).resolve().parents[1]
input_path=ROOT/'data'/'Arctic_Panel_With_Energy_Prices_v4.xlsx'
output_path=ROOT/'data'/'Arctic_Model_Ready_Dataset_v1.xlsx'
wb=SpreadsheetFile.import_xlsx(Blob.load(input_path))
source=wb.worksheets.get_item('Panel_2000_2024')
rows=source.get_range('A1:BC201').values
headers=rows[0]
idx={h:i for i,h in enumerate(headers)}

def val(row, col):
    return row[idx[col]] if col in idx and idx[col] < len(row) else None

def num(x):
    try:
        if x is None or x == '':
            return None
        return float(x)
    except Exception:
        return None

def ln(x):
    x=num(x)
    return math.log(x) if x is not None and x>0 else None

def prod(a,b):
    a=num(a); b=num(b)
    return a*b if a is not None and b is not None else None

# compute military expenditure growth by country from constant USD
prev_by_country={}
model_rows=[]
for r in rows[1:]:
    country=val(r,'country')
    year=int(val(r,'year'))
    mil_const=num(val(r,'military_exp_constant_usd'))
    prev=prev_by_country.get(country)
    growth=None
    if prev is not None and prev>0 and mil_const is not None:
        growth=(mil_const/prev - 1)*100
    if mil_const is not None:
        prev_by_country[country]=mil_const
    core_vars=['military_burden_pct_gdp','sea_ice_decline','energy_price_index_2010_100','gdp_per_capita','gdp_growth','inflation','trade_openness','population']
    core_present=all(num(val(r,c)) is not None for c in core_vars)
    scfi=num(val(r,'SCFI_model'))
    nafi=num(val(r,'NAFI_model'))
    rthreat=num(val(r,'RussiaThreat_model'))
    content_main=1 if core_present and scfi is not None else 0
    security_supp=1 if core_present and (nafi is not None or rthreat is not None) else 0
    external=1 if core_present else 0
    sea_decl=num(val(r,'sea_ice_decline'))
    energy_index=num(val(r,'energy_price_index_2010_100'))
    row=[
        country, year,
        external, content_main, security_supp,
        val(r,'arctic_eight'), val(r,'coastal_five'), val(r,'post_2022_war'), val(r,'nato_membership'), val(r,'post_nato_accession'), val(r,'russia_dummy'),
        val(r,'military_burden_pct_gdp'), mil_const, ln(mil_const), growth,
        val(r,'sea_ice_extent_m_km2'), sea_decl, val(r,'sea_ice_annual_mean_m_km2'), val(r,'sea_ice_annual_decline'),
        scfi, nafi, rthreat, val(r,'Resilience_model'), val(r,'CriticalInfra_model'),
        val(r,'brent_oil_price'), val(r,'europe_gas_price'), energy_index,
        val(r,'gdp_per_capita'), ln(val(r,'gdp_per_capita')), val(r,'gdp_growth'), val(r,'inflation'), val(r,'trade_openness'), val(r,'energy_rents_pct_gdp'), val(r,'oil_rents_pct_gdp'), val(r,'gas_rents_pct_gdp'), val(r,'population'), ln(val(r,'population')),
        prod(sea_decl, scfi), prod(scfi, val(r,'post_2022_war')), prod(nafi, val(r,'post_2022_war')), prod(rthreat, val(r,'post_2022_war')), prod(sea_decl, energy_index), prod(sea_decl, nafi),
        val(r,'sipri_data_quality_flag'), val(r,'wdi_data_quality_flag'), val(r,'sea_ice_source_flag'), val(r,'energy_price_source_flag'), val(r,'policy_source_file'), val(r,'security_source_file')
    ]
    model_rows.append(row)

model_headers=[
'country','year','sample_external_controls','sample_content_main','sample_security_supplement',
'arctic_eight','coastal_five','post_2022_war','nato_membership','post_nato_accession','russia_dummy',
'military_burden_pct_gdp','military_exp_constant_usd','ln_military_exp_constant_usd','military_exp_growth_pct',
'sea_ice_extent_m_km2','sea_ice_decline','sea_ice_annual_mean_m_km2','sea_ice_annual_decline',
'SCFI_model','NAFI_model','RussiaThreat_model','Resilience_model','CriticalInfra_model',
'brent_oil_price','europe_gas_price','energy_price_index_2010_100',
'gdp_per_capita','ln_gdp_per_capita','gdp_growth','inflation','trade_openness','energy_rents_pct_gdp','oil_rents_pct_gdp','gas_rents_pct_gdp','population','ln_population',
'sea_ice_decline_x_SCFI','SCFI_x_post2022','NAFI_x_post2022','RussiaThreat_x_post2022','sea_ice_decline_x_energy_price_index','sea_ice_decline_x_NAFI',
'sipri_data_quality_flag','wdi_data_quality_flag','sea_ice_source_flag','energy_price_source_flag','policy_source_file','security_source_file'
]

# Helper to recreate sheet (if exists, clear; otherwise add)
def get_clean_sheet(name):
    try:
        sh=wb.worksheets.get_item(name)
        # Clear a generous range
        sh.get_range('A1:AZ1000').clear({'contentsOnly': False})
        return sh
    except Exception:
        return wb.worksheets.add(name)

# write model dataset
sh=get_clean_sheet('Model_Dataset')
sh.get_range_by_indexes(0,0,1,len(model_headers)).values=[model_headers]
sh.get_range_by_indexes(1,0,len(model_rows),len(model_headers)).values=model_rows
try:
    sh.tables.add(sh.get_range_by_indexes(0,0,len(model_rows)+1,len(model_headers)), True, 'ModelDatasetTable')
except Exception:
    pass
# style
header=sh.get_range_by_indexes(0,0,1,len(model_headers))
header.format.fill='#0F3D5E'
header.format.font={'bold': True, 'color':'#FFFFFF'}
header.format.wrap_text=True
sh.freeze_panes.freeze_rows(1)
# widths
for col in range(len(model_headers)):
    rng=sh.get_range_by_indexes(0,col,len(model_rows)+1,1)
    if col in [0,47,48]:
        rng.format.column_width=24
    elif col in [46,45,44,43]:
        rng.format.column_width=18
    else:
        rng.format.column_width=15
# number formats
# col indexes
num_cols=list(range(11,43))
for c in num_cols:
    sh.get_range_by_indexes(1,c,len(model_rows),1).format.number_format='0.000'
for c in [1,2,3,4,5,6,7,8,9,10]:
    sh.get_range_by_indexes(1,c,len(model_rows),1).format.number_format='0'
# data bars for content and dv maybe
try:
    sh.get_range('L2:L201').conditional_formats.add_data_bar({'color':'#60A5FA','gradient':True})
    sh.get_range('T2:X201').conditional_formats.add_color_scale({'minColor':'#FEE2E2','midColor':'#FEF3C7','maxColor':'#DCFCE7'})
except Exception:
    pass

# sample summary
summary_headers=['metric','value','interpretation']
summary_rows=[]
def count_col(i, value=1):
    return sum(1 for r in model_rows if r[i]==value)
summary_rows += [
    ['Total country-year observations', len(model_rows), 'Arctic Eight, 2000-2024'],
    ['External controls sample', count_col(2), 'Rows with DV, sea ice, energy price and core macro controls'],
    ['Content main sample', count_col(3), 'Rows with external controls and SCFI_model available'],
    ['Security supplement sample', count_col(4), 'Rows with external controls and NAFI or RussiaThreat supplement available'],
    ['Countries', len(set(r[0] for r in model_rows)), 'USA, Canada, Russia, Norway, Denmark, Finland, Sweden, Iceland'],
]
# observations by country with content sample
countries=sorted(set(r[0] for r in model_rows))
for c in countries:
    summary_rows.append([f'{c} content sample obs', sum(1 for r in model_rows if r[0]==c and r[3]==1), 'SCFI_model non-missing within core sample'])
sh2=get_clean_sheet('Model_Readiness_Check')
sh2.get_range('A1:C1').values=[summary_headers]
sh2.get_range_by_indexes(1,0,len(summary_rows),3).values=summary_rows
sh2.get_range('A1:C1').format.fill='#065F46'
sh2.get_range('A1:C1').format.font={'bold': True, 'color':'#FFFFFF'}
sh2.get_range('A:C').format.autofit_columns()
try:
    sh2.tables.add(sh2.get_range_by_indexes(0,0,len(summary_rows)+1,3), True, 'ModelReadinessTable')
except Exception:
    pass

# missingness check
miss_headers=['variable','non_missing','missing','missing_pct','model_role','action']
roles={
'military_burden_pct_gdp':'dependent variable','sea_ice_decline':'main environmental IV','SCFI_model':'content-analysis IV','NAFI_model':'NATO/alliance framing','RussiaThreat_model':'Russia-threat framing','energy_price_index_2010_100':'energy value moderator','gdp_per_capita':'control','gdp_growth':'control','inflation':'control','trade_openness':'control','population':'control','energy_rents_pct_gdp':'energy/resource control','oil_rents_pct_gdp':'resource robustness','gas_rents_pct_gdp':'resource robustness'}
miss_rows=[]
for h in model_headers:
    vals=[r[model_headers.index(h)] for r in model_rows]
    non=sum(1 for v in vals if v not in [None,''])
    miss=len(vals)-non
    role=roles.get(h,'metadata / indicator / interaction')
    if miss==0:
        action='complete'
    elif h in ['SCFI_model','NAFI_model','RussiaThreat_model','Resilience_model','CriticalInfra_model']:
        action='expected: only available for assigned policy periods; use sample flags'
    else:
        action='check before final models or use reduced sample'
    miss_rows.append([h,non,miss,miss/len(vals),role,action])
sh3=get_clean_sheet('Missingness_Check')
sh3.get_range('A1:F1').values=[miss_headers]
sh3.get_range_by_indexes(1,0,len(miss_rows),6).values=miss_rows
sh3.get_range('A1:F1').format.fill='#7C2D12'
sh3.get_range('A1:F1').format.font={'bold': True, 'color':'#FFFFFF'}
sh3.get_range('D2:D100').format.number_format='0.0%'
sh3.get_range('A:F').format.autofit_columns()
try:
    sh3.tables.add(sh3.get_range_by_indexes(0,0,len(miss_rows)+1,6), True, 'MissingnessTable')
except Exception:
    pass
try:
    sh3.get_range('D2:D60').conditional_formats.add_color_scale({'minColor':'#DCFCE7','midColor':'#FEF3C7','maxColor':'#FECACA'})
except Exception:
    pass

# Model specs sheet
spec_headers=['model_id','sample_flag','dependent_variable','key_independent_variables','controls','fixed_effects','estimator_notes','purpose']
spec_rows=[
['M0_external_baseline','sample_external_controls','military_burden_pct_gdp','sea_ice_decline; energy_price_index_2010_100','ln_gdp_per_capita; gdp_growth; inflation; trade_openness; ln_population; energy_rents_pct_gdp','country FE; year FE','Two-way FE; clustered or Driscoll-Kraay SE','Baseline climate-energy-security model before content variables'],
['M1_content_main','sample_content_main','military_burden_pct_gdp','sea_ice_decline; SCFI_model','ln_gdp_per_capita; gdp_growth; inflation; trade_openness; ln_population; energy_price_index_2010_100','country FE; year FE','Two-way FE; content sample only','Tests whether strategic competition framing correlates with military burden'],
['M2_interaction','sample_content_main','military_burden_pct_gdp','sea_ice_decline; SCFI_model; sea_ice_decline_x_SCFI','same as M1','country FE; year FE','Interaction term; marginal effects required','Tests whether sea-ice decline matters more under strategic competition framing'],
['M3_post2022_framing','sample_content_main','military_burden_pct_gdp','SCFI_model; SCFI_x_post2022; post_2022_war','same as M1','country FE; year FE','Post-2022 term absorbed by year FE if full year FE included; use carefully','Tests Russia-Ukraine war break in strategic framing'],
['M4_security_supplement','sample_security_supplement','military_burden_pct_gdp','NAFI_model; RussiaThreat_model; NAFI_x_post2022; RussiaThreat_x_post2022','ln_gdp_per_capita; gdp_growth; inflation; trade_openness; ln_population','country FE; year FE','Supplement layer; not Arctic-specific strategy score','Tests NATO/alliance and Russia-threat security framing'],
['M5_robustness_excluding_Russia','sample_content_main & russia_dummy=0','military_burden_pct_gdp','sea_ice_decline; SCFI_model; sea_ice_decline_x_SCFI','same as M1','country FE; year FE','Non-Russia Arctic states only','Checks whether results are driven by Russia'],
['M6_coastal_five','sample_content_main & coastal_five=1','military_burden_pct_gdp','sea_ice_decline; SCFI_model; energy_price_index_2010_100','same as M1','country FE; year FE','Coastal states only','Tests Arctic littoral state logic']
]
sh4=get_clean_sheet('Model_Specifications')
sh4.get_range('A1:H1').values=[spec_headers]
sh4.get_range_by_indexes(1,0,len(spec_rows),8).values=spec_rows
sh4.get_range('A1:H1').format.fill='#1E3A8A'
sh4.get_range('A1:H1').format.font={'bold': True, 'color':'#FFFFFF'}
sh4.get_range('A:H').format.wrap_text=True
# set widths
for c,w in enumerate([24,24,24,42,42,22,44,44]):
    sh4.get_range_by_indexes(0,c,len(spec_rows)+1,1).format.column_width=w
try:
    sh4.tables.add(sh4.get_range_by_indexes(0,0,len(spec_rows)+1,8), True, 'ModelSpecsTable')
except Exception:
    pass

# Transformations sheet
trans_headers=['variable','formula_or_construction','interpretation','caution']
trans_rows=[
['ln_military_exp_constant_usd','ln(military_exp_constant_usd)','Scale-adjusted military expenditure level','Use as robustness DV, not main DV'],
['military_exp_growth_pct','100*((military_exp_constant_usd_t / military_exp_constant_usd_t-1)-1) within country','Annual real military expenditure growth','First year per country is missing'],
['ln_gdp_per_capita','ln(gdp_per_capita)','Economic development/capacity control','Use instead of raw GDP per capita in most models'],
['ln_population','ln(population)','Country scale control','Mostly time trend; country FE absorbs level differences'],
['sea_ice_decline','September sea ice extent_(t-1) - September sea ice extent_t','Positive value means a decline in September sea ice extent','Common Arctic-wide variable; not country-specific'],
['sea_ice_decline_x_SCFI','sea_ice_decline * SCFI_model','Tests conditional effect of melting under strategic competition framing','Requires content sample'],
['SCFI_x_post2022','SCFI_model * post_2022_war','Tests post-Russia-Ukraine war framing shift','Year FE can absorb post_2022 main effect'],
['NAFI_x_post2022','NAFI_model * post_2022_war','Tests post-war NATO/alliance framing shift','Security supplement layer only'],
['energy_price_index_2010_100','average of Brent oil and Europe gas price normalized to 2010=100','Captures energy value context','Avoid overfitting with both oil and gas in small sample']
]
sh5=get_clean_sheet('Variable_Transformations')
sh5.get_range('A1:D1').values=[trans_headers]
sh5.get_range_by_indexes(1,0,len(trans_rows),4).values=trans_rows
sh5.get_range('A1:D1').format.fill='#4C1D95'
sh5.get_range('A1:D1').format.font={'bold': True, 'color':'#FFFFFF'}
sh5.get_range('A:D').format.wrap_text=True
for c,w in enumerate([30,50,50,45]):
    sh5.get_range_by_indexes(0,c,len(trans_rows)+1,1).format.column_width=w
try:
    sh5.tables.add(sh5.get_range_by_indexes(0,0,len(trans_rows)+1,4), True, 'TransformationsTable')
except Exception:
    pass

# Update Next_Steps
sh6=get_clean_sheet('Next_Steps_Modeling')
steps=[['step','task','status','comment'],
       [1,'Run descriptive statistics and correlation matrix','Next','Use Model_Dataset and sample flags'],
       [2,'Estimate baseline two-way fixed effects model','Next','M0_external_baseline'],
       [3,'Estimate content model with SCFI','Pending','M1/M2 after checking sample size'],
       [4,'Run NATO/Russia-threat supplement model','Pending','M4; interpret as security supplement layer'],
       [5,'Robustness: exclude Russia and Iceland','Pending','Russia and Iceland are structurally unusual cases'],
       [6,'Prepare figures: sea ice trend, military burden by country, content score shifts','Pending','For article draft']]
sh6.get_range_by_indexes(0,0,len(steps),4).values=steps
sh6.get_range('A1:D1').format.fill='#111827'
sh6.get_range('A1:D1').format.font={'bold': True, 'color':'#FFFFFF'}
sh6.get_range('A:D').format.autofit_columns()

# verify snippets
print(wb.inspect({'kind':'table','range':'Model_Readiness_Check!A1:C15','include':'values','table_max_rows':20,'table_max_cols':5}).ndjson)
print(wb.inspect({'kind':'match','search_term':'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A','options':{'use_regex': True, 'max_results':50},'summary':'formula error scan'}).ndjson)
SpreadsheetFile.export_xlsx(wb).save(output_path)
print(output_path)
