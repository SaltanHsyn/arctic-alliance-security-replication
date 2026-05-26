import pandas as pd, numpy as np, math
import statsmodels.formula.api as smf

from pathlib import Path
ROOT=Path(__file__).resolve().parents[1]
INPUT=ROOT/'data'/'Arctic_Model_Ready_Dataset_v1.xlsx'
OUT=ROOT/'outputs'/'Arctic_Robustness_Models_v1.xlsx'
OUT.parent.mkdir(exist_ok=True)

df=pd.read_excel(INPUT, sheet_name='Model_Dataset')
for col in df.columns:
    if col not in ['country','policy_source_file','security_source_file','sipri_data_quality_flag','wdi_data_quality_flag','sea_ice_source_flag','energy_price_source_flag']:
        df[col]=pd.to_numeric(df[col], errors='coerce')
df['country']=df['country'].astype(str)
# Create lagged variables by country for DV and common annual variables.
df=df.sort_values(['country','year']).reset_index(drop=True)
for var in ['military_burden_pct_gdp','sea_ice_decline','energy_price_index_2010_100','brent_oil_price','europe_gas_price','SCFI_model','NAFI_model','RussiaThreat_model']:
    if var in df.columns:
        df[f'L1_{var}'] = df.groupby('country')[var].shift(1)
# Ensure interactions exist with lagged variables
df['L1_sea_ice_x_energy'] = df['L1_sea_ice_decline'] * df['L1_energy_price_index_2010_100']
df['post2022_x_nato'] = df['post_2022_war'] * df['nato_membership']
df['post2022_x_russia_dummy'] = df['post_2022_war'] * df['russia_dummy']

models=[]
coef_rows=[]
substantive_terms = [
    'sea_ice_decline','energy_price_index_2010_100','sea_ice_decline_x_energy_price_index',
    'post_2022_war','nato_membership','post_nato_accession','post2022_x_nato','post2022_x_russia_dummy',
    'ln_gdp_per_capita','gdp_growth','inflation','trade_openness','ln_population',
    'energy_rents_pct_gdp','oil_rents_pct_gdp','gas_rents_pct_gdp',
    'L1_sea_ice_decline','L1_energy_price_index_2010_100','L1_sea_ice_x_energy',
    'SCFI_model','NAFI_model','RussiaThreat_model','Resilience_model','CriticalInfra_model',
    'SCFI_x_post2022','NAFI_x_post2022','RussiaThreat_x_post2022'
]

def model_vars_from_formula(formula):
    # rough parse, used for complete-case counts only
    rhs=formula.split('~',1)[1]
    terms=[]
    for t in rhs.replace('+',' + ').split('+'):
        t=t.strip()
        if not t or t.startswith('C('):
            continue
        if ':' in t:
            terms.extend([x.strip() for x in t.split(':') if not x.strip().startswith('C(')])
        elif '*' in t:
            terms.extend([x.strip() for x in t.split('*') if not x.strip().startswith('C(')])
        else:
            terms.append(t)
    lhs=formula.split('~',1)[0].strip()
    return [lhs] + [t for t in terms if t in df.columns]

def run(model_id, label, formula, data, sample_def, caution):
    vars_needed=model_vars_from_formula(formula)
    sample_n_before=len(data)
    complete=data.dropna(subset=[v for v in vars_needed if v in data.columns]).copy()
    try:
        res=smf.ols(formula, data=data).fit(cov_type='HC3')
        status='OK'; warning=''
        n=int(res.nobs)
        summary={
            'model_id':model_id,'label':label,'formula':formula,'sample_definition':sample_def,
            'sample_n_before_complete_case':sample_n_before,'n_obs_used':n,
            'r_squared':float(res.rsquared),'adj_r_squared':float(res.rsquared_adj),'aic':float(res.aic),'bic':float(res.bic),
            'covariance':'HC3 robust','status':status,'warning':warning,'caution':caution
        }
        cis=res.conf_int()
        for term in res.params.index:
            coef_rows.append({
                'model_id':model_id,'term':term,'coef':float(res.params[term]),'std_err_HC3':float(res.bse[term]),
                't_value':float(res.tvalues[term]),'p_value':float(res.pvalues[term]),
                'ci_025':float(cis.loc[term,0]),'ci_975':float(cis.loc[term,1]),
                'substantive_term':'Yes' if term in substantive_terms else 'No'
            })
    except Exception as e:
        summary={
            'model_id':model_id,'label':label,'formula':formula,'sample_definition':sample_def,
            'sample_n_before_complete_case':sample_n_before,'n_obs_used':0,
            'r_squared':np.nan,'adj_r_squared':np.nan,'aic':np.nan,'bic':np.nan,
            'covariance':'HC3 robust','status':'ERROR','warning':str(e),'caution':caution
        }
    models.append(summary)

base_controls='sea_ice_decline + energy_price_index_2010_100 + ln_gdp_per_capita + gdp_growth + inflation + trade_openness + ln_population + post_2022_war'
base_fe=base_controls + ' + C(country)'
interaction='sea_ice_decline + energy_price_index_2010_100 + sea_ice_decline_x_energy_price_index + ln_gdp_per_capita + gdp_growth + inflation + trade_openness + ln_population + post_2022_war + C(country)'

run('R1','Full Arctic Eight: country FE baseline', f'military_burden_pct_gdp ~ {base_fe}', df, 'All Arctic Eight, 2000-2024', 'No year FE because sea ice and energy prices vary only by year.')
run('R2','Full Arctic Eight: energy interaction', f'military_burden_pct_gdp ~ {interaction}', df, 'All Arctic Eight, 2000-2024', 'Tests whether sea-ice decline has stronger effects when energy prices are high.')
run('R3','Iceland excluded', f'military_burden_pct_gdp ~ {interaction}', df[df.country!='Iceland'], 'All except Iceland', 'Iceland has no standing army; this checks whether results depend on that special case.')
run('R4','Russia excluded', f'military_burden_pct_gdp ~ {interaction}', df[df.country!='Russia'], 'All except Russia', 'Russia is structurally different and central rival case; exclusion test is essential.')
run('R5','Coastal Five only', f'military_burden_pct_gdp ~ {interaction}', df[df.coastal_five==1], 'USA, Canada, Russia, Norway, Denmark/Greenland', 'Direct Arctic coastal exposure; small N and fewer countries.')
run('R6','NATO states only', f'military_burden_pct_gdp ~ {interaction}', df[df.country!='Russia'], 'All NATO or post-NATO Arctic states; Russia excluded', 'Similar to Russia excluded; focuses alliance side.')
run('R7','Post-2022 only', 'military_burden_pct_gdp ~ sea_ice_decline + energy_price_index_2010_100 + ln_gdp_per_capita + gdp_growth + trade_openness + C(country)', df[df.post_2022_war==1], '2022-2024 only', 'Very small sample; only used as crisis-period sensitivity check.')
run('R8','Pre-2022 only', 'military_burden_pct_gdp ~ sea_ice_decline + energy_price_index_2010_100 + ln_gdp_per_capita + gdp_growth + trade_openness + C(country)', df[df.post_2022_war==0], '2000-2021 only', 'Pre-war baseline; excludes the post-2022 shock period.')
run('R9','Lagged environmental and energy variables', 'military_burden_pct_gdp ~ L1_sea_ice_decline + L1_energy_price_index_2010_100 + L1_sea_ice_x_energy + ln_gdp_per_capita + gdp_growth + inflation + trade_openness + ln_population + post_2022_war + C(country)', df, 'All Arctic Eight, 2001-2024 after lagging', 'Lagged model reduces simultaneity concerns but still exploratory.')
run('R10','Alternative DV: log military expenditure', 'ln_military_exp_constant_usd ~ sea_ice_decline + energy_price_index_2010_100 + sea_ice_decline_x_energy_price_index + ln_gdp_per_capita + gdp_growth + inflation + trade_openness + ln_population + post_2022_war + C(country)', df, 'All Arctic Eight where log military expenditure available', 'Uses spending level rather than burden; changes substantive interpretation.')
run('R11','Post-2022 x NATO interaction', f'military_burden_pct_gdp ~ {base_controls} + post2022_x_nato + C(country)', df, 'All Arctic Eight, 2000-2024', 'Tests whether post-2022 shift is stronger among NATO members; limited by NATO membership being mostly time-invariant.')
run('R12','Post-2022 x Russia dummy interaction', f'military_burden_pct_gdp ~ {base_controls} + post2022_x_russia_dummy + C(country)', df, 'All Arctic Eight, 2000-2024', 'Checks whether Russia drives the post-2022 shift differently from other states.')
# Content models (small N)
content=df[df.sample_content_main==1].copy()
run('C1','Content sample: strategic framing only', 'military_burden_pct_gdp ~ SCFI_model + NAFI_model + RussiaThreat_model + Resilience_model + CriticalInfra_model', content, 'Content-coded country-period observations only', 'Mechanism screen; not final causal model.')
run('C2','Content sample: post-2022 interactions', 'military_burden_pct_gdp ~ SCFI_model + NAFI_model + RussiaThreat_model + SCFI_x_post2022 + NAFI_x_post2022 + RussiaThreat_x_post2022', content, 'Content-coded observations only', 'Small sample; high overfitting risk.')
run('C3','Content sample: compact mechanism model', 'military_burden_pct_gdp ~ SCFI_model + NAFI_model + RussiaThreat_model + post_2022_war + nato_membership', content, 'Content-coded observations only', 'More parsimonious content model.')

summary_df=pd.DataFrame(models)
coefs_df=pd.DataFrame(coef_rows)
if not coefs_df.empty:
    coefs_df['significance'] = pd.cut(coefs_df['p_value'], bins=[-np.inf,0.01,0.05,0.10,np.inf], labels=['***','**','*','']).astype(str).replace('nan','')
sub_df=coefs_df[coefs_df['substantive_term']=='Yes'].copy() if not coefs_df.empty else pd.DataFrame()
# Direction stability: terms across R1-R12 (excluding content) for substantive variables
stability=[]
for term in substantive_terms:
    ss=sub_df[(sub_df.term==term)&(sub_df.model_id.str.startswith('R'))]
    if len(ss):
        stability.append({
            'term':term,'n_models':len(ss),'positive_count':int((ss.coef>0).sum()),'negative_count':int((ss.coef<0).sum()),
            'sig_p_lt_0_10_count':int((ss.p_value<0.10).sum()),'mean_coef':float(ss.coef.mean()),'min_p_value':float(ss.p_value.min())
        })
stability_df=pd.DataFrame(stability)
# Sample diagnostics
sample_rows=[]
samples={
    'Full Arctic Eight':df,
    'Iceland excluded':df[df.country!='Iceland'],
    'Russia excluded':df[df.country!='Russia'],
    'Coastal Five':df[df.coastal_five==1],
    'Post-2022 only':df[df.post_2022_war==1],
    'Pre-2022 only':df[df.post_2022_war==0],
    'Content-coded observations':content,
}
for name,d in samples.items():
    sample_rows.append({
        'sample':name,'n_obs':len(d),'countries':d.country.nunique(),'years_min':int(d.year.min()) if len(d) else None,'years_max':int(d.year.max()) if len(d) else None,
        'mean_military_burden':float(d.military_burden_pct_gdp.mean()) if len(d) else None,
        'post2022_share':float(d.post_2022_war.mean()) if len(d) else None
    })
sample_df=pd.DataFrame(sample_rows)
# Country shifts
country_shift=df.groupby('country').apply(lambda g: pd.Series({
    'mean_pre2022':g.loc[g.post_2022_war==0,'military_burden_pct_gdp'].mean(),
    'mean_post2022':g.loc[g.post_2022_war==1,'military_burden_pct_gdp'].mean(),
    'post_minus_pre':g.loc[g.post_2022_war==1,'military_burden_pct_gdp'].mean()-g.loc[g.post_2022_war==0,'military_burden_pct_gdp'].mean(),
    'n_pre':int((g.post_2022_war==0).sum()),'n_post':int((g.post_2022_war==1).sum())
})).reset_index()
# Red flags
warnings_df=pd.DataFrame([
    ['Common yearly regressors','Sea ice and global energy prices are identical across countries in a given year; full year fixed effects cannot be used with these variables.'],
    ['Small-N country panel','Only eight countries are available; all inference must be presented as exploratory/associational unless backed by strong theory.'],
    ['Post-2022 window is short','The post-war period contains only 2022-2024 in the current panel. Post-war models are sensitivity checks, not standalone proof.'],
    ['Content sample is limited','Content scores cover 41 country-year observations due to policy-regime coding. Content models should support mechanism discussion rather than serve as final causal identification.'],
    ['Iceland special case','Iceland lacks a standing army. Report Iceland-excluded robustness for military burden models.'],
    ['Russia special case','Russia is both an Arctic state and the principal threat reference after 2022. Russia-excluded models are required.'],
], columns=['issue','implication'])


# Write workbook with pandas/openpyxl (portable: no artifact_tool dependency)
dashboard=pd.DataFrame([
    ['Artifact','Arctic Robustness Models v1'],
    ['Input',str(INPUT)],
    ['Purpose','Stress-test baseline findings across country exclusions, coastal-only sample, pre/post-2022 periods, lagged regressors, alternative DV, and content-score mechanism models.'],
    ['Core dependent variable','military_burden_pct_gdp'],
    ['Main environmental mechanism','sea_ice_decline and sea_ice_decline_x_energy_price_index'],
    ['Main alliance/security mechanism','post_2022_war, NATO membership/accession, NAFI/RussiaThreat content scores'],
    ['Models estimated',len(summary_df)],
    ['Important caution','These models are robustness screens, not final publishable identification.'],
], columns=['item','value'])
next_steps=pd.DataFrame([
    ['1','Inspect coefficient stability for sea ice, energy interaction, and post-2022 terms.'],
    ['2','Use Iceland-excluded and Russia-excluded models as required robustness checks.'],
    ['3','Treat post-2022-only and content-score models as mechanism/sensitivity evidence due to small N.'],
    ['4','Choose 3-5 final specifications for paper tables.'],
    ['5','Prepare visualization: country post-2022 shifts, sea ice trend, energy price trend, and predicted military burden under energy interaction.'],
], columns=['step','action'])
with pd.ExcelWriter(OUT, engine='openpyxl') as writer:
    dashboard.to_excel(writer, sheet_name='Dashboard', index=False)
    summary_df.to_excel(writer, sheet_name='Robustness_Model_Summary', index=False)
    sub_df.to_excel(writer, sheet_name='Substantive_Coefficients', index=False)
    stability_df.to_excel(writer, sheet_name='Coefficient_Stability', index=False)
    coefs_df.to_excel(writer, sheet_name='All_Coefficients', index=False)
    sample_df.to_excel(writer, sheet_name='Sample_Diagnostics', index=False)
    country_shift.to_excel(writer, sheet_name='Country_Post2022_Shifts', index=False)
    summary_df[['model_id','label','formula','sample_definition','caution']].to_excel(writer, sheet_name='Model_Formulas', index=False)
    warnings_df.to_excel(writer, sheet_name='Method_Warnings', index=False)
    next_steps.to_excel(writer, sheet_name='Next_Steps', index=False)
print('saved', OUT)
