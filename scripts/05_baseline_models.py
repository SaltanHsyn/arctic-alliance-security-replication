import pandas as pd, numpy as np, re, json, math
import statsmodels.formula.api as smf
from statsmodels.tools.sm_exceptions import ConvergenceWarning
import warnings
from pathlib import Path

ROOT=Path(__file__).resolve().parents[1]
INPUT=ROOT/'data'/'Arctic_Model_Ready_Dataset_v1.xlsx'
OUT=ROOT/'outputs'/'Arctic_Panel_Baseline_Models_v1.xlsx'
OUT.parent.mkdir(exist_ok=True)

df=pd.read_excel(INPUT, sheet_name='Model_Dataset')
# Normalize names and types
for col in ['country']:
    df[col]=df[col].astype(str)
# Ensure numeric columns
for col in df.columns:
    if col not in ['country','policy_source_file','security_source_file','sipri_data_quality_flag','wdi_data_quality_flag','sea_ice_source_flag','energy_price_source_flag']:
        df[col]=pd.to_numeric(df[col], errors='ignore')

# Treat missing energy rents as 0 where WDI resource rents absent after 2022? No, leave missing for models that use it.
# Build alternate complete controls excluding energy_rents to keep 200 obs.

models=[]

def run_model(model_id, label, formula, data, sample_note):
    # drop complete cases for vars in formula automatically by statsmodels
    try:
        res = smf.ols(formula, data=data).fit(cov_type='HC3')
        n=int(res.nobs)
        r2=float(res.rsquared)
        adj=float(res.rsquared_adj)
        out={
            'model_id':model_id,
            'label':label,
            'formula':formula,
            'sample_note':sample_note,
            'n_obs':n,
            'r_squared':r2,
            'adj_r_squared':adj,
            'aic':float(res.aic),
            'bic':float(res.bic),
            'status':'OK',
            'warning':''
        }
        coef=[]
        for term in res.params.index:
            coef.append({
                'model_id':model_id,
                'term':term,
                'coef':res.params[term],
                'std_err_HC3':res.bse[term],
                't_value':res.tvalues[term],
                'p_value':res.pvalues[term],
                'ci_025':res.conf_int().loc[term,0],
                'ci_975':res.conf_int().loc[term,1]
            })
        return out,coef,res
    except Exception as e:
        return {'model_id':model_id,'label':label,'formula':formula,'sample_note':sample_note,'n_obs':0,'r_squared':np.nan,'adj_r_squared':np.nan,'aic':np.nan,'bic':np.nan,'status':'ERROR','warning':str(e)},[],None

# External controls models; do not include year FE because sea ice and energy prices are common yearly shocks.
base_controls = 'sea_ice_decline + energy_price_index_2010_100 + ln_gdp_per_capita + gdp_growth + inflation + trade_openness + ln_population + post_2022_war'
models_spec = [
    ('M1','Pooled baseline external model', f'military_burden_pct_gdp ~ {base_controls}', df, 'All Arctic Eight, pooled OLS, HC3 robust SE.'),
    ('M2','Country fixed effects baseline', f'military_burden_pct_gdp ~ {base_controls} + C(country)', df, 'All Arctic Eight; country dummies absorb time-invariant country differences.'),
    ('M3','Country FE with NATO accession', f'military_burden_pct_gdp ~ {base_controls} + nato_membership + post_nato_accession + C(country)', df, 'All Arctic Eight; NATO variables only identify within-country changes for Finland/Sweden.'),
    ('M4','Country FE with energy interaction', f'military_burden_pct_gdp ~ sea_ice_decline + energy_price_index_2010_100 + sea_ice_decline_x_energy_price_index + ln_gdp_per_capita + gdp_growth + inflation + trade_openness + ln_population + post_2022_war + C(country)', df, 'All Arctic Eight; tests conditional energy-price mechanism.'),
    ('M5','Coastal Five country FE', f'military_burden_pct_gdp ~ {base_controls} + C(country)', df[df['coastal_five']==1], 'Coastal Five only.'),
    ('M6','Russia excluded country FE', f'military_burden_pct_gdp ~ {base_controls} + C(country)', df[df['country']!='Russia'], 'Robustness excluding Russia.'),
    ('M7','External controls with resource rents', f'military_burden_pct_gdp ~ {base_controls} + energy_rents_pct_gdp + oil_rents_pct_gdp + gas_rents_pct_gdp + C(country)', df, 'Includes WDI resource rents; sample may shrink where WDI missing.'),
]

summary=[]; coefs=[]; fitted={}
for spec in models_spec:
    s,c,r=run_model(*spec)
    summary.append(s); coefs.extend(c); fitted[spec[0]]=r

# Content sample models (small n)
content=df[df['sample_content_main']==1].copy() if 'sample_content_main' in df.columns else df[df['SCFI_model'].notna()].copy()
content_controls = content.copy()
content_specs=[
    ('C1','Content-only pooled model', 'military_burden_pct_gdp ~ SCFI_model + NAFI_model + RussiaThreat_model + Resilience_model + CriticalInfra_model', content_controls, 'Small content-coded sample only; no FE.'),
    ('C2','Content plus post-2022 and NATO', 'military_burden_pct_gdp ~ SCFI_model + NAFI_model + RussiaThreat_model + Resilience_model + CriticalInfra_model + post_2022_war + nato_membership', content_controls, 'Small content-coded sample; screening only.'),
    ('C3','Content with country FE cautious', 'military_burden_pct_gdp ~ SCFI_model + NAFI_model + RussiaThreat_model + post_2022_war + C(country)', content_controls, 'Small content-coded sample with country dummies; high overfit risk.'),
]
for spec in content_specs:
    s,c,r=run_model(*spec)
    summary.append(s); coefs.extend(c); fitted[spec[0]]=r

summary_df=pd.DataFrame(summary)
coefs_df=pd.DataFrame(coefs)

# Significant coefficients of substantive variables only
substantive_terms = ['sea_ice_decline','energy_price_index_2010_100','post_2022_war','nato_membership','post_nato_accession','sea_ice_decline_x_energy_price_index','SCFI_model','NAFI_model','RussiaThreat_model','Resilience_model','CriticalInfra_model','energy_rents_pct_gdp','oil_rents_pct_gdp','gas_rents_pct_gdp']
sub_df=coefs_df[coefs_df['term'].isin(substantive_terms)].copy()
sub_df['significance'] = pd.cut(sub_df['p_value'], bins=[-np.inf,0.01,0.05,0.10,np.inf], labels=['***','**','*',''])

# Country and year summaries
country_summary = df.groupby('country').agg(
    n=('year','count'),
    mean_military_burden=('military_burden_pct_gdp','mean'),
    pre2022_military_burden=('military_burden_pct_gdp', lambda x: x[df.loc[x.index,'post_2022_war']==0].mean()),
    post2022_military_burden=('military_burden_pct_gdp', lambda x: x[df.loc[x.index,'post_2022_war']==1].mean()),
    mean_SCFI=('SCFI_model','mean'),
    mean_NAFI=('NAFI_model','mean'),
    mean_RussiaThreat=('RussiaThreat_model','mean')
).reset_index()
country_summary['post_minus_pre_military_burden'] = country_summary['post2022_military_burden'] - country_summary['pre2022_military_burden']

year_summary = df.groupby('year').agg(
    mean_military_burden=('military_burden_pct_gdp','mean'),
    sea_ice_extent_m_km2=('sea_ice_extent_m_km2','first'),
    sea_ice_decline=('sea_ice_decline','first'),
    energy_price_index_2010_100=('energy_price_index_2010_100','first')
).reset_index()

# Method warnings
warnings_df=pd.DataFrame([
    ['No year fixed effects in core models','Sea ice and energy prices vary only by year; full year FE would absorb their effects. Country FE and trend/post-2022 controls are used instead.'],
    ['Small number of countries','Arctic Eight contains only 8 countries. Inference should be interpreted cautiously; HC3 robust SE are used for screening.'],
    ['Content-score sample is small','Content-coded variables are available for 41 observations, not 200. Content models are exploratory and should be used primarily for mechanism evidence.'],
    ['Iceland special case','Iceland has no standing army; military burden is not directly comparable to other NATO states. Robustness should test Iceland excluded.'],
    ['Russia special case','Russia is the main rival/reference case and a structurally different political economy. Robustness excluding Russia is required.'],
    ['Norway 2025 document','Norway 2025 strategy is useful for post-war framing but sits outside 2000–2024 baseline period. Treat with caution in final article narrative.'],
], columns=['warning','implication'])


# Write workbook with pandas/openpyxl (portable: no artifact_tool dependency)
dashboard = pd.DataFrame([
    ['Project','Arctic Q1 Project — Baseline Panel Models'],
    ['Input dataset',str(INPUT)],
    ['Main dependent variable','military_burden_pct_gdp'],
    ['Primary environmental variable','sea_ice_decline'],
    ['Energy moderator','energy_price_index_2010_100 and sea_ice_decline_x_energy_price_index'],
    ['Main sample','Arctic Eight, 2000–2024, 200 country-year observations'],
    ['Content sample','41 content-coded observations; exploratory mechanism models only'],
    ['Important limitation','Core models do not use year fixed effects because common yearly regressors would be absorbed.'],
    ['Next step','Inspect robustness models.'],
], columns=['item','value'])
interp=[]
for mid in ['M1','M2','M3','M4','M5','M6','C1','C2','C3']:
    sub=sub_df[sub_df.model_id==mid] if not sub_df.empty else pd.DataFrame()
    for term in substantive_terms:
        row=sub[sub.term==term] if not sub.empty else pd.DataFrame()
        if not row.empty:
            r=row.iloc[0]
            interp.append([mid, term, 'positive' if r.coef>0 else 'negative', round(r.coef,4), round(r.p_value,4), str(r.significance), 'screening signal only'])
interpret_df=pd.DataFrame(interp, columns=['model_id','term','direction','coef','p_value','significance','note'])

with pd.ExcelWriter(OUT, engine='openpyxl') as writer:
    dashboard.to_excel(writer, sheet_name='Dashboard', index=False)
    summary_df.to_excel(writer, sheet_name='Model_Summary', index=False)
    sub_df.to_excel(writer, sheet_name='Substantive_Coefficients', index=False)
    coefs_df.to_excel(writer, sheet_name='All_Coefficients', index=False)
    country_summary.to_excel(writer, sheet_name='Country_Summary', index=False)
    year_summary.to_excel(writer, sheet_name='Annual_Trends', index=False)
    pd.DataFrame([{k:v for k,v in m.items() if k in ['model_id','label','formula','sample_note']} for m in summary]).to_excel(writer, sheet_name='Model_Specifications', index=False)
    warnings_df.to_excel(writer, sheet_name='Method_Warnings', index=False)
    interpret_df.to_excel(writer, sheet_name='Interpretation_Aid', index=False)
print('saved', OUT)
