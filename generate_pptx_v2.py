import io, zipfile, pandas as pd

def build_agency_pptx(df, template_path):
    # Nettoyage des colonnes
    df.columns = [c.strip() for c in df.columns]
    spend_col = "Integrated Spends" if "Integrated Spends" in df.columns else "Ad Spends"
    df[spend_col] = pd.to_numeric(df[spend_col], errors='coerce').fillna(0)

    agencies_stats = []
    for agency, group in df.groupby('Agency'):
        # Calcul NBB (uniquement WIN et DEPARTURE) 
        wins_sum = group[group['NewBiz'] == 'WIN'][spend_col].sum()
        deps_sum = group[group['NewBiz'] == 'DEPARTURE'][spend_col].sum()
        total_nbb = wins_sum - abs(deps_sum)

        # Calcul Rétention 
        total_retention = group[group['NewBiz'] == 'Retention'][spend_col].sum()

        # Filtrage Moves Significatifs (> 3M ou < -3M) pour Slide 3, 4 et Détails 
        significant = group[group[spend_col].abs() >= 3]
        
        wins_desc = " · ".join([f"{r['Advertiser']} {r[spend_col]:.0f}m" for _, r in significant[significant['NewBiz'] == 'WIN'].iterrows()])
        deps_desc = " · ".join([f"{r['Advertiser']} {abs(r[spend_col]):.0f}m" for _, r in significant[significant['NewBiz'] == 'DEPARTURE'].iterrows()])
        rets_desc = " · ".join([f"{r['Advertiser']} {r[spend_col]:.0f}m" for _, r in significant[significant['NewBiz'] == 'Retention'].iterrows()])

        agencies_stats.append({
            'agency': agency,
            'nbb': total_nbb,
            'retention': total_retention,
            'wins_desc': wins_desc if wins_desc else "-",
            'deps_desc': deps_desc if deps_desc else "-",
            'rets_desc': rets_desc if rets_desc else "-"
        })

    # Tri par NBB (Slide 7+) et Rétention (Slide 6) 
    nbb_ranking = sorted(agencies_stats, key=lambda x: x['nbb'], reverse=True)
    ret_ranking = sorted(agencies_stats, key=lambda x: x['retention'], reverse=True)

    # Simulation de modification PPTX (Structure XML) 
    with open(template_path, "rb") as f:
        zip_data = f.read()

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        with zipfile.ZipFile(io.BytesIO(zip_data)) as zin:
            for item in zin.infolist():
                zout.writestr(item.filename, zin.read(item.filename))
    
    return buf.getvalue()