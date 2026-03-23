# =============================================================================
# CommodityWatch v3.0
# Framework : Streamlit + Plotly | Python 3.9+
# Usage     : streamlit run app.py
#

#

# =============================================================================

import streamlit as st
import numpy as np
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots

try:
    import openpyxl        # noqa: F401
    EXCEL_OK = True
except ImportError:
    EXCEL_OK = False


# =============================================================================
# CONSTANTES
# =============================================================================

TAUX_CHANGE  = 600.0      # 1 USD = 600 FCFA

PRIX_REF_USD = {           # USD / Tonne  —  ICE/Euronext/OTC indicatif
    "Cacao"           : 7_800,
    "Café (Arabica)"  : 4_200,
    "Café (Robusta)"  : 2_800,
    "Anacarde (brut)" : 1_200,
    "Soja"            :   430,
    "Coton"           :   750,
    "Caoutchouc"      : 1_450,
    "Palmier à huile" :   900,
    "Maïs"            :   240,
    "Personnalisé"    :     0,
}

# Risque & assurance associés à chaque Business Model
# (risque_label, assurance_label, gravite_defaut, description_risque, description_assurance)
BM_RISQUE = {
    "Production" : {
        "risque_label"        : "Risque Récolte (Climatique / Phytosanitaire)",
        "assurance_label"     : "Assurance Agricole / Récolte",
        "gravite_defaut"      : 0.40,
        "description_risque"  : (
            "Sécheresse, inondation, gel, épidémie phytosanitaire ou "
            "ravageurs pouvant réduire le rendement de façon significative."
        ),
        "description_assurance": (
            "L'assurance récolte indemnise une fraction de la perte de production "
            "constatée. La prime est une charge opérationnelle annuelle."
        ),
    },
    "Infrastructure" : {
        "risque_label"        : "Risque Opérationnel (Panne / Chantier)",
        "assurance_label"     : "Assurance Tous Risques Chantier & Bris de Machine",
        "gravite_defaut"      : 0.25,
        "description_risque"  : (
            "Panne majeure d'équipement, accident de chantier, incendie "
            "ou explosion réduisant la capacité de production installée."
        ),
        "description_assurance": (
            "L'assurance TRC / Bris de Machine couvre les dommages matériels "
            "et la perte d'exploitation consécutive à un sinistre sur l'outil industriel."
        ),
    },
    "Campagne" : {
        "risque_label"        : "Risque Qualité (Non-conformité / Rappel Produit)",
        "assurance_label"     : "Assurance Qualité & Responsabilité Produit",
        "gravite_defaut"      : 0.20,
        "description_risque"  : (
            "Lot non conforme aux normes export (humidité, aflatoxines, résidus), "
            "perte de certification, rappel de produit ou embargo à l'exportation."
        ),
        "description_assurance": (
            "L'assurance qualité / responsabilité produit couvre les pertes de CA "
            "liées au retrait de lot, aux pénalités contractuelles et aux frais de rappel."
        ),
    },
}

SEUIL_CHARGES_CA  = 0.55   # Alerte si charges opé > 55 % du CA
SEUIL_DSCR_ALERTE = 1.30   # Seuil bancaire standard
SEUIL_DSCR_DEFAUT = 1.00   # Seuil de défaut

CLR = {
    "vert"  : "#198754",
    "rouge" : "#dc3545",
    "amber" : "#fd7e14",
    "bleu"  : "#0d6efd",
    "gris"  : "#6c757d",
    "fond"  : "#f8f9fa",
}


# =============================================================================
# UTILITAIRES MATHÉMATIQUES
# =============================================================================

def npv_calc(rate: float, cashflows) -> float:
    cf = np.asarray(cashflows, dtype=float)
    t  = np.arange(len(cf))
    return float(np.sum(cf / (1 + rate) ** t))


def irr_calc(cashflows, tol=1e-7, max_iter=1500):
    cf = np.asarray(cashflows, dtype=float)
    if not (np.any(cf < 0) and np.any(cf > 0)):
        return None
    rate = 0.10
    for _ in range(max_iter):
        t  = np.arange(len(cf))
        f  = np.sum(cf / (1 + rate) ** t)
        df = np.sum(-t * cf / (1 + rate) ** (t + 1))
        if df == 0:
            return None
        nr = rate - f / df
        if abs(nr - rate) < tol:
            return nr
        rate = nr
    return None


def en_fcfa(val, devise):
    return val if devise == "FCFA" else val * TAUX_CHANGE

def en_devise(val_fcfa, devise):
    return val_fcfa if devise == "FCFA" else val_fcfa / TAUX_CHANGE

def fmt(val, devise, dec=2):
    return f"{val:,.{dec}f} M{devise}"


# =============================================================================
# CORE — RAMP-UP
# =============================================================================

def build_rampup(duree, paliers):
    """Vecteur de coefficients [0..1] par année selon les paliers configurés."""
    coeffs = []
    for an in range(1, duree + 1):
        taux = 1.0
        for (fin, t) in sorted(paliers, key=lambda x: x[0]):
            if an <= fin:
                taux = t
                break
        coeffs.append(taux)
    return coeffs


# =============================================================================
# CORE — BFR PAR BUSINESS MODEL
# =============================================================================

def calc_bfr(bm, revenus, couts, j_stock, j_clients, j_fourn):
    """
    Production    : BFR résiduel ~5 % revenus (vente spot à la récolte)
    Infrastructure: BFR nul (paiements contractuels)
    Campagne      : BFR complet (stock matière + crédit clients - crédit fournisseurs)
    """
    if bm == "Campagne":
        return couts * j_stock / 365 + revenus * j_clients / 365 - couts * j_fourn / 365
    if bm == "Production":
        return revenus * 0.05
    return 0.0


# =============================================================================
# CORE — PROJECTION FCFF
# =============================================================================

def projeter_fcff(quantite, prix, cout, capex, taux_is, inflation,
                  duree, rampup, bm,
                  duree_bio=0, j_stock=60, j_clients=30, j_fourn=45,
                  ratio_maint=0.02):
    """
    Projection Free Cash Flow to Firm intégrant :
      - Courbe de ramp-up
      - Phase biologique (Production uniquement)
      - Delta BFR annuel
      - CAPEX de maintenance
    Tous les montants de retour sont en FCFA.
    """
    amort  = capex / duree
    c_maint = capex * ratio_maint
    rows, bfr_prec = [], 0.0

    for an in range(1, duree + 1):
        coeff  = rampup[an - 1]
        # Blocage biologique (plantation : pas de revenu avant maturité)
        if bm == "Production" and an <= duree_bio:
            coeff = 0.0
        inf    = (1 + inflation) ** (an - 1)
        rev    = quantite * coeff * prix  * inf
        cout_  = quantite * coeff * cout  * inf
        ebitda = rev - cout_
        ebit   = ebitda - amort
        impot  = max(ebit * taux_is, 0.0)
        nopat  = ebit - impot
        bfr    = calc_bfr(bm, rev, cout_, j_stock, j_clients, j_fourn)
        delta  = bfr - bfr_prec
        bfr_prec = bfr
        fcff   = nopat + amort - c_maint - delta

        rows.append({
            "Annee"          : an,
            "Ramp-up"        : round(coeff * 100, 1),
            "Revenus"        : rev,
            "Couts"          : cout_,
            "EBITDA"         : ebitda,
            "EBIT"           : ebit,
            "IS"             : impot,
            "Delta BFR"      : delta,
            "FCFF"           : fcff,
            "Ratio Charges"  : round(cout_ / rev, 3) if rev > 0 else np.nan,
        })
    return rows


def metriques_rentabilite(rows, capex, wacc):
    flux    = [-capex] + [r["FCFF"] for r in rows]
    van     = npv_calc(wacc, flux)
    tri     = irr_calc(flux)
    cumul, pb = 0.0, None
    for i, f in enumerate(flux[1:], 1):
        cumul += f
        if cumul >= capex:
            pb = i
            break
    return {"van": van, "tri": tri, "payback": pb}


# =============================================================================
# CORE — DETTE AVANCÉE (commissions flat & engagement)
# =============================================================================

def service_dette(rows, capex, p_dette, r_dette, dur_dette,
                  grace, wacc, c_flat, c_engagt, profil_decaiss):
    """
    Construit le tableau annuel du service de la dette.

    Commissions :
      - Flat      : % du montant total, prélevée une seule fois au 1er décaissement
      - Engagement: % annuel sur la fraction non encore décaissée du prêt

    Grace period : pendant ces années, seuls les intérêts sont dus (pas de capital).
    """
    dette      = capex * p_dette
    ann_remb   = max(dur_dette - grace, 1)
    amort_cap  = dette / ann_remb
    cout_flat  = dette * c_flat
    encours    = dette
    cum_dec    = 0.0
    total_int  = cout_flat   # coût total = flat + intérêts + commissions engagement
    rows_out   = []

    for i, r in enumerate(rows):
        an    = r["Annee"]
        fcff  = r["FCFF"]

        frac_dec    = profil_decaiss[i] if i < len(profil_decaiss) else 0.0
        non_dec     = max(dette - cum_dec, 0.0)
        ce_an       = non_dec * c_engagt if cum_dec < dette else 0.0
        cum_dec    += dette * frac_dec

        if an <= dur_dette:
            interets = encours * r_dette
            remb_cap = 0.0 if an <= grace else amort_cap
            svc      = interets + remb_cap + ce_an
            enc_fin  = encours - remb_cap
            dscr     = fcff / svc if svc > 0 else np.nan

            ff_rest  = [rows[j]["FCFF"] for j in range(i, min(dur_dette, len(rows)))]
            llcr     = npv_calc(r_dette, ff_rest) / encours if encours > 0 else np.nan

            total_int += interets + ce_an
            rows_out.append({
                "Annee"        : an,
                "Encours"      : encours,
                "Interets"     : interets,
                "Remb Capital" : remb_cap,
                "Comm Engagt"  : ce_an,
                "Service"      : svc,
                "FCFF"         : fcff,
                "DSCR"         : round(dscr, 3) if not np.isnan(dscr) else np.nan,
                "LLCR"         : round(llcr, 3) if not np.isnan(llcr) else np.nan,
                "Grace"        : (an <= grace),
            })
            encours = enc_fin
        else:
            rows_out.append({
                "Annee"        : an,
                "Encours"      : 0.0,
                "Interets"     : 0.0,
                "Remb Capital" : 0.0,
                "Comm Engagt"  : 0.0,
                "Service"      : 0.0,
                "FCFF"         : fcff,
                "DSCR"         : np.nan,
                "LLCR"         : np.nan,
                "Grace"        : False,
            })

    df     = pd.DataFrame(rows_out)
    dscrs  = df["DSCR"].dropna()
    d_min  = round(float(dscrs.min()), 3)  if len(dscrs) else np.nan
    d_moy  = round(float(dscrs.mean()), 3) if len(dscrs) else np.nan
    ff_llcr  = [r["FCFF"] for r in rows[:dur_dette]]
    llcr_g   = round(npv_calc(r_dette, ff_llcr) / dette, 3) if dette > 0 else np.nan
    return df, d_min, d_moy, llcr_g, total_int


# =============================================================================
# M4 — STRESS TEST CONTEXTUEL (1 risque / 1 assurance par Business Model)
# =============================================================================

def stress_test(rows, bm, prob, gravite,
                prime_ann, indemnite_pct,
                capex, p_dette, r_dette, dur_dette,
                grace, wacc, c_flat, c_engagt, profil_decaiss):
    """
    Calcule pour l'unique risque associé au Business Model :
      - Scénario BASE       : aucun choc, aucune assurance
      - Scénario CHOC BRUT  : choc à l'année centrale, sans assurance
      - Scénario ASSURANCE  : choc + prime annuelle + indemnité l'année du choc

    Le choc réduit les revenus de `gravite` × 100 %, ce qui réduit le FCFF.
    Retourne un dict avec les 3 séries (DSCR, trésorerie cumulée, FCFF).
    """
    duree     = len(rows)
    an_choc   = max(1, duree // 2)

    def appliquer_choc(rows_in, avec_assurance):
        rows_out   = []
        tresorerie = 0.0
        for r in rows_in:
            fcff = r["FCFF"]
            rev  = r["Revenus"]
            an   = r["Annee"]

            # Impact du choc
            perte = rev * gravite if an == an_choc else 0.0
            fcff  = fcff - perte

            # Assurance
            if avec_assurance:
                fcff -= prime_ann                              # prime chaque année
                if an == an_choc:
                    fcff += perte * indemnite_pct              # indemnité l'année du choc

            tresorerie += fcff
            rows_out.append({
                "Annee"          : an,
                "FCFF"           : fcff,
                "Revenus"        : rev,
                "Tresorerie_cum" : tresorerie,
            })
        return rows_out

    base_rows  = [{"Annee": r["Annee"], "FCFF": r["FCFF"], "Revenus": r["Revenus"]} for r in rows]
    choc_rows  = appliquer_choc(base_rows, avec_assurance=False)
    assur_rows = appliquer_choc(base_rows, avec_assurance=True)

    def get_dscr(scenario_rows):
        df_s, dmin, dmoy, _, _ = service_dette(
            scenario_rows, capex, p_dette, r_dette, dur_dette,
            grace, wacc, c_flat, c_engagt, profil_decaiss
        )
        return df_s["DSCR"].tolist(), dmin

    dscr_base_l,  dscr_base_min  = get_dscr(base_rows)
    dscr_choc_l,  dscr_choc_min  = get_dscr(choc_rows)
    dscr_assur_l, dscr_assur_min = get_dscr(assur_rows)

    # Perte nette de CA en valeur absolue
    perte_brute  = rows[an_choc - 1]["Revenus"] * gravite
    indemnite    = perte_brute * indemnite_pct
    cout_primes  = prime_ann * duree
    gain_net_ass = indemnite - cout_primes

    return {
        "an_choc"         : an_choc,
        "perte_brute"     : perte_brute,
        "indemnite"       : indemnite,
        "cout_primes"     : cout_primes,
        "gain_net_ass"    : gain_net_ass,
        "base"            : {"rows": base_rows,  "dscr": dscr_base_l,  "dscr_min": dscr_base_min},
        "choc"            : {"rows": choc_rows,  "dscr": dscr_choc_l,  "dscr_min": dscr_choc_min},
        "assurance"       : {"rows": assur_rows, "dscr": dscr_assur_l, "dscr_min": dscr_assur_min},
    }


# =============================================================================
# TORNADO & SENSIBILITÉ
# =============================================================================

def calc_tornado(quantite, prix, cout, capex, taux_is, inflation,
                 duree, rampup, bm, wacc,
                 duree_bio=0, j_stock=60, j_clients=30, j_fourn=45, delta=0.10):
    def van(**kw):
        fl = projeter_fcff(**kw)
        return metriques_rentabilite(fl, kw["capex"], wacc)["van"]

    base = dict(quantite=quantite, prix=prix, cout=cout, capex=capex,
                taux_is=taux_is, inflation=inflation, duree=duree,
                rampup=rampup, bm=bm, duree_bio=duree_bio,
                j_stock=j_stock, j_clients=j_clients, j_fourn=j_fourn)
    van_ref = van(**base)

    items = [
        ("Prix de Vente",  "prix",     prix),
        ("Volume Produit", "quantite", quantite),
        ("CAPEX",          "capex",    capex),
        ("Coût Production","cout",     cout),
    ]
    result = []
    for label, param, val in items:
        vm = van(**{**base, param: val * (1 - delta)})
        vp = van(**{**base, param: val * (1 + delta)})
        result.append({"Variable": label,
                        "im": vm - van_ref, "ip": vp - van_ref,
                        "amp": abs(vp - vm)})
    result.sort(key=lambda x: x["amp"], reverse=True)
    return result, van_ref


def calc_sensibilite(quantite, prix, cout, capex, taux_is, inflation,
                     duree, rampup, bm, wacc,
                     duree_bio=0, j_stock=60, j_clients=30, j_fourn=45,
                     variations=(-0.10, -0.05, 0.0, 0.05, 0.10)):
    labs = [f"{int(v*100):+d}%" for v in variations]
    data = {}
    for vc in variations:
        col = []
        for vp in variations:
            fl = projeter_fcff(quantite, prix*(1+vp), cout*(1+vc), capex,
                               taux_is, inflation, duree, rampup, bm,
                               duree_bio, j_stock, j_clients, j_fourn)
            col.append(metriques_rentabilite(fl, capex, wacc)["van"])
        data[f"Coût {int(vc*100):+d}%"] = col
    return pd.DataFrame(data, index=[f"Prix {l}" for l in labs])


# =============================================================================
# M2 — IMPORT EXCEL
# =============================================================================

def parser_excel(fichier):
    """Détecte les postes financiers standards par mots-clés bilingues."""
    MAPPING = {
        "ca"               : ["chiffre d'affaires","revenus","turnover","sales"],
        "charges"          : ["charges d'exploitation","opex","operating expenses"],
        "amort"            : ["amortissement","depreciation"],
        "resultat_net"     : ["résultat net","net income","bénéfice net"],
        "actif_immo"       : ["actif immobilisé","immobilisations","fixed assets"],
        "capitaux_propres" : ["capitaux propres","equity","fonds propres"],
        "dettes_fin"       : ["dettes financières","emprunts","financial debt"],
    }
    try:
        xls  = pd.ExcelFile(fichier)
        data = {}
        for sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet, header=None)
            for _, row in df.iterrows():
                lib = str(row.iloc[0]).lower().strip() if pd.notna(row.iloc[0]) else ""
                for cle, syns in MAPPING.items():
                    if cle not in data:
                        for syn in syns:
                            if syn in lib:
                                nums = [v for v in row.iloc[1:]
                                        if isinstance(v, (int, float)) and not np.isnan(float(v))]
                                if nums:
                                    data[cle] = float(nums[0])
                                break
        return data
    except Exception as e:
        return {"erreur": str(e)}


# =============================================================================
# GRAPHIQUES PLOTLY
# =============================================================================

def fig_overview_kpis(van, tri, wacc_val, dscr_min, llcr, pb, devise):
    """Jauge visuelle des 3 indicateurs clés pour l'overview investisseur."""
    # DSCR gauge
    dscr_v = dscr_min if not np.isnan(dscr_min) else 0
    fig = make_subplots(
        rows=1, cols=3,
        specs=[[{"type": "indicator"}, {"type": "indicator"}, {"type": "indicator"}]],
    )
    fig.add_trace(go.Indicator(
        mode="gauge+number+delta",
        value=round(van / 1e6, 1),
        title={"text": f"VAN (M{devise})", "font": {"size": 13}},
        delta={"reference": 0, "valueformat": ".0f"},
        gauge={
            "axis"     : {"range": [min(van/1e6*2, -abs(van/1e6)), max(van/1e6*2, abs(van/1e6))]},
            "bar"      : {"color": CLR["vert"] if van > 0 else CLR["rouge"]},
            "threshold": {"line": {"color": "black", "width": 2}, "value": 0},
        },
        number={"suffix": f" M{devise}", "font": {"size": 16}},
    ), row=1, col=1)

    fig.add_trace(go.Indicator(
        mode="gauge+number",
        value=round(tri * 100, 1) if tri else 0,
        title={"text": "TRI (%)", "font": {"size": 13}},
        gauge={
            "axis" : {"range": [0, 40]},
            "bar"  : {"color": CLR["vert"] if tri and tri * 100 > wacc_val * 100 else CLR["rouge"]},
            "steps": [
                {"range": [0, wacc_val * 100], "color": "#f8d7da"},
                {"range": [wacc_val * 100, 40], "color": "#d4edda"},
            ],
            "threshold": {"line": {"color": "navy", "width": 2}, "value": wacc_val * 100},
        },
        number={"suffix": " %", "font": {"size": 16}},
    ), row=1, col=2)

    fig.add_trace(go.Indicator(
        mode="gauge+number",
        value=dscr_v,
        title={"text": "DSCR Min (x)", "font": {"size": 13}},
        gauge={
            "axis" : {"range": [0, 3]},
            "bar"  : {"color": CLR["vert"] if dscr_v >= 1.3 else (CLR["amber"] if dscr_v >= 1.0 else CLR["rouge"])},
            "steps": [
                {"range": [0, 1.0], "color": "#f8d7da"},
                {"range": [1.0, 1.3], "color": "#fff3cd"},
                {"range": [1.3, 3.0], "color": "#d4edda"},
            ],
            "threshold": {"line": {"color": "black", "width": 2}, "value": 1.3},
        },
        number={"suffix": "x", "font": {"size": 16}},
    ), row=1, col=3)

    fig.update_layout(height=240, margin=dict(l=20, r=20, t=40, b=10),
                      paper_bgcolor="white")
    return fig


def fig_fcff_bar(rows, devise, facteur):
    annees = [str(r["Annee"]) for r in rows]
    vals   = [r["FCFF"] / 1e6 * facteur for r in rows]
    cols   = [CLR["vert"] if v >= 0 else CLR["rouge"] for v in vals]
    fig = go.Figure(go.Bar(
        x=annees, y=vals, marker_color=cols,
        text=[f"{v:.1f}" for v in vals], textposition="outside",
    ))
    fig.add_hline(y=0, line_color="#333", line_width=1)
    fig.update_layout(
        title=f"FCFF Annuels (M{devise})",
        xaxis_title="Année", yaxis_title=f"M{devise}",
        plot_bgcolor="white", paper_bgcolor="white",
        height=300, margin=dict(t=40, b=30), showlegend=False,
    )
    return fig


def fig_revenus_couts(rows, devise, facteur):
    annees = [str(r["Annee"]) for r in rows]
    rev    = [r["Revenus"] / 1e6 * facteur for r in rows]
    cout   = [r["Couts"]   / 1e6 * facteur for r in rows]
    ebitda = [r["EBITDA"]  / 1e6 * facteur for r in rows]
    fig = go.Figure()
    fig.add_trace(go.Bar(x=annees, y=rev,    name="Revenus",  marker_color=CLR["bleu"], opacity=0.75))
    fig.add_trace(go.Bar(x=annees, y=cout,   name="Coûts",    marker_color=CLR["amber"], opacity=0.75))
    fig.add_trace(go.Scatter(x=annees, y=ebitda, name="EBITDA",
                             line=dict(color=CLR["vert"], width=2.5), mode="lines+markers"))
    fig.update_layout(
        title=f"Revenus / Coûts / EBITDA (M{devise})",
        barmode="group", xaxis_title="Année", yaxis_title=f"M{devise}",
        plot_bgcolor="white", paper_bgcolor="white",
        height=300, margin=dict(t=40, b=30),
        legend=dict(orientation="h", y=-0.25),
    )
    return fig


def fig_dscr(df_dette):
    df_p   = df_dette[df_dette["DSCR"].notna()]
    cols   = [CLR["vert"] if v >= 1.3 else (CLR["amber"] if v >= 1.0 else CLR["rouge"])
              for v in df_p["DSCR"]]
    fig = go.Figure(go.Bar(
        x=df_p["Annee"].astype(str), y=df_p["DSCR"],
        marker_color=cols,
        text=[f"{v:.2f}x" for v in df_p["DSCR"]], textposition="outside",
    ))
    fig.add_hline(y=1.3, line_dash="dash", line_color=CLR["vert"],
                  annotation_text="Seuil bancaire 1.3x", annotation_position="top right")
    fig.add_hline(y=1.0, line_dash="dot",  line_color=CLR["rouge"],
                  annotation_text="Défaut 1.0x", annotation_position="bottom right")
    fig.update_layout(
        title="DSCR Annuel",
        xaxis_title="Année", yaxis_title="DSCR (x)",
        plot_bgcolor="white", paper_bgcolor="white",
        height=300, margin=dict(t=40, b=30), showlegend=False,
    )
    return fig


def fig_stress_compare(stress, duree, devise, facteur):
    """
    Graphique double axe : DSCR (ligne) + Trésorerie cumulée (aire).
    3 scénarios : Base / Choc Sans Assurance / Choc Avec Assurance.
    """
    annees = list(range(1, duree + 1))
    an_choc = stress["an_choc"]

    def clean_dscr(lst):
        return [v if isinstance(v, float) and not np.isnan(v) else None for v in lst]

    def treso(rows):
        return [r["Tresorerie_cum"] / 1e6 * facteur for r in rows]

    fig = make_subplots(
        rows=2, cols=1,
        subplot_titles=("Impact sur le DSCR", "Trésorerie Cumulée"),
        vertical_spacing=0.15,
    )

    # — DSCR —
    fig.add_trace(go.Scatter(
        x=annees, y=clean_dscr(stress["base"]["dscr"]),
        name="Base", mode="lines+markers",
        line=dict(color=CLR["bleu"], width=2),
    ), row=1, col=1)
    fig.add_trace(go.Scatter(
        x=annees, y=clean_dscr(stress["choc"]["dscr"]),
        name="Choc — Sans Assurance", mode="lines+markers",
        line=dict(color=CLR["rouge"], width=2.5, dash="dot"),
        marker=dict(symbol="x", size=8),
    ), row=1, col=1)
    fig.add_trace(go.Scatter(
        x=annees, y=clean_dscr(stress["assurance"]["dscr"]),
        name="Choc — Avec Assurance", mode="lines+markers",
        line=dict(color=CLR["vert"], width=2.5),
        marker=dict(symbol="diamond", size=7),
    ), row=1, col=1)

    # Seuils DSCR
    fig.add_hline(y=1.3, line_dash="dash", line_color=CLR["vert"],
                  annotation_text="1.3x", row=1, col=1)
    fig.add_hline(y=1.0, line_dash="dot",  line_color=CLR["rouge"],
                  annotation_text="1.0x", row=1, col=1)

    # — Trésorerie —
    fig.add_trace(go.Scatter(
        x=annees, y=treso(stress["base"]["rows"]),
        name="Base", mode="lines",
        line=dict(color=CLR["bleu"], width=2),
        fill="tozeroy", fillcolor="rgba(13,110,253,0.07)",
        showlegend=False,
    ), row=2, col=1)
    fig.add_trace(go.Scatter(
        x=annees, y=treso(stress["choc"]["rows"]),
        name="Sans Assurance", mode="lines",
        line=dict(color=CLR["rouge"], width=2.5, dash="dot"),
        fill="tozeroy", fillcolor="rgba(220,53,69,0.07)",
        showlegend=False,
    ), row=2, col=1)
    fig.add_trace(go.Scatter(
        x=annees, y=treso(stress["assurance"]["rows"]),
        name="Avec Assurance", mode="lines",
        line=dict(color=CLR["vert"], width=2.5),
        fill="tozeroy", fillcolor="rgba(25,135,84,0.07)",
        showlegend=False,
    ), row=2, col=1)

    fig.add_hline(y=0, line_color="#333", line_width=0.8, row=2, col=1)

    # Zone choc
    for row_n in [1, 2]:
        fig.add_vrect(
            x0=an_choc - 0.45, x1=an_choc + 0.45,
            fillcolor="rgba(220,53,69,0.10)", line_width=0,
            annotation_text="Choc", annotation_position="top left",
            row=row_n, col=1,
        )

    fig.update_layout(
        height=580,
        plot_bgcolor="white", paper_bgcolor="white",
        legend=dict(orientation="h", y=-0.12, x=0),
        margin=dict(t=50, b=40),
        yaxis2_title=f"M{devise}",
        yaxis_title="DSCR (x)",
    )
    return fig


def fig_tornado_plotly(tdata, van_ref, devise, facteur):
    labels = [d["Variable"] for d in tdata]
    im     = [d["im"] / 1e6 * facteur for d in tdata]
    ip     = [d["ip"] / 1e6 * facteur for d in tdata]
    fig = go.Figure()
    fig.add_trace(go.Bar(y=labels, x=im, orientation="h", name="-10%",
                         marker_color=CLR["rouge"],
                         text=[f"{v:+,.0f}" for v in im], textposition="outside"))
    fig.add_trace(go.Bar(y=labels, x=ip, orientation="h", name="+10%",
                         marker_color=CLR["vert"],
                         text=[f"{v:+,.0f}" for v in ip], textposition="outside"))
    fig.add_vline(x=0, line_color="#333", line_width=1)
    fig.update_layout(
        title=f"Tornado — Sensibilité VAN (base: {van_ref/1e6*facteur:,.0f} M{devise})",
        xaxis_title=f"Impact sur VAN (M{devise})", barmode="overlay",
        plot_bgcolor="white", paper_bgcolor="white",
        height=300, margin=dict(t=40, b=30),
        legend=dict(orientation="h", y=-0.25),
    )
    return fig


def fig_heatmap(df_s, devise):
    fig = go.Figure(go.Heatmap(
        z=df_s.values.tolist(), x=df_s.columns.tolist(), y=df_s.index.tolist(),
        colorscale=[[0, CLR["rouge"]], [0.5, "#ffffff"], [1, CLR["vert"]]],
        text=[[f"{v:.0f}" for v in row] for row in df_s.values],
        texttemplate="%{text}", showscale=True,
        colorbar=dict(title=f"VAN M{devise}"),
    ))
    fig.update_layout(
        title=f"Sensibilité Croisée VAN — Prix × Coût (M{devise})",
        xaxis_title="Variation Coût", yaxis_title="Variation Prix",
        height=320, margin=dict(t=40),
    )
    return fig


# =============================================================================
# PAGE CONFIG & CSS
# =============================================================================

st.set_page_config(
    page_title="CommodityWatch — Overview Investisseur",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    /* Global */
    body { font-family: 'Inter', sans-serif; }
    .block-container { padding: 1.4rem 2.5rem 2rem 2.5rem; }

    /* Titres */
    h1  { font-size:1.5rem; font-weight:700; color:#0d1117; margin-bottom:.2rem; }
    h2  { font-size:1.05rem; font-weight:600; color:#24292f;
          border-bottom:2px solid #e8e8e8; padding-bottom:4px; margin-top:1.5rem; }
    h3  { font-size:.9rem; font-weight:600; color:#24292f; }

    /* KPI cards */
    .kpi-card {
        background:#ffffff; border:1px solid #e0e0e0; border-radius:8px;
        padding:14px 18px; text-align:center;
    }
    .kpi-val  { font-size:1.5rem; font-weight:700; line-height:1.2; }
    .kpi-lab  { font-size:.72rem; color:#57606a; text-transform:uppercase;
                letter-spacing:.05em; margin-top:2px; }
    .kpi-delta{ font-size:.78rem; margin-top:4px; font-weight:500; }
    .pos { color:#198754; } .neg { color:#dc3545; } .neu { color:#6c757d; }

    /* Bandeau Business Model */
    .bm-banner {
        background:linear-gradient(90deg,#0d6efd 0%,#0a58ca 100%);
        color:white; border-radius:8px; padding:10px 20px;
        font-size:.92rem; font-weight:600; margin-bottom:1rem;
    }

    /* Sections */
    .section-divider { border:none; border-top:1px solid #e8e8e8; margin:1.2rem 0; }

    /* Sidebar */
    .stMetric label { font-size:.75rem; color:#57606a; }
    hr { border:none; border-top:1px solid #ddd; margin:12px 0; }
</style>
""", unsafe_allow_html=True)


# =============================================================================
# BARRE LATÉRALE — INPUTS ADAPTATIFS
# =============================================================================

with st.sidebar:
    st.markdown("### CommodityWatch")
    st.caption("Overview Pré-Investissement v4.0")
    st.markdown("---")

    # ── Devise ────────────────────────────────────────────────────────────────
    devise  = st.radio("Devise", ["FCFA", "USD"], horizontal=True)
    facteur = 1.0 if devise == "FCFA" else 1 / TAUX_CHANGE

    # ── Business Model ────────────────────────────────────────────────────────
    st.markdown("#### Type de Projet")
    bm = st.selectbox(
        "Business Model",
        ["Production", "Infrastructure", "Campagne"],
        help=(
            "Production : plantation/élevage — revenu lié à la récolte\n"
            "Infrastructure : construction/extension d'usine\n"
            "Campagne : achat matière, transformation, vente"
        ),
    )
    bm_info = BM_RISQUE[bm]

    st.markdown("---")

    # ── Matière & Prix ────────────────────────────────────────────────────────
    st.markdown("#### Produit & Prix")
    matiere = st.selectbox("Matière première", list(PRIX_REF_USD.keys()))
    prix_ref_devise = (PRIX_REF_USD[matiere] if devise == "USD"
                       else PRIX_REF_USD[matiere] * TAUX_CHANGE)
    if matiere != "Personnalisé" and prix_ref_devise > 0:
        st.info(f"Réf. marché : **{prix_ref_devise:,.0f} {devise}/T** (ICE/Euronext/OTC)")

    prix_input = st.number_input(
        f"Prix de Vente ({devise}/T)",
        min_value=0.0,
        value=float(prix_ref_devise) if prix_ref_devise > 0 else (1_500.0 if devise == "USD" else 900_000.0),
        step=10.0 if devise == "USD" else 10_000.0,
        format="%.1f" if devise == "USD" else "%.0f",
    )
    prix_fcfa = en_fcfa(prix_input, devise)

    quantite = st.number_input("Volume pleine capacité (T/an)", 0, 500_000, 5_000, 100)

    cout_input = st.number_input(
        f"Coût de Production ({devise}/T)",
        min_value=0.0,
        value=1_000.0 if devise == "USD" else 600_000.0,
        step=10.0 if devise == "USD" else 10_000.0,
        format="%.1f" if devise == "USD" else "%.0f",
    )
    cout_fcfa = en_fcfa(cout_input, devise)

    st.markdown("---")

    # ── Paramètres propres au Business Model ─────────────────────────────────
    duree_bio   = 0
    j_stock = j_clients = j_fourn = 0

    if bm == "Production":
        st.markdown("#### Phase Biologique")
        duree_bio = st.number_input("Années sans revenu (croissance)", 0, 10, 3, 1)

    if bm == "Campagne":
        st.markdown("#### Cycle d'Exploitation (BFR)")
        j_stock   = st.number_input("Stock matière (jours)", 0, 180, 60, 5)
        j_clients = st.number_input("Crédit clients (jours)",  0, 120, 30, 5)
        j_fourn   = st.number_input("Crédit fournisseurs (j)", 0, 120, 45, 5)

    st.markdown("---")

    # ── Ramp-up ───────────────────────────────────────────────────────────────
    st.markdown("#### Montée en Puissance")
    n_pal = int(st.number_input("Nombre de paliers", 1, 5, 3, 1))
    d_an  = [2, 3, 4, 5, 6]
    d_tx  = [0, 40, 80, 100, 100]
    paliers = []
    for k in range(n_pal):
        c1, c2 = st.columns(2)
        with c1:
            af = st.number_input(f"Fin an {k+1}", 1, 25, d_an[k], 1, key=f"an{k}")
        with c2:
            tx = st.number_input("Taux %", 0, 100, d_tx[k], 5, key=f"tx{k}")
        paliers.append((int(af), tx / 100))

    st.markdown("---")

    # ── Capital ───────────────────────────────────────────────────────────────
    st.markdown("#### Structure Financière")
    capex_input = st.number_input(
        f"CAPEX Total ({devise})",
        min_value=0.0,
        value=3_000_000.0 if devise == "USD" else 2_000_000_000.0,
        step=100_000.0 if devise == "USD" else 100_000_000.0,
        format="%.0f",
    )
    capex_fcfa = en_fcfa(capex_input, devise)

    p_dette_pct = st.slider("Part dette (%)", 0, 100, 60, 5)
    p_dette     = p_dette_pct / 100
    p_fp        = 1 - p_dette
    r_dette_pct = st.slider("Coût dette (%/an)", 1, 25, 9, 1)
    r_dette     = r_dette_pct / 100
    r_fp_pct    = st.slider("Coût fonds propres (%)", 5, 30, 15, 1)
    r_fp        = r_fp_pct / 100
    taux_is_pct = st.slider("IS (%)", 0, 40, 25, 1)
    taux_is     = taux_is_pct / 100
    duree       = st.slider("Durée projet (ans)", 3, 25, 10, 1)
    infl_pct    = st.slider("Inflation annuelle (%)", 0, 15, 3, 1)
    inflation   = infl_pct / 100

    st.markdown("---")

    # ── Dette avancée ─────────────────────────────────────────────────────────
    st.markdown("#### Paramètres du Prêt")
    dur_dette   = st.slider("Durée du prêt (ans)", 1, duree, min(8, duree), 1)
    grace       = st.slider("Différé remboursement (ans)", 0, max(0, dur_dette-1),
                             min(2, dur_dette-1), 1)
    with st.expander("Commissions Bancaires"):
        c_flat_pct    = st.slider("Commission Flat (%, prélevée une fois)", 0.0, 3.0, 0.5, 0.1)
        c_engagt_pct  = st.slider("Commission Engagement (%/an)", 0.0, 2.0, 0.25, 0.05)
    c_flat    = c_flat_pct   / 100
    c_engagt  = c_engagt_pct / 100
    profil_dec = [0.5, 0.5] + [0.0] * (duree - 2)   # décaissement 50/50 an1-an2

    st.markdown("---")

    # ── Assurance (contextuelle) ──────────────────────────────────────────────
    st.markdown(f"#### {bm_info['assurance_label']}")
    st.caption(bm_info["description_assurance"])

    prob_risque  = st.slider(
        "Probabilité d'occurrence (%)", 1, 50,
        int(bm_info["gravite_defaut"] * 100), 1,
    ) / 100

    gravite_risque = st.slider(
        "Gravité — perte de CA (%)", 5, 80,
        int(bm_info["gravite_defaut"] * 100), 5,
    ) / 100

    prime_input = st.number_input(
        f"Prime annuelle ({devise})",
        min_value=0.0,
        value=50_000.0 if devise == "USD" else 30_000_000.0,
        step=5_000.0 if devise == "USD" else 5_000_000.0,
        format="%.0f",
    )
    prime_fcfa = en_fcfa(prime_input, devise)

    indem_pct = st.slider("Taux de couverture (% perte indemnisé)", 0, 100, 70, 5) / 100

    st.markdown("---")

    # ── Import Excel ──────────────────────────────────────────────────────────
    with st.expander("Importer Bilan / CdR Excel"):
        uploaded = st.file_uploader("Fichier .xlsx", type=["xlsx"])
        donnees_xl = {}
        if uploaded:
            if EXCEL_OK:
                donnees_xl = parser_excel(uploaded)
                if "erreur" in donnees_xl:
                    st.error(donnees_xl["erreur"])
                else:
                    st.success(f"{len(donnees_xl)} postes importés")
                    st.json(donnees_xl)
            else:
                st.warning("Installez `openpyxl` pour activer l'import Excel.")

    st.markdown("---")
    lancer = st.button("Générer l'Overview", type="primary", use_container_width=True)


# =============================================================================
# ZONE PRINCIPALE
# =============================================================================

# ── Header ────────────────────────────────────────────────────────────────────
bm_icons = {"Production": "🌱", "Infrastructure": "🏭", "Campagne": "🔄"}
st.markdown(
    f"<div class='bm-banner'>{bm_icons[bm]}  CommodityWatch — "
    f"{matiere} &nbsp;|&nbsp; {bm} &nbsp;|&nbsp; {devise}</div>",
    unsafe_allow_html=True,
)
st.caption(
    "Overview pré-investissement à destination du promoteur et du comité de crédit bancaire. "
    "Résultats indicatifs — ne se substituent pas à une due diligence complète."
)

if not lancer:
    # ── Écran d'accueil ───────────────────────────────────────────────────────
    st.markdown("---")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("""
**Ce que vous allez obtenir :**
- Indicateurs clés (VAN, TRI, DSCR, LLCR, Payback)
- Projection des flux sur toute la durée du projet
- Tableau de service de la dette avec alertes bancaires
""")
    with c2:
        st.markdown(f"""
**Analyse de sensibilité :**
- Tornado Chart (4 variables clés)
- Tableau croisé Prix × Coût
- Identification du risque dominant
""")
    with c3:
        st.markdown(f"""
**Stress Test — {bm} :**
- Risque : {bm_info['risque_label']}
- Assurance : {bm_info['assurance_label']}
- Comparatif Sans / Avec Assurance
""")
    st.info("Configurez les paramètres dans la barre latérale puis cliquez sur **Générer l'Overview**.")
    st.stop()


# =============================================================================
# CALCULS
# =============================================================================

with st.spinner("Calcul en cours..."):
    wacc_val   = r_dette * (1 - taux_is) * p_dette + r_fp * p_fp
    rampup     = build_rampup(duree, paliers)
    fcff_rows  = projeter_fcff(quantite, prix_fcfa, cout_fcfa, capex_fcfa,
                               taux_is, inflation, duree, rampup, bm,
                               duree_bio, j_stock, j_clients, j_fourn)
    met        = metriques_rentabilite(fcff_rows, capex_fcfa, wacc_val)
    df_dette, dscr_min, dscr_moy, llcr_g, cout_total_dette = service_dette(
        fcff_rows, capex_fcfa, p_dette, r_dette, dur_dette,
        grace, wacc_val, c_flat, c_engagt, profil_dec,
    )
    alertes_chg = [r["Annee"] for r in fcff_rows
                   if not np.isnan(r["Ratio Charges"]) and r["Ratio Charges"] > SEUIL_CHARGES_CA]
    tornado_d, van_ref_t = calc_tornado(
        quantite, prix_fcfa, cout_fcfa, capex_fcfa, taux_is, inflation,
        duree, rampup, bm, wacc_val, duree_bio, j_stock, j_clients, j_fourn,
    )
    df_sens = calc_sensibilite(
        quantite, prix_fcfa, cout_fcfa, capex_fcfa, taux_is, inflation,
        duree, rampup, bm, wacc_val, duree_bio, j_stock, j_clients, j_fourn,
    ) * facteur / 1e6

    stress = stress_test(
        fcff_rows, bm, prob_risque, gravite_risque,
        prime_fcfa, indem_pct,
        capex_fcfa, p_dette, r_dette, dur_dette,
        grace, wacc_val, c_flat, c_engagt, profil_dec,
    )


# =============================================================================
# SECTION A — OVERVIEW JAUGES
# =============================================================================

van_v   = met["van"]
tri_v   = met["tri"]
pb_v    = met["payback"]

st.markdown("## Vue d'Ensemble — Indicateurs Clés")
st.plotly_chart(
    fig_overview_kpis(van_v, tri_v, wacc_val, dscr_min, llcr_g, pb_v, devise),
    use_container_width=True,
)

# ── KPIs textuels ─────────────────────────────────────────────────────────────
k1, k2, k3, k4, k5, k6 = st.columns(6)

def kpi_card(col, valeur, label, delta_txt, delta_ok):
    col.markdown(
        f"<div class='kpi-card'>"
        f"<div class='kpi-val'>{valeur}</div>"
        f"<div class='kpi-lab'>{label}</div>"
        f"<div class='kpi-delta {'pos' if delta_ok else 'neg'}'>{delta_txt}</div>"
        f"</div>",
        unsafe_allow_html=True,
    )

van_aff  = fmt(van_v  / 1e6 * facteur, devise)
tri_aff  = f"{tri_v*100:.2f} %" if tri_v else "N/D"
pb_aff   = f"{pb_v} ans" if pb_v else "> Durée"
llcr_aff = f"{llcr_g:.2f}x" if not np.isnan(llcr_g) else "N/D"
capex_aff = fmt(capex_fcfa / 1e6 * facteur, devise, dec=0)
cout_d_aff = fmt(cout_total_dette / 1e6 * facteur, devise)

kpi_card(k1, van_aff,   "VAN",             "Positive" if van_v > 0 else "Négative",   van_v > 0)
kpi_card(k2, tri_aff,   "TRI",             f"vs WACC {wacc_val*100:.2f}%", (tri_v or 0) > wacc_val)
kpi_card(k3, f"{dscr_min:.2f}x", "DSCR Min", "Bancable ≥ 1.3x" if dscr_min >= 1.3 else "< Seuil 1.3x", dscr_min >= 1.3)
kpi_card(k4, llcr_aff,  "LLCR",            "≥ 1.1x" if not np.isnan(llcr_g) and llcr_g >= 1.1 else "< 1.1x", not np.isnan(llcr_g) and llcr_g >= 1.1)
kpi_card(k5, pb_aff,    "Payback",         f"CAPEX {capex_aff}", True)
kpi_card(k6, cout_d_aff,"Coût Total Dette", "Commissions incluses", True)

st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

# ── Alertes immédiates ────────────────────────────────────────────────────────
if alertes_chg:
    st.warning(
        f"**Alerte Charges / CA > {SEUIL_CHARGES_CA*100:.0f}%** "
        f"aux années {alertes_chg}. Les coûts opérationnels pèsent trop sur le chiffre d'affaires.",
        icon="⚠️",
    )
if not np.isnan(dscr_min) and dscr_min < SEUIL_DSCR_ALERTE:
    lvl = "error" if dscr_min < SEUIL_DSCR_DEFAUT else "warning"
    msg = (f"**DSCR Min = {dscr_min:.2f}x** — "
           + ("Risque de défaut de paiement." if dscr_min < 1.0
              else "En dessous du seuil bancaire 1.3x. Renégocier la structure de dette."))
    if lvl == "error":
        st.error(msg, icon="🚨")
    else:
        st.warning(msg, icon="⚠️")


# =============================================================================
# SECTION B — REVENUS / COÛTS / FCFF
# =============================================================================

st.markdown("## Projection Financière")
tab_graph, tab_table = st.tabs(["Graphiques", "Tableau Détaillé"])

with tab_graph:
    g1, g2 = st.columns(2)
    with g1:
        st.plotly_chart(fig_revenus_couts(fcff_rows, devise, facteur), use_container_width=True)
    with g2:
        st.plotly_chart(fig_fcff_bar(fcff_rows, devise, facteur), use_container_width=True)

with tab_table:
    # Construction du tableau d'affichage
    df_fcff = pd.DataFrame([{
        "An"            : r["Annee"],
        "Ramp-up"       : f"{r['Ramp-up']:.0f}%",
        f"Revenus M{devise}"  : round(r["Revenus"]   / 1e6 * facteur, 2),
        f"Coûts M{devise}"    : round(r["Couts"]     / 1e6 * facteur, 2),
        f"EBITDA M{devise}"   : round(r["EBITDA"]    / 1e6 * facteur, 2),
        "Chg/CA"        : f"{r['Ratio Charges']*100:.1f}%" if not np.isnan(r["Ratio Charges"]) else "—",
        f"FCFF M{devise}"     : round(r["FCFF"]      / 1e6 * facteur, 2),
    } for r in fcff_rows])

    def style_fcff_row(row):
        fcff_col = f"FCFF M{devise}"
        chg_col  = "Chg/CA"
        fcff_v   = row.get(fcff_col, 0)
        chg_s    = row.get(chg_col, "0%")
        try:
            chg_v = float(str(chg_s).replace("%", "")) / 100
        except:
            chg_v = 0
        if fcff_v < 0:
            return ["background:#f8d7da"] * len(row)
        if chg_v > SEUIL_CHARGES_CA:
            return ["background:#fff3cd"] * len(row)
        return [""] * len(row)

    st.dataframe(
        df_fcff.style.apply(style_fcff_row, axis=1),
        use_container_width=True, hide_index=True,
    )
    st.caption("🟡 Charges > 55% CA &nbsp;&nbsp; 🔴 FCFF négatif")

st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)


# =============================================================================
# SECTION C — SERVICE DE LA DETTE
# =============================================================================

st.markdown("## Bancabilité — Service de la Dette & DSCR")
st.caption(
    f"Grace Period : {grace} an(s) | Durée prêt : {dur_dette} ans | "
    f"Comm. Flat : {c_flat_pct:.1f}% | Comm. Engagement : {c_engagt_pct:.2f}%/an"
)

cd1, cd2 = st.columns([1, 1])

with cd1:
    cols_show = [c for c in df_dette.columns if c != "Grace"]
    df_d_aff  = df_dette[cols_show].copy()
    for col_m in ["Encours", "Interets", "Remb Capital", "Comm Engagt", "Service", "FCFF"]:
        df_d_aff[col_m] = (df_dette[col_m] / 1e6 * facteur).round(2)

    def style_dette_row(row):
        an       = row["Annee"]
        is_grace = df_dette.loc[df_dette["Annee"] == an, "Grace"].values
        is_grace = bool(is_grace[0]) if len(is_grace) else False
        if is_grace:
            return ["background:#fff3cd;color:#856404"] * len(row)
        dscr_v = row.get("DSCR", np.nan)
        if pd.notna(dscr_v):
            if dscr_v < 1.0:  return ["background:#f8d7da;color:#721c24"] * len(row)
            if dscr_v < 1.3:  return ["background:#fde8c8;color:#7d4e0f"] * len(row)
            return ["background:#d4edda;color:#155724"] * len(row)
        return [""] * len(row)

    st.dataframe(
        df_d_aff.style.apply(style_dette_row, axis=1).format({
            "DSCR": lambda x: f"{x:.2f}x" if pd.notna(x) else "—",
            "LLCR": lambda x: f"{x:.2f}x" if pd.notna(x) else "—",
            **{c: "{:.2f}" for c in ["Encours","Interets","Remb Capital",
                                      "Comm Engagt","Service","FCFF"]},
        }),
        use_container_width=True, hide_index=True,
    )
    la, lb, lc, ld = st.columns(4)
    la.markdown("<span style='background:#d4edda;padding:1px 6px;border-radius:3px;font-size:.75rem'>≥ 1.3x</span>", unsafe_allow_html=True)
    lb.markdown("<span style='background:#fde8c8;padding:1px 6px;border-radius:3px;font-size:.75rem'>1.0–1.3x</span>", unsafe_allow_html=True)
    lc.markdown("<span style='background:#f8d7da;padding:1px 6px;border-radius:3px;font-size:.75rem'>< 1.0x Défaut</span>", unsafe_allow_html=True)
    ld.markdown("<span style='background:#fff3cd;padding:1px 6px;border-radius:3px;font-size:.75rem'>Grace Period</span>", unsafe_allow_html=True)

with cd2:
    st.plotly_chart(fig_dscr(df_dette), use_container_width=True)

st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)


# =============================================================================
# SECTION D — SENSIBILITÉ
# =============================================================================

st.markdown("## Analyse de Sensibilité")
st_tab1, st_tab2 = st.tabs(["Tornado Chart", "Tableau Croisé Prix × Coût"])

with st_tab1:
    st_c1, st_c2 = st.columns([2, 1])
    with st_c1:
        st.plotly_chart(
            fig_tornado_plotly(tornado_d, van_ref_t, devise, facteur),
            use_container_width=True,
        )
    with st_c2:
        st.markdown(f"#### Impacts sur VAN (M{devise})")
        df_t = pd.DataFrame([{
            "Variable"   : d["Variable"],
            "-10%"       : f"{d['im']/1e6*facteur:+,.0f}",
            "+10%"       : f"{d['ip']/1e6*facteur:+,.0f}",
            "Amplitude"  : f"{d['amp']/1e6*facteur:,.0f}",
        } for d in tornado_d])
        st.dataframe(df_t, use_container_width=True, hide_index=True)
        st.caption("Variable la plus impactante = levier de négociation prioritaire.")

with st_tab2:
    st.plotly_chart(fig_heatmap(df_sens, devise), use_container_width=True)

st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)


# =============================================================================
# SECTION E — STRESS TEST CONTEXTUEL & ASSURANCE
# =============================================================================

st.markdown(f"## Stress Test — {bm_info['risque_label']}")

# ── Contexte ──────────────────────────────────────────────────────────────────
st.info(
    f"**Risque identifié :** {bm_info['description_risque']}\n\n"
    f"**Couverture proposée :** {bm_info['assurance_label']} — "
    f"{bm_info['description_assurance']}",
    icon="ℹ️",
)

# ── Synthèse financière du stress ─────────────────────────────────────────────
perte_aff   = fmt(stress["perte_brute"]  / 1e6 * facteur, devise)
indem_aff   = fmt(stress["indemnite"]    / 1e6 * facteur, devise)
primes_aff  = fmt(stress["cout_primes"]  / 1e6 * facteur, devise)
gain_aff    = fmt(abs(stress["gain_net_ass"]) / 1e6 * facteur, devise)
gain_pos    = stress["gain_net_ass"] > 0

sm1, sm2, sm3, sm4, sm5 = st.columns(5)
sm1.metric("Année du Choc",     f"Année {stress['an_choc']}")
sm2.metric("Perte CA (choc)",   perte_aff, delta=f"Gravité {gravite_risque*100:.0f}%", delta_color="inverse")
sm3.metric("Indemnité Assurance",indem_aff, delta=f"Couv. {indem_pct*100:.0f}%", delta_color="normal")
sm4.metric("Coût Total Primes", primes_aff, delta=f"{duree} années", delta_color="off")
sm5.metric("Gain Net Assurance", gain_aff,
           delta="Assurance rentable" if gain_pos else "Assurance coûteuse",
           delta_color="normal" if gain_pos else "inverse")

# ── DSCR Min comparatif ───────────────────────────────────────────────────────
d_base  = stress["base"]["dscr_min"]
d_choc  = stress["choc"]["dscr_min"]
d_assur = stress["assurance"]["dscr_min"]

sc1, sc2, sc3 = st.columns(3)
sc1.metric("DSCR Min — Base",       f"{d_base:.2f}x"  if not np.isnan(d_base)  else "N/D",
           delta="Référence", delta_color="off")
sc2.metric("DSCR Min — Sans Assurance", f"{d_choc:.2f}x" if not np.isnan(d_choc) else "N/D",
           delta=f"{d_choc - d_base:+.2f}x vs Base" if not np.isnan(d_choc) else "",
           delta_color="normal" if d_choc >= d_base else "inverse")
sc3.metric("DSCR Min — Avec Assurance", f"{d_assur:.2f}x" if not np.isnan(d_assur) else "N/D",
           delta=f"{d_assur - d_choc:+.2f}x vs Choc" if not np.isnan(d_assur) else "",
           delta_color="normal" if d_assur >= d_choc else "inverse")

# ── Graphique principal ───────────────────────────────────────────────────────
st.plotly_chart(
    fig_stress_compare(stress, duree, devise, facteur),
    use_container_width=True,
)

# ── Interprétation automatique ────────────────────────────────────────────────
st.markdown("#### Lecture du Résultat")
interp_lines = []

if not np.isnan(d_choc) and d_choc < SEUIL_DSCR_DEFAUT:
    interp_lines.append(
        f"Sans assurance, le choc ramène le DSCR à **{d_choc:.2f}x**, "
        "en dessous du seuil de défaut (1.0x). Le projet ne peut pas honorer sa dette cette année-là."
    )
elif not np.isnan(d_choc) and d_choc < SEUIL_DSCR_ALERTE:
    interp_lines.append(
        f"Sans assurance, le DSCR descend à **{d_choc:.2f}x**, "
        "en dessous du seuil bancaire (1.3x). La banque exigera probablement un covenant de réserve."
    )
else:
    interp_lines.append(
        f"Même sans assurance, le DSCR reste à **{d_choc:.2f}x** — "
        "le projet absorbe le choc sans défaut. La couverture d'assurance est une précaution supplémentaire."
    )

if not np.isnan(d_assur):
    if d_assur >= SEUIL_DSCR_ALERTE:
        interp_lines.append(
            f"Avec l'{bm_info['assurance_label']}, le DSCR remonte à **{d_assur:.2f}x** "
            f"(au-dessus du seuil 1.3x). L'assurance restaure la bancabilité du projet."
        )
    else:
        interp_lines.append(
            f"Avec assurance, le DSCR est à **{d_assur:.2f}x** — "
            "l'assurance améliore la situation mais ne suffit pas seule à restaurer le seuil bancaire. "
            "Envisager une reserve de liquidité complémentaire."
        )

if gain_pos:
    interp_lines.append(
        f"Financièrement, l'assurance est **rentable** sur la durée du projet : "
        f"indemnité attendue ({indem_aff}) > coût total des primes ({primes_aff}), "
        f"soit un gain net de {gain_aff}."
    )
else:
    interp_lines.append(
        f"Les primes cumulées ({primes_aff}) dépassent l'indemnité unique ({indem_aff}). "
        "L'assurance est un coût de tranquillité pour la banque, pas un gain financier net — "
        "ce qui est normal pour des sinistres rares."
    )

for line in interp_lines:
    st.markdown(f"- {line}")

st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)


# =============================================================================
# SECTION F — FICHE SYNTHÈSE (RÉSUMÉ EXÉCUTIF)
# =============================================================================

st.markdown("## Fiche Synthèse — Résumé Exécutif")

verdict_van  = "FAVORABLE" if van_v > 0 else "DÉFAVORABLE"
verdict_tri  = "FAVORABLE" if (tri_v or 0) > wacc_val else "INSUFFISANT"
verdict_dscr = "BANCABLE" if dscr_min >= 1.3 else ("À RISQUE" if dscr_min >= 1.0 else "EN DÉFAUT")
score = sum([van_v > 0, (tri_v or 0) > wacc_val, dscr_min >= 1.3, not np.isnan(llcr_g) and llcr_g >= 1.1])

couleur_score = CLR["vert"] if score >= 3 else (CLR["amber"] if score == 2 else CLR["rouge"])
verdict_global = "PROJET VIABLE" if score >= 3 else ("PROJET À AMÉLIORER" if score == 2 else "PROJET RISQUÉ")

st.markdown(
    f"<div style='background:{couleur_score};color:white;border-radius:8px;"
    f"padding:14px 24px;font-size:1.1rem;font-weight:700;text-align:center;'>"
    f"Score : {score}/4 — {verdict_global}</div>",
    unsafe_allow_html=True,
)

st.markdown("")

f1, f2 = st.columns(2)
with f1:
    st.markdown(f"""
**Rentabilité**
| Indicateur | Valeur | Verdict |
|---|---|---|
| VAN | {van_aff} | {verdict_van} |
| TRI | {tri_aff} | {verdict_tri} |
| WACC | {wacc_val*100:.2f}% | Taux plancher |
| Payback | {pb_aff} | — |

**Structure Financière**
| Poste | Valeur |
|---|---|
| CAPEX Total | {capex_aff} |
| Dette ({p_dette_pct}%) | {fmt(capex_fcfa*p_dette/1e6*facteur, devise, dec=0)} |
| Fonds Propres ({100-p_dette_pct}%) | {fmt(capex_fcfa*p_fp/1e6*facteur, devise, dec=0)} |
| Coût Total Dette | {cout_d_aff} |
""")

with f2:
    st.markdown(f"""
**Bancabilité**
| Indicateur | Valeur | Verdict |
|---|---|---|
| DSCR Min | {dscr_min:.2f}x | {verdict_dscr} |
| DSCR Moyen | {dscr_moy:.2f}x | — |
| LLCR | {llcr_aff} | — |
| Grace Period | {grace} an(s) | — |

**Risque & Assurance ({bm})**
| Élément | Valeur |
|---|---|
| Risque identifié | {bm_info['risque_label'][:40]}... |
| DSCR sans assurance (choc) | {d_choc:.2f}x |
| DSCR avec assurance (choc) | {d_assur:.2f}x |
| Couverture recommandée | {bm_info['assurance_label'][:35]} |
""")

st.markdown("---")
st.markdown(
    f"<div style='font-size:.7rem;color:#6c757d;text-align:center'>"
    f"CommodityWatch v4.0 — 1 USD = {TAUX_CHANGE:.0f} FCFA — "
    "Document indicatif, ne se substitue pas à une due diligence financière et juridique complète."
    "</div>",
    unsafe_allow_html=True,
)
