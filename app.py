# =============================================================================
# CommodityWatch v4.1 — Overview Pré-Investissement
# Framework : Streamlit + Plotly | Python 3.9+
# Usage     : streamlit run app.py
#
# Philosophie : outil de première impression — l'analyste, le banquier ou le
#               promoteur obtient en 5 minutes une lecture complète du projet
#               avant toute modélisation Excel approfondie.
#
# Architecture :
#   1 Business Model = 1 logique financière = 1 risque = 1 assurance
#
#   Production    → Plantation/Élevage   → Assurance Récolte
#   Infrastructure→ Construction/Usine   → Assurance Tous Risques Chantier
#   Campagne      → Transformation/Vente → Assurance Qualité / Rappel Produit
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

TAUX_CHANGE = 600.0   # 1 USD = 600 FCFA

PRIX_REF_USD = {
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

# Un seul risque par Business Model — logique métier ancrée dans le code
BM_CONFIG = {
    "Production": {
        "icon"               : "🌱",
        "description"        : "Plantation / Élevage — revenu lié à la récolte",
        "risque_titre"       : "Risque Récolte",
        "risque_desc"        : (
            "Sécheresse, inondation, épidémie phytosanitaire ou ravageurs "
            "pouvant réduire le rendement de façon significative."
        ),
        "assurance_titre"    : "Assurance Agricole / Récolte",
        "assurance_desc"     : (
            "Indemnise une fraction de la perte de production constatée. "
            "La prime est une charge annuelle d'exploitation."
        ),
        "gravite_defaut"     : 0.40,
        "bfr_actif"          : False,
        "bio_actif"          : True,
    },
    "Infrastructure": {
        "icon"               : "🏭",
        "description"        : "Construction / Extension d'usine — CAPEX dominant",
        "risque_titre"       : "Risque Opérationnel (Panne / Chantier)",
        "risque_desc"        : (
            "Panne majeure d'équipement, accident de chantier, incendie "
            "réduisant la capacité de production installée."
        ),
        "assurance_titre"    : "Assurance TRC & Bris de Machine",
        "assurance_desc"     : (
            "Couvre les dommages matériels et la perte d'exploitation "
            "consécutive à un sinistre sur l'outil industriel."
        ),
        "gravite_defaut"     : 0.25,
        "bfr_actif"          : False,
        "bio_actif"          : False,
    },
    "Campagne": {
        "icon"               : "🔄",
        "description"        : "Achat matière — Transformation — Vente",
        "risque_titre"       : "Risque Qualité / Non-conformité",
        "risque_desc"        : (
            "Lot non conforme aux normes export, perte de certification, "
            "rappel produit ou embargo à l'exportation."
        ),
        "assurance_titre"    : "Assurance Qualité & Responsabilité Produit",
        "assurance_desc"     : (
            "Couvre les pertes de CA liées au retrait de lot, aux pénalités "
            "contractuelles et aux frais de rappel ou de décontamination."
        ),
        "gravite_defaut"     : 0.20,
        "bfr_actif"          : True,
        "bio_actif"          : False,
    },
}

SEUIL_CHARGES  = 0.55   # Alerte charges / CA
SEUIL_BANCAIRE = 1.30   # DSCR seuil standard
SEUIL_DEFAUT   = 1.00

CLR = {
    "vert"  : "#198754",
    "rouge" : "#dc3545",
    "amber" : "#fd7e14",
    "bleu"  : "#0d6efd",
    "gris"  : "#6c757d",
}


# =============================================================================
# UTILITAIRES MATHÉMATIQUES
# =============================================================================

def npv_calc(rate, cashflows):
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


def fcfa(val, devise):
    """Convertit en FCFA depuis la devise de saisie."""
    return val if devise == "FCFA" else val * TAUX_CHANGE


def affiche(val_fcfa, devise, dec=2):
    """Convertit pour affichage dans la devise choisie."""
    return val_fcfa if devise == "FCFA" else val_fcfa / TAUX_CHANGE


def fmt(val, devise, dec=1):
    return f"{val:,.{dec}f} M{devise}"


# =============================================================================
# MOTEUR — RAMP-UP
# =============================================================================

def build_rampup(duree, paliers):
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
# MOTEUR — BFR (Campagne uniquement)
# =============================================================================

def calc_bfr(bm, revenus, couts, j_stock, j_clients, j_fourn):
    if bm == "Campagne":
        return couts * j_stock / 365 + revenus * j_clients / 365 - couts * j_fourn / 365
    if bm == "Production":
        return revenus * 0.05
    return 0.0


# =============================================================================
# MOTEUR — PROJECTION FCFF
# =============================================================================

def projeter_fcff(quantite, prix, cout, capex, taux_is, inflation,
                  duree, rampup, bm,
                  duree_bio=0, j_stock=60, j_clients=30, j_fourn=45,
                  ratio_maint=0.02):
    """
    Retourne une liste de dicts, un par année.
    Tous les montants sont en FCFA.
    """
    amort   = capex / duree
    c_maint = capex * ratio_maint
    rows    = []
    bfr_prec = 0.0

    for an in range(1, duree + 1):
        coeff = rampup[an - 1]
        if bm == "Production" and an <= duree_bio:
            coeff = 0.0
        inf   = (1 + inflation) ** (an - 1)
        rev   = quantite * coeff * prix  * inf
        cout_ = quantite * coeff * cout  * inf
        ebitda = rev - cout_
        ebit   = ebitda - amort
        impot  = max(ebit * taux_is, 0.0)
        nopat  = ebit - impot
        bfr    = calc_bfr(bm, rev, cout_, j_stock, j_clients, j_fourn)
        delta  = bfr - bfr_prec
        bfr_prec = bfr
        fcff   = nopat + amort - c_maint - delta

        rows.append({
            "Annee"         : an,
            "Ramp_up"       : coeff * 100,
            "Revenus"       : rev,
            "Couts"         : cout_,
            "EBITDA"        : ebitda,
            "EBIT"          : ebit,
            "IS"            : impot,
            "Delta_BFR"     : delta,
            "FCFF"          : fcff,
            "Ratio_Charges" : (cout_ / rev) if rev > 0 else np.nan,
        })
    return rows


def metriques(rows, capex, wacc):
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
# MOTEUR — SERVICE DE LA DETTE AVANCÉ
# =============================================================================

def service_dette(rows, capex, p_dette, r_dette, dur_dette,
                  grace, wacc, c_flat, c_engagt, profil_dec):
    """
    Tableau annuel du service de la dette.
    Inclut commission flat (prélevée une fois) et commission d'engagement
    (sur la fraction non décaissée chaque année).
    """
    dette     = capex * p_dette
    ann_remb  = max(dur_dette - grace, 1)
    amort_cap = dette / ann_remb
    c_flat_mt = dette * c_flat   # montant commission flat
    encours   = dette
    cum_dec   = 0.0
    total_int = c_flat_mt
    out       = []

    for i, r in enumerate(rows):
        an   = r["Annee"]
        fcff = r["FCFF"]

        frac_dec  = profil_dec[i] if i < len(profil_dec) else 0.0
        non_dec   = max(dette - cum_dec, 0.0)
        ce_an     = non_dec * c_engagt if cum_dec < dette else 0.0
        cum_dec  += dette * frac_dec

        if an <= dur_dette:
            interets = encours * r_dette
            remb_cap = 0.0 if an <= grace else amort_cap
            svc      = interets + remb_cap + ce_an
            enc_fin  = encours - remb_cap
            dscr     = fcff / svc if svc > 0 else np.nan

            ff_rest  = [rows[j]["FCFF"] for j in range(i, min(dur_dette, len(rows)))]
            llcr     = npv_calc(r_dette, ff_rest) / encours if encours > 0 else np.nan

            total_int += interets + ce_an
            out.append({
                "Annee"        : an,
                "Encours"      : encours,
                "Interets"     : interets,
                "Remb_Capital" : remb_cap,
                "Comm_Engagt"  : ce_an,
                "Service"      : svc,
                "FCFF"         : fcff,
                "DSCR"         : round(dscr, 3) if not np.isnan(dscr) else np.nan,
                "LLCR"         : round(llcr, 3) if not np.isnan(llcr) else np.nan,
                "Grace"        : (an <= grace),
            })
            encours = enc_fin
        else:
            out.append({
                "Annee"        : an,
                "Encours"      : 0.0,
                "Interets"     : 0.0,
                "Remb_Capital" : 0.0,
                "Comm_Engagt"  : 0.0,
                "Service"      : 0.0,
                "FCFF"         : fcff,
                "DSCR"         : np.nan,
                "LLCR"         : np.nan,
                "Grace"        : False,
            })

    df     = pd.DataFrame(out)
    dscrs  = df["DSCR"].dropna()
    d_min  = round(float(dscrs.min()),  3) if len(dscrs) else np.nan
    d_moy  = round(float(dscrs.mean()), 3) if len(dscrs) else np.nan
    ff_l   = [r["FCFF"] for r in rows[:dur_dette]]
    llcr_g = round(npv_calc(r_dette, ff_l) / dette, 3) if dette > 0 else np.nan
    return df, d_min, d_moy, llcr_g, total_int


# =============================================================================
# MOTEUR — STRESS TEST (Sans probabilité, 2 scénarios : Choc / Avec Assurance)
# =============================================================================

def stress_test(rows, gravite, prime_ann, indem_pct,
                capex, p_dette, r_dette, dur_dette,
                grace, wacc, c_flat, c_engagt, profil_dec):
    """
    Calcule 3 scénarios discrets :
      BASE      : flux projetés nominaux, sans perturbation
      CHOC      : choc concentré à l'année centrale (perte revenus × gravité),
                  sans assurance
      ASSURANCE : même choc + prime annuelle en charge + indemnité à l'année du choc

    Retourne pour chaque scénario :
      - rows avec Tresorerie_cum
      - DSCR annuel
      - DSCR minimum
    """
    duree   = len(rows)
    an_choc = max(1, duree // 2)

    def appliquer(avec_assurance, avec_choc):
        """Construit la série de flux pour un scénario donné."""
        out   = []
        treso = 0.0
        for r in rows:
            an   = r["Annee"]
            fcff = r["FCFF"]
            rev  = r["Revenus"]

            perte = rev * gravite if (avec_choc and an == an_choc) else 0.0
            fcff -= perte
            if avec_assurance:
                fcff -= prime_ann
                if an == an_choc:
                    fcff += perte * indem_pct

            treso += fcff
            out.append({
                "Annee"          : an,
                "FCFF"           : fcff,
                "Revenus"        : rev,
                "Tresorerie_cum" : treso,
            })
        return out

    base_rows  = appliquer(avec_assurance=False, avec_choc=False)
    choc_rows  = appliquer(avec_assurance=False, avec_choc=True)
    assur_rows = appliquer(avec_assurance=True,  avec_choc=True)

    def dscr_pour(scenario_rows):
        df_s, dmin, _, _, _ = service_dette(
            scenario_rows, capex, p_dette, r_dette, dur_dette,
            grace, wacc, c_flat, c_engagt, profil_dec,
        )
        return df_s["DSCR"].tolist(), dmin

    d_base_l,  d_base_min  = dscr_pour(base_rows)
    d_choc_l,  d_choc_min  = dscr_pour(choc_rows)
    d_assur_l, d_assur_min = dscr_pour(assur_rows)

    perte_brute = rows[an_choc - 1]["Revenus"] * gravite
    indemnite   = perte_brute * indem_pct
    cout_primes = prime_ann * duree

    return {
        "an_choc"      : an_choc,
        "perte_brute"  : perte_brute,
        "indemnite"    : indemnite,
        "cout_primes"  : cout_primes,
        "gain_net"     : indemnite - cout_primes,
        "base"         : {"rows": base_rows,  "dscr": d_base_l,  "dmin": d_base_min},
        "choc"         : {"rows": choc_rows,  "dscr": d_choc_l,  "dmin": d_choc_min},
        "assurance"    : {"rows": assur_rows, "dscr": d_assur_l, "dmin": d_assur_min},
    }


# =============================================================================
# MOTEUR — TORNADO & SENSIBILITÉ
# =============================================================================

def calc_tornado(quantite, prix, cout, capex, taux_is, inflation,
                 duree, rampup, bm, wacc,
                 duree_bio=0, j_stock=60, j_clients=30, j_fourn=45, delta=0.10):
    def van(**kw):
        return metriques(projeter_fcff(**kw), kw["capex"], wacc)["van"]

    base = dict(quantite=quantite, prix=prix, cout=cout, capex=capex,
                taux_is=taux_is, inflation=inflation, duree=duree,
                rampup=rampup, bm=bm, duree_bio=duree_bio,
                j_stock=j_stock, j_clients=j_clients, j_fourn=j_fourn)
    van_ref = van(**base)

    result = []
    for label, param, val in [
        ("Prix de Vente",   "prix",     prix),
        ("Volume Produit",  "quantite", quantite),
        ("CAPEX",           "capex",    capex),
        ("Coût Production", "cout",     cout),
    ]:
        vm = van(**{**base, param: val * (1 - delta)})
        vp = van(**{**base, param: val * (1 + delta)})
        result.append({"Variable": label, "im": vm - van_ref,
                        "ip": vp - van_ref, "amp": abs(vp - vm)})
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
            col.append(metriques(fl, capex, wacc)["van"])
        data[f"Coût {int(vc*100):+d}%"] = col
    return pd.DataFrame(data, index=[f"Prix {l}" for l in labs])


# =============================================================================
# IMPORT EXCEL
# =============================================================================

def parser_excel(fichier):
    MAPPING = {
        "ca"              : ["chiffre d'affaires", "revenus", "turnover", "sales"],
        "charges"         : ["charges d'exploitation", "opex", "operating expenses"],
        "amort"           : ["amortissement", "depreciation"],
        "resultat_net"    : ["résultat net", "net income", "bénéfice net"],
        "actif_immo"      : ["actif immobilisé", "immobilisations", "fixed assets"],
        "capitaux_propres": ["capitaux propres", "equity", "fonds propres"],
        "dettes_fin"      : ["dettes financières", "emprunts", "financial debt"],
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
                                        if isinstance(v, (int, float))
                                        and not np.isnan(float(v))]
                                if nums:
                                    data[cle] = float(nums[0])
                                break
        return data
    except Exception as e:
        return {"erreur": str(e)}


# =============================================================================
# GRAPHIQUES — épurés, orientés overview
# =============================================================================

def fig_jauges(van, tri, wacc_val, dscr_min, devise):
    """3 jauges : VAN / TRI vs WACC / DSCR min — lecture immédiate."""
    dscr_v = dscr_min if not np.isnan(dscr_min) else 0.0
    tri_v  = (tri * 100) if tri else 0.0
    van_m  = van / 1e6

    fig = make_subplots(
        rows=1, cols=3,
        specs=[[{"type": "indicator"}] * 3],
        horizontal_spacing=0.05,
    )

    # VAN
    van_range = max(abs(van_m) * 2, 100)
    fig.add_trace(go.Indicator(
        mode="gauge+number+delta",
        value=round(van_m, 1),
        title={"text": f"VAN (M{devise})", "font": {"size": 12}},
        delta={"reference": 0},
        gauge={
            "axis" : {"range": [-van_range, van_range],
                      "tickformat": ".0f"},
            "bar"  : {"color": CLR["vert"] if van > 0 else CLR["rouge"]},
            "threshold": {"line": {"color": "black", "width": 2}, "value": 0},
        },
        number={"font": {"size": 18}},
    ), row=1, col=1)

    # TRI vs WACC
    fig.add_trace(go.Indicator(
        mode="gauge+number",
        value=round(tri_v, 1),
        title={"text": "TRI (%)", "font": {"size": 12}},
        gauge={
            "axis" : {"range": [0, max(tri_v * 1.5, 30)]},
            "bar"  : {"color": CLR["vert"] if tri_v > wacc_val * 100 else CLR["rouge"]},
            "steps": [
                {"range": [0, wacc_val * 100], "color": "#fde8e8"},
                {"range": [wacc_val * 100, max(tri_v * 1.5, 30)], "color": "#e6f4ea"},
            ],
            "threshold": {"line": {"color": "navy", "width": 2},
                          "value": wacc_val * 100},
        },
        number={"suffix": "%", "font": {"size": 18}},
    ), row=1, col=2)

    # DSCR
    fig.add_trace(go.Indicator(
        mode="gauge+number",
        value=round(dscr_v, 2),
        title={"text": "DSCR Min (x)", "font": {"size": 12}},
        gauge={
            "axis" : {"range": [0, max(dscr_v * 1.5, 2.5)]},
            "bar"  : {"color": (CLR["vert"] if dscr_v >= 1.3
                                else CLR["amber"] if dscr_v >= 1.0
                                else CLR["rouge"])},
            "steps": [
                {"range": [0, 1.0], "color": "#fde8e8"},
                {"range": [1.0, 1.3], "color": "#fff8e1"},
                {"range": [1.3, max(dscr_v * 1.5, 2.5)], "color": "#e6f4ea"},
            ],
            "threshold": {"line": {"color": "black", "width": 2}, "value": 1.3},
        },
        number={"suffix": "x", "font": {"size": 18}},
    ), row=1, col=3)

    fig.update_layout(
        height=220,
        margin=dict(l=10, r=10, t=30, b=5),
        paper_bgcolor="white",
    )
    return fig


def fig_stress_sans_avec(stress, duree, devise, facteur):
    """
    UN seul graphique, 2 sous-graphes verticaux :
      Haut : DSCR annuel — 3 lignes (Base / Choc sans assurance / Choc avec assurance)
      Bas  : Trésorerie cumulée — mêmes 3 scénarios en aire

    Pas de courbe de probabilité. Lecture directe de l'impact de l'assurance.
    """
    annees  = list(range(1, duree + 1))
    an_choc = stress["an_choc"]

    def clean(lst):
        return [v if (isinstance(v, float) and not np.isnan(v)) else None for v in lst]

    def treso(rows):
        return [r["Tresorerie_cum"] / 1e6 * facteur for r in rows]

    fig = make_subplots(
        rows=2, cols=1,
        subplot_titles=("DSCR annuel", f"Trésorerie cumulée (M{devise})"),
        vertical_spacing=0.16,
        row_heights=[0.5, 0.5],
    )

    TRACES = [
        ("Base (nominal)",          "base",      CLR["bleu"],  "solid",  None),
        ("Choc — Sans Assurance",   "choc",      CLR["rouge"], "dot",    "x"),
        ("Choc — Avec Assurance",   "assurance", CLR["vert"],  "solid",  "diamond"),
    ]

    for label, cle, couleur, dash, symbol in TRACES:
        # DSCR
        fig.add_trace(go.Scatter(
            x=annees, y=clean(stress[cle]["dscr"]),
            mode="lines+markers",
            name=label,
            line=dict(color=couleur, width=2, dash=dash),
            marker=dict(size=6, symbol=symbol or "circle"),
            legendgroup=label,
        ), row=1, col=1)

        # Trésorerie
        fig.add_trace(go.Scatter(
            x=annees, y=treso(stress[cle]["rows"]),
            mode="lines",
            name=label,
            line=dict(color=couleur, width=2, dash=dash),
            fill="tozeroy",
            fillcolor=f"rgba({int(couleur[1:3],16)},{int(couleur[3:5],16)},{int(couleur[5:7],16)},0.08)",
            legendgroup=label,
            showlegend=False,
        ), row=2, col=1)

    # Seuils DSCR
    fig.add_hline(y=1.3, line_dash="dash", line_color=CLR["vert"],
                  line_width=1.2,
                  annotation_text="Seuil bancaire 1.3x",
                  annotation_font_size=10,
                  row=1, col=1)
    fig.add_hline(y=1.0, line_dash="dot", line_color=CLR["rouge"],
                  line_width=1.2,
                  annotation_text="Défaut 1.0x",
                  annotation_font_size=10,
                  row=1, col=1)

    # Ligne zéro trésorerie
    fig.add_hline(y=0, line_color="#333", line_width=0.8, row=2, col=1)

    # Zone de l'année du choc (les 2 sous-graphes)
    for r_n in [1, 2]:
        fig.add_vrect(
            x0=an_choc - 0.45, x1=an_choc + 0.45,
            fillcolor="rgba(220,53,69,0.10)", line_width=0,
            annotation_text=f"Choc An {an_choc}",
            annotation_position="top left",
            annotation_font_size=9,
            row=r_n, col=1,
        )

    fig.update_layout(
        height=520,
        plot_bgcolor="white",
        paper_bgcolor="white",
        margin=dict(t=50, b=20, l=10, r=10),
        legend=dict(orientation="h", y=-0.08, x=0, font_size=11),
        yaxis_title="DSCR (x)",
        yaxis2_title=f"M{devise}",
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
    .block-container { padding: 1.2rem 2rem 2rem 2rem; }
    h1  { font-size: 1.45rem; font-weight: 700; color: #0d1117; }
    h2  { font-size: 1.0rem;  font-weight: 600; color: #24292f;
          border-bottom: 2px solid #e8e8e8; padding-bottom: 4px;
          margin-top: 1.4rem; }
    h3  { font-size: .88rem;  font-weight: 600; color: #24292f; }
    .stMetric label { font-size: .72rem; color: #57606a;
                      text-transform: uppercase; letter-spacing: .04em; }
    hr  { border: none; border-top: 1px solid #e8e8e8; margin: 1rem 0; }

    /* Bandeau Business Model */
    .bm-banner {
        background: linear-gradient(90deg, #0d6efd, #0a58ca);
        color: white; border-radius: 8px; padding: 10px 20px;
        font-size: .9rem; font-weight: 600; margin-bottom: .8rem;
    }

    /* Score verdict */
    .verdict-box {
        border-radius: 8px; padding: 12px 20px;
        font-size: 1.05rem; font-weight: 700;
        text-align: center; color: white; margin-bottom: .6rem;
    }

    /* Tableau overview compact */
    .overview-table { font-size: .82rem; }
</style>
""", unsafe_allow_html=True)


# =============================================================================
# BARRE LATÉRALE
# =============================================================================

with st.sidebar:
    st.markdown("### CommodityWatch v4.1")
    st.caption("Overview Pré-Investissement")
    st.markdown("---")

    devise  = st.radio("Devise", ["FCFA", "USD"], horizontal=True)
    facteur = 1.0 if devise == "FCFA" else 1.0 / TAUX_CHANGE

    st.markdown("---")
    st.markdown("#### Type de Projet")
    bm = st.selectbox(
        "Business Model",
        ["Production", "Infrastructure", "Campagne"],
    )
    cfg = BM_CONFIG[bm]
    st.caption(cfg["description"])

    # Phase biologique (Production uniquement)
    duree_bio = 0
    if cfg["bio_actif"]:
        duree_bio = st.number_input(
            "Années sans revenu (croissance)", 0, 10, 3, 1,
        )

    # BFR (Campagne uniquement)
    j_stock = j_clients = j_fourn = 0
    if cfg["bfr_actif"]:
        st.markdown("**Cycle d'Exploitation**")
        j_stock   = st.number_input("Stock matière (jours)",    0, 180, 60, 5)
        j_clients = st.number_input("Crédit clients (jours)",   0, 120, 30, 5)
        j_fourn   = st.number_input("Crédit fournisseurs (j.)", 0, 120, 45, 5)

    st.markdown("---")
    st.markdown("#### Produit & Prix")

    matiere     = st.selectbox("Matière première", list(PRIX_REF_USD.keys()))
    prix_ref    = PRIX_REF_USD[matiere] * (1 if devise == "USD" else TAUX_CHANGE)
    if prix_ref > 0:
        st.info(f"Réf. marché : **{prix_ref:,.0f} {devise}/T** (ICE/Euronext/OTC)")

    prix_input  = st.number_input(
        f"Prix de Vente ({devise}/T)",
        min_value=0.0,
        value=float(prix_ref) if prix_ref > 0 else (1_500.0 if devise == "USD" else 900_000.0),
        step=10.0 if devise == "USD" else 10_000.0,
        format="%.1f" if devise == "USD" else "%.0f",
    )
    prix_fcfa_v = fcfa(prix_input, devise)

    quantite = st.number_input("Volume pleine capacité (T/an)", 0, 500_000, 5_000, 100)

    cout_input = st.number_input(
        f"Coût de Production ({devise}/T)",
        min_value=0.0,
        value=1_000.0 if devise == "USD" else 600_000.0,
        step=10.0 if devise == "USD" else 10_000.0,
        format="%.1f" if devise == "USD" else "%.0f",
    )
    cout_fcfa_v = fcfa(cout_input, devise)

    st.markdown("---")
    st.markdown("#### Montée en Puissance (Ramp-up)")
    n_pal   = int(st.number_input("Paliers", 1, 5, 3, 1))
    d_an_df = [2, 3, 4, 5, 6]
    d_tx_df = [0, 40, 80, 100, 100]
    paliers = []
    for k in range(n_pal):
        c1, c2 = st.columns(2)
        with c1:
            af = st.number_input(f"Fin an", 1, 25, d_an_df[k], 1, key=f"an{k}")
        with c2:
            tx = st.number_input("%", 0, 100, d_tx_df[k], 5, key=f"tx{k}")
        paliers.append((int(af), tx / 100))

    st.markdown("---")
    st.markdown("#### Structure Financière")

    capex_input = st.number_input(
        f"CAPEX Total ({devise})",
        min_value=0.0,
        value=3_000_000.0 if devise == "USD" else 2_000_000_000.0,
        step=100_000.0 if devise == "USD" else 100_000_000.0,
        format="%.0f",
    )
    capex_fcfa_v = fcfa(capex_input, devise)

    p_dette_pct = st.slider("Part dette (% CAPEX)", 0, 100, 60, 5)
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
    st.markdown("#### Paramètres du Prêt")
    dur_dette  = st.slider("Durée prêt (ans)",       1, duree, min(8, duree), 1)
    grace      = st.slider("Différé remboursement (ans)", 0, max(0, dur_dette - 1),
                            min(2, dur_dette - 1), 1)
    with st.expander("Commissions Bancaires"):
        c_flat_pct   = st.slider("Commission Flat (%, prélevée une fois)", 0.0, 3.0, 0.5, 0.1)
        c_engagt_pct = st.slider("Commission Engagement (%/an)", 0.0, 2.0, 0.25, 0.05)
    c_flat   = c_flat_pct   / 100
    c_engagt = c_engagt_pct / 100
    profil_dec = [0.5, 0.5] + [0.0] * (duree - 2)

    st.markdown("---")

    # Assurance — label et défauts adaptés au Business Model
    st.markdown(f"#### {cfg['assurance_titre']}")
    st.caption(cfg["assurance_desc"])

    gravite = st.slider(
        f"Gravité du sinistre — perte de CA (%)",
        5, 80, int(cfg["gravite_defaut"] * 100), 5,
    ) / 100

    prime_input = st.number_input(
        f"Prime annuelle ({devise})",
        min_value=0.0,
        value=50_000.0 if devise == "USD" else 30_000_000.0,
        step=5_000.0  if devise == "USD" else 5_000_000.0,
        format="%.0f",
    )
    prime_fcfa_v = fcfa(prime_input, devise)

    indem_pct = st.slider(
        "Taux de couverture (% de la perte indemnisé)", 0, 100, 70, 5,
    ) / 100

    st.markdown("---")
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
                st.warning("Installez `openpyxl` pour l'import Excel.")

    st.markdown("---")
    lancer = st.button("Générer l'Overview", type="primary", use_container_width=True)


# =============================================================================
# ZONE PRINCIPALE — HEADER
# =============================================================================

st.markdown(
    f"<div class='bm-banner'>"
    f"{cfg['icon']}  CommodityWatch — {matiere} &nbsp;|&nbsp; "
    f"{bm} &nbsp;|&nbsp; {devise}"
    f"</div>",
    unsafe_allow_html=True,
)
st.caption(
    "Overview pré-investissement — première lecture à chaud pour promoteur, "
    "analyste et comité de crédit bancaire."
)

if not lancer:
    c1, c2, c3 = st.columns(3)
    c1.info(f"**{cfg['icon']} {bm}**\n\n{cfg['description']}")
    c2.info(f"**Risque couvert**\n\n{cfg['risque_titre']}")
    c3.info(f"**Assurance associée**\n\n{cfg['assurance_titre']}")
    st.info(
        "Renseignez les paramètres dans la barre latérale "
        "puis cliquez sur **Générer l'Overview**.",
        icon="👈",
    )
    st.stop()


# =============================================================================
# CALCULS
# =============================================================================

with st.spinner("Calcul en cours..."):
    wacc_val  = r_dette * (1 - taux_is) * p_dette + r_fp * p_fp
    rampup    = build_rampup(duree, paliers)

    fcff_rows = projeter_fcff(
        quantite, prix_fcfa_v, cout_fcfa_v, capex_fcfa_v,
        taux_is, inflation, duree, rampup, bm,
        duree_bio, j_stock, j_clients, j_fourn,
    )

    met = metriques(fcff_rows, capex_fcfa_v, wacc_val)

    df_dette, dscr_min, dscr_moy, llcr_g, cout_total_dette = service_dette(
        fcff_rows, capex_fcfa_v, p_dette, r_dette, dur_dette,
        grace, wacc_val, c_flat, c_engagt, profil_dec,
    )

    alertes_chg = [r["Annee"] for r in fcff_rows
                   if not np.isnan(r["Ratio_Charges"])
                   and r["Ratio_Charges"] > SEUIL_CHARGES]

    tornado_d, van_ref_t = calc_tornado(
        quantite, prix_fcfa_v, cout_fcfa_v, capex_fcfa_v,
        taux_is, inflation, duree, rampup, bm, wacc_val,
        duree_bio, j_stock, j_clients, j_fourn,
    )

    df_sens = calc_sensibilite(
        quantite, prix_fcfa_v, cout_fcfa_v, capex_fcfa_v,
        taux_is, inflation, duree, rampup, bm, wacc_val,
        duree_bio, j_stock, j_clients, j_fourn,
    ) * facteur / 1e6

    stress = stress_test(
        fcff_rows, gravite, prime_fcfa_v, indem_pct,
        capex_fcfa_v, p_dette, r_dette, dur_dette,
        grace, wacc_val, c_flat, c_engagt, profil_dec,
    )


# =============================================================================
# SECTION 1 — VERDICT GLOBAL + JAUGES
# =============================================================================

van_v  = met["van"]
tri_v  = met["tri"]
pb_v   = met["payback"]

score = sum([
    van_v > 0,
    (tri_v or 0) > wacc_val,
    not np.isnan(dscr_min) and dscr_min >= SEUIL_BANCAIRE,
    not np.isnan(llcr_g)   and llcr_g   >= 1.10,
])

VERDICTS = {
    4: ("#198754", "PROJET VIABLE — 4/4 critères satisfaits"),
    3: ("#198754", "PROJET FAVORABLE — 3/4 critères satisfaits"),
    2: ("#fd7e14", "PROJET À AMÉLIORER — 2/4 critères satisfaits"),
    1: ("#dc3545", "PROJET RISQUÉ — 1/4 critères satisfaits"),
    0: ("#dc3545", "PROJET NON VIABLE — aucun critère satisfait"),
}
v_couleur, v_texte = VERDICTS[score]

st.markdown(
    f"<div class='verdict-box' style='background:{v_couleur}'>{v_texte}</div>",
    unsafe_allow_html=True,
)

# Jauges
st.plotly_chart(
    fig_jauges(van_v, tri_v, wacc_val, dscr_min, devise),
    use_container_width=True,
)

# Alertes immédiates
if alertes_chg:
    st.warning(
        f"Charges opérationnelles > {SEUIL_CHARGES*100:.0f}% du CA "
        f"aux années **{alertes_chg}**.",
        icon="⚠️",
    )
if not np.isnan(dscr_min):
    if dscr_min < SEUIL_DEFAUT:
        st.error(
            f"DSCR Min = **{dscr_min:.2f}x** — Risque de défaut de paiement. "
            "Le projet ne peut pas honorer son service de dette dans le scénario de base.",
            icon="🚨",
        )
    elif dscr_min < SEUIL_BANCAIRE:
        st.warning(
            f"DSCR Min = **{dscr_min:.2f}x** — En dessous du seuil bancaire standard (1.3x). "
            "Renégocier la maturité ou la franchise.",
            icon="⚠️",
        )

st.markdown("---")


# =============================================================================
# SECTION 2 — TABLEAU OVERVIEW (tout sur un seul tableau synthétique)
# =============================================================================

st.markdown("## Tableau de Synthèse")

def v(val_fcfa, dec=1):
    """Valeur convertie pour affichage en M devise."""
    return round(affiche(val_fcfa, devise) / 1e6, dec)

# Couleur indicateur
def ind(ok):
    return "✅" if ok else "❌"

capex_a = v(capex_fcfa_v, 0)
dette_a = v(capex_fcfa_v * p_dette, 0)
fp_a    = v(capex_fcfa_v * p_fp, 0)
van_a   = v(van_v)
tri_a   = f"{tri_v*100:.2f}%" if tri_v else "N/D"
pb_a    = f"{pb_v} ans" if pb_v else "> durée"
cd_a    = v(cout_total_dette)

data_synthese = {
    "Rubrique": [
        "─── RENTABILITÉ ───",
        "Valeur Actuelle Nette (VAN)",
        "Taux de Rendement Interne (TRI)",
        "Coût Moyen Pondéré (WACC)",
        "Délai de Récupération (Payback)",
        "",
        "─── FINANCEMENT ───",
        f"CAPEX Total",
        f"Dont Dette ({p_dette_pct}%)",
        f"Dont Fonds Propres ({100-p_dette_pct}%)",
        "Coût Total de la Dette",
        "Durée du Prêt / Différé",
        "",
        "─── BANCABILITÉ ───",
        "DSCR Minimum",
        "DSCR Moyen",
        "LLCR",
        "",
        "─── RISQUE & ASSURANCE ───",
        "Risque identifié",
        "Perte CA en cas de sinistre",
        "Indemnité estimée",
        "Coût cumulé des primes",
        "Gain net de l'assurance",
    ],
    "Valeur": [
        "",
        f"{van_a:,.1f} M{devise}",
        tri_a,
        f"{wacc_val*100:.2f}%",
        pb_a,
        "",
        "",
        f"{capex_a:,.0f} M{devise}",
        f"{dette_a:,.0f} M{devise}",
        f"{fp_a:,.0f} M{devise}",
        f"{cd_a:,.1f} M{devise}",
        f"{dur_dette} ans / {grace} an(s)",
        "",
        "",
        f"{dscr_min:.2f}x" if not np.isnan(dscr_min) else "N/D",
        f"{dscr_moy:.2f}x" if not np.isnan(dscr_moy) else "N/D",
        f"{llcr_g:.2f}x"   if not np.isnan(llcr_g)   else "N/D",
        "",
        "",
        cfg["risque_titre"],
        f"{v(stress['perte_brute']):.1f} M{devise}",
        f"{v(stress['indemnite']):.1f} M{devise}",
        f"{v(stress['cout_primes']):.1f} M{devise}",
        f"{v(stress['gain_net']):.1f} M{devise}",
    ],
    "Verdict": [
        "", ind(van_v > 0), ind((tri_v or 0) > wacc_val),
        "─", "─",
        "",
        "", "─", "─", "─", "─", "─",
        "",
        "",
        ind(not np.isnan(dscr_min) and dscr_min >= 1.3),
        ind(not np.isnan(dscr_moy) and dscr_moy >= 1.3),
        ind(not np.isnan(llcr_g)   and llcr_g   >= 1.1),
        "",
        "",
        "─",
        "─",
        "─",
        "─",
        ind(stress["gain_net"] > 0),
    ],
}

df_synth = pd.DataFrame(data_synthese)

def style_synthese(row):
    if str(row["Rubrique"]).startswith("───"):
        return ["background:#f0f4f8;font-weight:700;color:#0d1117"] * len(row)
    if row["Rubrique"] == "":
        return ["background:white;border:none"] * len(row)
    return [""] * len(row)

st.dataframe(
    df_synth.style.apply(style_synthese, axis=1),
    use_container_width=True,
    hide_index=True,
    height=680,
)

st.markdown("---")


# =============================================================================
# SECTION 3 — PROJECTION FCFF (tableau compact)
# =============================================================================

st.markdown("## Projection des Flux de Trésorerie")

# Tableau
df_proj = pd.DataFrame([{
    "An"                   : r["Annee"],
    "Cap. (%)"             : f"{r['Ramp_up']:.0f}%",
    f"Revenus M{devise}"   : round(affiche(r["Revenus"],  devise) / 1e6, 2),
    f"Coûts M{devise}"     : round(affiche(r["Couts"],    devise) / 1e6, 2),
    f"EBITDA M{devise}"    : round(affiche(r["EBITDA"],   devise) / 1e6, 2),
    "Chg/CA"               : f"{r['Ratio_Charges']*100:.1f}%" if not np.isnan(r["Ratio_Charges"]) else "—",
    f"FCFF M{devise}"      : round(affiche(r["FCFF"],     devise) / 1e6, 2),
} for r in fcff_rows])

def style_proj(row):
    fcff_col = f"FCFF M{devise}"
    chg_col  = "Chg/CA"
    fcff_v   = row.get(fcff_col, 0)
    try:
        chg_v = float(str(row.get(chg_col, "0%")).replace("%", "")) / 100
    except:
        chg_v = 0
    if fcff_v < 0:
        return ["background:#f8d7da"] * len(row)
    if chg_v > SEUIL_CHARGES:
        return ["background:#fff3cd"] * len(row)
    return [""] * len(row)

st.dataframe(
    df_proj.style.apply(style_proj, axis=1),
    use_container_width=True, hide_index=True,
)
st.caption("🟡 Charges > 55% CA &nbsp;&nbsp; 🔴 FCFF négatif")

st.markdown("---")


# =============================================================================
# SECTION 4 — SERVICE DE LA DETTE (tableau bancaire)
# =============================================================================

st.markdown("## Service de la Dette & Bancabilité")
st.caption(
    f"Grace Period : **{grace} an(s)** | Durée prêt : **{dur_dette} ans** | "
    f"Comm. Flat : **{c_flat_pct:.1f}%** | Comm. Engagement : **{c_engagt_pct:.2f}%/an**"
)

cols_show = [c for c in df_dette.columns if c != "Grace"]
df_d      = df_dette[cols_show].copy()
for col_m in ["Encours", "Interets", "Remb_Capital", "Comm_Engagt", "Service", "FCFF"]:
    df_d[col_m] = (df_dette[col_m] / 1e6 * facteur).round(2)

def style_dette(row):
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
    df_d.style.apply(style_dette, axis=1).format({
        "DSCR": lambda x: f"{x:.2f}x" if pd.notna(x) else "—",
        "LLCR": lambda x: f"{x:.2f}x" if pd.notna(x) else "—",
        **{c: "{:.2f}" for c in ["Encours","Interets","Remb_Capital",
                                   "Comm_Engagt","Service","FCFF"]},
    }),
    use_container_width=True, hide_index=True,
)
la, lb, lc, ld = st.columns(4)
la.markdown("<span style='background:#d4edda;padding:1px 7px;border-radius:3px;font-size:.76rem'>DSCR ≥ 1.3x</span>", unsafe_allow_html=True)
lb.markdown("<span style='background:#fde8c8;padding:1px 7px;border-radius:3px;font-size:.76rem'>1.0 – 1.3x</span>", unsafe_allow_html=True)
lc.markdown("<span style='background:#f8d7da;padding:1px 7px;border-radius:3px;font-size:.76rem'>< 1.0x Défaut</span>", unsafe_allow_html=True)
ld.markdown("<span style='background:#fff3cd;padding:1px 7px;border-radius:3px;font-size:.76rem'>Grace Period</span>", unsafe_allow_html=True)

st.markdown("---")


# =============================================================================
# SECTION 5 — SENSIBILITÉ (Tornado + Tableau croisé)
# =============================================================================

st.markdown("## Analyse de Sensibilité")
t1, t2 = st.tabs(["Variables Clés (Tornado)", "Tableau Croisé Prix × Coût"])

with t1:
    c_tor, c_tab = st.columns([3, 2])
    with c_tor:
        labels = [d["Variable"] for d in tornado_d]
        im     = [d["im"] / 1e6 * facteur for d in tornado_d]
        ip     = [d["ip"] / 1e6 * facteur for d in tornado_d]
        fig_t  = go.Figure()
        fig_t.add_trace(go.Bar(y=labels, x=im, orientation="h", name="-10%",
                               marker_color=CLR["rouge"],
                               text=[f"{v:+,.0f}" for v in im],
                               textposition="outside"))
        fig_t.add_trace(go.Bar(y=labels, x=ip, orientation="h", name="+10%",
                               marker_color=CLR["vert"],
                               text=[f"{v:+,.0f}" for v in ip],
                               textposition="outside"))
        fig_t.add_vline(x=0, line_color="#333", line_width=1)
        fig_t.update_layout(
            title=f"Tornado — Impact sur la VAN (M{devise})",
            xaxis_title=f"M{devise}", barmode="overlay",
            plot_bgcolor="white", paper_bgcolor="white",
            height=280, margin=dict(t=40, b=20),
            legend=dict(orientation="h", y=-0.3),
        )
        st.plotly_chart(fig_t, use_container_width=True)

    with c_tab:
        st.markdown(f"#### Impact sur VAN (M{devise})")
        df_t = pd.DataFrame([{
            "Variable"  : d["Variable"],
            "-10%"      : f"{d['im']/1e6*facteur:+,.0f}",
            "+10%"      : f"{d['ip']/1e6*facteur:+,.0f}",
            "Écart"     : f"{d['amp']/1e6*facteur:,.0f}",
        } for d in tornado_d])
        st.dataframe(df_t, use_container_width=True, hide_index=True)
        st.caption(
            "La variable en tête de liste est le **levier de négociation prioritaire** "
            "avant signature."
        )

with t2:
    fig_h = go.Figure(go.Heatmap(
        z=df_sens.values.tolist(),
        x=df_sens.columns.tolist(),
        y=df_sens.index.tolist(),
        colorscale=[[0, CLR["rouge"]], [0.5, "#ffffff"], [1, CLR["vert"]]],
        text=[[f"{v:.0f}" for v in row] for row in df_sens.values],
        texttemplate="%{text}",
        colorbar=dict(title=f"VAN M{devise}"),
    ))
    fig_h.update_layout(
        title=f"Sensibilité Croisée VAN — Prix × Coût (M{devise})",
        xaxis_title="Variation Coût", yaxis_title="Variation Prix",
        height=320, margin=dict(t=40),
    )
    st.plotly_chart(fig_h, use_container_width=True)

st.markdown("---")


# =============================================================================
# SECTION 6 — STRESS TEST & ASSURANCE
# =============================================================================

st.markdown(f"## Stress Test — {cfg['risque_titre']}")

st.info(
    f"**Risque :** {cfg['risque_desc']}\n\n"
    f"**Couverture :** {cfg['assurance_titre']} — {cfg['assurance_desc']}",
    icon="ℹ️",
)

# ── Tableau comparatif des 3 scénarios ───────────────────────────────────────
d_base  = stress["base"]["dmin"]
d_choc  = stress["choc"]["dmin"]
d_assur = stress["assurance"]["dmin"]

def dscr_badge(val):
    if np.isnan(val): return "N/D"
    if val >= 1.3:    return f"✅ {val:.2f}x"
    if val >= 1.0:    return f"⚠️ {val:.2f}x"
    return f"❌ {val:.2f}x"

def treso_fin(rows):
    return rows[-1]["Tresorerie_cum"] / 1e6 * facteur if rows else 0.0

df_comp = pd.DataFrame({
    "Critère": [
        "DSCR Minimum",
        "Trésorerie cumulée finale",
        "Perte CA (choc)",
        "Indemnité reçue",
        "Coût total primes",
        "Gain net assurance",
    ],
    "Base (nominal)": [
        dscr_badge(d_base),
        f"{treso_fin(stress['base']['rows']):,.1f} M{devise}",
        "—", "—", "—", "—",
    ],
    "Choc — Sans Assurance": [
        dscr_badge(d_choc),
        f"{treso_fin(stress['choc']['rows']):,.1f} M{devise}",
        f"{v(stress['perte_brute']):.1f} M{devise}",
        "—",
        "—",
        "—",
    ],
    "Choc — Avec Assurance": [
        dscr_badge(d_assur),
        f"{treso_fin(stress['assurance']['rows']):,.1f} M{devise}",
        f"{v(stress['perte_brute']):.1f} M{devise}",
        f"{v(stress['indemnite']):.1f} M{devise}",
        f"{v(stress['cout_primes']):.1f} M{devise}",
        f"{'✅' if stress['gain_net']>0 else '⚠️'} {v(stress['gain_net']):.1f} M{devise}",
    ],
})

def style_comp(row):
    if "DSCR" in str(row["Critère"]):
        return ["font-weight:600"] * len(row)
    return [""] * len(row)

st.dataframe(
    df_comp.style.apply(style_comp, axis=1),
    use_container_width=True, hide_index=True,
)

# ── Un seul graphique : DSCR + Trésorerie — 3 scénarios ─────────────────────
st.plotly_chart(
    fig_stress_sans_avec(stress, duree, devise, facteur),
    use_container_width=True,
)

# ── Interprétation en langage clair ──────────────────────────────────────────
st.markdown("#### Lecture rapide")

lignes = []

if not np.isnan(d_choc):
    if d_choc < SEUIL_DEFAUT:
        lignes.append(
            f"Sans assurance, un sinistre à l'année {stress['an_choc']} ferait chuter le DSCR "
            f"à **{d_choc:.2f}x** — en dessous du seuil de défaut. "
            "Le projet ne peut plus rembourser sa dette cette année-là."
        )
    elif d_choc < SEUIL_BANCAIRE:
        lignes.append(
            f"Sans assurance, le DSCR tomberait à **{d_choc:.2f}x** "
            f"(sous le seuil bancaire de 1.3x). "
            "La banque exigerait probablement un compte de réserve ou un covenant de couverture."
        )
    else:
        lignes.append(
            f"Même sans assurance, le projet absorbe le choc "
            f"avec un DSCR de **{d_choc:.2f}x**. "
            "L'assurance reste une précaution recommandée, non une nécessité absolue."
        )

if not np.isnan(d_assur):
    gain_dscr = d_assur - d_choc if not np.isnan(d_choc) else 0
    if d_assur >= SEUIL_BANCAIRE:
        lignes.append(
            f"Avec la **{cfg['assurance_titre']}**, le DSCR remonte à **{d_assur:.2f}x** "
            f"({gain_dscr:+.2f}x vs choc sans assurance) — la bancabilité est restaurée."
        )
    else:
        lignes.append(
            f"Avec assurance, le DSCR atteint **{d_assur:.2f}x** "
            f"({gain_dscr:+.2f}x) — amélioré mais encore sous le seuil de 1.3x. "
            "Envisager une réserve de liquidité complémentaire."
        )

if stress["gain_net"] > 0:
    lignes.append(
        f"Financièrement, l'assurance est rentable sur {duree} ans : "
        f"l'indemnité estimée ({v(stress['indemnite']):.1f} M{devise}) "
        f"dépasse le coût total des primes ({v(stress['cout_primes']):.1f} M{devise})."
    )
else:
    lignes.append(
        f"Le coût des primes sur {duree} ans ({v(stress['cout_primes']):.1f} M{devise}) "
        f"dépasse l'indemnité unique ({v(stress['indemnite']):.1f} M{devise}). "
        "C'est normal pour un sinistre rare : la valeur de l'assurance est la sécurité "
        "qu'elle procure au banquier, pas le gain financier attendu."
    )

for lg in lignes:
    st.markdown(f"- {lg}")

st.markdown("---")
st.markdown(
    f"<div style='font-size:.70rem;color:#6c757d;text-align:center'>"
    f"CommodityWatch v4.1 — 1 USD = {TAUX_CHANGE:.0f} FCFA — "
    "Document indicatif pré-investissement. "
    "Ne se substitue pas à une due diligence financière et juridique complète."
    "</div>",
    unsafe_allow_html=True,
)
