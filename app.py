# =============================================================================
# CommodityWatch - Outil de Pre-Scoring de Projets Agro-Industriels
# Framework : Streamlit | Python 3.9+
# Usage     : streamlit run app.py
# =============================================================================

import streamlit as st
import numpy as np
import numpy_financial as npf
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from itertools import product

# =============================================================================
# CONFIGURATION GLOBALE DE LA PAGE
# =============================================================================

st.set_page_config(
    page_title="CommodityWatch | Pre-Scoring",
    page_icon=None,
    layout="wide",
    initial_sidebar_state="expanded",
)

# Injection de CSS minimaliste pour affiner l'apparence
st.markdown(
    """
    <style>
        /* Typographie et espacements */
        h1 { font-size: 1.6rem; font-weight: 700; color: #0d1117; }
        h2 { font-size: 1.1rem; font-weight: 600; color: #24292f; border-bottom: 1px solid #e0e0e0; padding-bottom: 6px; }
        h3 { font-size: 0.95rem; font-weight: 600; color: #24292f; }
        .stMetric label { font-size: 0.78rem; color: #57606a; text-transform: uppercase; letter-spacing: 0.04em; }
        .stMetric .metric-container { background: #f6f8fa; border-radius: 6px; padding: 12px; }
        /* Séparateur section */
        hr { border: none; border-top: 1px solid #e0e0e0; margin: 20px 0; }
        /* Réduction du padding principal */
        .block-container { padding-top: 2rem; padding-bottom: 2rem; }
    </style>
    """,
    unsafe_allow_html=True,
)


# =============================================================================
# FONCTIONS METIER (MOTEUR DE CALCUL)
# =============================================================================

def calculer_wacc(part_dette, cout_dette, taux_is, part_fonds_propres, cout_fonds_propres):
    """
    Calcule le Coût Moyen Pondéré du Capital (WACC) avec bouclier fiscal.

    WACC = Kd * (1 - IS) * (D/V) + Ke * (E/V)

    Parameters
    ----------
    part_dette            : float  - Proportion de la dette dans le capital (0 à 1)
    cout_dette            : float  - Coût brut de la dette (annuel, 0 à 1)
    taux_is               : float  - Taux d'imposition sur les sociétés (0 à 1)
    part_fonds_propres    : float  - Proportion des fonds propres dans le capital (0 à 1)
    cout_fonds_propres    : float  - Coût des fonds propres / CAPM (0 à 1)

    Returns
    -------
    float : WACC en valeur décimale
    """
    cout_dette_apres_is = cout_dette * (1 - taux_is)
    wacc = (cout_dette_apres_is * part_dette) + (cout_fonds_propres * part_fonds_propres)
    return wacc


def projeter_fcff(
    quantite, prix_vente, cout_production,
    capex_initial, taux_is, taux_inflation,
    duree, ratio_maintenance=0.02
):
    """
    Génère la projection annuelle des Free Cash Flows to Firm (FCFF).

    Hypothèses simplifiées pour le pre-scoring :
      - Revenus = Quantité * Prix ajusté à l'inflation
      - Couts = Quantité * Cout ajusté à l'inflation
      - EBITDA = Revenus - Couts
      - Amortissement linéaire du CAPEX sur la durée du projet
      - EBIT = EBITDA - Amortissement
      - NOPAT = EBIT * (1 - IS)  [si EBIT > 0, sinon 0]
      - CAPEX de maintenance = ratio_maintenance * CAPEX initial (annuel)
      - FCFF = NOPAT + Amortissement - CAPEX maintenance
      (Variation du BFR simplifiée à 0 pour le pre-scoring)

    Returns
    -------
    list of dict : Une entrée par année avec toutes les composantes du FCFF
    """
    amortissement_annuel = capex_initial / duree
    capex_maintenance = capex_initial * ratio_maintenance
    resultats = []

    for annee in range(1, duree + 1):
        facteur_inflation = (1 + taux_inflation) ** (annee - 1)
        revenus = quantite * prix_vente * facteur_inflation
        couts = quantite * cout_production * facteur_inflation
        ebitda = revenus - couts
        ebit = ebitda - amortissement_annuel
        impot = max(ebit * taux_is, 0)  # Pas d'impot si EBIT negatif
        nopat = ebit - impot
        fcff = nopat + amortissement_annuel - capex_maintenance

        resultats.append({
            "Annee": annee,
            "Revenus (MFCFA)": round(revenus / 1e6, 2),
            "Couts Prod. (MFCFA)": round(couts / 1e6, 2),
            "EBITDA (MFCFA)": round(ebitda / 1e6, 2),
            "Amortissement (MFCFA)": round(amortissement_annuel / 1e6, 2),
            "EBIT (MFCFA)": round(ebit / 1e6, 2),
            "Impot IS (MFCFA)": round(impot / 1e6, 2),
            "FCFF (MFCFA)": round(fcff / 1e6, 2),
        })

    return resultats


def calculer_metriques(fcff_liste, capex_initial, wacc):
    """
    Calcule VAN, TRI et Payback Period a partir des FCFF projetes.

    Parameters
    ----------
    fcff_liste   : list of dict - Résultat de projeter_fcff()
    capex_initial: float        - Investissement initial en FCFA
    wacc         : float        - Taux d'actualisation

    Returns
    -------
    dict : VAN (MFCFA), TRI (%), Payback (annees)
    """
    flux = [-capex_initial] + [r["FCFF (MFCFA)"] * 1e6 for r in fcff_liste]
    van = npf.npv(wacc, flux)
    try:
        tri = npf.irr(flux)
    except Exception:
        tri = None

    # Payback : cumul des flux positifs jusqu'au remboursement du CAPEX
    cumul = 0
    payback = None
    for i, f in enumerate(flux[1:], 1):
        cumul += f
        if cumul >= capex_initial:
            payback = i
            break

    return {
        "VAN (MFCFA)": round(van / 1e6, 2),
        "TRI (%)": round(tri * 100, 2) if tri is not None else None,
        "Payback (annees)": payback,
    }


def calculer_dscr(
    fcff_liste, capex_initial, part_dette, cout_dette, duree_dette=None
):
    """
    Calcule le DSCR (Debt Service Coverage Ratio) annuel.

    DSCR = FCFF / Service de la dette annuel
    Service de la dette = amortissement du capital + interets

    Hypothèse : Remboursement lineaire du principal sur duree_dette.

    Returns
    -------
    list of float : DSCR par annee (NaN apres remboursement total)
    float         : DSCR moyen sur la duree du pret
    """
    montant_dette = capex_initial * part_dette
    if duree_dette is None:
        duree_dette = len(fcff_liste)

    amort_capital = montant_dette / duree_dette
    dscr_annuels = []

    for i, r in enumerate(fcff_liste):
        if i < duree_dette:
            capital_restant = montant_dette - i * amort_capital
            interets = capital_restant * cout_dette
            service_dette = amort_capital + interets
            dscr = (r["FCFF (MFCFA)"] * 1e6) / service_dette if service_dette > 0 else np.nan
        else:
            dscr = np.nan  # Plus de dette
        dscr_annuels.append(round(dscr, 2) if not np.isnan(dscr) else np.nan)

    valeurs_valides = [d for d in dscr_annuels if not np.isnan(d)]
    dscr_moyen = round(np.mean(valeurs_valides), 2) if valeurs_valides else np.nan
    return dscr_annuels, dscr_moyen


def analyse_sensibilite_van(
    quantite, prix_vente, cout_production,
    capex_initial, taux_is, taux_inflation,
    duree, wacc, variations=(-0.10, -0.05, 0.0, 0.05, 0.10)
):
    """
    Génère un tableau croisé de sensibilité de la VAN en faisant varier
    simultanement le Prix de Vente et le Cout de Production.

    Returns
    -------
    pd.DataFrame : Tableau croisé (index = Prix, colonnes = Cout)
    """
    index_labels = [f"{int(v*100):+d}%" for v in variations]
    data = {}

    for v_cout in variations:
        col = []
        cout_adj = cout_production * (1 + v_cout)
        for v_prix in variations:
            prix_adj = prix_vente * (1 + v_prix)
            flux_temp = projeter_fcff(
                quantite, prix_adj, cout_adj,
                capex_initial, taux_is, taux_inflation, duree
            )
            metriques_temp = calculer_metriques(flux_temp, capex_initial, wacc)
            col.append(metriques_temp["VAN (MFCFA)"])
        data[f"Cout {int(v_cout*100):+d}%"] = col

    df = pd.DataFrame(data, index=[f"Prix {l}" for l in index_labels])
    return df


def simulation_monte_carlo(
    quantite, prix_vente, cout_production,
    capex_initial, taux_is, taux_inflation,
    duree, wacc,
    part_dette, cout_dette,
    sigma_prix, sigma_quantite,
    n_iterations=1000
):
    """
    Simulation de Monte Carlo sur la VAN et le DSCR.

    Les variables Prix de Vente et Quantite suivent une distribution normale
    centree sur leur valeur de base, avec les ecarts-types fournis.

    Returns
    -------
    np.array : Distribution des VAN simulees (MFCFA)
    float    : Probabilite de VAN < 0
    float    : Probabilite de DSCR moyen < 1
    """
    vans = []
    prob_dscr = 0

    for _ in range(n_iterations):
        prix_sim = np.random.normal(prix_vente, prix_vente * sigma_prix)
        qte_sim = max(np.random.normal(quantite, quantite * sigma_quantite), 0)

        flux_sim = projeter_fcff(
            qte_sim, prix_sim, cout_production,
            capex_initial, taux_is, taux_inflation, duree
        )
        met_sim = calculer_metriques(flux_sim, capex_initial, wacc)
        vans.append(met_sim["VAN (MFCFA)"])

        _, dscr_moy = calculer_dscr(flux_sim, capex_initial, part_dette, cout_dette)
        if not np.isnan(dscr_moy) and dscr_moy < 1:
            prob_dscr += 1

    vans_arr = np.array(vans)
    prob_van_neg = round((vans_arr < 0).mean() * 100, 1)
    prob_dscr_def = round((prob_dscr / n_iterations) * 100, 1)
    return vans_arr, prob_van_neg, prob_dscr_def


# =============================================================================
# INTERFACE STREAMLIT - BARRE LATERALE (INPUTS)
# =============================================================================

with st.sidebar:
    st.markdown("## CommodityWatch")
    st.markdown("**Pre-Scoring de Projet Agro-Industriel**")
    st.markdown("---")

    st.markdown("### Parametres de Marche")

    matiere_premiere = st.selectbox(
        "Matiere Premiere",
        options=["Cafe", "Cacao", "Anacarde", "Soja"],
    )

    quantite = st.number_input(
        "Quantite Produite Annuelle (Tonnes)",
        min_value=0,
        value=5000,
        step=100,
    )

    prix_vente = st.number_input(
        "Prix de Vente Unitaire (FCFA/Tonne)",
        min_value=0,
        value=900_000,
        step=10_000,
    )

    cout_production = st.number_input(
        "Cout de Production Unitaire (FCFA/Tonne)",
        min_value=0,
        value=600_000,
        step=10_000,
    )

    st.markdown("---")
    st.markdown("### Investissement & Structure du Capital")

    capex_initial = st.number_input(
        "CAPEX Initial Total (FCFA)",
        min_value=0,
        value=2_000_000_000,
        step=100_000_000,
        format="%d",
    )

    part_dette_pct = st.slider(
        "Part de Dette (% du CAPEX)",
        min_value=0, max_value=100, value=60, step=5,
    )
    part_dette = part_dette_pct / 100
    part_fonds_propres = 1 - part_dette

    cout_dette_pct = st.slider(
        "Cout de la Dette (% Annuel)",
        min_value=1, max_value=25, value=9, step=1,
    )
    cout_dette = cout_dette_pct / 100

    cout_fp_pct = st.slider(
        "Cout des Fonds Propres - CAPM (% Annuel)",
        min_value=5, max_value=30, value=15, step=1,
    )
    cout_fonds_propres = cout_fp_pct / 100

    taux_is_pct = st.slider(
        "Taux d'Imposition (IS %)",
        min_value=0, max_value=40, value=25, step=1,
    )
    taux_is = taux_is_pct / 100

    duree = st.slider(
        "Duree du Projet (Annees)",
        min_value=3, max_value=25, value=10, step=1,
    )

    taux_inflation_pct = st.slider(
        "Taux d'Inflation Annuel Moyen (%)",
        min_value=0, max_value=15, value=3, step=1,
    )
    taux_inflation = taux_inflation_pct / 100

    st.markdown("---")
    st.markdown("### Simulation Monte Carlo")

    sigma_prix_pct = st.slider(
        "Ecart-type Prix de Vente (%)",
        min_value=1, max_value=30, value=10, step=1,
    )
    sigma_prix = sigma_prix_pct / 100

    sigma_qte_pct = st.slider(
        "Ecart-type Quantite Produite (%)",
        min_value=1, max_value=30, value=8, step=1,
    )
    sigma_quantite = sigma_qte_pct / 100

    n_iterations = st.selectbox(
        "Nombre d'Iterations",
        options=[500, 1000, 2000, 5000],
        index=1,
    )

    st.markdown("---")
    lancer = st.button("Lancer l'Analyse", type="primary", use_container_width=True)


# =============================================================================
# INTERFACE STREAMLIT - ZONE PRINCIPALE (OUTPUTS)
# =============================================================================

st.markdown(f"## Tableau de Bord Pre-Scoring — {matiere_premiere}")
st.markdown(
    "Outil d'analyse financiere rapide destiné à évaluer la viabilité d'un projet avant modélisation complète."
)
st.markdown("---")

if not lancer:
    st.info(
        "Renseignez les parametres dans la barre laterale puis cliquez sur "
        "**Lancer l'Analyse** pour afficher les résultats."
    )
    st.stop()


# =============================================================================
# CALCULS PRINCIPAUX
# =============================================================================

with st.spinner("Calculs en cours..."):

    # 1. WACC
    wacc = calculer_wacc(
        part_dette, cout_dette, taux_is,
        part_fonds_propres, cout_fonds_propres
    )

    # 2. FCFF projetes
    fcff_liste = projeter_fcff(
        quantite, prix_vente, cout_production,
        capex_initial, taux_is, taux_inflation, duree
    )

    # 3. Metriques de rentabilite
    metriques = calculer_metriques(fcff_liste, capex_initial, wacc)

    # 4. DSCR
    dscr_annuels, dscr_moyen = calculer_dscr(
        fcff_liste, capex_initial, part_dette, cout_dette
    )

    # 5. Sensibilite
    df_sensibilite = analyse_sensibilite_van(
        quantite, prix_vente, cout_production,
        capex_initial, taux_is, taux_inflation,
        duree, wacc
    )

    # 6. Monte Carlo
    vans_mc, prob_van_neg, prob_dscr_def = simulation_monte_carlo(
        quantite, prix_vente, cout_production,
        capex_initial, taux_is, taux_inflation,
        duree, wacc,
        part_dette, cout_dette,
        sigma_prix, sigma_quantite,
        n_iterations
    )


# =============================================================================
# SECTION 1 : INDICATEURS CLES (KPIs)
# =============================================================================

st.markdown("## Indicateurs Cles de Performance")

van_val = metriques["VAN (MFCFA)"]
tri_val = metriques["TRI (%)"]
payback_val = metriques["Payback (annees)"]

col1, col2, col3, col4, col5 = st.columns(5)

with col1:
    st.metric(
        label="VAN",
        value=f"{van_val:,.0f} MFCFA",
        delta="Positive" if van_val > 0 else "Negative",
        delta_color="normal" if van_val > 0 else "inverse",
    )

with col2:
    st.metric(
        label="TRI",
        value=f"{tri_val:.2f} %" if tri_val is not None else "N/D",
        delta=f"WACC : {wacc*100:.2f} %",
        delta_color="normal" if (tri_val or 0) > wacc * 100 else "inverse",
    )

with col3:
    st.metric(
        label="WACC",
        value=f"{wacc*100:.2f} %",
    )

with col4:
    st.metric(
        label="DSCR Moyen",
        value=f"{dscr_moyen:.2f}x" if not np.isnan(dscr_moyen) else "N/D",
        delta="Satisfaisant" if dscr_moyen >= 1.2 else "Insuffisant",
        delta_color="normal" if dscr_moyen >= 1.2 else "inverse",
    )

with col5:
    st.metric(
        label="Payback",
        value=f"{payback_val} ans" if payback_val else "> Duree Projet",
        delta_color="off",
    )

st.markdown("---")


# =============================================================================
# SECTION 2 : TABLEAU DES FCFF PROJETES
# =============================================================================

st.markdown("## Projection des Flux de Tresorerie (FCFF)")
st.caption(
    "Montants en millions de FCFA (MFCFA). Inflation appliquee annuellement aux revenus et couts."
)

df_fcff = pd.DataFrame(fcff_liste)
df_fcff["DSCR"] = dscr_annuels
df_fcff["DSCR"] = df_fcff["DSCR"].apply(
    lambda x: f"{x:.2f}x" if not (isinstance(x, float) and np.isnan(x)) else "-"
)

# Mise en forme conditionnelle du FCFF
def colorer_fcff(val):
    if isinstance(val, (int, float)):
        color = "#d4edda" if val >= 0 else "#f8d7da"
        return f"background-color: {color};"
    return ""

styled_df = df_fcff.style.applymap(colorer_fcff, subset=["FCFF (MFCFA)"])
st.dataframe(styled_df, use_container_width=True, hide_index=True)

st.markdown("---")


# =============================================================================
# SECTION 3 : ANALYSE DE SENSIBILITE
# =============================================================================

st.markdown("## Analyse de Sensibilite de la VAN (MFCFA)")
st.caption(
    "Variation croisee du Prix de Vente (lignes) et du Cout de Production (colonnes). "
    "Base = valeurs saisies. Cellules vertes : VAN positive."
)

def colorer_sensibilite(val):
    if isinstance(val, (int, float)):
        if val > 0:
            intensite = min(int(abs(val) / (abs(df_sensibilite.values).max() + 1e-9) * 80), 80)
            return f"background-color: rgba(25, 135, 84, {intensite/100 + 0.1}); color: white;"
        else:
            intensite = min(int(abs(val) / (abs(df_sensibilite.values).max() + 1e-9) * 80), 80)
            return f"background-color: rgba(220, 53, 69, {intensite/100 + 0.1}); color: white;"
    return ""

styled_sens = df_sensibilite.style.applymap(colorer_sensibilite).format("{:.0f}")
st.dataframe(styled_sens, use_container_width=True)

st.markdown("---")


# =============================================================================
# SECTION 4 : SIMULATION DE MONTE CARLO
# =============================================================================

st.markdown("## Analyse de Risque — Simulation de Monte Carlo")
st.caption(
    f"{n_iterations:,} iterations. Prix et Quantite simulés selon une distribution normale."
)

col_mc1, col_mc2, col_mc3 = st.columns([2, 1, 1])

with col_mc1:
    # Histogramme de la distribution des VAN
    fig, ax = plt.subplots(figsize=(7, 3.5))
    ax.hist(vans_mc, bins=50, color="#0d6efd", alpha=0.75, edgecolor="white", linewidth=0.4)
    ax.axvline(0, color="#dc3545", linewidth=1.5, linestyle="--", label="VAN = 0")
    ax.axvline(np.percentile(vans_mc, 5), color="#fd7e14", linewidth=1.2,
               linestyle=":", label="Percentile 5%")
    ax.axvline(np.mean(vans_mc), color="#198754", linewidth=1.2,
               linestyle="-.", label=f"Moyenne : {np.mean(vans_mc):.0f} MFCFA")
    ax.set_xlabel("VAN (MFCFA)", fontsize=9)
    ax.set_ylabel("Frequence", fontsize=9)
    ax.set_title("Distribution des VAN Simulees", fontsize=10, fontweight="bold", pad=10)
    ax.legend(fontsize=8, framealpha=0.6)
    ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"{x:,.0f}"))
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    plt.tight_layout()
    st.pyplot(fig)
    plt.close(fig)

with col_mc2:
    st.markdown("#### Probabilites de Defaut")
    st.metric(
        label="P(VAN < 0)",
        value=f"{prob_van_neg} %",
        delta="Risque eleve" if prob_van_neg > 30 else "Risque maitrise",
        delta_color="inverse" if prob_van_neg > 30 else "normal",
    )
    st.metric(
        label="P(DSCR < 1)",
        value=f"{prob_dscr_def} %",
        delta="Risque eleve" if prob_dscr_def > 30 else "Risque maitrise",
        delta_color="inverse" if prob_dscr_def > 30 else "normal",
    )

with col_mc3:
    st.markdown("#### Statistiques VAN")
    stats_mc = {
        "Statistique": ["Moyenne", "Mediane", "Ecart-type", "Pct. 5%", "Pct. 95%"],
        "Valeur (MFCFA)": [
            f"{np.mean(vans_mc):,.0f}",
            f"{np.median(vans_mc):,.0f}",
            f"{np.std(vans_mc):,.0f}",
            f"{np.percentile(vans_mc, 5):,.0f}",
            f"{np.percentile(vans_mc, 95):,.0f}",
        ],
    }
    st.dataframe(pd.DataFrame(stats_mc), use_container_width=True, hide_index=True)

st.markdown("---")
st.markdown(
    "<div style='font-size:0.75rem; color:#57606a;'>"
    "CommodityWatch v1.0 — Outil de pre-scoring interne. "
    "Les résultats sont des estimations indicatives et ne constituent pas un avis financier définitif."
    "</div>",
    unsafe_allow_html=True,
)
