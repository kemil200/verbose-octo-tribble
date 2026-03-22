# =============================================================================
# CommodityWatch v2.0 - Outil de Pre-Scoring de Projets Agro-Industriels
# Framework : Streamlit | Python 3.9+
# Usage     : streamlit run app.py
#
# Nouveautes v2.0 :
#   - Pilier 1 : Courbe de Ramp-up parametrable + Grace Period (differe dette)
#   - Pilier 2 : Tableau de service de la dette avec DSCR et LLCR annuels,
#                alerte visuelle si DSCR < 1.3x
#   - Pilier 3 : Tornado Chart (sensibilite VAN sur 4 variables cles)
#   - Pilier 4 : Monte Carlo avec risque de queue (choc climatique)
# =============================================================================

import streamlit as st
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker


# =============================================================================
# FONCTIONS MATHEMATIQUES (remplacement numpy_financial)
# =============================================================================

def npv(rate, cashflows):
    """VAN : somme des flux actualises au taux fourni (annee 0 = t=0)."""
    cashflows = np.asarray(cashflows, dtype=float)
    indices = np.arange(len(cashflows))
    return np.sum(cashflows / (1 + rate) ** indices)


def irr(cashflows, tol=1e-7, max_iter=1000):
    """
    TRI par methode de Newton-Raphson.
    Retourne None si la convergence echoue ou si le profil de flux est invalide.
    """
    cashflows = np.asarray(cashflows, dtype=float)
    if not (np.any(cashflows < 0) and np.any(cashflows > 0)):
        return None
    rate = 0.1
    for _ in range(max_iter):
        indices = np.arange(len(cashflows))
        f  = np.sum(cashflows / (1 + rate) ** indices)
        df = np.sum(-indices * cashflows / (1 + rate) ** (indices + 1))
        if df == 0:
            return None
        rate_new = rate - f / df
        if abs(rate_new - rate) < tol:
            return rate_new
        rate = rate_new
    return None


# =============================================================================
# CONFIGURATION GLOBALE DE LA PAGE
# =============================================================================

st.set_page_config(
    page_title="CommodityWatch | Pre-Scoring v2",
    page_icon=None,
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
        h1 { font-size: 1.6rem; font-weight: 700; color: #0d1117; }
        h2 { font-size: 1.1rem; font-weight: 600; color: #24292f;
             border-bottom: 1px solid #e0e0e0; padding-bottom: 6px; }
        h3 { font-size: 0.95rem; font-weight: 600; color: #24292f; }
        .stMetric label { font-size: 0.78rem; color: #57606a;
                          text-transform: uppercase; letter-spacing: 0.04em; }
        hr  { border: none; border-top: 1px solid #e0e0e0; margin: 20px 0; }
        .block-container { padding-top: 2rem; padding-bottom: 2rem; }
    </style>
    """,
    unsafe_allow_html=True,
)


# =============================================================================
# PILIER 1 — COURBE DE RAMP-UP
# =============================================================================

def construire_rampup(duree, rampup_config):
    """
    Genere un vecteur de coefficients de production (0.0 a 1.0) sur la duree
    du projet, en fonction de la configuration de montee en puissance.

    Parameters
    ----------
    duree         : int  - Nombre total d'annees du projet
    rampup_config : list of (int, float) - [(annee_fin, taux), ...]
                    Exemple : [(2, 0.0), (3, 0.40), (4, 0.80)]
                    annees 1-2 : 0%, annee 3 : 40%, annee 4 : 80%, annees 5+ : 100%

    Returns
    -------
    list of float : Coefficients annuels de production
    """
    coeffs = []
    for annee in range(1, duree + 1):
        taux = 1.0
        for (annee_fin, t) in sorted(rampup_config, key=lambda x: x[0]):
            if annee <= annee_fin:
                taux = t
                break
        coeffs.append(taux)
    return coeffs


# =============================================================================
# MOTEUR DE CALCUL FINANCIER
# =============================================================================

def calculer_wacc(part_dette, cout_dette, taux_is, part_fonds_propres, cout_fonds_propres):
    """WACC avec bouclier fiscal : Kd*(1-IS)*(D/V) + Ke*(E/V)."""
    return (cout_dette * (1 - taux_is) * part_dette) + (cout_fonds_propres * part_fonds_propres)


def projeter_fcff(
    quantite, prix_vente, cout_production,
    capex_initial, taux_is, taux_inflation,
    duree, coeffs_rampup, ratio_maintenance=0.02
):
    """
    Projection des FCFF avec courbe de ramp-up.

    Logique annuelle :
      Revenus = Quantite * CoeffRampup * Prix * Facteur_Inflation
      Couts   = Quantite * CoeffRampup * CoutProd * Facteur_Inflation
      EBITDA  = Revenus - Couts
      Amort.  = CAPEX / Duree  (lineaire, independant du ramp-up)
      EBIT    = EBITDA - Amort.
      IS      = max(EBIT * taux_IS, 0)
      NOPAT   = EBIT - IS
      FCFF    = NOPAT + Amort. - CAPEX_maintenance
    """
    amortissement_annuel = capex_initial / duree
    capex_maintenance    = capex_initial * ratio_maintenance
    resultats = []

    for annee in range(1, duree + 1):
        coeff          = coeffs_rampup[annee - 1]
        fact_inflation = (1 + taux_inflation) ** (annee - 1)

        revenus = quantite * coeff * prix_vente      * fact_inflation
        couts   = quantite * coeff * cout_production * fact_inflation
        ebitda  = revenus - couts
        ebit    = ebitda - amortissement_annuel
        impot   = max(ebit * taux_is, 0)
        nopat   = ebit - impot
        fcff    = nopat + amortissement_annuel - capex_maintenance

        resultats.append({
            "Annee"                  : annee,
            "Ramp-up (%)"            : round(coeff * 100, 0),
            "Revenus (MFCFA)"        : round(revenus / 1e6, 2),
            "Couts Prod. (MFCFA)"   : round(couts / 1e6, 2),
            "EBITDA (MFCFA)"        : round(ebitda / 1e6, 2),
            "Amortissement (MFCFA)" : round(amortissement_annuel / 1e6, 2),
            "EBIT (MFCFA)"          : round(ebit / 1e6, 2),
            "Impot IS (MFCFA)"      : round(impot / 1e6, 2),
            "FCFF (MFCFA)"          : round(fcff / 1e6, 2),
        })

    return resultats


def calculer_metriques(fcff_liste, capex_initial, wacc):
    """VAN, TRI et Payback a partir des FCFF."""
    flux    = [-capex_initial] + [r["FCFF (MFCFA)"] * 1e6 for r in fcff_liste]
    van_val = npv(wacc, flux)
    tri_val = irr(flux)

    cumul   = 0
    payback = None
    for i, f in enumerate(flux[1:], 1):
        cumul += f
        if cumul >= capex_initial:
            payback = i
            break

    return {
        "VAN (MFCFA)"     : round(van_val / 1e6, 2),
        "TRI (%)"         : round(tri_val * 100, 2) if tri_val is not None else None,
        "Payback (annees)": payback,
    }


# =============================================================================
# PILIER 2 — SERVICE DE LA DETTE, DSCR ET LLCR
# =============================================================================

def calculer_service_dette(
    fcff_liste, capex_initial, part_dette,
    cout_dette, duree_dette, grace_period, wacc
):
    """
    Tableau annuel du service de la dette avec DSCR et LLCR.

    Grace Period : pendant grace_period annees, seuls les interets sont dus.
    Apres grace period : remboursement lineaire du capital sur
    (duree_dette - grace_period) annees.

    DSCR annuel = FCFF / Service_Dette_annuel
    LLCR annuel = NPV(FCFF futurs sur duree pret) / Encours

    Returns
    -------
    df        : pd.DataFrame - Tableau annuel complet
    dscr_min  : float        - DSCR minimum sur la duree du pret
    dscr_moy  : float        - DSCR moyen sur la duree du pret
    llcr_global : float      - LLCR calcule a t=0
    """
    montant_dette        = capex_initial * part_dette
    annees_remboursement = duree_dette - grace_period
    amort_capital        = (montant_dette / annees_remboursement
                            if annees_remboursement > 0 else 0)

    rows    = []
    encours = montant_dette

    for i, r in enumerate(fcff_liste):
        annee = r["Annee"]
        fcff  = r["FCFF (MFCFA)"] * 1e6

        if annee <= duree_dette:
            interets = encours * cout_dette

            if annee <= grace_period:
                remboursement_capital = 0.0
            else:
                remboursement_capital = amort_capital

            service_dette = interets + remboursement_capital
            encours_fin   = encours - remboursement_capital
            dscr_an       = fcff / service_dette if service_dette > 0 else np.nan

            # LLCR : NPV des FCFF restants / encours courant
            fcff_futurs = [fcff_liste[j]["FCFF (MFCFA)"] * 1e6
                           for j in range(i, min(duree_dette, len(fcff_liste)))]
            van_futurs  = npv(cout_dette, fcff_futurs) if fcff_futurs else 0.0
            llcr_an     = van_futurs / encours if encours > 0 else np.nan

            rows.append({
                "Annee"                  : annee,
                "Encours Debut (MFCFA)"  : round(encours / 1e6, 2),
                "Interets (MFCFA)"       : round(interets / 1e6, 2),
                "Remb. Capital (MFCFA)"  : round(remboursement_capital / 1e6, 2),
                "Service Total (MFCFA)"  : round(service_dette / 1e6, 2),
                "FCFF (MFCFA)"           : round(fcff / 1e6, 2),
                "DSCR"                   : round(dscr_an, 2) if not np.isnan(dscr_an) else np.nan,
                "LLCR"                   : round(llcr_an, 2) if not np.isnan(llcr_an) else np.nan,
                "Grace Period"           : (annee <= grace_period),
            })
            encours = encours_fin
        else:
            rows.append({
                "Annee"                  : annee,
                "Encours Debut (MFCFA)"  : 0.0,
                "Interets (MFCFA)"       : 0.0,
                "Remb. Capital (MFCFA)"  : 0.0,
                "Service Total (MFCFA)"  : 0.0,
                "FCFF (MFCFA)"           : round(fcff / 1e6, 2),
                "DSCR"                   : np.nan,
                "LLCR"                   : np.nan,
                "Grace Period"           : False,
            })

    df       = pd.DataFrame(rows)
    dscrs    = df["DSCR"].dropna()
    dscr_min = round(dscrs.min(), 2)  if len(dscrs) else np.nan
    dscr_moy = round(dscrs.mean(), 2) if len(dscrs) else np.nan

    # LLCR global a t=0
    fcff_pour_llcr = [r["FCFF (MFCFA)"] * 1e6 for r in fcff_liste[:duree_dette]]
    llcr_global    = (round(npv(cout_dette, fcff_pour_llcr) / montant_dette, 2)
                      if montant_dette > 0 else np.nan)

    return df, dscr_min, dscr_moy, llcr_global


# =============================================================================
# PILIER 3 — TORNADO CHART
# =============================================================================

def tornado_van(
    quantite, prix_vente, cout_production, capex_initial,
    taux_is, taux_inflation, duree, wacc, coeffs_rampup,
    delta=0.10
):
    """
    Impact sur la VAN d'une variation de +/- delta (defaut 10%)
    sur chacune des 4 variables cles : Prix, Quantite, CAPEX, OPEX.

    Returns
    -------
    list of dict : Trie par amplitude decroissante
    float        : VAN de reference
    """
    def calc_van(**kw):
        fl = projeter_fcff(**kw)
        return calculer_metriques(fl, kw["capex_initial"], wacc)["VAN (MFCFA)"]

    base_kw = dict(
        quantite=quantite, prix_vente=prix_vente,
        cout_production=cout_production, capex_initial=capex_initial,
        taux_is=taux_is, taux_inflation=taux_inflation,
        duree=duree, coeffs_rampup=coeffs_rampup
    )
    van_ref = calc_van(**base_kw)

    sensibilites = [
        ("Prix de Vente",     "prix_vente",       prix_vente),
        ("Quantite Produite", "quantite",          quantite),
        ("CAPEX",             "capex_initial",     capex_initial),
        ("OPEX (Cout Prod.)", "cout_production",   cout_production),
    ]

    resultats = []
    for label, param, val_base in sensibilites:
        kw_m = {**base_kw, param: val_base * (1 - delta)}
        kw_p = {**base_kw, param: val_base * (1 + delta)}
        van_m = calc_van(**kw_m)
        van_p = calc_van(**kw_p)
        resultats.append({
            "Variable"    : label,
            "VAN_moins"   : van_m,
            "VAN_plus"    : van_p,
            "Impact_moins": van_m - van_ref,
            "Impact_plus" : van_p - van_ref,
            "Amplitude"   : abs(van_p - van_m),
        })

    resultats.sort(key=lambda x: x["Amplitude"], reverse=True)
    return resultats, van_ref


def tracer_tornado(tornado_data, van_ref):
    """Trace le graphique Tornado horizontal."""
    labels    = [d["Variable"]      for d in tornado_data]
    impacts_m = [d["Impact_moins"]  for d in tornado_data]
    impacts_p = [d["Impact_plus"]   for d in tornado_data]

    fig, ax = plt.subplots(figsize=(8, max(3, len(labels) * 0.9 + 1)))
    y = np.arange(len(labels))

    for i, (im, ip) in enumerate(zip(impacts_m, impacts_p)):
        ax.barh(i, im, height=0.5, color="#dc3545", alpha=0.82, label="-10%" if i == 0 else "")
        ax.barh(i, ip, height=0.5, color="#198754", alpha=0.82, label="+10%" if i == 0 else "")
        offset = max(abs(im), abs(ip)) * 0.03
        ax.text(im - offset, i, f"{im:+,.0f}", va="center", ha="right", fontsize=8, color="#dc3545")
        ax.text(ip + offset, i, f"{ip:+,.0f}", va="center", ha="left",  fontsize=8, color="#198754")

    ax.axvline(0, color="#24292f", linewidth=0.9)
    ax.set_yticks(y)
    ax.set_yticklabels(labels, fontsize=9)
    ax.set_xlabel("Impact sur la VAN (MFCFA)", fontsize=9)
    ax.set_title(f"Tornado — Sensibilite de la VAN (base : {van_ref:,.0f} MFCFA)",
                 fontsize=10, fontweight="bold", pad=10)
    ax.legend(fontsize=8, framealpha=0.6, loc="lower right")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"{x:,.0f}"))
    plt.tight_layout()
    return fig


# =============================================================================
# PILIER 4 — MONTE CARLO AVEC RISQUE DE QUEUE
# =============================================================================

def simulation_monte_carlo(
    quantite, prix_vente, cout_production,
    capex_initial, taux_is, taux_inflation,
    duree, wacc, coeffs_rampup,
    part_dette, cout_dette, duree_dette, grace_period,
    sigma_prix, sigma_quantite,
    prob_choc, amplitude_choc,
    n_iterations=1000
):
    """
    Monte Carlo avec distribution a queue epaisse (risque de queue).

    A chaque iteration :
      - Prix de Vente ~ N(mu, sigma_prix)
      - Quantite      ~ N(mu, sigma_quantite), puis multipliee par (1 - amplitude_choc)
        avec probabilite prob_choc (choc climatique de Bernoulli)

    Ce mecanisme produit une distribution bimodale refletant les projets agricoles
    exposes aux aleas climatiques extremes.

    Returns
    -------
    vans_arr      : np.array - VAN simulees (MFCFA)
    prob_van_neg  : float    - P(VAN < 0) en %
    prob_dscr_def : float    - P(DSCR_min < 1) en %
    n_chocs       : int      - Nombre de chocs survenu dans la simulation
    """
    vans       = []
    n_dscr_def = 0
    n_chocs    = 0

    for _ in range(n_iterations):
        choc      = np.random.rand() < prob_choc
        mult_choc = (1 - amplitude_choc) if choc else 1.0
        if choc:
            n_chocs += 1

        prix_sim = np.random.normal(prix_vente, prix_vente * sigma_prix)
        qte_sim  = max(
            np.random.normal(quantite, quantite * sigma_quantite) * mult_choc,
            0.0
        )

        fl  = projeter_fcff(qte_sim, prix_sim, cout_production,
                            capex_initial, taux_is, taux_inflation,
                            duree, coeffs_rampup)
        met = calculer_metriques(fl, capex_initial, wacc)
        vans.append(met["VAN (MFCFA)"])

        _, dscr_min_sim, _, _ = calculer_service_dette(
            fl, capex_initial, part_dette,
            cout_dette, duree_dette, grace_period, wacc
        )
        if not np.isnan(dscr_min_sim) and dscr_min_sim < 1.0:
            n_dscr_def += 1

    vans_arr      = np.array(vans)
    prob_van_neg  = round((vans_arr < 0).mean() * 100, 1)
    prob_dscr_def = round((n_dscr_def / n_iterations) * 100, 1)
    return vans_arr, prob_van_neg, prob_dscr_def, n_chocs


# =============================================================================
# TABLEAU CROISE DE SENSIBILITE (conserve de v1)
# =============================================================================

def analyse_sensibilite_van(
    quantite, prix_vente, cout_production,
    capex_initial, taux_is, taux_inflation,
    duree, wacc, coeffs_rampup,
    variations=(-0.10, -0.05, 0.0, 0.05, 0.10)
):
    """Tableau croise Prix x Cout de Production sur la VAN."""
    index_labels = [f"{int(v*100):+d}%" for v in variations]
    data = {}
    for v_cout in variations:
        col      = []
        cout_adj = cout_production * (1 + v_cout)
        for v_prix in variations:
            prix_adj = prix_vente * (1 + v_prix)
            fl = projeter_fcff(quantite, prix_adj, cout_adj,
                               capex_initial, taux_is, taux_inflation,
                               duree, coeffs_rampup)
            col.append(calculer_metriques(fl, capex_initial, wacc)["VAN (MFCFA)"])
        data[f"Cout {int(v_cout*100):+d}%"] = col
    return pd.DataFrame(data, index=[f"Prix {l}" for l in index_labels])


# =============================================================================
# BARRE LATERALE — INPUTS
# =============================================================================

with st.sidebar:
    st.markdown("## CommodityWatch")
    st.markdown("**Pre-Scoring Agro-Industriel v2.0**")
    st.markdown("---")

    st.markdown("### Parametres de Marche")

    matiere_premiere = st.selectbox(
        "Matiere Premiere",
        options=["Cafe", "Cacao", "Anacarde", "Soja"],
    )
    quantite = st.number_input(
        "Quantite a Pleine Capacite (Tonnes/an)",
        min_value=0, value=5000, step=100,
    )
    prix_vente = st.number_input(
        "Prix de Vente Unitaire (FCFA/Tonne)",
        min_value=0, value=900_000, step=10_000,
    )
    cout_production = st.number_input(
        "Cout de Production Unitaire (FCFA/Tonne)",
        min_value=0, value=600_000, step=10_000,
    )

    st.markdown("---")
    st.markdown("### Courbe de Montee en Puissance (Ramp-up)")
    st.caption(
        "Definissez le taux de production (%) pour chaque palier d'annees. "
        "Au-dela du dernier palier defini, 100% est applique automatiquement."
    )

    n_paliers = st.number_input(
        "Nombre de paliers de ramp-up",
        min_value=1, max_value=5, value=3, step=1
    )
    rampup_config = []
    defaults_an = [2, 3, 4, 5, 6]
    defaults_tx = [0, 40, 80, 100, 100]
    for k in range(int(n_paliers)):
        c1, c2 = st.columns(2)
        with c1:
            an_fin = st.number_input(
                f"Palier {k+1} — Fin annee",
                min_value=1, max_value=25,
                value=defaults_an[k],
                step=1, key=f"ru_an_{k}"
            )
        with c2:
            taux_ru = st.number_input(
                "Taux (%)",
                min_value=0, max_value=100,
                value=defaults_tx[k],
                step=5, key=f"ru_tx_{k}"
            )
        rampup_config.append((int(an_fin), taux_ru / 100))

    st.markdown("---")
    st.markdown("### Structure du Capital")

    capex_initial = st.number_input(
        "CAPEX Initial Total (FCFA)",
        min_value=0, value=2_000_000_000, step=100_000_000, format="%d",
    )
    part_dette_pct     = st.slider("Part de Dette (% du CAPEX)", 0, 100, 60, 5)
    part_dette         = part_dette_pct / 100
    part_fonds_propres = 1 - part_dette

    cout_dette_pct = st.slider("Cout de la Dette (% Annuel)", 1, 25, 9, 1)
    cout_dette     = cout_dette_pct / 100

    cout_fp_pct        = st.slider("Cout des Fonds Propres - CAPM (% Annuel)", 5, 30, 15, 1)
    cout_fonds_propres = cout_fp_pct / 100

    taux_is_pct = st.slider("Taux d'Imposition (IS %)", 0, 40, 25, 1)
    taux_is     = taux_is_pct / 100

    duree = st.slider("Duree du Projet (Annees)", 3, 25, 10, 1)

    taux_inflation_pct = st.slider("Inflation Annuelle Moyenne (%)", 0, 15, 3, 1)
    taux_inflation     = taux_inflation_pct / 100

    st.markdown("---")
    st.markdown("### Parametres de la Dette")

    duree_dette = st.slider(
        "Duree du Pret (Annees)",
        min_value=1, max_value=duree, value=min(8, duree), step=1,
    )
    grace_period = st.slider(
        "Differe de Remboursement — Grace Period (Annees)",
        min_value=0, max_value=max(0, duree_dette - 1), value=min(2, duree_dette - 1), step=1,
    )

    st.markdown("---")
    st.markdown("### Simulation Monte Carlo")

    sigma_prix_pct = st.slider("Ecart-type Prix de Vente (%)", 1, 30, 10, 1)
    sigma_prix     = sigma_prix_pct / 100

    sigma_qte_pct  = st.slider("Ecart-type Quantite Produite (%)", 1, 30, 8, 1)
    sigma_quantite = sigma_qte_pct / 100

    st.markdown("**Risque de Queue — Choc Climatique**")
    prob_choc_pct = st.slider(
        "Probabilite de choc annuel (%)", 1, 30, 5, 1,
        help="Probabilite qu'un evenement climatique severe survienne."
    )
    prob_choc = prob_choc_pct / 100

    amplitude_choc_pct = st.slider(
        "Amplitude du choc sur la Quantite (%)", 10, 80, 40, 5,
        help="Reduction de la production en cas de choc (ex: 40% = perte de 40% du rendement)."
    )
    amplitude_choc = amplitude_choc_pct / 100

    n_iterations = st.selectbox(
        "Nombre d'Iterations", options=[500, 1000, 2000, 5000], index=1,
    )

    st.markdown("---")
    lancer = st.button("Lancer l'Analyse", type="primary", use_container_width=True)


# =============================================================================
# ZONE PRINCIPALE — HEADER
# =============================================================================

st.markdown(f"## Tableau de Bord Pre-Scoring — {matiere_premiere}")
st.markdown(
    "Outil d'analyse financiere rapide pour evaluer la viabilite d'un projet "
    "agro-industriel avant modelisation complete."
)
st.markdown("---")

if not lancer:
    st.info(
        "Renseignez les parametres dans la barre laterale puis cliquez sur "
        "**Lancer l'Analyse** pour afficher les resultats."
    )
    st.stop()


# =============================================================================
# CALCULS PRINCIPAUX
# =============================================================================

with st.spinner("Calculs en cours..."):

    coeffs_rampup = construire_rampup(duree, rampup_config)

    wacc = calculer_wacc(part_dette, cout_dette, taux_is,
                         part_fonds_propres, cout_fonds_propres)

    fcff_liste = projeter_fcff(
        quantite, prix_vente, cout_production,
        capex_initial, taux_is, taux_inflation,
        duree, coeffs_rampup
    )

    metriques = calculer_metriques(fcff_liste, capex_initial, wacc)

    df_dette, dscr_min, dscr_moy, llcr_global = calculer_service_dette(
        fcff_liste, capex_initial, part_dette,
        cout_dette, duree_dette, grace_period, wacc
    )

    tornado_data, van_ref = tornado_van(
        quantite, prix_vente, cout_production, capex_initial,
        taux_is, taux_inflation, duree, wacc, coeffs_rampup
    )

    df_sensibilite = analyse_sensibilite_van(
        quantite, prix_vente, cout_production,
        capex_initial, taux_is, taux_inflation,
        duree, wacc, coeffs_rampup
    )

    vans_mc, prob_van_neg, prob_dscr_def, n_chocs_sim = simulation_monte_carlo(
        quantite, prix_vente, cout_production,
        capex_initial, taux_is, taux_inflation,
        duree, wacc, coeffs_rampup,
        part_dette, cout_dette, duree_dette, grace_period,
        sigma_prix, sigma_quantite,
        prob_choc, amplitude_choc,
        n_iterations
    )


# =============================================================================
# SECTION 1 — KPIs
# =============================================================================

st.markdown("## Indicateurs Cles de Performance")

van_val     = metriques["VAN (MFCFA)"]
tri_val     = metriques["TRI (%)"]
payback_val = metriques["Payback (annees)"]

col1, col2, col3, col4, col5, col6 = st.columns(6)

with col1:
    st.metric("VAN",
              f"{van_val:,.0f} MFCFA",
              delta="Positive" if van_val > 0 else "Negative",
              delta_color="normal" if van_val > 0 else "inverse")
with col2:
    st.metric("TRI",
              f"{tri_val:.2f} %" if tri_val is not None else "N/D",
              delta=f"WACC : {wacc*100:.2f} %",
              delta_color="normal" if (tri_val or 0) > wacc * 100 else "inverse")
with col3:
    st.metric("WACC", f"{wacc*100:.2f} %")
with col4:
    st.metric("DSCR Min",
              f"{dscr_min:.2f}x" if not np.isnan(dscr_min) else "N/D",
              delta="Bancable" if dscr_min >= 1.3 else "Sous seuil 1.3x",
              delta_color="normal" if dscr_min >= 1.3 else "inverse")
with col5:
    st.metric("LLCR",
              f"{llcr_global:.2f}x" if not np.isnan(llcr_global) else "N/D",
              delta="Satisfaisant" if llcr_global >= 1.1 else "Insuffisant",
              delta_color="normal" if llcr_global >= 1.1 else "inverse")
with col6:
    st.metric("Payback",
              f"{payback_val} ans" if payback_val else "> Duree Projet")

st.markdown("---")


# =============================================================================
# SECTION 2 — PROJECTION FCFF + RAMP-UP
# =============================================================================

st.markdown("## Projection des Flux de Tresorerie (FCFF) avec Ramp-up")
st.caption("Montants en MFCFA. Colonne Ramp-up (%) = taux de capacite applique chaque annee.")

df_fcff = pd.DataFrame(fcff_liste)

def colorer_fcff(val):
    if isinstance(val, (int, float)):
        return "background-color: #d4edda;" if val >= 0 else "background-color: #f8d7da;"
    return ""

def colorer_rampup(val):
    if isinstance(val, (int, float)) and val < 100:
        return "background-color: #fff3cd; color: #856404;"
    return ""

styled_fcff = (
    df_fcff.style
    .applymap(colorer_fcff,   subset=["FCFF (MFCFA)"])
    .applymap(colorer_rampup, subset=["Ramp-up (%)"])
    .format({col: "{:.2f}" for col in df_fcff.columns if col != "Annee"})
)
st.dataframe(styled_fcff, use_container_width=True, hide_index=True)

st.markdown("---")


# =============================================================================
# SECTION 3 — SERVICE DE LA DETTE, DSCR & LLCR
# =============================================================================

st.markdown("## Bancabilite — Service de la Dette, DSCR et LLCR")
st.caption(
    f"Seuil d'alerte DSCR : **1.3x**  |  Grace Period : **{grace_period} an(s)**  |  "
    f"Duree du pret : **{duree_dette} an(s)**"
)

cols_affichage = [c for c in df_dette.columns if c != "Grace Period"]
df_affichage   = df_dette[cols_affichage].copy()

def styler_dette(row):
    styles = [""] * len(row)
    annee  = row["Annee"]
    grace  = df_dette.loc[df_dette["Annee"] == annee, "Grace Period"].values
    is_grace = len(grace) > 0 and grace[0]

    if is_grace:
        return [f"background-color: #fff3cd; color: #856404;"] * len(row)

    dscr_val = row.get("DSCR", np.nan)
    if pd.notna(dscr_val):
        if dscr_val < 1.0:
            return ["background-color: #f8d7da; color: #721c24;"] * len(row)
        elif dscr_val < 1.3:
            return ["background-color: #fde8c8; color: #7d4e0f;"] * len(row)
        else:
            return ["background-color: #d4edda; color: #155724;"] * len(row)

    return styles

styled_dette = (
    df_affichage.style
    .apply(styler_dette, axis=1)
    .format({
        "Encours Debut (MFCFA)" : "{:.2f}",
        "Interets (MFCFA)"      : "{:.2f}",
        "Remb. Capital (MFCFA)" : "{:.2f}",
        "Service Total (MFCFA)" : "{:.2f}",
        "FCFF (MFCFA)"          : "{:.2f}",
        "DSCR" : lambda x: f"{x:.2f}x" if pd.notna(x) else "—",
        "LLCR" : lambda x: f"{x:.2f}x" if pd.notna(x) else "—",
    })
)
st.dataframe(styled_dette, use_container_width=True, hide_index=True)

leg1, leg2, leg3, leg4 = st.columns(4)
leg1.markdown("<span style='background:#d4edda;padding:2px 8px;border-radius:4px;font-size:0.8rem;'>DSCR &ge; 1.3x</span>", unsafe_allow_html=True)
leg2.markdown("<span style='background:#fde8c8;padding:2px 8px;border-radius:4px;font-size:0.8rem;'>1.0 &le; DSCR &lt; 1.3x</span>", unsafe_allow_html=True)
leg3.markdown("<span style='background:#f8d7da;padding:2px 8px;border-radius:4px;font-size:0.8rem;'>DSCR &lt; 1.0 — Defaut</span>", unsafe_allow_html=True)
leg4.markdown("<span style='background:#fff3cd;padding:2px 8px;border-radius:4px;font-size:0.8rem;'>Grace Period</span>", unsafe_allow_html=True)

st.markdown("---")


# =============================================================================
# SECTION 4 — TORNADO CHART
# =============================================================================

st.markdown("## Analyse de Sensibilite — Tornado Chart")
st.caption(
    "Variation de +/- 10% appliquee separement sur chaque variable cle. "
    "L'amplitude de la barre reflete l'elasticite de la VAN."
)

col_t1, col_t2 = st.columns([2, 1])

with col_t1:
    fig_tornado = tracer_tornado(tornado_data, van_ref)
    st.pyplot(fig_tornado)
    plt.close(fig_tornado)

with col_t2:
    st.markdown("#### Tableau des Impacts (MFCFA)")
    df_tornado = pd.DataFrame([
        {
            "Variable"    : d["Variable"],
            "Impact -10%" : f"{d['Impact_moins']:+,.0f}",
            "Impact +10%" : f"{d['Impact_plus']:+,.0f}",
            "Amplitude"   : f"{d['Amplitude']:,.0f}",
        }
        for d in tornado_data
    ])
    st.dataframe(df_tornado, use_container_width=True, hide_index=True)

st.markdown("---")


# =============================================================================
# SECTION 5 — TABLEAU CROISE DE SENSIBILITE
# =============================================================================

st.markdown("## Tableau Croise de Sensibilite de la VAN (MFCFA)")
st.caption("Prix de Vente (lignes) x Cout de Production (colonnes). Cellules colorees par magnitude.")

def colorer_sensibilite(val):
    if isinstance(val, (int, float)):
        max_abs = max(abs(df_sensibilite.values.max()), abs(df_sensibilite.values.min()), 1e-9)
        intensite = min(int(abs(val) / max_abs * 80), 80)
        if val > 0:
            return f"background-color: rgba(25, 135, 84, {intensite/100+0.1}); color: white;"
        else:
            return f"background-color: rgba(220, 53, 69, {intensite/100+0.1}); color: white;"
    return ""

styled_sens = df_sensibilite.style.applymap(colorer_sensibilite).format("{:.0f}")
st.dataframe(styled_sens, use_container_width=True)

st.markdown("---")


# =============================================================================
# SECTION 6 — MONTE CARLO AVEC RISQUE DE QUEUE
# =============================================================================

st.markdown("## Analyse de Risque — Monte Carlo avec Risque de Queue")
st.caption(
    f"{n_iterations:,} iterations  |  Choc climatique : prob. {prob_choc_pct}%, "
    f"amplitude -{amplitude_choc_pct}% sur la quantite  |  "
    f"Chocs survenus : {n_chocs_sim} ({n_chocs_sim/n_iterations*100:.1f}%)"
)

col_mc1, col_mc2, col_mc3 = st.columns([2, 1, 1])

with col_mc1:
    fig_mc, ax_mc = plt.subplots(figsize=(7, 3.8))
    ax_mc.hist(vans_mc, bins=60, color="#0d6efd", alpha=0.72,
               edgecolor="white", linewidth=0.3, label="Distribution VAN")
    ax_mc.axvline(0, color="#dc3545", linewidth=1.5, linestyle="--", label="VAN = 0")
    ax_mc.axvline(np.percentile(vans_mc, 5), color="#fd7e14", linewidth=1.2,
                  linestyle=":", label=f"Pct. 5% : {np.percentile(vans_mc, 5):,.0f}")
    ax_mc.axvline(np.mean(vans_mc), color="#198754", linewidth=1.2,
                  linestyle="-.", label=f"Moyenne : {np.mean(vans_mc):,.0f}")
    ax_mc.set_xlabel("VAN (MFCFA)", fontsize=9)
    ax_mc.set_ylabel("Frequence", fontsize=9)
    ax_mc.set_title("Distribution des VAN Simulees (avec risque de queue)",
                    fontsize=10, fontweight="bold", pad=10)
    ax_mc.legend(fontsize=8, framealpha=0.6)
    ax_mc.xaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"{x:,.0f}"))
    ax_mc.spines["top"].set_visible(False)
    ax_mc.spines["right"].set_visible(False)
    plt.tight_layout()
    st.pyplot(fig_mc)
    plt.close(fig_mc)

with col_mc2:
    st.markdown("#### Probabilites de Defaut")
    st.metric("P(VAN < 0)", f"{prob_van_neg} %",
              delta="Risque eleve" if prob_van_neg > 30 else "Risque maitrise",
              delta_color="inverse" if prob_van_neg > 30 else "normal")
    st.metric("P(DSCR_min < 1)", f"{prob_dscr_def} %",
              delta="Risque eleve" if prob_dscr_def > 30 else "Risque maitrise",
              delta_color="inverse" if prob_dscr_def > 30 else "normal")
    st.metric("Freq. Choc Climatique",
              f"{n_chocs_sim/n_iterations*100:.1f} %",
              delta=f"Cible : {prob_choc_pct} %",
              delta_color="off")

with col_mc3:
    st.markdown("#### Statistiques VAN")
    stats_mc = {
        "Statistique"    : ["Moyenne", "Mediane", "Ecart-type",
                            "Pct. 5%", "Pct. 25%", "Pct. 75%", "Pct. 95%"],
        "Valeur (MFCFA)" : [
            f"{np.mean(vans_mc):,.0f}",
            f"{np.median(vans_mc):,.0f}",
            f"{np.std(vans_mc):,.0f}",
            f"{np.percentile(vans_mc,  5):,.0f}",
            f"{np.percentile(vans_mc, 25):,.0f}",
            f"{np.percentile(vans_mc, 75):,.0f}",
            f"{np.percentile(vans_mc, 95):,.0f}",
        ],
    }
    st.dataframe(pd.DataFrame(stats_mc), use_container_width=True, hide_index=True)

st.markdown("---")
st.markdown(
    "<div style='font-size:0.75rem; color:#57606a;'>"
    "CommodityWatch v2.0 — Outil de pre-scoring interne. "
    "Les resultats sont des estimations indicatives et ne constituent pas un avis financier definitif."
    "</div>",
    unsafe_allow_html=True,
)
