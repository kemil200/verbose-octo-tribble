# =============================================================================
# CommodityWatch v3.0 — Standard Institutionnel International
# Framework : Streamlit + Plotly | Python 3.9+
# Usage     : streamlit run app.py
#
# Modules v3.0 :
#   M1 — Sélecteur de Business Model (Campagne / Production / Infrastructure)
#   M2 — Internationalisation USD/FCFA + Import Excel + Prix indexés marchés
#   M3 — Moteur de Dette Avancé (commissions flat & engagement, alerte charges)
#   M4 — Stress Test discret (3 événements) + Couverture Assurance
#   M5 — Visualisation Plotly : impact Sans/Avec Assurance sur DSCR & Trésorerie
# =============================================================================

import io
import streamlit as st
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots

# ── Dépendance optionnelle pour l'import Excel ──────────────────────────────
try:
    import openpyxl  # noqa: F401
    EXCEL_OK = True
except ImportError:
    EXCEL_OK = False


# =============================================================================
# CONSTANTES GLOBALES
# =============================================================================

TAUX_CHANGE_FCFA_USD = 600.0   # 1 USD ≈ 600 FCFA (ajustable)

# Prix de référence internationaux indicatifs (USD/Tonne)
# Source : ICE Futures / Euronext — à mettre à jour manuellement
PRIX_MARCHE_USD = {
    "Café (Arabica)"  : 4_200,   # ICE C contract
    "Café (Robusta)"  : 2_800,   # Euronext Robusta
    "Cacao"           : 7_800,   # ICE Cocoa
    "Anacarde (brut)" : 1_200,   # Marché OTC Afrique de l'Ouest
    "Soja"            :   430,   # CBOT Soja
    "Coton"           :   750,   # ICE Cotton
    "Caoutchouc"      : 1_450,   # SGX RSS3
    "Personnalisé"    :     0,
}

# Seuil d'alerte charges / CA (ratio coûts opérationnels)
SEUIL_CHARGES_CA = 0.55

# Couleurs institutionnelles
CLR_VERT  = "#198754"
CLR_ROUGE = "#dc3545"
CLR_AMBER = "#fd7e14"
CLR_BLEU  = "#0d6efd"
CLR_GRIS  = "#6c757d"


# =============================================================================
# UTILITAIRES MATHÉMATIQUES (sans dépendance numpy_financial)
# =============================================================================

def npv_calc(rate: float, cashflows: list) -> float:
    """VAN : ∑ CF_t / (1+r)^t  pour t = 0..n."""
    cf = np.asarray(cashflows, dtype=float)
    t  = np.arange(len(cf))
    return float(np.sum(cf / (1 + rate) ** t))


def irr_calc(cashflows: list, tol: float = 1e-7, max_iter: int = 1500) -> float | None:
    """TRI par Newton-Raphson. Retourne None si pas de convergence."""
    cf = np.asarray(cashflows, dtype=float)
    if not (np.any(cf < 0) and np.any(cf > 0)):
        return None
    rate = 0.10
    for _ in range(max_iter):
        t   = np.arange(len(cf))
        f   = np.sum(cf / (1 + rate) ** t)
        df  = np.sum(-t * cf / (1 + rate) ** (t + 1))
        if df == 0:
            return None
        new = rate - f / df
        if abs(new - rate) < tol:
            return new
        rate = new
    return None


def convertir(valeur: float, vers: str) -> float:
    """Convertit entre FCFA et USD selon le sens demandé."""
    if vers == "USD":
        return valeur / TAUX_CHANGE_FCFA_USD
    return valeur * TAUX_CHANGE_FCFA_USD   # vers FCFA


def fmt_montant(valeur: float, devise: str, millions: bool = True) -> str:
    """Formate un montant avec l'unité correcte."""
    unite = f"M{devise}" if millions else devise
    return f"{valeur:,.2f} {unite}"


# =============================================================================
# M1 — RAMP-UP
# =============================================================================

def construire_rampup(duree: int, paliers: list) -> list:
    """
    Construit le vecteur de coefficients de production annuels [0..1].
    paliers : liste de (annee_fin, taux_decimal)
    """
    coeffs = []
    for annee in range(1, duree + 1):
        taux = 1.0
        for (fin, t) in sorted(paliers, key=lambda x: x[0]):
            if annee <= fin:
                taux = t
                break
        coeffs.append(taux)
    return coeffs


# =============================================================================
# M1 — PARAMÈTRES SPÉCIFIQUES PAR BUSINESS MODEL
# =============================================================================

def parametres_bfr(business_model: str, revenus_annuels: float,
                   couts_annuels: float, jours_stock: int,
                   jours_clients: int, jours_fournisseurs: int) -> float:
    """
    Calcule la variation de BFR selon le Business Model.

    Campagne     : BFR fort (stock matière + crédit clients)
    Production   : BFR faible (vente spot à la récolte)
    Infrastructure: BFR négligeable (paiements contractuels)

    BFR = Stock + Clients - Fournisseurs
    """
    if business_model == "Campagne":
        stock         = couts_annuels  * jours_stock         / 365
        clients       = revenus_annuels * jours_clients       / 365
        fournisseurs  = couts_annuels  * jours_fournisseurs   / 365
        return stock + clients - fournisseurs
    elif business_model == "Production":
        return revenus_annuels * 0.05   # BFR résiduel 5 % des revenus
    else:   # Infrastructure
        return 0.0


def ajustement_biologique(annee: int, duree_croissance: int) -> float:
    """
    Pour le modèle Production (plantation) : les actifs biologiques
    (arbres, plants) ne génèrent aucun revenu pendant la phase de croissance.
    Retourne 0 pendant la phase, 1 ensuite.
    """
    return 0.0 if annee <= duree_croissance else 1.0


# =============================================================================
# M2 — PRIX INDEXÉS MARCHÉS INTERNATIONAUX
# =============================================================================

def prix_reference_marche(matiere: str, devise: str) -> float:
    """
    Retourne le prix de référence de marché en devise choisie.
    Les prix de base sont en USD/Tonne.
    """
    prix_usd = PRIX_MARCHE_USD.get(matiere, 0.0)
    if devise == "USD":
        return prix_usd
    return prix_usd * TAUX_CHANGE_FCFA_USD


def parser_excel_financier(fichier) -> dict:
    """
    Importe un fichier Excel et tente de mapper les postes standards
    du Bilan et du Compte de Résultat.

    Colonnes attendues (nom flexible, détection par mots-clés) :
      Compte de Résultat : Chiffre d'affaires, Charges d'exploitation,
                           Amortissements, Résultat net
      Bilan              : Actif immobilisé, Capitaux propres, Dettes financières

    Retourne un dict avec les valeurs extraites ou None si non trouvé.
    """
    try:
        xls  = pd.ExcelFile(fichier)
        data = {}

        MAPPING = {
            "ca"          : ["chiffre d'affaires", "revenus", "turnover", "sales"],
            "charges"     : ["charges d'exploitation", "opex", "couts", "operating expenses"],
            "amort"       : ["amortissement", "depreciation", "dotation"],
            "resultat_net": ["résultat net", "net income", "bénéfice net", "profit net"],
            "actif_immo"  : ["actif immobilisé", "immobilisations", "fixed assets"],
            "capitaux_propres": ["capitaux propres", "equity", "fonds propres"],
            "dettes_fin"  : ["dettes financières", "emprunts", "financial debt"],
        }

        for sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet, header=None)
            # Cherche dans la première colonne (libellés) des mots-clés
            for idx, row in df.iterrows():
                libelle = str(row.iloc[0]).lower().strip() if pd.notna(row.iloc[0]) else ""
                for cle, synonymes in MAPPING.items():
                    if cle not in data:
                        for syn in synonymes:
                            if syn in libelle:
                                # Prend la première valeur numérique de la ligne
                                vals = [v for v in row.iloc[1:] if isinstance(v, (int, float)) and not np.isnan(v)]
                                if vals:
                                    data[cle] = float(vals[0])
                                break
        return data
    except Exception as e:
        return {"erreur": str(e)}


# =============================================================================
# M3 — MOTEUR DE DETTE AVANCÉ
# =============================================================================

def calculer_wacc(part_dette: float, cout_dette: float, taux_is: float,
                  part_fp: float, cout_fp: float) -> float:
    """WACC = Kd*(1-IS)*(D/V) + Ke*(E/V)."""
    return cout_dette * (1 - taux_is) * part_dette + cout_fp * part_fp


def calculer_service_dette_avance(
    fcff_liste       : list,
    capex_initial    : float,
    part_dette       : float,
    cout_dette       : float,
    duree_dette      : int,
    grace_period     : int,
    wacc             : float,
    commission_flat  : float,   # % du montant total, prélevé à t=0
    commission_engagt: float,   # % annuel sur part non décaissée
    profil_decaiss   : list,    # fraction du CAPEX décaissée chaque année [0..1]
) -> tuple:
    """
    Service de la dette avec commissions institutionnelles.

    commission_flat     : prélevée une seule fois au premier décaissement
    commission_engagement : sur la fraction non encore décaissée chaque année

    Retourne (DataFrame, dscr_min, dscr_moy, llcr_global, cout_total_dette)
    """
    montant_dette         = capex_initial * part_dette
    cout_flat             = montant_dette * commission_flat
    annees_remboursement  = max(duree_dette - grace_period, 1)
    amort_capital         = montant_dette / annees_remboursement

    rows           = []
    encours        = montant_dette
    cumul_decaisse = 0.0
    cout_total_int = cout_flat  # on commence avec la commission flat

    for i, r in enumerate(fcff_liste):
        annee = r["Annee"]
        fcff  = r["FCFF (M)"] * 1e6

        # Décaissement progressif du CAPEX
        frac_decaiss       = profil_decaiss[i] if i < len(profil_decaiss) else 0.0
        decaisse_annee     = montant_dette * frac_decaiss
        non_decaisse       = montant_dette - cumul_decaisse
        comm_engagt_annee  = non_decaisse * commission_engagt
        cumul_decaisse    += decaisse_annee

        if annee <= duree_dette:
            interets = encours * cout_dette
            remb_cap = 0.0 if annee <= grace_period else amort_capital

            # Commission d'engagement seulement pendant la phase de décaissement
            comm_annee   = comm_engagt_annee if cumul_decaisse < montant_dette else 0.0
            service_tot  = interets + remb_cap + comm_annee
            encours_fin  = encours - remb_cap
            dscr_an      = fcff / service_tot if service_tot > 0 else np.nan

            fcff_rest    = [fcff_liste[j]["FCFF (M)"] * 1e6
                            for j in range(i, min(duree_dette, len(fcff_liste)))]
            van_rest     = npv_calc(cout_dette, fcff_rest) if fcff_rest else 0.0
            llcr_an      = van_rest / encours if encours > 0 else np.nan

            cout_total_int += interets + comm_annee
            rows.append({
                "Annee"              : annee,
                "Encours (M)"        : round(encours / 1e6, 2),
                "Interets (M)"       : round(interets / 1e6, 2),
                "Remb. Capital (M)"  : round(remb_cap / 1e6, 2),
                "Comm. Engagt (M)"   : round(comm_annee / 1e6, 3),
                "Service Total (M)"  : round(service_tot / 1e6, 2),
                "FCFF (M)"           : round(fcff / 1e6, 2),
                "DSCR"               : round(dscr_an, 2) if not np.isnan(dscr_an) else np.nan,
                "LLCR"               : round(llcr_an, 2) if not np.isnan(llcr_an) else np.nan,
                "Grace Period"       : (annee <= grace_period),
            })
            encours = encours_fin
        else:
            rows.append({
                "Annee"              : annee,
                "Encours (M)"        : 0.0,
                "Interets (M)"       : 0.0,
                "Remb. Capital (M)"  : 0.0,
                "Comm. Engagt (M)"   : 0.0,
                "Service Total (M)"  : 0.0,
                "FCFF (M)"           : round(fcff / 1e6, 2),
                "DSCR"               : np.nan,
                "LLCR"               : np.nan,
                "Grace Period"       : False,
            })

    df       = pd.DataFrame(rows)
    dscrs    = df["DSCR"].dropna()
    dscr_min = round(dscrs.min(), 2)  if len(dscrs) else np.nan
    dscr_moy = round(dscrs.mean(), 2) if len(dscrs) else np.nan

    fcff_llcr   = [r["FCFF (M)"] * 1e6 for r in fcff_liste[:duree_dette]]
    llcr_global = (round(npv_calc(cout_dette, fcff_llcr) / montant_dette, 2)
                   if montant_dette > 0 else np.nan)

    return df, dscr_min, dscr_moy, llcr_global, cout_total_int


# =============================================================================
# M3 — ALERTE CHARGES / CA
# =============================================================================

def verifier_ratio_charges(fcff_liste: list, seuil: float = SEUIL_CHARGES_CA) -> list:
    """
    Retourne la liste des années où Coûts / Revenus > seuil.
    Indicateur institutionnel : coûts opérationnels > 55 % du CA = signal d'alerte.
    """
    alertes = []
    for r in fcff_liste:
        rev = r.get("Revenus (M)", 0)
        cout = r.get("Couts (M)", 0)
        if rev > 0 and (cout / rev) > seuil:
            alertes.append(r["Annee"])
    return alertes


# =============================================================================
# M1 — PROJECTION FCFF (moteur central, Business Model aware)
# =============================================================================

def projeter_fcff(
    quantite        : float,
    prix_vente      : float,
    cout_production : float,
    capex_initial   : float,
    taux_is         : float,
    taux_inflation  : float,
    duree           : int,
    coeffs_rampup   : list,
    business_model  : str,
    jours_stock     : int   = 60,
    jours_clients   : int   = 30,
    jours_fournisseurs: int = 45,
    duree_croissance: int   = 0,
    ratio_maintenance: float = 0.02,
) -> list:
    """
    Projection FCFF intégrant le Business Model, le BFR et les immobilisations biologiques.

    FCFF = NOPAT + Amortissement - CAPEX_maintenance - Delta_BFR

    Delta_BFR est calculé chaque année comme la variation du BFR par rapport à N-1.
    Pour le modèle Production, un ajustement biologique bloque les revenus
    pendant la phase de croissance des plants.
    """
    amort_annuel   = capex_initial / duree
    capex_maint    = capex_initial * ratio_maintenance
    resultats      = []
    bfr_precedent  = 0.0

    for annee in range(1, duree + 1):
        coeff       = coeffs_rampup[annee - 1]
        fact_inf    = (1 + taux_inflation) ** (annee - 1)

        # Modèle Production : blocage revenus phase biologique
        if business_model == "Production":
            coeff_bio = ajustement_biologique(annee, duree_croissance)
            coeff     = coeff * coeff_bio

        revenus = quantite * coeff * prix_vente      * fact_inf
        couts   = quantite * coeff * cout_production * fact_inf
        ebitda  = revenus - couts
        ebit    = ebitda - amort_annuel
        impot   = max(ebit * taux_is, 0.0)
        nopat   = ebit - impot

        # BFR
        bfr_courant = parametres_bfr(
            business_model, revenus, couts,
            jours_stock, jours_clients, jours_fournisseurs
        )
        delta_bfr  = bfr_courant - bfr_precedent
        bfr_precedent = bfr_courant

        fcff = nopat + amort_annuel - capex_maint - delta_bfr

        resultats.append({
            "Annee"            : annee,
            "Ramp-up (%)"      : round(coeff * 100, 1),
            "Revenus (M)"      : round(revenus / 1e6, 3),
            "Couts (M)"        : round(couts   / 1e6, 3),
            "Ratio Charges/CA" : round(couts / revenus, 3) if revenus > 0 else np.nan,
            "EBITDA (M)"       : round(ebitda / 1e6, 3),
            "Amort. (M)"       : round(amort_annuel / 1e6, 3),
            "EBIT (M)"         : round(ebit   / 1e6, 3),
            "IS (M)"           : round(impot  / 1e6, 3),
            "Delta BFR (M)"    : round(delta_bfr / 1e6, 3),
            "FCFF (M)"         : round(fcff   / 1e6, 3),
        })

    return resultats


def calculer_metriques(fcff_liste: list, capex_initial: float, wacc: float) -> dict:
    """VAN, TRI, Payback."""
    flux    = [-capex_initial] + [r["FCFF (M)"] * 1e6 for r in fcff_liste]
    van_val = npv_calc(wacc, flux)
    tri_val = irr_calc(flux)

    cumul, payback = 0.0, None
    for i, f in enumerate(flux[1:], 1):
        cumul += f
        if cumul >= capex_initial:
            payback = i
            break

    return {
        "VAN (M)"         : round(van_val / 1e6, 2),
        "TRI (%)"         : round(tri_val * 100, 2) if tri_val else None,
        "Payback (annees)": payback,
    }


# =============================================================================
# M4 — STRESS TEST DISCRET + COUVERTURE ASSURANCE
# =============================================================================

def definir_evenements_stress() -> dict:
    """
    Catalogue des 3 événements discrets de stress.
    Structure : {nom: {description, prob_defaut, gravite_defaut}}
    """
    return {
        "Choc Climatique"    : {
            "description" : "Sécheresse, inondation ou épidémie phytosanitaire",
            "prob_defaut"  : 0.08,
            "gravite_defaut": 0.40,
        },
        "Choc Opérationnel"  : {
            "description" : "Panne majeure d'équipement ou grève prolongée",
            "prob_defaut"  : 0.12,
            "gravite_defaut": 0.20,
        },
        "Choc Qualité"       : {
            "description" : "Lot non conforme, perte de certification, rappel produit",
            "prob_defaut"  : 0.06,
            "gravite_defaut": 0.25,
        },
    }


def simuler_stress_test(
    fcff_liste       : list,
    evenements       : dict,
    prime_assurance  : float,   # prime annuelle en FCFA (ou USD)
    indemnite_pct    : float,   # % de la perte couverte par l'assurance
    capex_initial    : float,
    wacc             : float,
    part_dette       : float,
    cout_dette       : float,
    duree_dette      : int,
    grace_period     : int,
    commission_flat  : float,
    commission_engagt: float,
    profil_decaiss   : list,
) -> dict:
    """
    Pour chaque événement discret, calcule :
      - L'impact sur les revenus (perte = gravité × CA)
      - Les flux ajustés Sans Assurance et Avec Assurance
      - Le DSCR annuel dans chaque scénario
      - La trésorerie cumulée dans chaque scénario

    Hypothèse de modélisation :
      L'événement survient à son année de probabilité maximale (milieu de projet).
      Son impact dure 1 an (événement unique concentré).
      L'assurance rembourse indemnite_pct × perte_revenus la même année,
      mais prélève la prime chaque année sur toute la durée du projet.

    Retourne un dict de résultats par événement.
    """
    duree          = len(fcff_liste)
    annee_choc     = max(1, duree // 2)   # choc au milieu du projet
    resultats      = {}

    for nom, params in evenements.items():
        prob     = params["prob"]
        gravite  = params["gravite"]
        perte_ca = None   # calculée dynamiquement

        rows_sans = []
        rows_avec = []
        tresorerie_sans = 0.0
        tresorerie_avec = 0.0

        # --- Recalcul FCFF ajusté pour chaque scénario ---
        for r in fcff_liste:
            annee  = r["Annee"]
            fcff   = r["FCFF (M)"] * 1e6
            rev    = r["Revenus (M)"] * 1e6

            # Impact du choc à l'année définie
            if annee == annee_choc:
                perte_ca     = rev * gravite          # perte brute
                impact_fcff  = -perte_ca              # réduit le FCFF
            else:
                perte_ca    = 0.0
                impact_fcff = 0.0

            # Sans assurance
            fcff_sans        = fcff + impact_fcff
            tresorerie_sans += fcff_sans
            rows_sans.append({
                "Annee"          : annee,
                "FCFF (M)"       : round(fcff_sans / 1e6, 3),
                "Tresorerie (M)" : round(tresorerie_sans / 1e6, 3),
            })

            # Avec assurance : prime annuelle en charge, indemnité si choc
            indemnite       = perte_ca * indemnite_pct if annee == annee_choc else 0.0
            fcff_avec       = fcff + impact_fcff - prime_assurance + indemnite
            tresorerie_avec += fcff_avec
            rows_avec.append({
                "Annee"          : annee,
                "FCFF (M)"       : round(fcff_avec / 1e6, 3),
                "Tresorerie (M)" : round(tresorerie_avec / 1e6, 3),
            })

        # DSCR pour chaque scénario
        def dscr_serie(rows_scenario):
            df_tmp, dmin, dmoy, _, _ = calculer_service_dette_avance(
                rows_scenario, capex_initial, part_dette, cout_dette,
                duree_dette, grace_period, wacc,
                commission_flat, commission_engagt, profil_decaiss
            )
            return df_tmp["DSCR"].tolist(), dmin, dmoy

        df_sans  = [{"Annee": r["Annee"], "FCFF (M)": r["FCFF (M)"]} for r in rows_sans]
        df_avec  = [{"Annee": r["Annee"], "FCFF (M)": r["FCFF (M)"]} for r in rows_avec]

        dscr_sans_liste, dscr_min_sans, _ = dscr_serie(df_sans)
        dscr_avec_liste, dscr_min_avec, _ = dscr_serie(df_avec)

        resultats[nom] = {
            "probabilite"     : prob,
            "gravite"         : gravite,
            "annee_choc"      : annee_choc,
            "sans_assurance"  : rows_sans,
            "avec_assurance"  : rows_avec,
            "dscr_sans"       : dscr_sans_liste,
            "dscr_avec"       : dscr_avec_liste,
            "dscr_min_sans"   : dscr_min_sans,
            "dscr_min_avec"   : dscr_min_avec,
        }

    return resultats


# =============================================================================
# M5 — VISUALISATIONS PLOTLY
# =============================================================================

def fig_fcff_waterfall(fcff_liste: list, devise: str) -> go.Figure:
    """Graphique en cascade (waterfall) des FCFF annuels."""
    annees = [str(r["Annee"]) for r in fcff_liste]
    vals   = [r["FCFF (M)"] for r in fcff_liste]
    colors = [CLR_VERT if v >= 0 else CLR_ROUGE for v in vals]

    fig = go.Figure(go.Bar(
        x=annees, y=vals,
        marker_color=colors,
        text=[f"{v:.1f}" for v in vals],
        textposition="outside",
    ))
    fig.update_layout(
        title=f"FCFF Projetés (M{devise})",
        xaxis_title="Année", yaxis_title=f"M{devise}",
        plot_bgcolor="white", paper_bgcolor="white",
        font=dict(size=11), height=380,
        showlegend=False,
    )
    fig.add_hline(y=0, line_width=1.2, line_color="#24292f")
    return fig


def fig_dscr_annuel(df_dette: pd.DataFrame) -> go.Figure:
    """Graphique DSCR annuel avec seuils 1.0x et 1.3x."""
    df_plot = df_dette[df_dette["DSCR"].notna()].copy()
    colors  = [
        CLR_VERT  if v >= 1.3 else (CLR_AMBER if v >= 1.0 else CLR_ROUGE)
        for v in df_plot["DSCR"]
    ]

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=df_plot["Annee"].astype(str),
        y=df_plot["DSCR"],
        marker_color=colors,
        name="DSCR annuel",
        text=[f"{v:.2f}x" for v in df_plot["DSCR"]],
        textposition="outside",
    ))
    fig.add_hline(y=1.3, line_dash="dash", line_color=CLR_VERT,
                  annotation_text="Seuil bancaire 1.3x", annotation_position="right")
    fig.add_hline(y=1.0, line_dash="dot",  line_color=CLR_ROUGE,
                  annotation_text="Seuil défaut 1.0x",   annotation_position="right")
    fig.update_layout(
        title="DSCR Annuel — Service de la Dette",
        xaxis_title="Année", yaxis_title="DSCR (x)",
        plot_bgcolor="white", paper_bgcolor="white",
        height=380, showlegend=False,
    )
    return fig


def fig_stress_dscr(stress_result: dict, nom_evt: str, duree: int) -> go.Figure:
    """Compare DSCR Sans vs Avec Assurance pour un événement."""
    annees = list(range(1, duree + 1))
    d_sans = stress_result[nom_evt]["dscr_sans"]
    d_avec = stress_result[nom_evt]["dscr_avec"]
    annee_choc = stress_result[nom_evt]["annee_choc"]

    # Remplace NaN par None pour Plotly
    def clean(lst):
        return [v if (isinstance(v, float) and not np.isnan(v)) else None for v in lst]

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=annees, y=clean(d_sans),
        mode="lines+markers", name="Sans Assurance",
        line=dict(color=CLR_ROUGE, width=2, dash="dot"),
        marker=dict(size=6),
    ))
    fig.add_trace(go.Scatter(
        x=annees, y=clean(d_avec),
        mode="lines+markers", name="Avec Assurance",
        line=dict(color=CLR_VERT, width=2),
        marker=dict(size=6),
    ))
    fig.add_hline(y=1.3, line_dash="dash", line_color=CLR_AMBER,
                  annotation_text="Seuil 1.3x")
    fig.add_hline(y=1.0, line_dash="dot",  line_color=CLR_ROUGE,
                  annotation_text="Seuil défaut")
    fig.add_vrect(
        x0=annee_choc - 0.4, x1=annee_choc + 0.4,
        fillcolor="rgba(220,53,69,0.10)", line_width=0,
        annotation_text="Choc", annotation_position="top left",
    )
    fig.update_layout(
        title=f"Impact sur le DSCR — {nom_evt}",
        xaxis_title="Année", yaxis_title="DSCR (x)",
        plot_bgcolor="white", paper_bgcolor="white",
        height=380, legend=dict(orientation="h", y=-0.2),
    )
    return fig


def fig_stress_tresorerie(stress_result: dict, nom_evt: str, devise: str) -> go.Figure:
    """Compare trésorerie cumulée Sans vs Avec Assurance."""
    sans = stress_result[nom_evt]["sans_assurance"]
    avec = stress_result[nom_evt]["avec_assurance"]
    annees = [r["Annee"] for r in sans]
    annee_choc = stress_result[nom_evt]["annee_choc"]

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=annees, y=[r["Tresorerie (M)"] for r in sans],
        mode="lines+markers", name="Sans Assurance",
        line=dict(color=CLR_ROUGE, width=2, dash="dot"),
        fill="tozeroy", fillcolor="rgba(220,53,69,0.07)",
    ))
    fig.add_trace(go.Scatter(
        x=annees, y=[r["Tresorerie (M)"] for r in avec],
        mode="lines+markers", name="Avec Assurance",
        line=dict(color=CLR_VERT, width=2),
        fill="tozeroy", fillcolor="rgba(25,135,84,0.07)",
    ))
    fig.add_hline(y=0, line_color="#24292f", line_width=1)
    fig.add_vrect(
        x0=annee_choc - 0.4, x1=annee_choc + 0.4,
        fillcolor="rgba(220,53,69,0.10)", line_width=0,
        annotation_text="Choc", annotation_position="top left",
    )
    fig.update_layout(
        title=f"Trésorerie Cumulée — {nom_evt} (M{devise})",
        xaxis_title="Année", yaxis_title=f"M{devise}",
        plot_bgcolor="white", paper_bgcolor="white",
        height=380, legend=dict(orientation="h", y=-0.2),
    )
    return fig


def fig_tornado(tornado_data: list, van_ref: float, devise: str) -> go.Figure:
    """Tornado Chart horizontal — impact sur la VAN."""
    labels   = [d["Variable"]      for d in tornado_data]
    imp_m    = [d["Impact_moins"]  for d in tornado_data]
    imp_p    = [d["Impact_plus"]   for d in tornado_data]

    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=labels, x=imp_m, orientation="h",
        name="-10%", marker_color=CLR_ROUGE,
        text=[f"{v:+,.0f}" for v in imp_m], textposition="outside",
    ))
    fig.add_trace(go.Bar(
        y=labels, x=imp_p, orientation="h",
        name="+10%", marker_color=CLR_VERT,
        text=[f"{v:+,.0f}" for v in imp_p], textposition="outside",
    ))
    fig.add_vline(x=0, line_color="#24292f", line_width=1.2)
    fig.update_layout(
        title=f"Tornado — Sensibilité de la VAN (base : {van_ref:,.0f} M{devise})",
        xaxis_title=f"Impact sur la VAN (M{devise})",
        barmode="overlay", plot_bgcolor="white", paper_bgcolor="white",
        height=360, legend=dict(orientation="h", y=-0.2),
    )
    return fig


def fig_sensibilite_heatmap(df_sens: pd.DataFrame, devise: str) -> go.Figure:
    """Heatmap interactive de sensibilité croisée Prix × Coût."""
    fig = go.Figure(go.Heatmap(
        z=df_sens.values.tolist(),
        x=df_sens.columns.tolist(),
        y=df_sens.index.tolist(),
        colorscale=[
            [0.0,  "#dc3545"],
            [0.5,  "#ffffff"],
            [1.0,  "#198754"],
        ],
        text=[[f"{v:.0f}" for v in row] for row in df_sens.values],
        texttemplate="%{text}",
        showscale=True,
        colorbar=dict(title=f"VAN (M{devise})"),
    ))
    fig.update_layout(
        title=f"Sensibilité Croisée VAN — Prix × Coût (M{devise})",
        xaxis_title="Variation Coût de Production",
        yaxis_title="Variation Prix de Vente",
        height=380,
    )
    return fig


# =============================================================================
# TORNADO — CALCUL
# =============================================================================

def calcul_tornado(
    quantite, prix_vente, cout_production, capex_initial,
    taux_is, taux_inflation, duree, wacc, coeffs_rampup,
    business_model, jours_stock, jours_clients, jours_fournisseurs,
    duree_croissance, delta=0.10,
):
    def van(**kw):
        fl = projeter_fcff(**kw)
        return calculer_metriques(fl, kw["capex_initial"], wacc)["VAN (M)"]

    base = dict(
        quantite=quantite, prix_vente=prix_vente,
        cout_production=cout_production, capex_initial=capex_initial,
        taux_is=taux_is, taux_inflation=taux_inflation, duree=duree,
        coeffs_rampup=coeffs_rampup, business_model=business_model,
        jours_stock=jours_stock, jours_clients=jours_clients,
        jours_fournisseurs=jours_fournisseurs, duree_croissance=duree_croissance,
    )
    van_ref = van(**base)

    sensibilites = [
        ("Prix de Vente",   "prix_vente",       prix_vente),
        ("Volume Produit",  "quantite",          quantite),
        ("CAPEX",           "capex_initial",     capex_initial),
        ("OPEX (Coût/T)",   "cout_production",   cout_production),
    ]

    resultats = []
    for label, param, val_base in sensibilites:
        van_m = van(**{**base, param: val_base * (1 - delta)})
        van_p = van(**{**base, param: val_base * (1 + delta)})
        resultats.append({
            "Variable"     : label,
            "Impact_moins" : van_m - van_ref,
            "Impact_plus"  : van_p - van_ref,
            "Amplitude"    : abs(van_p - van_m),
        })

    resultats.sort(key=lambda x: x["Amplitude"], reverse=True)
    return resultats, van_ref


# =============================================================================
# SENSIBILITÉ CROISÉE
# =============================================================================

def analyse_sensibilite(
    quantite, prix_vente, cout_production, capex_initial,
    taux_is, taux_inflation, duree, wacc, coeffs_rampup,
    business_model, jours_stock, jours_clients, jours_fournisseurs,
    duree_croissance, variations=(-0.10, -0.05, 0.0, 0.05, 0.10),
):
    labels = [f"{int(v*100):+d}%" for v in variations]
    data   = {}
    for v_c in variations:
        col = []
        for v_p in variations:
            fl = projeter_fcff(
                quantite, prix_vente * (1 + v_p), cout_production * (1 + v_c),
                capex_initial, taux_is, taux_inflation, duree, coeffs_rampup,
                business_model, jours_stock, jours_clients,
                jours_fournisseurs, duree_croissance,
            )
            col.append(calculer_metriques(fl, capex_initial, wacc)["VAN (M)"])
        data[f"Coût {int(v_c*100):+d}%"] = col
    return pd.DataFrame(data, index=[f"Prix {l}" for l in labels])


# =============================================================================
# PAGE CONFIG & CSS
# =============================================================================

st.set_page_config(
    page_title="CommodityWatch | Standard Institutionnel v3",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    h1  { font-size:1.55rem; font-weight:700; color:#0d1117; }
    h2  { font-size:1.05rem; font-weight:600; color:#24292f;
          border-bottom:1px solid #e0e0e0; padding-bottom:5px; margin-top:1.4rem; }
    h3  { font-size:0.9rem;  font-weight:600; color:#24292f; }
    .stMetric label { font-size:0.75rem; color:#57606a;
                      text-transform:uppercase; letter-spacing:.04em; }
    .block-container { padding-top:1.6rem; padding-bottom:2rem; }
    .stAlert { border-radius:6px; }
    hr { border:none; border-top:1px solid #e8e8e8; margin:18px 0; }
</style>
""", unsafe_allow_html=True)


# =============================================================================
# BARRE LATÉRALE
# =============================================================================

with st.sidebar:
    st.markdown("## CommodityWatch v3")
    st.markdown("**Standard Institutionnel International**")
    st.markdown("---")

    # ── M2 : Devise ──────────────────────────────────────────────────────────
    devise = st.radio("Devise de travail", ["FCFA", "USD"], horizontal=True)
    M = 1e6  # facteur millions

    st.markdown("---")

    # ── M1 : Business Model ──────────────────────────────────────────────────
    st.markdown("### Business Model")
    business_model = st.selectbox(
        "Type de projet",
        ["Campagne", "Production", "Infrastructure"],
        help=(
            "Campagne : achat/transformation/vente avec fort BFR.\n"
            "Production : plantation avec phase biologique de croissance.\n"
            "Infrastructure : construction/extension d'usine, CAPEX dominant."
        ),
    )

    # Paramètre spécifique Production
    duree_croissance = 0
    if business_model == "Production":
        duree_croissance = st.number_input(
            "Phase de croissance biologique (années sans revenu)",
            min_value=0, max_value=10, value=3, step=1,
        )

    st.markdown("---")

    # ── M2 : Matière première & prix marché ──────────────────────────────────
    st.markdown("### Matière Première & Prix")

    matieres = list(PRIX_MARCHE_USD.keys())
    matiere  = st.selectbox("Matière première", matieres)

    prix_marche_ref = prix_reference_marche(matiere, devise)
    if matiere != "Personnalisé":
        st.info(
            f"Référence marché ({matiere}) : "
            f"{prix_marche_ref:,.0f} {devise}/T "
            f"— Source ICE/Euronext/OTC"
        )
        utiliser_prix_marche = st.checkbox("Indexer sur ce prix de référence", value=False)
    else:
        utiliser_prix_marche = False

    prix_vente_input = st.number_input(
        f"Prix de Vente ({devise}/Tonne)",
        min_value=0.0,
        value=float(prix_marche_ref) if utiliser_prix_marche and prix_marche_ref > 0 else (
            1_500.0 if devise == "USD" else 900_000.0
        ),
        step=10.0 if devise == "USD" else 10_000.0,
        format="%.2f" if devise == "USD" else "%.0f",
    )
    # Normalisation en FCFA pour les calculs internes
    prix_vente_fcfa = prix_vente_input if devise == "FCFA" else prix_vente_input * TAUX_CHANGE_FCFA_USD

    quantite = st.number_input(
        "Quantité à pleine capacité (Tonnes/an)",
        min_value=0, value=5000, step=100,
    )

    cout_prod_input = st.number_input(
        f"Coût de Production ({devise}/Tonne)",
        min_value=0.0,
        value=1_000.0 if devise == "USD" else 600_000.0,
        step=10.0 if devise == "USD" else 10_000.0,
        format="%.2f" if devise == "USD" else "%.0f",
    )
    cout_production_fcfa = cout_prod_input if devise == "FCFA" else cout_prod_input * TAUX_CHANGE_FCFA_USD

    # ── BFR (Campagne uniquement) ─────────────────────────────────────────────
    jours_stock = jours_clients = jours_fournisseurs = 0
    if business_model == "Campagne":
        with st.expander("Paramètres BFR (Cycle d'Exploitation)"):
            jours_stock        = st.number_input("Stock matière (jours)", 0, 180, 60, 5)
            jours_clients      = st.number_input("Crédit clients (jours)", 0, 120, 30, 5)
            jours_fournisseurs = st.number_input("Crédit fournisseurs (jours)", 0, 120, 45, 5)

    st.markdown("---")

    # ── Ramp-up ───────────────────────────────────────────────────────────────
    st.markdown("### Montée en Puissance (Ramp-up)")
    n_paliers = int(st.number_input("Nombre de paliers", 1, 5, 3, 1))
    defaults_an = [2, 3, 4, 5, 6]
    defaults_tx = [0, 40, 80, 100, 100]
    rampup_config = []
    for k in range(n_paliers):
        c1, c2 = st.columns(2)
        with c1:
            an_fin = st.number_input(f"Palier {k+1} — Fin année", 1, 25,
                                     defaults_an[k], 1, key=f"ru_an_{k}")
        with c2:
            tx_ru = st.number_input("Taux (%)", 0, 100, defaults_tx[k], 5, key=f"ru_tx_{k}")
        rampup_config.append((int(an_fin), tx_ru / 100))

    st.markdown("---")

    # ── Capital & WACC ────────────────────────────────────────────────────────
    st.markdown("### Structure du Capital")
    capex_input = st.number_input(
        f"CAPEX Total ({devise})",
        min_value=0.0,
        value=3_000_000.0 if devise == "USD" else 2_000_000_000.0,
        step=100_000.0 if devise == "USD" else 100_000_000.0,
        format="%.0f",
    )
    capex_fcfa = capex_input if devise == "FCFA" else capex_input * TAUX_CHANGE_FCFA_USD

    part_dette_pct     = st.slider("Part dette (% CAPEX)", 0, 100, 60, 5)
    part_dette         = part_dette_pct / 100
    part_fp            = 1 - part_dette
    cout_dette_pct     = st.slider("Coût dette (% annuel)", 1, 25, 9, 1)
    cout_dette         = cout_dette_pct / 100
    cout_fp_pct        = st.slider("Coût fonds propres / CAPM (%)", 5, 30, 15, 1)
    cout_fp            = cout_fp_pct / 100
    taux_is_pct        = st.slider("IS (%)", 0, 40, 25, 1)
    taux_is            = taux_is_pct / 100
    duree              = st.slider("Durée projet (années)", 3, 25, 10, 1)
    taux_inflation_pct = st.slider("Inflation annuelle (%)", 0, 15, 3, 1)
    taux_inflation     = taux_inflation_pct / 100

    st.markdown("---")

    # ── M3 : Moteur de dette avancé ───────────────────────────────────────────
    st.markdown("### Paramètres de la Dette (Avancé)")
    duree_dette   = st.slider("Durée du prêt (années)", 1, duree, min(8, duree), 1)
    grace_period  = st.slider("Différé de remboursement — Grace Period (années)",
                               0, max(0, duree_dette - 1), min(2, duree_dette - 1), 1)

    with st.expander("Commissions Bancaires"):
        comm_flat_pct   = st.slider("Commission Flat (% du montant, prélevée une fois)", 0.0, 3.0, 0.5, 0.1)
        comm_engagt_pct = st.slider("Commission d'Engagement (% annuel, part non décaissée)", 0.0, 2.0, 0.25, 0.05)

    commission_flat   = comm_flat_pct   / 100
    commission_engagt = comm_engagt_pct / 100

    # Profil de décaissement CAPEX (simplifié : 50% an1, 50% an2)
    profil_decaiss = [0.5, 0.5] + [0.0] * (duree - 2)

    st.markdown("---")

    # ── M2 : Import Excel ────────────────────────────────────────────────────
    st.markdown("### Import Financier (Excel)")
    uploaded = st.file_uploader(
        "Importer Bilan / Compte de Résultat (.xlsx)",
        type=["xlsx"],
        help="Le fichier doit contenir les libellés standards en colonne A.",
    )
    donnees_excel = {}
    if uploaded:
        if EXCEL_OK:
            donnees_excel = parser_excel_financier(uploaded)
            if "erreur" in donnees_excel:
                st.error(f"Erreur lecture Excel : {donnees_excel['erreur']}")
            else:
                st.success(f"{len(donnees_excel)} postes importés depuis Excel.")
                with st.expander("Données importées"):
                    st.json(donnees_excel)
        else:
            st.warning("Module `openpyxl` non installé. Ajoutez-le à requirements.txt.")

    st.markdown("---")

    # ── M4 : Stress Test ─────────────────────────────────────────────────────
    st.markdown("### Stress Test & Assurance")

    catalogue_evt   = definir_evenements_stress()
    evenements_conf = {}

    for nom, defauts in catalogue_evt.items():
        with st.expander(nom):
            st.caption(defauts["description"])
            prob_evt = st.slider(
                "Probabilité d'occurrence (%)", 1, 50,
                int(defauts["prob_defaut"] * 100), 1,
                key=f"prob_{nom}"
            )
            grav_evt = st.slider(
                "Gravité — perte de CA (%)", 5, 80,
                int(defauts["gravite_defaut"] * 100), 5,
                key=f"grav_{nom}"
            )
            evenements_conf[nom] = {
                "prob"    : prob_evt / 100,
                "gravite" : grav_evt / 100,
                "description": defauts["description"],
            }

    st.markdown("**Couverture Assurance**")
    prime_assurance_input = st.number_input(
        f"Prime annuelle ({devise})",
        min_value=0.0,
        value=50_000.0 if devise == "USD" else 30_000_000.0,
        step=5_000.0 if devise == "USD" else 5_000_000.0,
        format="%.0f",
    )
    prime_assurance_fcfa = (prime_assurance_input if devise == "FCFA"
                            else prime_assurance_input * TAUX_CHANGE_FCFA_USD)
    indemnite_pct = st.slider(
        "Taux de couverture assurance (% de la perte indemnisé)", 0, 100, 70, 5
    ) / 100

    st.markdown("---")
    lancer = st.button("Lancer l'Analyse", type="primary", use_container_width=True)


# =============================================================================
# ZONE PRINCIPALE — HEADER
# =============================================================================

st.markdown(f"## CommodityWatch — {matiere} | {business_model} | {devise}")
st.caption(
    "Outil de pre-scoring institutionnel pour projets agro-industriels. "
    "Résultats indicatifs — ne se substituent pas à une due diligence complète."
)
st.markdown("---")

if not lancer:
    col_i1, col_i2, col_i3 = st.columns(3)
    with col_i1:
        st.info("**M1 — Business Model**\nSélectionnez le type de projet dans la sidebar.", icon="🏭")
    with col_i2:
        st.info("**M4 — Stress Test**\nConfigurez les 3 événements de risque discrets.", icon="⚡")
    with col_i3:
        st.info("**M5 — Visualisation**\nComparaison Sans/Avec Assurance sur DSCR & Trésorerie.", icon="📊")
    st.stop()


# =============================================================================
# CALCULS CENTRAUX
# =============================================================================

with st.spinner("Moteur financier en cours..."):

    wacc          = calculer_wacc(part_dette, cout_dette, taux_is, part_fp, cout_fp)
    coeffs_rampup = construire_rampup(duree, rampup_config)

    fcff_liste = projeter_fcff(
        quantite, prix_vente_fcfa, cout_production_fcfa,
        capex_fcfa, taux_is, taux_inflation, duree, coeffs_rampup,
        business_model, jours_stock, jours_clients, jours_fournisseurs,
        duree_croissance,
    )

    # Conversion pour affichage
    facteur_affichage = 1.0 if devise == "FCFA" else (1 / TAUX_CHANGE_FCFA_USD)
    fcff_affichage = []
    for r in fcff_liste:
        ra = {k: (round(v * facteur_affichage, 3) if isinstance(v, float) and k not in ("Annee", "Ramp-up (%)") else v)
              for k, v in r.items()}
        fcff_affichage.append(ra)

    metriques = calculer_metriques(fcff_liste, capex_fcfa, wacc)
    van_affich = round(metriques["VAN (M)"] * facteur_affichage, 2)
    capex_affich = round(capex_fcfa / 1e6 * facteur_affichage, 2)

    df_dette, dscr_min, dscr_moy, llcr_global, cout_total_dette = calculer_service_dette_avance(
        fcff_liste, capex_fcfa, part_dette, cout_dette,
        duree_dette, grace_period, wacc,
        commission_flat, commission_engagt, profil_decaiss,
    )
    # Conversion tableau dette
    for col_m in ["Encours (M)", "Interets (M)", "Remb. Capital (M)",
                  "Comm. Engagt (M)", "Service Total (M)", "FCFF (M)"]:
        df_dette[col_m] = df_dette[col_m] * facteur_affichage

    alertes_charges = verifier_ratio_charges(fcff_liste, SEUIL_CHARGES_CA)

    tornado_data, van_ref_tornado = calcul_tornado(
        quantite, prix_vente_fcfa, cout_production_fcfa, capex_fcfa,
        taux_is, taux_inflation, duree, wacc, coeffs_rampup,
        business_model, jours_stock, jours_clients, jours_fournisseurs, duree_croissance,
    )
    van_ref_affich = round(van_ref_tornado * facteur_affichage, 2)
    for d in tornado_data:
        d["Impact_moins"] *= facteur_affichage
        d["Impact_plus"]  *= facteur_affichage
        d["Amplitude"]    *= facteur_affichage

    df_sens = analyse_sensibilite(
        quantite, prix_vente_fcfa, cout_production_fcfa, capex_fcfa,
        taux_is, taux_inflation, duree, wacc, coeffs_rampup,
        business_model, jours_stock, jours_clients, jours_fournisseurs, duree_croissance,
    )
    df_sens = df_sens * facteur_affichage

    # Stress Test
    stress_results = simuler_stress_test(
        fcff_liste, evenements_conf,
        prime_assurance_fcfa, indemnite_pct,
        capex_fcfa, wacc,
        part_dette, cout_dette, duree_dette, grace_period,
        commission_flat, commission_engagt, profil_decaiss,
    )


# =============================================================================
# SECTION 1 — KPIs
# =============================================================================

st.markdown("## Indicateurs Clés de Performance")

k1, k2, k3, k4, k5, k6, k7 = st.columns(7)

with k1:
    st.metric("VAN", f"{van_affich:,.0f} M{devise}",
              delta="Positive" if van_affich > 0 else "Négative",
              delta_color="normal" if van_affich > 0 else "inverse")
with k2:
    tri_val = metriques["TRI (%)"]
    st.metric("TRI", f"{tri_val:.2f} %" if tri_val else "N/D",
              delta=f"WACC {wacc*100:.2f} %",
              delta_color="normal" if (tri_val or 0) > wacc * 100 else "inverse")
with k3:
    st.metric("WACC", f"{wacc*100:.2f} %")
with k4:
    st.metric("DSCR Min",
              f"{dscr_min:.2f}x" if not np.isnan(dscr_min) else "N/D",
              delta="Bancable" if dscr_min >= 1.3 else "< Seuil 1.3x",
              delta_color="normal" if dscr_min >= 1.3 else "inverse")
with k5:
    st.metric("LLCR",
              f"{llcr_global:.2f}x" if not np.isnan(llcr_global) else "N/D",
              delta_color="off")
with k6:
    pb = metriques["Payback (annees)"]
    st.metric("Payback", f"{pb} ans" if pb else "> Durée")
with k7:
    cout_total_affich = round(cout_total_dette / 1e6 * facteur_affichage, 2)
    st.metric("Coût Total Dette",
              f"{cout_total_affich:,.1f} M{devise}",
              delta="Inclut commissions", delta_color="off")

# Alertes
if alertes_charges:
    st.warning(
        f"Alerte Charges / CA > {SEUIL_CHARGES_CA*100:.0f}% "
        f"aux années : **{alertes_charges}**. "
        "Revoir la structure de coûts opérationnels.",
        icon="⚠️",
    )
if not np.isnan(dscr_min) and dscr_min < 1.3:
    st.error(
        f"DSCR Min = {dscr_min:.2f}x — En dessous du seuil bancaire standard (1.3x). "
        "Le projet présente un risque de refinancement.",
        icon="🚨",
    )

st.markdown("---")


# =============================================================================
# SECTION 2 — PROJECTION FCFF
# =============================================================================

st.markdown(f"## Projection FCFF — Modèle {business_model}")

col_table, col_graph = st.columns([1, 1])

with col_table:
    st.caption(f"Montants en M{devise}. BFR intégré. Couleurs : vert = FCFF positif.")

    def style_fcff(row):
        if row.get("FCFF (M)", 0) < 0:
            return ["background-color:#f8d7da"] * len(row)
        if row.get("Ratio Charges/CA", 0) > SEUIL_CHARGES_CA:
            return ["background-color:#fff3cd"] * len(row)
        return [""] * len(row)

    df_display = pd.DataFrame(fcff_affichage)
    styled = (
        df_display.style
        .apply(style_fcff, axis=1)
        .format({c: "{:.2f}" for c in df_display.columns
                 if c not in ("Annee",) and df_display[c].dtype == float})
    )
    st.dataframe(styled, use_container_width=True, hide_index=True)

with col_graph:
    st.plotly_chart(fig_fcff_waterfall(fcff_affichage, devise),
                    use_container_width=True)

st.markdown("---")


# =============================================================================
# SECTION 3 — SERVICE DE LA DETTE AVANCÉ
# =============================================================================

st.markdown("## Bancabilité — Service de la Dette, DSCR & LLCR")
st.caption(
    f"Comm. Flat : {comm_flat_pct:.1f}% prélevée au décaissement | "
    f"Comm. Engagement : {comm_engagt_pct:.2f}%/an sur part non décaissée | "
    f"Grace Period : {grace_period} an(s) | Durée : {duree_dette} ans"
)

col_dt, col_dg = st.columns([1, 1])

with col_dt:
    cols_show = [c for c in df_dette.columns if c != "Grace Period"]
    df_aff    = df_dette[cols_show].copy()

    def style_dette(row):
        annee    = row["Annee"]
        is_grace = df_dette.loc[df_dette["Annee"] == annee, "Grace Period"].values
        is_grace = bool(is_grace[0]) if len(is_grace) else False
        if is_grace:
            return ["background-color:#fff3cd;color:#856404"] * len(row)
        dscr_v = row.get("DSCR", np.nan)
        if pd.notna(dscr_v):
            if dscr_v < 1.0:
                return ["background-color:#f8d7da;color:#721c24"] * len(row)
            if dscr_v < 1.3:
                return ["background-color:#fde8c8;color:#7d4e0f"] * len(row)
            return ["background-color:#d4edda;color:#155724"] * len(row)
        return [""] * len(row)

    styled_d = (
        df_aff.style
        .apply(style_dette, axis=1)
        .format({
            **{c: "{:.2f}" for c in cols_show if c not in ("Annee",)
               and df_aff[c].dtype == float},
            "DSCR": lambda x: f"{x:.2f}x" if pd.notna(x) else "—",
            "LLCR": lambda x: f"{x:.2f}x" if pd.notna(x) else "—",
        })
    )
    st.dataframe(styled_d, use_container_width=True, hide_index=True)

    # Légende
    l1, l2, l3, l4 = st.columns(4)
    l1.markdown("<span style='background:#d4edda;padding:1px 6px;border-radius:3px;font-size:.78rem'>DSCR ≥ 1.3x</span>", unsafe_allow_html=True)
    l2.markdown("<span style='background:#fde8c8;padding:1px 6px;border-radius:3px;font-size:.78rem'>1.0–1.3x</span>", unsafe_allow_html=True)
    l3.markdown("<span style='background:#f8d7da;padding:1px 6px;border-radius:3px;font-size:.78rem'>< 1.0x Défaut</span>", unsafe_allow_html=True)
    l4.markdown("<span style='background:#fff3cd;padding:1px 6px;border-radius:3px;font-size:.78rem'>Grace Period</span>", unsafe_allow_html=True)

with col_dg:
    st.plotly_chart(fig_dscr_annuel(df_dette), use_container_width=True)

st.markdown("---")


# =============================================================================
# SECTION 4 — ANALYSE DE SENSIBILITÉ
# =============================================================================

st.markdown("## Analyse de Sensibilité")

tab_tornado, tab_heatmap = st.tabs(["Tornado Chart", "Tableau Croisé"])

with tab_tornado:
    col_tf, col_tt = st.columns([2, 1])
    with col_tf:
        st.plotly_chart(fig_tornado(tornado_data, van_ref_affich, devise),
                        use_container_width=True)
    with col_tt:
        st.markdown(f"#### Impacts sur la VAN (M{devise})")
        df_t = pd.DataFrame([{
            "Variable"    : d["Variable"],
            "Impact -10%" : f"{d['Impact_moins']:+,.0f}",
            "Impact +10%" : f"{d['Impact_plus']:+,.0f}",
            "Amplitude"   : f"{d['Amplitude']:,.0f}",
        } for d in tornado_data])
        st.dataframe(df_t, use_container_width=True, hide_index=True)

with tab_heatmap:
    st.plotly_chart(fig_sensibilite_heatmap(df_sens, devise), use_container_width=True)

st.markdown("---")


# =============================================================================
# SECTION 5 — STRESS TEST & ASSURANCE
# =============================================================================

st.markdown("## Stress Test — Événements Discrets & Couverture Assurance")
st.caption(
    f"Prime assurance : {prime_assurance_input:,.0f} {devise}/an | "
    f"Taux de couverture : {indemnite_pct*100:.0f}% de la perte indemnisée"
)

# Résumé des événements
df_evt_resume = pd.DataFrame([
    {
        "Événement"    : nom,
        "Description"  : evenements_conf[nom]["description"],
        "Probabilité"  : f"{evenements_conf[nom]['prob']*100:.0f}%",
        "Gravité (CA)" : f"{evenements_conf[nom]['gravite']*100:.0f}%",
        "Année du choc": stress_results[nom]["annee_choc"],
        "DSCR Min Sans": f"{stress_results[nom]['dscr_min_sans']:.2f}x" if not np.isnan(stress_results[nom]['dscr_min_sans']) else "—",
        "DSCR Min Avec": f"{stress_results[nom]['dscr_min_avec']:.2f}x" if not np.isnan(stress_results[nom]['dscr_min_avec']) else "—",
    }
    for nom in stress_results
])

def style_resume_stress(row):
    dscr_s = row.get("DSCR Min Sans", "—")
    dscr_a = row.get("DSCR Min Avec", "—")
    try:
        val_s = float(dscr_s.replace("x",""))
        base = ["background-color:#f8d7da"] if val_s < 1.3 else ["background-color:#fff3cd"] if val_s < 1.5 else [""]
    except:
        base = [""]
    return base * len(row)

st.dataframe(
    df_evt_resume.style.apply(style_resume_stress, axis=1),
    use_container_width=True, hide_index=True
)

# Détail par événement
for nom in stress_results:
    with st.expander(f"Détail — {nom}", expanded=False):
        g1, g2 = st.columns(2)
        with g1:
            st.plotly_chart(
                fig_stress_dscr(stress_results, nom, duree),
                use_container_width=True
            )
        with g2:
            st.plotly_chart(
                fig_stress_tresorerie(stress_results, nom, devise),
                use_container_width=True
            )

        # Tableau comparatif
        df_comp = pd.DataFrame({
            "Année"             : [r["Annee"] for r in stress_results[nom]["sans_assurance"]],
            f"Tréso. Sans (M{devise})": [round(r["Tresorerie (M)"] * facteur_affichage, 2)
                                          for r in stress_results[nom]["sans_assurance"]],
            f"Tréso. Avec (M{devise})": [round(r["Tresorerie (M)"] * facteur_affichage, 2)
                                          for r in stress_results[nom]["avec_assurance"]],
            "DSCR Sans"         : [f"{v:.2f}x" if isinstance(v, float) and not np.isnan(v) else "—"
                                   for v in stress_results[nom]["dscr_sans"]],
            "DSCR Avec"         : [f"{v:.2f}x" if isinstance(v, float) and not np.isnan(v) else "—"
                                   for v in stress_results[nom]["dscr_avec"]],
        })

        def style_comp(row):
            an = row["Année"]
            if an == stress_results[nom]["annee_choc"]:
                return ["background-color:#fff3cd"] * len(row)
            return [""] * len(row)

        st.dataframe(
            df_comp.style.apply(style_comp, axis=1),
            use_container_width=True, hide_index=True
        )

        # Valeur ajoutée de l'assurance
        gain_dscr = (stress_results[nom]["dscr_min_avec"] - stress_results[nom]["dscr_min_sans"])
        if not np.isnan(gain_dscr):
            couleur = "success" if gain_dscr > 0 else "error"
            signe   = "+" if gain_dscr > 0 else ""
            st.info(
                f"L'assurance améliore le DSCR minimum de **{signe}{gain_dscr:.2f}x** "
                f"({stress_results[nom]['dscr_min_sans']:.2f}x → "
                f"{stress_results[nom]['dscr_min_avec']:.2f}x) "
                f"pour un coût annuel de {prime_assurance_input:,.0f} {devise}.",
                icon="ℹ️",
            )

st.markdown("---")
st.markdown(
    f"<div style='font-size:.72rem;color:#6c757d'>"
    f"CommodityWatch v3.0 — Taux de change : 1 USD = {TAUX_CHANGE_FCFA_USD:.0f} FCFA — "
    "Résultats indicatifs, ne se substituent pas à une due diligence financière complète."
    "</div>",
    unsafe_allow_html=True,
)
