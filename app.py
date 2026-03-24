# =============================================================================
# CommodityWatch v5.0 — Overview Pré-Investissement
# Framework : Streamlit + Plotly | Python 3.9+
# Usage     : streamlit run commodity_watch_v5.py
#
# Public cible : Équipes DFI (IFC, Proparco, BAD, AFD, BOAD),
#                Analystes crédit banque commerciale,
#                Head of Agro-Industry
#
# Architecture :
#   1 Business Model = 1 logique financière = 1 risque = 1 assurance
#   Structure dette : Senior + Mezzanine + Quasi-Equity (3 tranches)
#   Sorties : P&L pro forma · Bilan simplifié · FCFF · FCFE
#             Waterfall · DSCR/LLCR/ICR/Gearing · Stress test multi-années
#             Export PDF mémo teaser (reportlab)
# =============================================================================

import io
import math
import streamlit as st
import numpy as np
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import date

try:
    import openpyxl        # noqa: F401
    EXCEL_OK = True
except ImportError:
    EXCEL_OK = False

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                    Table, TableStyle, HRFlowable)
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
    PDF_OK = True
except ImportError:
    PDF_OK = False


# =============================================================================
# CONSTANTES
# =============================================================================

TAUX_CHANGE = 600.0   # 1 USD = 600 FCFA

BM_CONFIG = {
    "Production": {
        "icon"            : "🌱",
        "description"     : "Plantation / Élevage — revenu lié à la récolte",
        "risque_titre"    : "Risque Récolte",
        "risque_desc"     : "Sécheresse, inondation, épidémie phytosanitaire ou ravageurs réduisant le rendement.",
        "assurance_titre" : "Assurance Agricole / Récolte",
        "assurance_desc"  : "Indemnise une fraction de la perte de production constatée.",
        "gravite_defaut"  : 0.40,
        "bfr_actif"       : False,
        "bio_actif"       : True,
    },
    "Infrastructure": {
        "icon"            : "🏭",
        "description"     : "Construction / Extension d'usine — CAPEX dominant",
        "risque_titre"    : "Risque Opérationnel (Panne / Chantier)",
        "risque_desc"     : "Panne majeure d'équipement, accident de chantier, incendie.",
        "assurance_titre" : "Assurance TRC & Bris de Machine",
        "assurance_desc"  : "Couvre dommages matériels et perte d'exploitation consécutive.",
        "gravite_defaut"  : 0.25,
        "bfr_actif"       : False,
        "bio_actif"       : False,
    },
    "Campagne": {
        "icon"            : "🔄",
        "description"     : "Achat matière — Transformation — Vente",
        "risque_titre"    : "Risque Qualité / Non-conformité",
        "risque_desc"     : "Lot non conforme, perte de certification, rappel produit ou embargo.",
        "assurance_titre" : "Assurance Qualité & Responsabilité Produit",
        "assurance_desc"  : "Couvre pertes CA liées au retrait de lot et pénalités contractuelles.",
        "gravite_defaut"  : 0.20,
        "bfr_actif"       : True,
        "bio_actif"       : False,
    },
}

# Fiscalité OHADA par pays (IS nominal)
PAYS_FISCAL = {
    "Générique OHADA" : {"is": 0.25, "tva": 0.18},
    "Côte d'Ivoire"   : {"is": 0.25, "tva": 0.18},
    "Togo"            : {"is": 0.27, "tva": 0.18},
    "Sénégal"         : {"is": 0.30, "tva": 0.18},
    "Cameroun"        : {"is": 0.30, "tva": 0.1925},
    "Ghana"           : {"is": 0.25, "tva": 0.15},
    "Mali"            : {"is": 0.30, "tva": 0.18},
    "Burkina Faso"    : {"is": 0.27, "tva": 0.18},
}

SEUIL_CHARGES  = 0.55
SEUIL_BANCAIRE = 1.30
SEUIL_DEFAUT   = 1.00
DSRA_MOIS      = 6      # mois de service senior en réserve

CLR = {
    "vert"  : "#198754",
    "rouge" : "#dc3545",
    "amber" : "#fd7e14",
    "bleu"  : "#0d6efd",
    "gris"  : "#6c757d",
    "navy"  : "#0a3d62",
}


# =============================================================================
# UTILITAIRES
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
    return val if devise == "FCFA" else val * TAUX_CHANGE


def affiche(val_fcfa, devise):
    return val_fcfa if devise == "FCFA" else val_fcfa / TAUX_CHANGE


def m(val_fcfa, devise, dec=2):
    """Valeur en millions de la devise choisie."""
    return round(affiche(val_fcfa, devise) / 1e6, dec)


def fmt_m(val_fcfa, devise, dec=1):
    return f"{m(val_fcfa, devise, dec):,.{dec}f} M{devise}"


# =============================================================================
# RAMP-UP
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
# BFR
# =============================================================================

def calc_bfr(bm, revenus, couts, j_stock, j_clients, j_fourn):
    if bm == "Campagne":
        return couts * j_stock / 365 + revenus * j_clients / 365 - couts * j_fourn / 365
    if bm == "Production":
        return revenus * 0.05
    return 0.0


# =============================================================================
# PROJECTION FCFF + P&L + BILAN
# =============================================================================

def projeter(quantite, prix, cout, capex, taux_is,
             inflation, duree, rampup, bm,
             duree_bio=0, j_stock=60, j_clients=30, j_fourn=45,
             ratio_maint=0.02,
             # 3 tranches de dette
             dette_senior=0, r_senior=0.09, dur_senior=8, grace_senior=2,
             dette_mezz=0,   r_mezz=0.12,   dur_mezz=6,   grace_mezz=2,
             dette_qe=0,     r_qe=0.15,     dur_qe=5,     grace_qe=1,
             zone_franche=False, duree_exo=5):
    """
    Retourne rows : liste de dicts par année.
    Inclut P&L complet, bilan simplifié, FCFF, FCFE, ratios bancaires.
    """
    amort_capex = capex / duree
    c_maint     = capex * ratio_maint
    rows        = []
    bfr_prec    = 0.0
    immo_net    = capex

    # Calcul DSRA = 6 mois de service senior annuel moyen (réserve de trésorerie)
    svc_senior_moy = (dette_senior * r_senior +
                      (dette_senior / max(dur_senior - grace_senior, 1))) if dette_senior > 0 else 0
    dsra = svc_senior_moy * DSRA_MOIS / 12

    # Amortissements dette par tranche
    ann_senior = max(dur_senior - grace_senior, 1)
    ann_mezz   = max(dur_mezz   - grace_mezz,   1)
    ann_qe     = max(dur_qe     - grace_qe,      1)

    enc_senior = dette_senior
    enc_mezz   = dette_mezz
    enc_qe     = dette_qe

    for an in range(1, duree + 1):
        coeff = rampup[an - 1]
        if bm == "Production" and an <= duree_bio:
            coeff = 0.0
        inf   = (1 + inflation) ** (an - 1)
        rev   = quantite * coeff * prix  * inf
        cout_ = quantite * coeff * cout  * inf

        ebitda  = rev - cout_
        immo_net = max(immo_net - amort_capex, 0)
        ebit    = ebitda - amort_capex

        # Intérêts par tranche
        int_senior = enc_senior * r_senior if an <= dur_senior else 0.0
        int_mezz   = enc_mezz   * r_mezz   if an <= dur_mezz   else 0.0
        int_qe     = enc_qe     * r_qe     if an <= dur_qe     else 0.0
        int_total  = int_senior + int_mezz + int_qe

        # ICR (Interest Coverage Ratio)
        icr = ebitda / int_total if int_total > 0 else np.nan

        # EBT & IS
        ebt = ebit - int_total
        if zone_franche and an <= duree_exo:
            is_eff = 0.0
        else:
            is_eff = taux_is
        impot   = max(ebt * is_eff, 0.0)
        rn      = ebt - impot  # Résultat net

        # NOPAT (pour FCFF)
        nopat = ebit * (1 - is_eff) if ebit >= 0 else ebit

        # BFR
        bfr   = calc_bfr(bm, rev, cout_, j_stock, j_clients, j_fourn)
        delta = bfr - bfr_prec
        bfr_prec = bfr

        # FCFF
        fcff = nopat + amort_capex - c_maint - delta

        # Remboursements capital
        remb_s = (dette_senior / ann_senior) if (an > grace_senior and an <= dur_senior) else 0.0
        remb_m = (dette_mezz   / ann_mezz)   if (an > grace_mezz   and an <= dur_mezz)   else 0.0
        remb_q = (dette_qe     / ann_qe)     if (an > grace_qe     and an <= dur_qe)     else 0.0

        svc_senior = int_senior + remb_s
        svc_mezz   = int_mezz   + remb_m
        svc_qe     = int_qe     + remb_q
        svc_total  = svc_senior + svc_mezz + svc_qe

        # FCFE = FCFF - service total de la dette
        fcfe = fcff - svc_total

        # DSCR (sur service total)
        dscr = fcff / svc_total if svc_total > 0 else np.nan

        # DSCR Senior uniquement (critère bancaire strict)
        dscr_s = fcff / svc_senior if svc_senior > 0 else np.nan

        # Mise à jour encours
        enc_senior -= remb_s
        enc_mezz   -= remb_m
        enc_qe     -= remb_q

        dette_totale = max(enc_senior, 0) + max(enc_mezz, 0) + max(enc_qe, 0)

        # Gearing
        fp_comptable = capex - dette_senior - dette_mezz - dette_qe  # FP initiaux
        gearing = dette_totale / fp_comptable if fp_comptable > 0 else np.nan

        # Dette nette / EBITDA
        dn_ebitda = dette_totale / ebitda if ebitda > 0 else np.nan

        rows.append({
            # Identification
            "Annee"          : an,
            "Ramp_up"        : coeff * 100,
            # P&L
            "Revenus"        : rev,
            "Couts"          : cout_,
            "EBITDA"         : ebitda,
            "Amort"          : amort_capex,
            "EBIT"           : ebit,
            "Int_Senior"     : int_senior,
            "Int_Mezz"       : int_mezz,
            "Int_QE"         : int_qe,
            "Int_Total"      : int_total,
            "EBT"            : ebt,
            "IS"             : impot,
            "RN"             : rn,
            # Cash flows
            "FCFF"           : fcff,
            "Svc_Senior"     : svc_senior,
            "Svc_Mezz"       : svc_mezz,
            "Svc_QE"         : svc_qe,
            "Svc_Total"      : svc_total,
            "FCFE"           : fcfe,
            # Bilan simplifié
            "Immo_Net"       : immo_net,
            "BFR"            : bfr,
            "DSRA"           : dsra,
            "Enc_Senior"     : max(enc_senior, 0),
            "Enc_Mezz"       : max(enc_mezz, 0),
            "Enc_QE"         : max(enc_qe, 0),
            "Dette_Totale"   : dette_totale,
            # Ratios
            "Ratio_Charges"  : (cout_ / rev) if rev > 0 else np.nan,
            "DSCR"           : round(dscr,   3) if not np.isnan(dscr)   else np.nan,
            "DSCR_Senior"    : round(dscr_s, 3) if not np.isnan(dscr_s) else np.nan,
            "ICR"            : round(icr,    2) if not np.isnan(icr)    else np.nan,
            "Gearing"        : round(gearing, 2) if not np.isnan(gearing) else np.nan,
            "DN_EBITDA"      : round(dn_ebitda, 2) if not np.isnan(dn_ebitda) else np.nan,
            "Marge_EBITDA"   : (ebitda / rev) if rev > 0 else np.nan,
            "Marge_RN"       : (rn / rev)     if rev > 0 else np.nan,
        })
    return rows


def metriques(rows, capex, wacc):
    flux_fcff = [-capex] + [r["FCFF"] for r in rows]
    flux_fcfe = [-(capex * (1 - sum([0]) / capex))] + [r["FCFE"] for r in rows]
    van  = npv_calc(wacc, flux_fcff)
    tri  = irr_calc(flux_fcff)
    # TRI actionnaire (sur FCFE)
    fp_init = rows[0]["Immo_Net"] - rows[0]["Dette_Totale"] if rows else capex
    flux_eq = [-max(fp_init, capex * 0.1)] + [r["FCFE"] for r in rows]
    tri_eq  = irr_calc(flux_eq)

    cumul, pb = 0.0, None
    for i, f in enumerate(flux_fcff[1:], 1):
        cumul += f
        if cumul >= capex:
            pb = i
            break

    dscrs = [r["DSCR"] for r in rows if not np.isnan(r["DSCR"])]
    dscr_min = min(dscrs) if dscrs else np.nan
    dscr_moy = sum(dscrs) / len(dscrs) if dscrs else np.nan

    icrs = [r["ICR"] for r in rows if not np.isnan(r.get("ICR", np.nan))]
    icr_min = min(icrs) if icrs else np.nan

    return {
        "van": van, "tri": tri, "tri_eq": tri_eq,
        "payback": pb, "dscr_min": dscr_min,
        "dscr_moy": dscr_moy, "icr_min": icr_min,
    }


# =============================================================================
# LLCR
# =============================================================================

def calc_llcr(rows, dette_senior, r_senior):
    if dette_senior <= 0:
        return np.nan
    ff = [r["FCFF"] for r in rows if r["Enc_Senior"] > 0 or r["Annee"] == 1]
    if not ff:
        return np.nan
    return round(npv_calc(r_senior, ff) / dette_senior, 3)


# =============================================================================
# WATERFALL — DSRA + priorité senior → mezz → QE → dividendes
# =============================================================================

def waterfall(rows):
    """
    Retourne pour chaque année :
      FCFF → DSRA top-up → Service Senior → Service Mezz → Service QE → Dividende résiduel
    """
    out = []
    dsra_reserve = rows[0]["DSRA"] if rows else 0.0
    dsra_actuel  = 0.0

    for r in rows:
        fcff = r["FCFF"]
        dsra_cible = dsra_reserve

        # 1. Top-up DSRA si nécessaire
        topup = max(dsra_cible - dsra_actuel, 0)
        topup = min(topup, fcff)
        dsra_actuel += topup
        dispo = fcff - topup

        # 2. Service Senior (prioritaire)
        svc_s = r["Svc_Senior"]
        pay_s = min(svc_s, max(dispo, 0))
        shortfall_s = svc_s - pay_s
        dispo -= pay_s

        # 3. Service Mezzanine
        svc_m = r["Svc_Mezz"]
        pay_m = min(svc_m, max(dispo, 0))
        shortfall_m = svc_m - pay_m
        dispo -= pay_m

        # 4. Service Quasi-Equity
        svc_q = r["Svc_QE"]
        pay_q = min(svc_q, max(dispo, 0))
        shortfall_q = svc_q - pay_q
        dispo -= pay_q

        # 5. Dividende résiduel
        dividende = max(dispo, 0)

        out.append({
            "Annee"        : r["Annee"],
            "FCFF"         : fcff,
            "DSRA_topup"   : topup,
            "Pay_Senior"   : pay_s,
            "Shortfall_S"  : shortfall_s,
            "Pay_Mezz"     : pay_m,
            "Shortfall_M"  : shortfall_m,
            "Pay_QE"       : pay_q,
            "Shortfall_Q"  : shortfall_q,
            "Dividende"    : dividende,
        })
    return out


# =============================================================================
# STRESS TEST MULTI-ANNÉES
# =============================================================================

def stress_test(rows, gravite, prime_ann, indem_pct, duree_choc=1):
    """
    3 scénarios × 2 durées de choc (1 an / 2 ans consécutifs).
    Retourne base, choc, assurance pour la durée_choc choisie.
    """
    duree   = len(rows)
    an_choc = max(1, duree // 2)

    def appliquer(avec_assurance, avec_choc):
        out, treso = [], 0.0
        for r in rows:
            an   = r["Annee"]
            fcff = r["FCFF"]
            rev  = r["Revenus"]
            en_choc = avec_choc and (an_choc <= an < an_choc + duree_choc)
            perte = rev * gravite if en_choc else 0.0
            fcff -= perte
            if avec_assurance:
                fcff -= prime_ann
                if en_choc:
                    fcff += perte * indem_pct
            treso += fcff
            out.append({
                "Annee": an, "FCFF": fcff,
                "Revenus": rev, "Tresorerie_cum": treso,
            })
        return out

    base_r  = appliquer(False, False)
    choc_r  = appliquer(False, True)
    assur_r = appliquer(True,  True)

    def dscr_min_for(sc_rows):
        dscrs = []
        for r_sc, r_orig in zip(sc_rows, rows):
            svc = r_orig["Svc_Total"]
            if svc > 0:
                dscrs.append(r_sc["FCFF"] / svc)
        return (min(dscrs) if dscrs else np.nan), dscrs

    b_min, b_dscr = dscr_min_for(base_r)
    c_min, c_dscr = dscr_min_for(choc_r)
    a_min, a_dscr = dscr_min_for(assur_r)

    perte_brute = rows[an_choc - 1]["Revenus"] * gravite * duree_choc
    indemnite   = perte_brute * indem_pct
    cout_primes = prime_ann * duree

    return {
        "an_choc"     : an_choc,
        "duree_choc"  : duree_choc,
        "perte_brute" : perte_brute,
        "indemnite"   : indemnite,
        "cout_primes" : cout_primes,
        "gain_net"    : indemnite - cout_primes,
        "base"        : {"rows": base_r,  "dscr": b_dscr, "dmin": b_min},
        "choc"        : {"rows": choc_r,  "dscr": c_dscr, "dmin": c_min},
        "assurance"   : {"rows": assur_r, "dscr": a_dscr, "dmin": a_min},
    }


# =============================================================================
# SENSIBILITÉ / TORNADO
# =============================================================================

def calc_tornado(quantite, prix, cout, capex, taux_is, inflation,
                 duree, rampup, bm, wacc, kwargs_proj, delta=0.10):
    def van(**kw):
        return metriques(projeter(**kw), kw["capex"], wacc)["van"]

    base = dict(quantite=quantite, prix=prix, cout=cout, capex=capex,
                taux_is=taux_is, inflation=inflation, duree=duree,
                rampup=rampup, bm=bm, **kwargs_proj)
    van_ref = van(**base)
    result  = []
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
                     duree, rampup, bm, wacc, kwargs_proj,
                     variations=(-0.10, -0.05, 0.0, 0.05, 0.10)):
    labs = [f"{int(v*100):+d}%" for v in variations]
    data = {}
    for vc in variations:
        col = []
        for vp in variations:
            fl = projeter(quantite, prix*(1+vp), cout*(1+vc), capex,
                          taux_is, inflation, duree, rampup, bm, **kwargs_proj)
            col.append(metriques(fl, capex, wacc)["van"])
        data[f"Coût {int(vc*100):+d}%"] = col
    return pd.DataFrame(data, index=[f"Prix {l}" for l in labs])


# =============================================================================
# GRAPHIQUES
# =============================================================================

def fig_jauges(van, tri, tri_eq, wacc_val, dscr_min, icr_min, devise):
    dscr_v = dscr_min if not np.isnan(dscr_min) else 0.0
    icr_v  = icr_min  if (icr_min and not np.isnan(icr_min)) else 0.0
    tri_v  = (tri * 100) if tri else 0.0
    tri_eq_v = (tri_eq * 100) if tri_eq else 0.0
    van_m  = van / 1e6

    fig = make_subplots(
        rows=1, cols=4,
        specs=[[{"type": "indicator"}] * 4],
        horizontal_spacing=0.04,
    )

    van_range = max(abs(van_m) * 2, 100)
    fig.add_trace(go.Indicator(
        mode="gauge+number+delta",
        value=round(van_m, 1),
        title={"text": f"VAN (M{devise})", "font": {"size": 11}},
        delta={"reference": 0},
        gauge={
            "axis": {"range": [-van_range, van_range], "tickformat": ".0f"},
            "bar":  {"color": CLR["vert"] if van > 0 else CLR["rouge"]},
            "threshold": {"line": {"color": "black", "width": 2}, "value": 0},
        },
        number={"font": {"size": 16}},
    ), row=1, col=1)

    fig.add_trace(go.Indicator(
        mode="gauge+number",
        value=round(tri_v, 1),
        title={"text": "TRI Projet (%)", "font": {"size": 11}},
        gauge={
            "axis": {"range": [0, max(tri_v * 1.5, 30)]},
            "bar":  {"color": CLR["vert"] if tri_v > wacc_val * 100 else CLR["rouge"]},
            "steps": [
                {"range": [0, wacc_val * 100], "color": "#fde8e8"},
                {"range": [wacc_val * 100, max(tri_v * 1.5, 30)], "color": "#e6f4ea"},
            ],
            "threshold": {"line": {"color": "navy", "width": 2}, "value": wacc_val * 100},
        },
        number={"suffix": "%", "font": {"size": 16}},
    ), row=1, col=2)

    fig.add_trace(go.Indicator(
        mode="gauge+number",
        value=round(tri_eq_v, 1),
        title={"text": "TRI Actionnaire (%)", "font": {"size": 11}},
        gauge={
            "axis": {"range": [0, max(tri_eq_v * 1.5, 30)]},
            "bar":  {"color": CLR["vert"] if tri_eq_v > 15 else CLR["amber"]},
            "threshold": {"line": {"color": "navy", "width": 2}, "value": 15},
        },
        number={"suffix": "%", "font": {"size": 16}},
    ), row=1, col=3)

    fig.add_trace(go.Indicator(
        mode="gauge+number",
        value=round(dscr_v, 2),
        title={"text": "DSCR Min (x)", "font": {"size": 11}},
        gauge={
            "axis": {"range": [0, max(dscr_v * 1.5, 2.5)]},
            "bar":  {"color": (CLR["vert"] if dscr_v >= 1.3
                               else CLR["amber"] if dscr_v >= 1.0
                               else CLR["rouge"])},
            "steps": [
                {"range": [0, 1.0], "color": "#fde8e8"},
                {"range": [1.0, 1.3], "color": "#fff8e1"},
                {"range": [1.3, max(dscr_v * 1.5, 2.5)], "color": "#e6f4ea"},
            ],
            "threshold": {"line": {"color": "black", "width": 2}, "value": 1.3},
        },
        number={"suffix": "x", "font": {"size": 16}},
    ), row=1, col=4)

    fig.update_layout(
        height=230, margin=dict(l=10, r=10, t=30, b=5),
        paper_bgcolor="white",
    )
    return fig


def fig_waterfall_chart(wf_rows, devise, facteur):
    annees   = [r["Annee"] for r in wf_rows]
    fcff_v   = [r["FCFF"]      / 1e6 * facteur for r in wf_rows]
    senior_v = [r["Pay_Senior"]/ 1e6 * facteur for r in wf_rows]
    mezz_v   = [r["Pay_Mezz"]  / 1e6 * facteur for r in wf_rows]
    qe_v     = [r["Pay_QE"]    / 1e6 * facteur for r in wf_rows]
    div_v    = [r["Dividende"] / 1e6 * facteur for r in wf_rows]

    fig = go.Figure()
    fig.add_trace(go.Bar(name="Service Senior", x=annees, y=senior_v, marker_color=CLR["rouge"]))
    fig.add_trace(go.Bar(name="Service Mezz",   x=annees, y=mezz_v,   marker_color=CLR["amber"]))
    fig.add_trace(go.Bar(name="Service QE",     x=annees, y=qe_v,     marker_color="#9b59b6"))
    fig.add_trace(go.Bar(name="Dividende",      x=annees, y=div_v,    marker_color=CLR["vert"]))
    fig.add_trace(go.Scatter(name="FCFF", x=annees, y=fcff_v,
                              mode="lines+markers", line=dict(color=CLR["bleu"], width=2)))
    fig.update_layout(
        barmode="stack",
        title=f"Waterfall — Emploi du FCFF par tranche (M{devise})",
        xaxis_title="Année", yaxis_title=f"M{devise}",
        height=360, plot_bgcolor="white", paper_bgcolor="white",
        margin=dict(t=50, b=20),
        legend=dict(orientation="h", y=-0.2),
    )
    return fig


def fig_pl_chart(rows, devise, facteur):
    annees  = [r["Annee"] for r in rows]
    rev_v   = [r["Revenus"] / 1e6 * facteur for r in rows]
    ebitda_v= [r["EBITDA"]  / 1e6 * facteur for r in rows]
    ebit_v  = [r["EBIT"]    / 1e6 * facteur for r in rows]
    rn_v    = [r["RN"]      / 1e6 * facteur for r in rows]

    fig = go.Figure()
    fig.add_trace(go.Bar(name="Revenus", x=annees, y=rev_v, marker_color="#adb5bd", opacity=0.6))
    fig.add_trace(go.Scatter(name="EBITDA", x=annees, y=ebitda_v, mode="lines+markers",
                              line=dict(color=CLR["bleu"], width=2)))
    fig.add_trace(go.Scatter(name="EBIT",   x=annees, y=ebit_v,   mode="lines+markers",
                              line=dict(color=CLR["amber"], width=2, dash="dot")))
    fig.add_trace(go.Scatter(name="Rés. Net", x=annees, y=rn_v,   mode="lines+markers",
                              line=dict(color=CLR["vert"], width=2)))
    fig.add_hline(y=0, line_color="#333", line_width=0.8)
    fig.update_layout(
        title=f"P&L Pro Forma (M{devise})",
        xaxis_title="Année", yaxis_title=f"M{devise}",
        height=340, plot_bgcolor="white", paper_bgcolor="white",
        margin=dict(t=50, b=20),
        legend=dict(orientation="h", y=-0.2),
    )
    return fig


def fig_ratios(rows):
    annees   = [r["Annee"] for r in rows]
    dscr_v   = [r["DSCR"]    if not np.isnan(r["DSCR"])    else None for r in rows]
    dscr_s_v = [r["DSCR_Senior"] if not np.isnan(r["DSCR_Senior"]) else None for r in rows]
    icr_v    = [r["ICR"]     if not np.isnan(r.get("ICR", np.nan)) else None for r in rows]
    dn_v     = [r["DN_EBITDA"] if not np.isnan(r.get("DN_EBITDA", np.nan)) else None for r in rows]

    fig = make_subplots(rows=1, cols=2,
                        subplot_titles=("DSCR & ICR", "Dette Nette / EBITDA"))

    fig.add_trace(go.Scatter(x=annees, y=dscr_v,   name="DSCR Total",
                              line=dict(color=CLR["bleu"], width=2)), row=1, col=1)
    fig.add_trace(go.Scatter(x=annees, y=dscr_s_v, name="DSCR Senior",
                              line=dict(color=CLR["navy"], width=2, dash="dot")), row=1, col=1)
    fig.add_trace(go.Scatter(x=annees, y=icr_v,    name="ICR",
                              line=dict(color=CLR["amber"], width=2)), row=1, col=1)
    fig.add_hline(y=1.3, line_dash="dash", line_color=CLR["vert"],  line_width=1,
                  annotation_text="1.3x", row=1, col=1)
    fig.add_hline(y=1.0, line_dash="dot",  line_color=CLR["rouge"], line_width=1,
                  annotation_text="1.0x", row=1, col=1)

    fig.add_trace(go.Scatter(x=annees, y=dn_v, name="DN/EBITDA",
                              fill="tozeroy", line=dict(color=CLR["rouge"], width=2),
                              fillcolor="rgba(220,53,69,0.08)"), row=1, col=2)
    fig.add_hline(y=4.0, line_dash="dash", line_color=CLR["amber"], line_width=1,
                  annotation_text="Seuil 4x", row=1, col=2)

    fig.update_layout(
        height=320, plot_bgcolor="white", paper_bgcolor="white",
        margin=dict(t=40, b=10),
        legend=dict(orientation="h", y=-0.25),
    )
    return fig


def fig_stress(stress, duree, devise, facteur):
    annees  = list(range(1, duree + 1))
    an_choc = stress["an_choc"]

    def clean(lst):
        return [v if (isinstance(v, float) and not np.isnan(v)) else None for v in lst]

    def treso(rows):
        return [r["Tresorerie_cum"] / 1e6 * facteur for r in rows]

    fig = make_subplots(rows=2, cols=1,
                        subplot_titles=("DSCR annuel", f"Trésorerie cumulée (M{devise})"),
                        vertical_spacing=0.18, row_heights=[0.5, 0.5])

    TRACES = [
        ("Base",              "base",      CLR["bleu"],  "solid"),
        ("Choc sans assurance","choc",     CLR["rouge"], "dot"),
        ("Choc avec assurance","assurance",CLR["vert"],  "solid"),
    ]
    for label, cle, couleur, dash in TRACES:
        fig.add_trace(go.Scatter(
            x=annees, y=clean(stress[cle]["dscr"]),
            mode="lines+markers", name=label,
            line=dict(color=couleur, width=2, dash=dash),
            legendgroup=label,
        ), row=1, col=1)
        fig.add_trace(go.Scatter(
            x=annees, y=treso(stress[cle]["rows"]),
            mode="lines", name=label,
            line=dict(color=couleur, width=2, dash=dash),
            fill="tozeroy",
            fillcolor=f"rgba({int(couleur[1:3],16)},{int(couleur[3:5],16)},{int(couleur[5:7],16)},0.08)",
            legendgroup=label, showlegend=False,
        ), row=2, col=1)

    fig.add_hline(y=1.3, line_dash="dash", line_color=CLR["vert"],
                  line_width=1.2, annotation_text="Seuil bancaire 1.3x",
                  annotation_font_size=9, row=1, col=1)
    fig.add_hline(y=1.0, line_dash="dot",  line_color=CLR["rouge"],
                  line_width=1.2, annotation_text="Défaut", row=1, col=1)
    fig.add_hline(y=0, line_color="#333", line_width=0.8, row=2, col=1)

    choc_end = an_choc + stress["duree_choc"] - 1
    for rn in [1, 2]:
        fig.add_vrect(x0=an_choc - 0.45, x1=choc_end + 0.45,
                      fillcolor="rgba(220,53,69,0.10)", line_width=0,
                      annotation_text=f"Choc An {an_choc}–{choc_end}",
                      annotation_font_size=9, row=rn, col=1)

    fig.update_layout(
        height=540, plot_bgcolor="white", paper_bgcolor="white",
        margin=dict(t=50, b=20),
        legend=dict(orientation="h", y=-0.08, font_size=11),
    )
    return fig


# =============================================================================
# EXPORT PDF — MÉMO TEASER (reportlab)
# =============================================================================

def generer_pdf(projet_nom, matiere, bm_label, devise, facteur,
                rows, met, stress, capex_fcfa, p_dette, cfg,
                dette_s, dette_m, dette_q, wacc_val, llcr):
    if not PDF_OK:
        return None

    buf    = io.BytesIO()
    doc    = SimpleDocTemplate(buf, pagesize=A4,
                                leftMargin=2*cm, rightMargin=2*cm,
                                topMargin=2*cm, bottomMargin=2*cm)
    styles = getSampleStyleSheet()
    W      = A4[0] - 4*cm

    # Styles personnalisés
    s_title  = ParagraphStyle("t", parent=styles["Title"],   fontSize=16, spaceAfter=4,
                               textColor=colors.HexColor("#0a3d62"))
    s_sub    = ParagraphStyle("s", parent=styles["Normal"],  fontSize=9,  textColor=colors.grey)
    s_h1     = ParagraphStyle("h1",parent=styles["Heading2"],fontSize=11, textColor=colors.HexColor("#0a3d62"),
                               spaceBefore=10, spaceAfter=4)
    s_body   = ParagraphStyle("b", parent=styles["Normal"],  fontSize=8.5,leading=13)
    s_warn   = ParagraphStyle("w", parent=styles["Normal"],  fontSize=8,
                               textColor=colors.HexColor("#856404"),
                               backColor=colors.HexColor("#fff3cd"))
    s_ok     = ParagraphStyle("ok",parent=styles["Normal"],  fontSize=8,
                               textColor=colors.HexColor("#155724"),
                               backColor=colors.HexColor("#d4edda"))
    s_err    = ParagraphStyle("er",parent=styles["Normal"],  fontSize=8,
                               textColor=colors.HexColor("#721c24"),
                               backColor=colors.HexColor("#f8d7da"))

    story = []

    # ── EN-TÊTE ──────────────────────────────────────────────────────────────
    story.append(Paragraph(f"CommodityWatch — Mémo Teaser Investisseur", s_title))
    story.append(Paragraph(
        f"{projet_nom}  ·  {matiere}  ·  {bm_label}  ·  {devise}  ·  {date.today().strftime('%d/%m/%Y')}",
        s_sub))
    story.append(HRFlowable(width=W, thickness=1.5, color=colors.HexColor("#0a3d62"),
                             spaceAfter=8))

    # ── VERDICT ──────────────────────────────────────────────────────────────
    van_v   = met["van"]
    tri_v   = met["tri"] or 0
    dscr_m  = met["dscr_min"]
    score = sum([
        van_v > 0,
        tri_v > wacc_val,
        not np.isnan(dscr_m) and dscr_m >= 1.3,
        not np.isnan(llcr)   and llcr   >= 1.1,
    ])
    VERDICTS_T = {
        4: ("PROJET VIABLE — 4/4 critères",      s_ok),
        3: ("PROJET FAVORABLE — 3/4 critères",    s_ok),
        2: ("PROJET À AMÉLIORER — 2/4 critères", s_warn),
        1: ("PROJET RISQUÉ — 1/4 critères",       s_err),
        0: ("PROJET NON VIABLE — 0/4 critères",   s_err),
    }
    v_txt, v_sty = VERDICTS_T[score]
    story.append(Paragraph(f"  {v_txt}  ", v_sty))
    story.append(Spacer(1, 8))

    # ── KPIs CLÉS (tableau 2 lignes × 4 cols) ────────────────────────────────
    story.append(Paragraph("Indicateurs Clés de Performance", s_h1))

    def fv(val, suf="", dec=1):
        if val is None or (isinstance(val, float) and np.isnan(val)):
            return "N/D"
        return f"{val:,.{dec}f}{suf}"

    van_a   = m(van_v, devise, 1)
    tri_a   = fv(tri_v * 100, "%")
    tri_eq_a= fv((met.get("tri_eq") or 0) * 100, "%")
    pb_a    = f"{met['payback']} ans" if met["payback"] else "> durée"
    dscr_a  = fv(dscr_m, "x", 2)
    icr_a   = fv(met.get("icr_min"), "x", 2)
    llcr_a  = fv(llcr, "x", 2)
    wacc_a  = fv(wacc_val * 100, "%")

    kpi_data = [
        ["VAN", f"{van_a:,.1f} M{devise}", "TRI Projet", tri_a],
        ["TRI Actionnaire", tri_eq_a,       "Payback",    pb_a],
        ["DSCR Minimum",   dscr_a,          "LLCR",       llcr_a],
        ["ICR Minimum",    icr_a,           "WACC",       wacc_a],
    ]
    kpi_style = TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#f8f9fa")),
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#e9ecef")),
        ("BACKGROUND", (2, 0), (2, -1), colors.HexColor("#e9ecef")),
        ("FONTSIZE",   (0, 0), (-1, -1), 8.5),
        ("FONTNAME",   (1, 0), (1, -1), "Helvetica-Bold"),
        ("FONTNAME",   (3, 0), (3, -1), "Helvetica-Bold"),
        ("GRID",       (0, 0), (-1, -1), 0.5, colors.HexColor("#dee2e6")),
        ("VALIGN",     (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING",(0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING",(0,0),(-1,-1),  4),
    ])
    kpi_tbl = Table(kpi_data, colWidths=[W*0.22, W*0.28, W*0.22, W*0.28])
    kpi_tbl.setStyle(kpi_style)
    story.append(kpi_tbl)
    story.append(Spacer(1, 8))

    # ── STRUCTURE FINANCIÈRE ─────────────────────────────────────────────────
    story.append(Paragraph("Structure de Financement", s_h1))
    capex_m  = m(capex_fcfa, devise, 0)
    det_s_m  = m(dette_s,    devise, 0)
    det_m_m  = m(dette_m,    devise, 0)
    det_q_m  = m(dette_q,    devise, 0)
    fp_m     = m(capex_fcfa - dette_s - dette_m - dette_q, devise, 0)

    fin_data = [
        ["CAPEX Total",    f"{capex_m:,.0f} M{devise}", "Fonds Propres",    f"{fp_m:,.0f} M{devise}"],
        ["Dette Senior",   f"{det_s_m:,.0f} M{devise}", "Mezzanine",        f"{det_m_m:,.0f} M{devise}"],
        ["Quasi-Equity",   f"{det_q_m:,.0f} M{devise}", "DSRA (réserve)",   f"{m(rows[0]['DSRA'],devise,1):,.1f} M{devise}"],
    ]
    fin_tbl = Table(fin_data, colWidths=[W*0.22, W*0.28, W*0.22, W*0.28])
    fin_tbl.setStyle(kpi_style)
    story.append(fin_tbl)
    story.append(Spacer(1, 8))

    # ── TABLEAU P&L SYNTHÉTIQUE ──────────────────────────────────────────────
    story.append(Paragraph("P&L Pro Forma — Synthèse (M" + devise + ")", s_h1))
    hdr = ["An", "Revenus", "EBITDA", "EBIT", "Int.", "RN", "FCFF", "DSCR"]
    pl_data = [hdr]
    for r in rows:
        pl_data.append([
            str(r["Annee"]),
            f"{m(r['Revenus'], devise, 1):,.1f}",
            f"{m(r['EBITDA'],  devise, 1):,.1f}",
            f"{m(r['EBIT'],    devise, 1):,.1f}",
            f"{m(r['Int_Total'],devise,1):,.1f}",
            f"{m(r['RN'],      devise, 1):,.1f}",
            f"{m(r['FCFF'],    devise, 1):,.1f}",
            f"{r['DSCR']:.2f}x" if not np.isnan(r["DSCR"]) else "—",
        ])
    pl_style = TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0a3d62")),
        ("TEXTCOLOR",  (0, 0), (-1, 0), colors.white),
        ("FONTNAME",   (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE",   (0, 0), (-1, -1), 7.5),
        ("GRID",       (0, 0), (-1, -1), 0.4, colors.HexColor("#dee2e6")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1),
         [colors.white, colors.HexColor("#f8f9fa")]),
        ("VALIGN",     (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN",      (1, 0), (-1, -1), "RIGHT"),
        ("LEFTPADDING",(0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING",(0,0),(-1,-1),  3),
    ])
    cw = W / len(hdr)
    pl_tbl = Table(pl_data, colWidths=[cw] * len(hdr))
    pl_tbl.setStyle(pl_style)
    story.append(pl_tbl)
    story.append(Spacer(1, 8))

    # ── STRESS TEST ──────────────────────────────────────────────────────────
    story.append(Paragraph("Stress Test Assurance", s_h1))
    d_b = stress["base"]["dmin"]
    d_c = stress["choc"]["dmin"]
    d_a = stress["assurance"]["dmin"]

    def badge(v):
        if np.isnan(v): return "N/D"
        return f"{v:.2f}x"

    st_data = [
        ["Scénario",          "DSCR Min", "Perte CA",         "Indemnité",        "Primes cumulées"],
        ["Base",              badge(d_b), "—",                "—",                "—"],
        ["Choc sans assurance",badge(d_c),f"{m(stress['perte_brute'],devise,1):.1f} M{devise}","—","—"],
        ["Choc avec assurance",badge(d_a),f"{m(stress['perte_brute'],devise,1):.1f} M{devise}",
         f"{m(stress['indemnite'],devise,1):.1f} M{devise}",
         f"{m(stress['cout_primes'],devise,1):.1f} M{devise}"],
    ]
    st_style = TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0a3d62")),
        ("TEXTCOLOR",  (0, 0), (-1, 0), colors.white),
        ("FONTNAME",   (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE",   (0, 0), (-1, -1), 7.5),
        ("GRID",       (0, 0), (-1, -1), 0.4, colors.HexColor("#dee2e6")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1),
         [colors.white, colors.HexColor("#fff3cd"), colors.HexColor("#d4edda")]),
        ("VALIGN",     (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING",(0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING",(0,0),(-1,-1),  3),
    ])
    st_tbl = Table(st_data, colWidths=[W*0.24, W*0.14, W*0.20, W*0.20, W*0.22])
    st_tbl.setStyle(st_style)
    story.append(st_tbl)
    story.append(Spacer(1, 8))

    # ── RISQUE & ASSURANCE ───────────────────────────────────────────────────
    story.append(Paragraph("Risque Identifié & Couverture", s_h1))
    story.append(Paragraph(f"<b>Risque :</b> {cfg['risque_titre']} — {cfg['risque_desc']}", s_body))
    story.append(Spacer(1, 4))
    story.append(Paragraph(f"<b>Couverture :</b> {cfg['assurance_titre']} — {cfg['assurance_desc']}", s_body))
    story.append(Spacer(1, 10))

    # ── DISCLAIMER ───────────────────────────────────────────────────────────
    story.append(HRFlowable(width=W, thickness=0.5, color=colors.grey, spaceAfter=4))
    story.append(Paragraph(
        f"CommodityWatch v5.0 — Document indicatif pré-investissement. "
        f"1 USD = {TAUX_CHANGE:.0f} FCFA. Ne se substitue pas à une due diligence "
        "financière et juridique complète. Confidentiel.",
        ParagraphStyle("disc", parent=styles["Normal"], fontSize=7, textColor=colors.grey)))

    doc.build(story)
    buf.seek(0)
    return buf.getvalue()


# =============================================================================
# PAGE CONFIG & CSS
# =============================================================================

st.set_page_config(
    page_title="CommodityWatch v5.0 — DFI / Banque",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    .block-container { padding: 1.2rem 2rem 2rem 2rem; }
    h1  { font-size: 1.45rem; font-weight: 700; color: #0d1117; }
    h2  { font-size: 1.0rem; font-weight: 600; color: #24292f;
          border-bottom: 2px solid #e8e8e8; padding-bottom: 4px; margin-top:1.4rem; }
    h3  { font-size: .88rem; font-weight: 600; color: #24292f; }
    .stMetric label { font-size: .72rem; color: #57606a;
                      text-transform: uppercase; letter-spacing: .04em; }
    hr  { border: none; border-top: 1px solid #e8e8e8; margin: 1rem 0; }
    .bm-banner {
        background: linear-gradient(90deg, #0a3d62, #1a5276);
        color: white; border-radius: 8px; padding: 10px 20px;
        font-size: .9rem; font-weight: 600; margin-bottom: .8rem;
    }
    .verdict-box {
        border-radius: 8px; padding: 12px 20px;
        font-size: 1.05rem; font-weight: 700;
        text-align: center; color: white; margin-bottom: .6rem;
    }
    .tranche-card {
        background: #f8f9fa; border-left: 4px solid #0a3d62;
        border-radius: 6px; padding: 8px 14px; margin-bottom: 8px;
        font-size: .82rem;
    }
</style>
""", unsafe_allow_html=True)


# =============================================================================
# BARRE LATÉRALE
# =============================================================================

with st.sidebar:
    st.markdown("### CommodityWatch v5.0")
    st.caption("DFI · Banque commerciale · BOAD")
    st.markdown("---")

    # Nom du projet
    projet_nom = st.text_input("Nom du projet", value="Projet Agri Alpha")

    devise  = st.radio("Devise", ["FCFA", "USD"], horizontal=True)
    facteur = 1.0 if devise == "FCFA" else 1.0 / TAUX_CHANGE

    st.markdown("---")
    st.markdown("#### Business Model")
    bm = st.selectbox("Type de Projet", ["Production", "Infrastructure", "Campagne"])
    cfg = BM_CONFIG[bm]
    st.caption(cfg["description"])

    duree_bio = 0
    if cfg["bio_actif"]:
        duree_bio = st.number_input("Années sans revenu (croissance)", 0, 10, 3, 1)

    j_stock = j_clients = j_fourn = 0
    if cfg["bfr_actif"]:
        st.markdown("**Cycle d'Exploitation**")
        j_stock   = st.number_input("Stock matière (jours)",    0, 180, 60, 5)
        j_clients = st.number_input("Crédit clients (jours)",   0, 120, 30, 5)
        j_fourn   = st.number_input("Crédit fournisseurs (j.)", 0, 120, 45, 5)

    st.markdown("---")
    st.markdown("#### Produit & Prix")
    matiere = st.selectbox("Matière première", [
        "Cacao","Café (Arabica)","Café (Robusta)","Anacarde (brut)",
        "Soja","Coton","Caoutchouc","Palmier à huile","Maïs","Personnalisé",
    ])
    prix_input = st.number_input(
        f"Prix de Vente ({devise}/T)", min_value=0.0,
        value=7800.0 if devise == "USD" else 4_680_000.0,
        step=10.0 if devise == "USD" else 10_000.0,
        format="%.1f" if devise == "USD" else "%.0f",
    )
    prix_fcfa_v = fcfa(prix_input, devise)

    quantite = st.number_input("Volume pleine capacité (T/an)", 0, 500_000, 5_000, 100)

    cout_input = st.number_input(
        f"Coût de Production ({devise}/T)", min_value=0.0,
        value=1_000.0 if devise == "USD" else 600_000.0,
        step=10.0 if devise == "USD" else 10_000.0,
        format="%.1f" if devise == "USD" else "%.0f",
    )
    cout_fcfa_v = fcfa(cout_input, devise)

    st.markdown("---")
    st.markdown("#### Ramp-up")
    n_pal   = int(st.number_input("Paliers", 1, 5, 3, 1))
    d_an_df = [2, 3, 4, 5, 6]
    d_tx_df = [0, 40, 80, 100, 100]
    paliers = []
    for k in range(n_pal):
        c1, c2 = st.columns(2)
        with c1:
            af = st.number_input("Fin an", 1, 25, d_an_df[k], 1, key=f"an{k}")
        with c2:
            tx = st.number_input("%", 0, 100, d_tx_df[k], 5, key=f"tx{k}")
        paliers.append((int(af), tx / 100))

    st.markdown("---")
    st.markdown("#### CAPEX & Durée")
    capex_input = st.number_input(
        f"CAPEX Total ({devise})", min_value=0.0,
        value=3_000_000.0 if devise == "USD" else 2_000_000_000.0,
        step=100_000.0 if devise == "USD" else 100_000_000.0,
        format="%.0f",
    )
    capex_fcfa_v = fcfa(capex_input, devise)

    duree       = st.slider("Durée projet (ans)", 3, 25, 12, 1)
    infl_pct    = st.slider("Inflation annuelle (%)", 0, 15, 3, 1)
    inflation   = infl_pct / 100

    # ── DETTE SENIOR ─────────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("#### 🏦 Tranche Senior")
    p_senior_pct = st.slider("Part Senior (% CAPEX)", 0, 80, 50, 5)
    r_senior_pct = st.slider("Taux Senior (%/an)", 1, 20, 9, 1)
    dur_senior   = st.slider("Durée Senior (ans)", 1, duree, min(8, duree), 1)
    grace_senior = st.slider("Différé Senior (ans)", 0, max(dur_senior-1,1),
                              min(2, dur_senior-1), 1)
    dette_senior_fcfa = capex_fcfa_v * p_senior_pct / 100
    r_senior          = r_senior_pct / 100

    # ── MEZZANINE ─────────────────────────────────────────────────────────────
    st.markdown("#### 📊 Tranche Mezzanine")
    p_mezz_pct   = st.slider("Part Mezz (% CAPEX)", 0, 40, 15, 5)
    r_mezz_pct   = st.slider("Taux Mezz (%/an)", 1, 25, 12, 1)
    dur_mezz     = st.slider("Durée Mezz (ans)", 1, duree, min(6, duree), 1)
    grace_mezz   = st.slider("Différé Mezz (ans)", 0, max(dur_mezz-1,1),
                              min(2, dur_mezz-1), 1)
    dette_mezz_fcfa = capex_fcfa_v * p_mezz_pct / 100
    r_mezz          = r_mezz_pct / 100

    # ── QUASI-EQUITY ──────────────────────────────────────────────────────────
    st.markdown("#### 💼 Quasi-Equity")
    p_qe_pct   = st.slider("Part QE (% CAPEX)", 0, 20, 5, 5)
    r_qe_pct   = st.slider("Taux QE (%/an)", 5, 30, 15, 1)
    dur_qe     = st.slider("Durée QE (ans)", 1, duree, min(5, duree), 1)
    grace_qe   = st.slider("Différé QE (ans)", 0, max(dur_qe-1,1),
                            min(1, dur_qe-1), 1)
    dette_qe_fcfa  = capex_fcfa_v * p_qe_pct / 100
    r_qe           = r_qe_pct / 100

    p_fp_pct = 100 - p_senior_pct - p_mezz_pct - p_qe_pct
    if p_fp_pct < 0:
        st.error(f"⚠️ Les tranches dépassent 100% du CAPEX ({100-p_fp_pct:.0f}% alloué). Réduire les parts.")

    # ── PARAMÈTRES FISCAUX ────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("#### Paramètres Fiscaux (OHADA)")
    pays = st.selectbox("Pays / Référentiel", list(PAYS_FISCAL.keys()))
    taux_is_defaut = PAYS_FISCAL[pays]["is"]
    taux_is_pct = st.slider("IS nominal (%)", 0, 40, int(taux_is_defaut * 100), 1)
    taux_is     = taux_is_pct / 100
    zone_franche = st.checkbox("Exonération zone franche / code investissements")
    duree_exo    = 0
    if zone_franche:
        duree_exo = st.number_input("Durée exonération IS (ans)", 1, 15, 5, 1)

    # ── WACC ─────────────────────────────────────────────────────────────────
    r_fp_pct = st.slider("Coût fonds propres (%)", 5, 30, 15, 1)
    r_fp     = r_fp_pct / 100

    # ── ASSURANCE ─────────────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown(f"#### {cfg['assurance_titre']}")
    st.caption(cfg["assurance_desc"])
    gravite = st.slider("Gravité sinistre — perte CA (%)", 5, 80,
                         int(cfg["gravite_defaut"] * 100), 5) / 100
    prime_input = st.number_input(
        f"Prime annuelle ({devise})", min_value=0.0,
        value=50_000.0 if devise == "USD" else 30_000_000.0,
        step=5_000.0 if devise == "USD" else 5_000_000.0,
        format="%.0f",
    )
    prime_fcfa_v = fcfa(prime_input, devise)
    indem_pct    = st.slider("Taux couverture (% perte indemnisé)", 0, 100, 70, 5) / 100
    duree_choc   = st.radio("Durée du choc stress test", [1, 2], horizontal=True,
                             format_func=lambda x: f"{x} an{'s' if x > 1 else ''}")

    st.markdown("---")
    lancer = st.button("▶  Générer l'Overview", type="primary", use_container_width=True)


# =============================================================================
# HEADER PRINCIPAL
# =============================================================================

st.markdown(
    f"<div class='bm-banner'>"
    f"{cfg['icon']}  CommodityWatch v5.0 — {projet_nom} &nbsp;|&nbsp; "
    f"{matiere} &nbsp;|&nbsp; {bm} &nbsp;|&nbsp; {devise}"
    f"</div>",
    unsafe_allow_html=True,
)
st.caption("Overview pré-investissement · Public : DFI · Banque commerciale · BOAD")

if not lancer:
    c1, c2, c3 = st.columns(3)
    c1.info(f"**{cfg['icon']} {bm}**\n\n{cfg['description']}")
    c2.info(f"**Risque couvert**\n\n{cfg['risque_titre']}")
    c3.info(f"**Assurance associée**\n\n{cfg['assurance_titre']}")
    st.info("Renseignez les paramètres puis cliquez sur **▶ Générer l'Overview**.", icon="👈")
    st.stop()


# =============================================================================
# CALCULS
# =============================================================================

with st.spinner("Modélisation en cours…"):
    # WACC pondéré 3 tranches + FP
    p_s = p_senior_pct / 100
    p_m = p_mezz_pct   / 100
    p_q = p_qe_pct     / 100
    p_f = max(p_fp_pct, 0) / 100

    wacc_val = (r_senior * (1 - taux_is) * p_s
              + r_mezz   * (1 - taux_is) * p_m
              + r_qe                     * p_q
              + r_fp                     * p_f)

    rampup = build_rampup(duree, paliers)

    kwargs_proj = dict(
        duree_bio=duree_bio, j_stock=j_stock, j_clients=j_clients, j_fourn=j_fourn,
        dette_senior=dette_senior_fcfa, r_senior=r_senior,
        dur_senior=dur_senior, grace_senior=grace_senior,
        dette_mezz=dette_mezz_fcfa, r_mezz=r_mezz,
        dur_mezz=dur_mezz, grace_mezz=grace_mezz,
        dette_qe=dette_qe_fcfa, r_qe=r_qe,
        dur_qe=dur_qe, grace_qe=grace_qe,
        zone_franche=zone_franche, duree_exo=duree_exo,
    )

    rows = projeter(
        quantite, prix_fcfa_v, cout_fcfa_v, capex_fcfa_v,
        taux_is, inflation, duree, rampup, bm,
        **kwargs_proj,
    )

    met  = metriques(rows, capex_fcfa_v, wacc_val)
    llcr = calc_llcr(rows, dette_senior_fcfa, r_senior)
    wf   = waterfall(rows)

    stress = stress_test(
        rows, gravite, prime_fcfa_v, indem_pct, duree_choc=int(duree_choc),
    )

    alertes_chg = [r["Annee"] for r in rows
                   if not np.isnan(r["Ratio_Charges"]) and r["Ratio_Charges"] > SEUIL_CHARGES]

    tornado_d, _ = calc_tornado(
        quantite, prix_fcfa_v, cout_fcfa_v, capex_fcfa_v,
        taux_is, inflation, duree, rampup, bm, wacc_val, kwargs_proj,
    )

    df_sens = calc_sensibilite(
        quantite, prix_fcfa_v, cout_fcfa_v, capex_fcfa_v,
        taux_is, inflation, duree, rampup, bm, wacc_val, kwargs_proj,
    ) * facteur / 1e6


# =============================================================================
# SECTION 1 — VERDICT + JAUGES
# =============================================================================

van_v   = met["van"]
tri_v   = met["tri"] or 0
dscr_m  = met["dscr_min"]
tri_eq  = met.get("tri_eq") or 0
icr_min = met.get("icr_min", np.nan)

score = sum([
    van_v > 0,
    tri_v > wacc_val,
    not np.isnan(dscr_m) and dscr_m >= SEUIL_BANCAIRE,
    not np.isnan(llcr)   and llcr   >= 1.10,
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

st.plotly_chart(
    fig_jauges(van_v, tri_v, tri_eq, wacc_val, dscr_m, icr_min, devise),
    use_container_width=True,
)

# Alertes immédiates
if alertes_chg:
    st.warning(f"Charges opérationnelles > {SEUIL_CHARGES*100:.0f}% du CA aux années **{alertes_chg}**.", icon="⚠️")
if not np.isnan(dscr_m):
    if dscr_m < SEUIL_DEFAUT:
        st.error(f"DSCR Min = **{dscr_m:.2f}x** — Risque de défaut. Service de la dette non couvert.", icon="🚨")
    elif dscr_m < SEUIL_BANCAIRE:
        st.warning(f"DSCR Min = **{dscr_m:.2f}x** — Sous le seuil bancaire (1.3x). Renégocier maturité ou grâce.", icon="⚠️")
if not np.isnan(icr_min) and icr_min < 2.0:
    st.warning(f"ICR Min = **{icr_min:.2f}x** — Coverage des intérêts faible (seuil DFI : 2.0x).", icon="⚠️")

st.markdown("---")

# Résumé structure dette (bandeau compact)
col1, col2, col3, col4 = st.columns(4)
col1.markdown(f"<div class='tranche-card'>🏦 <b>Senior</b> {p_senior_pct}% — {r_senior_pct}%/an — {dur_senior}a</div>", unsafe_allow_html=True)
col2.markdown(f"<div class='tranche-card'>📊 <b>Mezz</b> {p_mezz_pct}% — {r_mezz_pct}%/an — {dur_mezz}a</div>", unsafe_allow_html=True)
col3.markdown(f"<div class='tranche-card'>💼 <b>Quasi-Eq.</b> {p_qe_pct}% — {r_qe_pct}%/an — {dur_qe}a</div>", unsafe_allow_html=True)
col4.markdown(f"<div class='tranche-card'>💰 <b>FP</b> {max(p_fp_pct,0)}% — coût {r_fp_pct}%</div>", unsafe_allow_html=True)

st.markdown("---")


# =============================================================================
# SECTION 2 — TABLEAU DE SYNTHÈSE COMPLET
# =============================================================================

st.markdown("## Tableau de Synthèse")

def ind(ok):
    return "✅" if ok else "❌"

def vf(val_fcfa, dec=1):
    return round(affiche(val_fcfa, devise) / 1e6, dec)

fp_total = capex_fcfa_v - dette_senior_fcfa - dette_mezz_fcfa - dette_qe_fcfa
dsra_val = rows[0]["DSRA"] if rows else 0

synth = {
    "Rubrique": [
        "─── RENTABILITÉ ───",
        "VAN (Projet)",
        "TRI Projet",
        "TRI Actionnaire",
        "WACC",
        "Payback",
        "",
        "─── FINANCEMENT ───",
        "CAPEX Total",
        "  Dont Senior",
        "  Dont Mezzanine",
        "  Dont Quasi-Equity",
        "  Dont Fonds Propres",
        "DSRA (réserve 6 mois)",
        "",
        "─── BANCABILITÉ ───",
        "DSCR Minimum",
        "DSCR Senior Minimum",
        "DSCR Moyen",
        "LLCR",
        "ICR Minimum",
        "Gearing (Dte/FP) — An 1",
        "Dette Nette/EBITDA — An 1",
        "",
        "─── FISCAL ───",
        "Référentiel",
        "IS applicable",
        "Zone franche",
        "",
        "─── RISQUE & ASSURANCE ───",
        "Risque identifié",
        "Perte CA (choc)",
        "Indemnité estimée",
        "Coût cumulé primes",
        "Gain net assurance",
    ],
    "Valeur": [
        "",
        f"{vf(van_v):,.1f} M{devise}",
        f"{tri_v*100:.2f}%" if tri_v else "N/D",
        f"{tri_eq*100:.2f}%" if tri_eq else "N/D",
        f"{wacc_val*100:.2f}%",
        f"{met['payback']} ans" if met["payback"] else "> durée",
        "",
        "",
        f"{vf(capex_fcfa_v,0):,.0f} M{devise}",
        f"{vf(dette_senior_fcfa,0):,.0f} M{devise} ({p_senior_pct}%)",
        f"{vf(dette_mezz_fcfa,0):,.0f} M{devise} ({p_mezz_pct}%)",
        f"{vf(dette_qe_fcfa,0):,.0f} M{devise} ({p_qe_pct}%)",
        f"{vf(fp_total,0):,.0f} M{devise} ({max(p_fp_pct,0)}%)",
        f"{vf(dsra_val):,.1f} M{devise}",
        "",
        "",
        f"{dscr_m:.2f}x" if not np.isnan(dscr_m) else "N/D",
        f"{min([r['DSCR_Senior'] for r in rows if not np.isnan(r['DSCR_Senior'])], default=float('nan')):.2f}x",
        f"{met['dscr_moy']:.2f}x" if not np.isnan(met['dscr_moy']) else "N/D",
        f"{llcr:.2f}x" if not np.isnan(llcr) else "N/D",
        f"{icr_min:.2f}x" if (icr_min and not np.isnan(icr_min)) else "N/D",
        f"{rows[0]['Gearing']:.2f}x" if rows and not np.isnan(rows[0]['Gearing']) else "N/D",
        f"{rows[0]['DN_EBITDA']:.2f}x" if rows and not np.isnan(rows[0]['DN_EBITDA']) else "N/D",
        "",
        "",
        pays,
        f"{taux_is*100:.0f}%" + (" (nominal)" if not zone_franche else f" → 0% sur {duree_exo} ans"),
        "Oui" if zone_franche else "Non",
        "",
        "",
        cfg["risque_titre"],
        f"{vf(stress['perte_brute']):.1f} M{devise}",
        f"{vf(stress['indemnite']):.1f} M{devise}",
        f"{vf(stress['cout_primes']):.1f} M{devise}",
        f"{vf(stress['gain_net']):.1f} M{devise}",
    ],
    "Verdict": [
        "", ind(van_v > 0), ind(tri_v > wacc_val), ind(tri_eq > 0.15),
        "─", "─",
        "",
        "", "─", "─", "─", "─", "─", "─",
        "",
        "",
        ind(not np.isnan(dscr_m) and dscr_m >= 1.3),
        ind(not np.isnan(dscr_m) and dscr_m >= 1.3),
        ind(not np.isnan(met['dscr_moy']) and met['dscr_moy'] >= 1.3),
        ind(not np.isnan(llcr) and llcr >= 1.1),
        ind(not (icr_min and not np.isnan(icr_min)) or icr_min >= 2.0),
        "─", "─",
        "",
        "", "─", "─", "─",
        "",
        "", "─", "─", "─", "─", ind(stress["gain_net"] > 0),
    ],
}

df_synth = pd.DataFrame(synth)

def style_synthese(row):
    if str(row["Rubrique"]).startswith("───"):
        return ["background:#e8f0fe;font-weight:700;color:#0a3d62"] * len(row)
    if row["Rubrique"] == "":
        return ["background:white;border:none"] * len(row)
    return [""] * len(row)

st.dataframe(
    df_synth.style.apply(style_synthese, axis=1),
    use_container_width=True, hide_index=True, height=870,
)

st.markdown("---")


# =============================================================================
# SECTION 3 — P&L PRO FORMA + GRAPHIQUE
# =============================================================================

st.markdown("## P&L Pro Forma & Flux de Trésorerie")

tab_pl, tab_proj = st.tabs(["P&L Complet", "Tableau FCFF / FCFE"])

with tab_pl:
    df_pl = pd.DataFrame([{
        "An"                     : r["Annee"],
        "Cap.(%)"                : f"{r['Ramp_up']:.0f}%",
        f"CA M{devise}"          : round(affiche(r["Revenus"], devise)/1e6, 2),
        f"EBITDA M{devise}"      : round(affiche(r["EBITDA"],  devise)/1e6, 2),
        "Mg.EBITDA"              : f"{r['Marge_EBITDA']*100:.1f}%" if not np.isnan(r["Marge_EBITDA"]) else "—",
        f"EBIT M{devise}"        : round(affiche(r["EBIT"],    devise)/1e6, 2),
        f"Int. M{devise}"        : round(affiche(r["Int_Total"],devise)/1e6, 2),
        f"EBT M{devise}"         : round(affiche(r["EBT"],     devise)/1e6, 2),
        f"IS M{devise}"          : round(affiche(r["IS"],      devise)/1e6, 2),
        f"Rés.Net M{devise}"     : round(affiche(r["RN"],      devise)/1e6, 2),
        "Mg.RN"                  : f"{r['Marge_RN']*100:.1f}%" if not np.isnan(r.get("Marge_RN", np.nan)) else "—",
        "ICR"                    : f"{r['ICR']:.1f}x" if not np.isnan(r.get("ICR", np.nan)) else "—",
    } for r in rows])

    def style_pl(row):
        rn_col = f"Rés.Net M{devise}"
        if row.get(rn_col, 0) < 0:
            return ["background:#f8d7da"] * len(row)
        return [""] * len(row)

    st.dataframe(df_pl.style.apply(style_pl, axis=1),
                 use_container_width=True, hide_index=True)
    st.plotly_chart(fig_pl_chart(rows, devise, facteur), use_container_width=True)

with tab_proj:
    df_proj = pd.DataFrame([{
        "An"                   : r["Annee"],
        f"FCFF M{devise}"      : round(affiche(r["FCFF"],  devise)/1e6, 2),
        f"Svc.Senior M{devise}": round(affiche(r["Svc_Senior"],devise)/1e6, 2),
        f"Svc.Mezz M{devise}"  : round(affiche(r["Svc_Mezz"],  devise)/1e6, 2),
        f"Svc.QE M{devise}"    : round(affiche(r["Svc_QE"],    devise)/1e6, 2),
        f"FCFE M{devise}"      : round(affiche(r["FCFE"],  devise)/1e6, 2),
        "DSCR"                 : f"{r['DSCR']:.2f}x"       if not np.isnan(r["DSCR"])       else "—",
        "DSCR Senior"          : f"{r['DSCR_Senior']:.2f}x"if not np.isnan(r["DSCR_Senior"])else "—",
        "DN/EBITDA"            : f"{r['DN_EBITDA']:.1f}x"  if not np.isnan(r.get("DN_EBITDA",np.nan)) else "—",
    } for r in rows])

    def style_proj(row):
        fcff_col = f"FCFF M{devise}"
        fcfe_col = f"FCFE M{devise}"
        if row.get(fcff_col, 0) < 0 or row.get(fcfe_col, 0) < 0:
            return ["background:#f8d7da"] * len(row)
        return [""] * len(row)

    st.dataframe(df_proj.style.apply(style_proj, axis=1),
                 use_container_width=True, hide_index=True)
    st.caption("🔴 FCFF ou FCFE négatif")

st.markdown("---")


# =============================================================================
# SECTION 4 — BILAN SIMPLIFIÉ
# =============================================================================

st.markdown("## Bilan Simplifié Prévisionnel")

df_bilan = pd.DataFrame([{
    "An"                       : r["Annee"],
    f"Immo.Net M{devise}"      : round(affiche(r["Immo_Net"],     devise)/1e6, 1),
    f"BFR M{devise}"           : round(affiche(r["BFR"],          devise)/1e6, 1),
    f"DSRA M{devise}"          : round(affiche(r["DSRA"],         devise)/1e6, 1),
    f"Enc.Senior M{devise}"    : round(affiche(r["Enc_Senior"],   devise)/1e6, 1),
    f"Enc.Mezz M{devise}"      : round(affiche(r["Enc_Mezz"],     devise)/1e6, 1),
    f"Enc.QE M{devise}"        : round(affiche(r["Enc_QE"],       devise)/1e6, 1),
    f"Dette Tot. M{devise}"    : round(affiche(r["Dette_Totale"], devise)/1e6, 1),
    "Gearing"                  : f"{r['Gearing']:.2f}x" if not np.isnan(r["Gearing"]) else "—",
} for r in rows])

st.dataframe(df_bilan, use_container_width=True, hide_index=True)
st.markdown("---")


# =============================================================================
# SECTION 5 — WATERFALL & RATIOS BANCAIRES
# =============================================================================

st.markdown("## Waterfall & Ratios Bancaires")

col_wf, col_rt = st.columns([3, 2])

with col_wf:
    st.plotly_chart(fig_waterfall_chart(wf, devise, facteur), use_container_width=True)

with col_rt:
    st.plotly_chart(fig_ratios(rows), use_container_width=True)

# Tableau waterfall détaillé
with st.expander("Détail Waterfall annuel"):
    df_wf = pd.DataFrame([{
        "An"                      : r["Annee"],
        f"FCFF M{devise}"         : round(affiche(r["FCFF"],       devise)/1e6, 2),
        f"DSRA top-up M{devise}"  : round(affiche(r["DSRA_topup"], devise)/1e6, 2),
        f"Pay.Senior M{devise}"   : round(affiche(r["Pay_Senior"], devise)/1e6, 2),
        f"Shortfall S M{devise}"  : round(affiche(r["Shortfall_S"],devise)/1e6, 2),
        f"Pay.Mezz M{devise}"     : round(affiche(r["Pay_Mezz"],   devise)/1e6, 2),
        f"Pay.QE M{devise}"       : round(affiche(r["Pay_QE"],     devise)/1e6, 2),
        f"Dividende M{devise}"    : round(affiche(r["Dividende"],  devise)/1e6, 2),
    } for r in wf])

    def style_wf(row):
        sf_col = f"Shortfall S M{devise}"
        if row.get(sf_col, 0) > 0:
            return ["background:#f8d7da"] * len(row)
        return [""] * len(row)

    st.dataframe(df_wf.style.apply(style_wf, axis=1),
                 use_container_width=True, hide_index=True)
    st.caption("🔴 Shortfall = service senior non couvert par le FCFF disponible")

    la, lb, lc = st.columns(3)
    la.markdown("<span style='background:#d4edda;padding:1px 7px;border-radius:3px;font-size:.76rem'>DSCR ≥ 1.3x</span>", unsafe_allow_html=True)
    lb.markdown("<span style='background:#fde8c8;padding:1px 7px;border-radius:3px;font-size:.76rem'>1.0 – 1.3x</span>", unsafe_allow_html=True)
    lc.markdown("<span style='background:#f8d7da;padding:1px 7px;border-radius:3px;font-size:.76rem'>< 1.0x / Shortfall</span>", unsafe_allow_html=True)

st.markdown("---")


# =============================================================================
# SECTION 6 — SENSIBILITÉ (TORNADO + HEATMAP)
# =============================================================================

st.markdown("## Analyse de Sensibilité")
t1, t2 = st.tabs(["Variables Clés (Tornado)", "Croisé Prix × Coût"])

with t1:
    c_tor, c_tab = st.columns([3, 2])
    with c_tor:
        labels = [d["Variable"] for d in tornado_d]
        im     = [d["im"] / 1e6 * facteur for d in tornado_d]
        ip     = [d["ip"] / 1e6 * facteur for d in tornado_d]
        fig_t  = go.Figure()
        fig_t.add_trace(go.Bar(y=labels, x=im, orientation="h", name="-10%",
                               marker_color=CLR["rouge"],
                               text=[f"{v:+,.0f}" for v in im], textposition="outside"))
        fig_t.add_trace(go.Bar(y=labels, x=ip, orientation="h", name="+10%",
                               marker_color=CLR["vert"],
                               text=[f"{v:+,.0f}" for v in ip], textposition="outside"))
        fig_t.add_vline(x=0, line_color="#333", line_width=1)
        fig_t.update_layout(
            title=f"Tornado — Impact sur VAN (M{devise})",
            xaxis_title=f"M{devise}", barmode="overlay",
            plot_bgcolor="white", paper_bgcolor="white",
            height=280, margin=dict(t=40, b=20),
            legend=dict(orientation="h", y=-0.3),
        )
        st.plotly_chart(fig_t, use_container_width=True)
    with c_tab:
        st.markdown(f"#### Impact VAN (M{devise})")
        st.dataframe(pd.DataFrame([{
            "Variable": d["Variable"],
            "-10%"    : f"{d['im']/1e6*facteur:+,.0f}",
            "+10%"    : f"{d['ip']/1e6*facteur:+,.0f}",
            "Écart"   : f"{d['amp']/1e6*facteur:,.0f}",
        } for d in tornado_d]), use_container_width=True, hide_index=True)
        st.caption("Variable en tête = **levier de négociation prioritaire** avant signature.")

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
        title=f"VAN — Prix × Coût (M{devise})",
        xaxis_title="Variation Coût", yaxis_title="Variation Prix",
        height=320, margin=dict(t=40),
    )
    st.plotly_chart(fig_h, use_container_width=True)

st.markdown("---")


# =============================================================================
# SECTION 7 — STRESS TEST & ASSURANCE
# =============================================================================

st.markdown(f"## Stress Test — {cfg['risque_titre']}")
st.info(
    f"**Risque :** {cfg['risque_desc']}\n\n"
    f"**Couverture :** {cfg['assurance_titre']} — {cfg['assurance_desc']}",
    icon="ℹ️",
)

def dscr_badge(val):
    if np.isnan(val): return "N/D"
    if val >= 1.3:    return f"✅ {val:.2f}x"
    if val >= 1.0:    return f"⚠️ {val:.2f}x"
    return f"❌ {val:.2f}x"

def treso_fin(sc_rows):
    return sc_rows[-1]["Tresorerie_cum"] / 1e6 * facteur if sc_rows else 0.0

df_comp = pd.DataFrame({
    "Critère": [
        "DSCR Minimum", "Trésorerie cumulée finale",
        "Perte CA (choc)", "Indemnité reçue",
        "Coût total primes", "Gain net assurance",
    ],
    "Base": [
        dscr_badge(stress["base"]["dmin"]),
        f"{treso_fin(stress['base']['rows']):,.1f} M{devise}",
        "—", "—", "—", "—",
    ],
    "Choc sans assurance": [
        dscr_badge(stress["choc"]["dmin"]),
        f"{treso_fin(stress['choc']['rows']):,.1f} M{devise}",
        f"{vf(stress['perte_brute']):.1f} M{devise}", "—", "—", "—",
    ],
    "Choc avec assurance": [
        dscr_badge(stress["assurance"]["dmin"]),
        f"{treso_fin(stress['assurance']['rows']):,.1f} M{devise}",
        f"{vf(stress['perte_brute']):.1f} M{devise}",
        f"{vf(stress['indemnite']):.1f} M{devise}",
        f"{vf(stress['cout_primes']):.1f} M{devise}",
        f"{'✅' if stress['gain_net']>0 else '⚠️'} {vf(stress['gain_net']):.1f} M{devise}",
    ],
})
st.dataframe(df_comp, use_container_width=True, hide_index=True)
st.plotly_chart(fig_stress(stress, duree, devise, facteur), use_container_width=True)

# Interprétation
st.markdown("#### Lecture rapide")
d_choc  = stress["choc"]["dmin"]
d_assur = stress["assurance"]["dmin"]
lignes  = []
if not np.isnan(d_choc):
    if d_choc < SEUIL_DEFAUT:
        lignes.append(f"Sans assurance, le choc à l'année {stress['an_choc']} fait chuter le DSCR à **{d_choc:.2f}x** — risque de défaut confirmé.")
    elif d_choc < SEUIL_BANCAIRE:
        lignes.append(f"Sans assurance, le DSCR tombe à **{d_choc:.2f}x** — sous le covenant bancaire de 1.3x. Un DSRA ou une garantie d'État serait nécessaire.")
    else:
        lignes.append(f"Le projet absorbe le choc sans assurance (DSCR = **{d_choc:.2f}x**). L'assurance reste recommandée pour satisfaire les exigences DFI.")
if not np.isnan(d_assur):
    gain = d_assur - (d_choc if not np.isnan(d_choc) else 0)
    txt  = f"Avec la **{cfg['assurance_titre']}**, le DSCR remonte à **{d_assur:.2f}x** ({gain:+.2f}x)."
    txt += " Bancabilité restaurée ✅" if d_assur >= SEUIL_BANCAIRE else " Toujours sous 1.3x — réserve complémentaire conseillée."
    lignes.append(txt)
if stress["gain_net"] > 0:
    lignes.append(f"Rentabilité de l'assurance sur {duree} ans : indemnité ({vf(stress['indemnite']):.1f} M{devise}) > primes ({vf(stress['cout_primes']):.1f} M{devise}).")
else:
    lignes.append(f"Coût primes ({vf(stress['cout_primes']):.1f} M{devise}) > indemnité ({vf(stress['indemnite']):.1f} M{devise}). Valeur de l'assurance = sécurité bancaire, pas gain financier attendu.")
for lg in lignes:
    st.markdown(f"- {lg}")

st.markdown("---")


# =============================================================================
# SECTION 8 — EXPORT PDF
# =============================================================================

st.markdown("## Export — Mémo Teaser Investisseur")

if not PDF_OK:
    st.warning("Installez `reportlab` pour activer l'export PDF : `pip install reportlab`")
else:
    if st.button("📄 Générer le Mémo PDF (1-2 pages)", type="secondary"):
        with st.spinner("Génération du mémo PDF…"):
            pdf_bytes = generer_pdf(
                projet_nom=projet_nom, matiere=matiere, bm_label=bm,
                devise=devise, facteur=facteur, rows=rows, met=met,
                stress=stress, capex_fcfa=capex_fcfa_v, p_dette=p_senior_pct/100,
                cfg=cfg, dette_s=dette_senior_fcfa, dette_m=dette_mezz_fcfa,
                dette_q=dette_qe_fcfa, wacc_val=wacc_val, llcr=llcr,
            )
        if pdf_bytes:
            st.download_button(
                label="⬇️ Télécharger le Mémo PDF",
                data=pdf_bytes,
                file_name=f"CommodityWatch_{projet_nom.replace(' ','_')}.pdf",
                mime="application/pdf",
            )
            st.success("Mémo généré avec succès.")

st.markdown("---")
st.markdown(
    f"<div style='font-size:.70rem;color:#6c757d;text-align:center'>"
    f"CommodityWatch v5.0 — 1 USD = {TAUX_CHANGE:.0f} FCFA — "
    "Document indicatif pré-investissement. "
    "Ne se substitue pas à une due diligence financière et juridique complète. "
    "Confidentiel — Usage interne."
    "</div>",
    unsafe_allow_html=True,
)
