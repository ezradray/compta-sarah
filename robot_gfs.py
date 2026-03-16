"""
Robot GFS — Génération Facture Salle
Basé sur la structure 2025 (modèle définitif)
Col A=date soirée, B=N°facture, C=ignoré, D=nom client,
E=adresse, F=CP, G=ville, H..AN=règlements, AO=TTC, AP=HT, AQ=TVA, AR=note
"""

import zipfile, re
from io import BytesIO
from datetime import datetime
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor, black, white

DARK   = HexColor("#1a1a1a")
GREY   = HexColor("#6c757d")
LGREY  = HexColor("#adb5bd")
XLIGHT = HexColor("#f4f4f4")
BLACK  = black

MOIS_NOMS = {
    1:"JANVIER",2:"FEVRIER",3:"MARS",4:"AVRIL",5:"MAI",6:"JUIN",
    7:"JUILLET",8:"AOUT",9:"SEPTEMBRE",10:"OCTOBRE",11:"NOVEMBRE",12:"DECEMBRE"
}

def fmt(v):
    return f"{v:,.2f}".replace(",","X").replace(".",",").replace("X"," ") + " €"

def thin_line(c, y, x1=10, x2=200):
    c.setStrokeColor(BLACK); c.setLineWidth(0.5)
    c.line(x1*mm, y*mm, x2*mm, y*mm)

# ── LECTURE EXCEL ─────────────────────────────────────────────
def lire_factures(fichier_io, annee):
    year_str = str(annee)
    try:
        df = pd.read_excel(fichier_io, sheet_name=year_str, header=None)
    except Exception as e:
        return None, f"Onglet '{year_str}' introuvable : {e}"

    # Trouver ligne header (contient "Date pièce" ou "Code Client")
    data_start = 22  # défaut 2025
    for i in range(min(25, len(df))):
        row = df.iloc[i]
        for v in row:
            if str(v).strip() in ('Date pièce','Date piece','FA25000001','FA24000002'):
                data_start = i
                break
        else:
            continue
        break

    # Trouver colonnes TTC/HT/TVA depuis la ligne header (ligne data_start - 1)
    col_ttc = col_ht = col_tva = col_note = None
    if data_start > 0:
        h = df.iloc[data_start - 1]
        for ci, v in enumerate(h):
            vs = str(v).strip().upper()
            if 'TTC' in vs:              col_ttc  = ci
            if vs in ('MONTANT H.T','H.T','HT','MONTANT HT'): col_ht = ci
            if 'TVA' in vs:              col_tva  = ci
            if 'COMMENT' in vs or 'NOTE' in vs: col_note = ci

    # Fallback colonnes 2025 connues
    if col_ttc  is None: col_ttc  = 39
    if col_ht   is None: col_ht   = 40
    if col_tva  is None: col_tva  = 41
    if col_note is None: col_note = 42

    factures = []
    for idx in range(data_start, len(df)):
        row = df.iloc[idx]

        # Col A = date soirée
        v0 = row.iloc[0]
        date_soiree = None
        try:
            date_soiree = pd.to_datetime(v0)
            if pd.isnull(date_soiree): date_soiree = None
        except:
            # Cas "16-17/04/25" ou "18-19/04/2025" → prendre première date
            if isinstance(v0, str):
                m = re.search(r'(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})', v0)
                if m:
                    try:
                        day = m.group(1).split('-')[-1] if '-' in m.group(1) else m.group(1)
                        yr  = m.group(3) if len(m.group(3))==4 else f"20{m.group(3)}"
                        date_soiree = pd.to_datetime(f"{day}/{m.group(2)}/{yr}", dayfirst=True)
                    except: pass

        if date_soiree is None or pd.isnull(date_soiree):
            continue

        # Col B = N° facture
        num_fa = str(row.iloc[1]).strip()
        if not num_fa.startswith('FA'):
            continue

        # Col C = ignoré
        # Col D = nom client
        client_nom = str(row.iloc[3]).strip() if len(row) > 3 else ''
        if client_nom in ('nan',''): client_nom = ''

        # Col E = adresse, F = CP, G = ville
        adr  = str(row.iloc[4]).strip() if len(row) > 4 else ''
        cp   = str(row.iloc[5]).strip() if len(row) > 5 else ''
        ville= str(row.iloc[6]).strip() if len(row) > 6 else ''
        if adr   in ('nan',''): adr   = ''
        if cp    in ('nan',''): cp    = ''
        if ville in ('nan',''): ville = ''

        # Construire adresse sur 2 lignes
        client_adr1 = adr
        client_adr2 = f"{cp} {ville}".strip() if cp or ville else ''

        # TTC / HT / TVA
        try: ttc = float(row.iloc[col_ttc])
        except: ttc = 0.0
        try: ht = float(row.iloc[col_ht])
        except: ht = 0.0
        try: tva = float(row.iloc[col_tva])
        except: tva = 0.0
        if ttc > 0 and ht == 0:
            ht  = round(ttc / 1.2, 2)
            tva = round(ttc - ht, 2)

        # Règlements : col 7..col_ttc-1, groupes de (montant, date, mode, banque) par 4
        reglements = []
        ci = 7
        while ci + 1 < col_ttc:
            try:
                mt = row.iloc[ci]
                if str(mt) in ('nan','') or pd.isnull(mt):
                    ci += 4; continue
                mt_f = float(mt)
                if mt_f <= 0:
                    ci += 4; continue
                # date règlement
                try:
                    dt_fmt = pd.to_datetime(row.iloc[ci+1]).strftime('%d/%m/%Y')
                except:
                    dt_fmt = ''
                # mode et banque
                mode  = str(row.iloc[ci+2]).strip() if ci+2 < col_ttc else ''
                banque= str(row.iloc[ci+3]).strip() if ci+3 < col_ttc else ''
                if mode   in ('nan','-',''): mode   = 'VIRT'
                if banque in ('nan','-',''): banque = ''
                # Libellé affiché
                if banque and banque != mode:
                    libelle = f"{banque} N° {mode}" if mode not in ('VIRT','CHQ','LCL','BNP','SG','S.G','C.E','CDC','B.POST','B. POST','B. POP') else f"{mode} {banque}".strip()
                else:
                    libelle = mode
                reglements.append((libelle, '', mt_f, dt_fmt))
            except:
                pass
            ci += 4

        # Note
        note = ''
        try:
            note = str(row.iloc[col_note]).strip()
            if note.lower() == 'nan': note = ''
        except: pass

        factures.append({
            'num_facture'  : num_fa,
            'date_facture' : date_soiree.strftime('%d/%m/%Y'),
            'date_soiree'  : date_soiree.strftime('%d/%m/%Y'),
            'date_sort'    : date_soiree,
            'client_nom'   : client_nom,
            'client_adr1'  : client_adr1,
            'client_adr2'  : client_adr2,
            'montant_ttc'  : round(ttc, 2),
            'montant_ht'   : round(ht, 2),
            'tva'          : round(tva, 2),
            'reglements'   : reglements,
            'note'         : note,
        })

    if not factures:
        return None, f"Aucune facture trouvée pour {annee}"

    # Tri chronologique strict
    factures.sort(key=lambda x: x['date_sort'])
    return factures, None


# ── GÉNÉRATION PDF ────────────────────────────────────────────
def generer_pdf(d):
    buf = BytesIO()
    c   = canvas.Canvas(buf, pagesize=A4)
    W, H = A4

    # BANDE TITRE
    c.setFillColor(DARK)
    c.rect(0, H-30*mm, W, 30*mm, fill=1, stroke=0)
    c.setFillColor(white); c.setFont("Helvetica-Bold", 20)
    c.drawCentredString(W/2, H-18*mm, "FACTURE")
    c.setFont("Helvetica", 9)
    c.drawCentredString(W/2, H-26*mm,
        f"Éditée en Euro  —  N° {d['num_facture']}  —  Date : {d['date_facture']}")

    # SOGEPA
    y0 = H - 40*mm
    c.setFillColor(DARK); c.setFont("Helvetica-Bold", 11)
    c.drawString(12*mm, y0, "SOGEPA")
    c.setFont("Helvetica", 8.5); c.setFillColor(BLACK)
    for ligne in ["SAS au capital de 15 000 €","29, RUE DU PROGRES",
                  "93107 MONTREUIL CEDEX","Tél : 01.48.70.41.26",
                  "SIRET : 43926359100018"]:
        y0 -= 5*mm
        c.drawString(12*mm, y0, ligne)

    # CLIENT
    c.setFillColor(XLIGHT)
    c.roundRect(108*mm, H-78*mm, 90*mm, 40*mm, 3*mm, fill=1, stroke=0)
    c.setFillColor(DARK); c.setFont("Helvetica-Bold", 8)
    c.drawString(112*mm, H-38*mm, "CLIENT :")
    c.setFont("Helvetica-Bold", 10); c.setFillColor(BLACK)
    c.drawString(112*mm, H-46*mm, d["client_nom"])
    c.setFont("Helvetica", 9)
    if d.get("client_adr1"):
        c.drawString(112*mm, H-53*mm, d["client_adr1"])
    if d.get("client_adr2"):
        c.drawString(112*mm, H-60*mm, d["client_adr2"])

    # INFOS PIÈCE
    thin_line(c, (H/mm)-84)
    x = 12
    for label, val in [("N° Pièce", d["num_facture"]), ("Date pièce", d["date_facture"]),
                        ("Montant TTC", fmt(d["montant_ttc"])), ("Échéance", d["date_facture"])]:
        c.setFillColor(GREY); c.setFont("Helvetica", 8)
        c.drawString(x*mm, (H/mm-89)*mm, label)
        c.setFillColor(DARK); c.setFont("Helvetica-Bold", 9)
        c.drawString(x*mm, (H/mm-95)*mm, val)
        x += 47
    thin_line(c, (H/mm)-99)

    # RÉFÉRENCE
    c.setFillColor(DARK); c.setFont("Helvetica-Bold", 9)
    c.drawString(12*mm, (H/mm-105)*mm, f"Réf. commande : SOIREE DU {d['date_soiree']}")
    thin_line(c, (H/mm)-109)

    # TABLEAU DÉSIGNATION
    yh = (H/mm-120)*mm
    c.setFillColor(DARK); c.rect(10*mm, yh, 190*mm, 8*mm, fill=1, stroke=0)
    c.setFillColor(white); c.setFont("Helvetica-Bold", 8.5)
    for txt, xp in [("CODE",12),("DÉSIGNATION",38),("QTÉ",130),("PU HT",150),("MONTANT HT",172)]:
        c.drawString(xp*mm, yh+2.5*mm, txt)

    yrow = yh - 9*mm
    c.setFillColor(XLIGHT); c.rect(10*mm, yrow, 190*mm, 9*mm, fill=1, stroke=0)
    c.setFillColor(BLACK); c.setFont("Helvetica", 9)
    c.drawString(12*mm,  yrow+2.8*mm, "LSER")
    c.drawString(38*mm,  yrow+2.8*mm, "LOCATION SALLE ESPACE ROYAL")
    c.drawString(131*mm, yrow+2.8*mm, "1,000")
    c.drawRightString(160*mm, yrow+2.8*mm, fmt(d["montant_ht"]))
    c.drawRightString(197*mm, yrow+2.8*mm, fmt(d["montant_ht"]))
    c.setFont("Helvetica-Oblique", 8.5); c.setFillColor(GREY)
    c.drawString(38*mm, yrow-5*mm,  f"Soirée du {d['date_soiree']}")
    c.drawString(38*mm, yrow-11*mm, "Location salle :")
    c.setFont("Helvetica", 8.5); c.setFillColor(BLACK)
    c.drawString(38*mm, yrow-17*mm, "6, RUE DE VALMY — 93107 MONTREUIL CEDEX")

    # RÈGLEMENTS
    yr = yrow - 35*mm
    c.setFillColor(DARK); c.rect(10*mm, yr, 190*mm, 7*mm, fill=1, stroke=0)
    c.setFillColor(white); c.setFont("Helvetica-Bold", 8.5)
    c.drawString(12*mm,  yr+2*mm, "RÈGLEMENTS REÇUS")
    c.drawString(110*mm, yr+2*mm, "MONTANT")
    c.drawString(155*mm, yr+2*mm, "DATE")
    alt = True
    for libelle, _, montant, date in d["reglements"]:
        yr -= 8*mm
        if alt:
            c.setFillColor(XLIGHT); c.rect(10*mm, yr, 190*mm, 8*mm, fill=1, stroke=0)
        alt = not alt
        c.setFillColor(BLACK); c.setFont("Helvetica", 9)
        c.drawString(12*mm,  yr+2.5*mm, libelle)
        c.drawRightString(125*mm, yr+2.5*mm, fmt(montant))
        c.drawString(155*mm, yr+2.5*mm, date)

    thin_line(c, (yr-4)/mm)

    # TOTAUX
    yt = yr - 10*mm
    for label, val in [("Montant HT", d["montant_ht"]),
                        ("TVA 20 %",   d["tva"]),
                        ("Montant TTC", d["montant_ttc"])]:
        c.setFont("Helvetica", 9); c.setFillColor(GREY)
        c.drawRightString(155*mm, yt, label)
        c.setFont("Helvetica-Bold", 9); c.setFillColor(BLACK)
        c.drawRightString(197*mm, yt, fmt(val))
        yt -= 7*mm

    yt -= 3*mm
    c.setFillColor(DARK)
    c.roundRect(120*mm, yt-4*mm, 78*mm, 12*mm, 2*mm, fill=1, stroke=0)
    c.setFillColor(white); c.setFont("Helvetica-Bold", 10)
    c.drawString(124*mm, yt+1*mm, "NET À PAYER")
    c.drawRightString(196*mm, yt+1*mm, fmt(d["montant_ttc"]))

    if d.get("note"):
        c.setFont("Helvetica-Oblique", 8); c.setFillColor(GREY)
        c.drawString(12*mm, yt+1*mm, f"Note : {d['note']}")

    # MENTIONS LÉGALES
    thin_line(c, 22)
    c.setFont("Helvetica", 7); c.setFillColor(GREY)
    c.drawCentredString(W/2, 17*mm,
        "En cas de retard de paiement, une pénalité de 3 fois le taux d'intérêt légal sera appliquée, "
        "ainsi qu'une indemnité forfaitaire de 40 € pour frais de recouvrement.")
    c.drawCentredString(W/2, 13*mm,
        "SOGEPA — SAS au capital de 15 000 € — SIRET 43926359100018 — TVA intracommunautaire : FR 12 439 263 591")

    c.save(); buf.seek(0)
    return buf


# ── FONCTION PRINCIPALE ───────────────────────────────────────
def run_gfs(fichier_io, annee):
    factures, err = lire_factures(fichier_io, annee)
    if err:
        return None, None, err

    buf_zip = BytesIO()
    with zipfile.ZipFile(buf_zip, 'w', zipfile.ZIP_DEFLATED) as zf:
        for f in factures:
            pdf = generer_pdf(f)
            nom_client = f['client_nom'].replace(' ','_').replace('/','_')[:30]
            date_fmt   = f['date_soiree'].replace('/','_')
            fname      = f"{f['num_facture']}_{nom_client}_{date_fmt}.pdf"
            zf.writestr(fname, pdf.read())

    buf_zip.seek(0)
    nom_zip = f"SOGEPA_FACTURES_SALLE_{annee}.zip"
    stats = {
        'nb'      : len(factures),
        'annee'   : annee,
        'nom_zip' : nom_zip,
        'premiere': factures[0]['num_facture'],
        'derniere': factures[-1]['num_facture'],
    }
    return buf_zip, nom_zip, stats
