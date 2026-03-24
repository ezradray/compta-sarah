import pandas as pd
import re
import zipfile
import calendar
from io import BytesIO
from datetime import date
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas as rl_canvas

w, h = A4
NAVY  = colors.HexColor('#1E3A5F')
LIGHT = colors.HexColor('#EEF2FF')
GOLD  = colors.HexColor('#C9A84C')
WHITE = colors.white
BLACK = colors.black

MOIS_NOMS = {
    1:'Janvier',2:'Février',3:'Mars',4:'Avril',5:'Mai',6:'Juin',
    7:'Juillet',8:'Août',9:'Septembre',10:'Octobre',11:'Novembre',12:'Décembre'
}
MOIS_COL = {1:7,2:8,3:9,4:10,5:11,6:12,7:13,8:14,9:15,10:16,11:17,12:18}

SOCIETES = {
    "SCI EMJJ": {
        "nom":"SCI EMJJ","adresse":"6, RUE DE VALMY - 93107 MONTREUIL CEDEX",
        "tel":"01.48.70.41.26","portable":"06.22.30.58.84","ville":"MONTREUIL",
        "sheet_kw":["EMJJ"],"sheet_exc":["MIRA"],"avis":True,
    },
    "SCI EMJJ MIRABEAU": {
        "nom":"SCI EMJJ","adresse":"6, RUE DE VALMY - 93107 MONTREUIL CEDEX",
        "tel":"01.48.70.41.26","portable":"06.22.30.58.84","ville":"MONTREUIL",
        "sheet_kw":["MIRA"],"sheet_exc":[],"avis":True,
    },
    "SCI MMA": {
        "nom":"SCI MMA","adresse":"29, RUE DU PROGRES - 93107 MONTREUIL CEDEX",
        "tel":"01.48.70.41.26","portable":"06.22.30.58.84","ville":"MONTREUIL",
        "sheet_kw":["MMA"],"sheet_exc":[],"avis":True,
    },
    "SCI AZM": {
        "nom":"SCI AZM","adresse":"6, RUE DE VALMY - 93100 MONTREUIL",
        "tel":"01.48.70.41.26","portable":"06.22.30.58.84","ville":"MONTREUIL",
        "sheet_kw":["AZM"],"sheet_exc":[],"avis":True,"trimestriel":True,
    },
    "SCI MAZ": {
        "nom":"SCI MAZ","adresse":"6, RUE DE VALMY - 93100 MONTREUIL",
        "tel":"01.48.70.41.26","portable":"06.22.30.58.84","ville":"MONTREUIL",
        "sheet_kw":["MAZ"],"sheet_exc":[],"avis":True,"trimestriel":True,
    },
    "SOGEPA": {
        "nom":"SOGEPA","adresse":"SAS au capital de 15.000 euros",
        "adresse2":"RCS BOBIGNY B 439 263 591",
        "adresse3":"29, RUE DU PROGRES - 93107 MONTREUIL CEDEX",
        "tel":"01.48.70.41.26","portable":"06.22.30.58.84","ville":"MONTREUIL",
        "sheet_kw":["SOGEPA"],"sheet_exc":[],"avis":True,"trimestriel":True,
    },
    "M.A LA GARENNE": {
        "nom":"M.A LA GARENNE","adresse":"","tel":"","portable":"","ville":"",
        "sheet_kw":["GARENNE"],"sheet_exc":[],"avis":False,
    },
}

def _safe(v):
    if v is None: return ''
    s=str(v).strip(); return '' if s.lower()=='nan' else s

def _num(v):
    try:
        f=float(v); return 0.0 if str(f)=='nan' else f
    except: return 0.0

def _fin_mois(mois,annee):
    return calendar.monthrange(annee,mois)[1]

def _fmt(v):
    return f"{abs(v):,.2f} \u20ac".replace(',',' ')

def _fmt_date(d):

    try:
        import re
        m=re.match(r'(\d{4})-(\d{2})-(\d{2})',str(d).strip())
        if m: return f"{m.group(3)}/{m.group(2)}/{m.group(1)}"
    except: pass
    return str(d).strip()

def lire_locataires(loyers_io, societe_key, mois, annee):
    cfg=SOCIETES.get(societe_key)
    if not cfg: return [],f"Societe inconnue : {societe_key}"
    wb=pd.ExcelFile(loyers_io)
    sheet=None
    for s in wb.sheet_names:
        su=s.upper()
        if all(k.upper() in su for k in cfg['sheet_kw']):
            if not any(e.upper() in su for e in cfg['sheet_exc']):
                sheet=s; break
    if not sheet: return [],f"Onglet introuvable pour {societe_key}"
    df=pd.read_excel(loyers_io,sheet_name=sheet,header=None)
    # Pour trimestriel : lire le 1er mois du trimestre
    _COL_TRIM = {6: MOIS_COL[4]-1, 9: MOIS_COL[7]-1, 12: MOIS_COL[10]-1}
    is_trim = cfg.get('trimestriel', False) and mois in _COL_TRIM
    col_m = _COL_TRIM[mois] if is_trim else MOIS_COL[mois]-1
    adresse_immeuble=''
    for i in range(min(10,len(df))):
        v=_safe(df.iloc[i,0])
        if re.search(r'RUE|AV\b|BD\b|IMPASSE|PLACE|CHEMIN',v,re.I):
            adresse_immeuble=v; break
    locataires=[]
    i=0
    while i<len(df):
        row=df.iloc[i]
        v0c = _safe(row.iloc[0]).replace('.','').replace('/','').replace(' ','').replace('-','')
        if (v0c.isdigit() or v0c[:1].isdigit()) and 'Locataire' in _safe(row.iloc[1]):
            b={
                'civilite':_safe(row.iloc[2]),'nom':_safe(row.iloc[3]),
                'adresse_bien':adresse_immeuble.split(' - ')[0].strip() if ' - ' in adresse_immeuble else adresse_immeuble,
                'cp_bien':adresse_immeuble.split(' - ')[1].strip() if ' - ' in adresse_immeuble else '',
                'adresse_loc':'','cp_ville_loc':'',
                'local':'','appartement':'','superficie':'',
                'loyer_hab':0,'prov_charges':0,'charges_supp':[],
                'solde_prec':0,'date_solde':'','caf':[],'reglements':[],'reglements_trim':[],'solde_prec_trim':0,'date_solde_trim':'',
                'net_a_payer':None,'a_paye':False,
            }
            if not b['nom']: i+=1; continue
            # col M-1 = mouvements du mois précédent
            col_prec = col_m - 1
            phase = 'infos'  # infos → solde1 → mouvements
            for k in range(i+1,min(i+50,len(df))):
                rk=df.iloc[k]
                cleft =_safe(rk.iloc[1])
                cvall =_safe(rk.iloc[2])
                cvall2=_safe(rk.iloc[3]) if len(rk)>3 else ''
                label =_safe(rk.iloc[3]) if len(rk)>3 else ''
                label5=_safe(rk.iloc[4]) if len(rk)>4 else ''
                vm    =rk.iloc[col_m]    if col_m    < len(rk) else None
                vm_n1 =rk.iloc[col_prec] if 0<=col_prec < len(rk) else None

                # BREAK : nouveau locataire ou Quittance envoyée
                v0b = _safe(rk.iloc[0]).replace('.','').replace('/','').replace(' ','').replace('-','')
                if (v0b.isdigit() or v0b[:1].isdigit()) and 'Locataire' in cleft:
                    break
                if 'Quittance' in label5:
                    break

                # LABEL DROITE — traité EN PRIORITÉ (peut coexister avec info gauche)
                # Détecter "Solde initial" en col B pour faire avancer la phase
                if 'Solde initial' in label or 'Solde initial' in cleft:
                    if phase=='infos': phase='solde1'

                if 'Solde pr' in label and 'initial' not in label.lower():
                    if phase=='infos':
                        phase='solde1'          # 1ère occurrence = solde initial → skip
                    elif phase=='solde1':
                        v_sp=_num(vm_n1)        # 2ème occurrence = solde réel M-1
                        if v_sp!=0:
                            b['solde_prec']=v_sp
                            b['date_solde']=f"01/{(mois-1 if mois>1 else 12):02d}/{annee if mois>1 else annee-1}"
                        phase='mouvements'

                elif phase in ('solde1','mouvements'):
                    if 'CAF' in label:
                        v=_num(vm_n1)
                        if v!=0:
                            date_caf=f"01/{mois-1 if mois>1 else 12:02d}/{annee if mois>1 else annee-1}"
                            if k+1<len(df):
                                d=_safe(df.iloc[k+1,col_prec])
                                if re.search(r'\d{4}',d): date_caf=_fmt_date(d)
                            b['caf'].append((date_caf,v))
                    elif 'glement locataire' in label:
                        v=_num(vm_n1)
                        if v!=0:
                            date_r=''; mode_r=''
                            if k+1<len(df):
                                d=_safe(df.iloc[k+1,col_prec])
                                if re.search(r'\d{4}',d): date_r=_fmt_date(d)
                            if k+2<len(df):
                                mode_r=_safe(df.iloc[k+2,col_prec])
                            if mode_r in('-','','nan'): mode_r='Virement'
                            b['reglements'].append((date_r,mode_r,v))
                    elif 'SOLDE LOCATAIRE' in label:
                        # Lire directement col S (index 18) — ne pas calculer
                        v18=_num(rk.iloc[18]) if len(rk)>18 else 0
                        b['net_a_payer']=v18

                elif phase=='infos' and label and 'TOTAL DU' not in label.upper() and 'SOLDE' not in label.upper():
                    # Toutes les lignes col D dans l'ordre naturel du tableau
                    v=_num(vm)
                    if v!=0:
                        # Éviter doublons — tout dans charges_supp pour respecter l'ordre Excel
                        if not any(l==label for l,_ in b['charges_supp']):
                            b['charges_supp'].append((label,v))
                        # Mettre à jour loyer_hab et prov_charges pour compatibilité quittance
                        if label in ('Loyer Habitation',"Loyer d'habitation"):
                            b['loyer_hab']=v
                        elif label in ('Prov. Charges','Provision Charges','Prov. pour charges'):
                            b['prov_charges']=v

                # INFO GAUCHE — uniquement infos fixes, PAS les montants
                if cleft=='Adresse':
                    b['adresse_loc']=cvall; b['cp_ville_loc']=cvall2
                elif cleft=='Local' and not b['local']:
                    b['local']=cvall
                elif cleft=='Appartement' and not b['appartement']:
                    b['appartement']=cvall
                elif cleft=='Superficie' and not b['superficie']:
                    b['superficie']=cvall

            # ── MOUVEMENTS TRIMESTRIELS ──────────────────────────
            if is_trim:
                _TRIM_MVTS = {
                    6:  {'solde_col': 6,  'reglt_cols': [6,7,8],   'date_avis': f'01/04/{annee}', 'date_solde': f'01/01/{annee}'},
                    9:  {'solde_col': 9,  'reglt_cols': [9,10,11],  'date_avis': f'01/07/{annee}', 'date_solde': f'01/04/{annee}'},
                    12: {'solde_col': 12, 'reglt_cols': [12,13,14], 'date_avis': f'01/10/{annee}', 'date_solde': f'01/07/{annee}'},
                }
                tm = _TRIM_MVTS[mois]
                b['date_avis_trim'] = tm['date_avis']
                # Chercher solde précédent et règlements dans le bloc
                for k2 in range(i+1, min(i+50, len(df))):
                    rk2   = df.iloc[k2]
                    lab2  = _safe(rk2.iloc[3])
                    if 'Quittance' in _safe(rk2.iloc[4] if len(rk2)>4 else ''): break
                    v0t = _safe(rk2.iloc[0]).replace('.','').replace('/','').replace(' ','').replace('-','')
                    if (v0t.isdigit() or v0t[:1].isdigit()) and 'Locataire' in _safe(rk2.iloc[1]): break
                    if 'Solde pr' in lab2:
                        try:
                            v = float(rk2.iloc[tm['solde_col']])
                            if v != 0:
                                b['solde_prec_trim'] = v
                                b['date_solde_trim'] = tm['date_solde']
                        except: pass
                    elif 'glement locataire' in lab2:
                        dates_row = df.iloc[k2+1] if k2+1 < len(df) else None
                        modes_row = df.iloc[k2+2] if k2+2 < len(df) else None
                        for col in tm['reglt_cols']:
                            try:
                                mt = float(rk2.iloc[col])
                                if mt != 0 and str(mt) != 'nan':
                                    try: dt = pd.to_datetime(dates_row.iloc[col]).strftime('%d/%m/%Y')
                                    except: dt = ''
                                    md = _safe(modes_row.iloc[col]) if modes_row is not None else 'Virement'
                                    if md in ('-','','nan'): md = 'Virement'
                                    b['reglements_trim'].append((dt, f"Règlement — {md}", mt))
                            except: pass
                        break

            if b['net_a_payer'] is None:
                total=b['loyer_hab']+b['prov_charges']+sum(v for _,v in b['charges_supp'])
                mvts=b['solde_prec']+sum(v for _,v in b['caf'])+sum(v for _,_,v in b['reglements'])
                b['net_a_payer']=total+mvts
            b['a_paye']=len(b['reglements'])>0 or any(v!=0 for _,v in b['caf'])
            locataires.append(b)
        i+=1
    return locataires,None

def _hline(c,y,x1=20,x2=190,lw=0.5):
    c.setStrokeColor(NAVY); c.setLineWidth(lw)
    c.line(x1*mm,y*mm,x2*mm,y*mm)

def _gold(c,y):
    c.setStrokeColor(GOLD); c.setLineWidth(1.5)
    c.line(20*mm,y*mm,190*mm,y*mm)

def _band(c,y,hh=6):
    c.setFillColor(NAVY); c.rect(20*mm,y*mm,170*mm,hh*mm,fill=1,stroke=0)

def generer_avis(bloc,cfg,mois,annee):
    fin=_fin_mois(mois,annee)
    today=date.today().strftime("%d/%m/%Y")
    debut_s=f"01/{mois:02d}/{annee}"; fin_s=f"{fin:02d}/{mois:02d}/{annee}"
    # Période trimestrielle
    _TRIM = {6:("2ème Trimestre","Avril","Juin"), 9:("3ème Trimestre","Juillet","Septembre"), 12:("4ème Trimestre","Octobre","Décembre")}
    is_trim = cfg.get('trimestriel', False) and mois in _TRIM
    if is_trim:
        t_label, t_debut, t_fin = _TRIM[mois]
        periode_str = f"{t_label} {annee}  —  {t_debut} à {t_fin} {annee}"
    else:
        periode_str = f"Période du {debut_s}  Au  {fin_s}"
    buf=BytesIO()
    c=rl_canvas.Canvas(buf,pagesize=A4)
    _band(c,268,hh=18)
    c.setFillColor(WHITE); c.setFont("Helvetica-Bold",14)
    c.drawCentredString(w/2,279*mm,"AVIS D'ÉCHÉANCE")
    c.setFont("Helvetica",9)
    c.drawCentredString(w/2,273*mm,periode_str)
    c.setFillColor(BLACK)
    c.setFont("Helvetica",9)
    c.drawRightString(188*mm,264*mm,f"{cfg['ville']}, le {today}")
    # Societe
    c.setFont("Helvetica-Bold",10); c.drawString(20*mm,258*mm,cfg['nom'])
    c.setFont("Helvetica",9)
    y_s=253
    for part in cfg['adresse'].split(' - '):
        c.drawString(20*mm,y_s*mm,part.strip()); y_s-=5
    if cfg.get('adresse2'): c.drawString(20*mm,y_s*mm,cfg['adresse2']); y_s-=5
    if cfg.get('adresse3'):
        for part in cfg['adresse3'].split(' - '):
            c.drawString(20*mm,y_s*mm,part.strip()); y_s-=5
    c.drawString(20*mm,y_s*mm,f"Tél : {cfg['tel']}"); y_s-=5
    c.drawString(20*mm,y_s*mm,f"Port. : {cfg['portable']}")
    y_gold = y_s - 5
    _gold(c, y_gold)
    # Locataire (ancré en haut, indépendant de la hauteur société)
    c.setFont("Helvetica",9); c.drawString(110*mm,258*mm,bloc['civilite'])
    c.setFont("Helvetica-Bold",9); c.drawString(110*mm,253*mm,bloc['nom'])
    c.setFont("Helvetica",9)
    c.drawString(110*mm,248*mm,bloc['adresse_loc'])
    c.drawString(110*mm,243*mm,bloc['cp_ville_loc'])

    # Bien — positionné dynamiquement sous le trait doré
    y_bien = y_gold - 8
    c.setFont("Helvetica-Bold",9); c.setFillColor(NAVY)
    c.drawString(20*mm, y_bien*mm, "Identification du bien\xa0:")
    c.setFont("Helvetica",9); c.setFillColor(BLACK)
    y_bien -= 5; c.drawString(20*mm, y_bien*mm, bloc['appartement'])
    y_bien -= 5; c.drawString(20*mm, y_bien*mm, bloc['local'])
    y_bien -= 5; c.drawString(20*mm, y_bien*mm, bloc['adresse_loc'])
    y_bien -= 5; c.drawString(20*mm, y_bien*mm, bloc['cp_ville_loc'])
    y_bien -= 8
    c.setFont("Helvetica-Oblique",8.5)
    c.drawString(20*mm, y_bien*mm, "Veuillez trouver ci-joint, votre avis d'échéance")
    y_bien -= 5
    msg_reglt = "Merci de nous adresser votre règlement le 1er du trimestre" if cfg.get("trimestriel") else "Merci de nous adresser votre règlement le 1er du mois"
    c.drawString(20*mm, y_bien*mm, msg_reglt)
    y_bien -= 5
    c.setFont("Helvetica-Bold",8.5); c.drawString(20*mm, y_bien*mm, "La comptabilité")
    y_bien -= 8
    # Tableau designation — positionné dynamiquement
    _band(c, y_bien); c.setFillColor(WHITE); c.setFont("Helvetica-Bold",9)
    c.drawString(22*mm,(y_bien+1.5)*mm,"DÉSIGNATION"); c.drawRightString(188*mm,(y_bien+1.5)*mm,"MONTANT")
    c.setFillColor(BLACK)
    # Ordre naturel Excel : tout dans charges_supp
    lignes=[(lbl,v) for lbl,v in bloc['charges_supp'] if v!=0]
    # Fallback mensuel si charges_supp vide
    if not lignes:
        if bloc['loyer_hab']!=0:    lignes.append(("Loyer Habitation",bloc['loyer_hab']))
        if bloc['prov_charges']!=0: lignes.append(("Prov. Charges",bloc['prov_charges']))
    y = y_bien - 5
    for i,(lbl,mt) in enumerate(lignes):
        if i%2==0:
            c.setFillColor(LIGHT); c.rect(20*mm,(y-4)*mm,170*mm,6.5*mm,fill=1,stroke=0)
        c.setFillColor(BLACK); c.setFont("Helvetica",9)
        c.drawString(22*mm,y*mm,lbl)
        # Montant négatif → afficher en rouge avec signe -
        if mt < 0:
            c.setFillColor(colors.HexColor('#cc0000'))
            c.drawRightString(188*mm,y*mm,"- " + _fmt(mt))
            c.setFillColor(BLACK)
        else:
            c.drawRightString(188*mm,y*mm,_fmt(mt))
        y-=7
    total_avis=sum(mt for _,mt in lignes)  # négatifs déduits du total
    c.setFillColor(LIGHT); c.rect(20*mm,(y-4)*mm,170*mm,6.5*mm,fill=1,stroke=0)
    c.setFillColor(NAVY); c.setFont("Helvetica-Bold",9)
    c.drawString(22*mm,y*mm,"TOTAL DE L'AVIS")
    c.drawRightString(188*mm,y*mm,_fmt(total_avis))
    c.setFillColor(BLACK); y-=12
    # Tableau mouvements
    _band(c,y); c.setFillColor(WHITE); c.setFont("Helvetica-Bold",8)
    c.drawString(22*mm,(y+1.5)*mm,"DATE")
    c.drawString(50*mm,(y+1.5)*mm,"LIBELLÉ")
    c.drawString(135*mm,(y+1.5)*mm,"DÉBIT")
    c.drawString(163*mm,(y+1.5)*mm,"CRÉDIT")
    c.setFillColor(BLACK)
    mouvs=[]
    if is_trim:
        # ── MODE TRIMESTRIEL ─────────────────────────────────
        # Solde précédent = col 1er mois trimestre N-1
        sp = bloc.get('solde_prec_trim', 0)
        if sp != 0:
            dt_sp = bloc.get('date_solde_trim', '')
            if sp > 0: mouvs.append((dt_sp, "Solde précédent", sp, None))
            else:      mouvs.append((dt_sp, "Solde précédent", None, sp))
        # Règlements des 3 mois du trimestre précédent
        for dt, lib, mt in bloc.get('reglements_trim', []):
            if mt < 0: mouvs.append((dt, lib, None, mt))
            else:      mouvs.append((dt, lib, mt, None))
        # Total avis → 1er mois du trimestre EN COURS
        date_avis = bloc.get('date_avis_trim', f'01/{mois:02d}/{annee}')
        mouvs.append((date_avis, "Total de l'avis ci-dessus", total_avis, None))
    else:
        # ── MODE MENSUEL ─────────────────────────────────────
        if bloc['solde_prec']!=0:
            dt_sp=bloc.get('date_solde',f"01/{mois:02d}/{annee}")
            if bloc['solde_prec']>0: mouvs.append((dt_sp,"Solde précédent",bloc['solde_prec'],None))
            else: mouvs.append((dt_sp,"Solde précédent",None,bloc['solde_prec']))
        for dt,v in bloc['caf']:
            if v>0: mouvs.append((dt,"Virement CAF",v,None))
            else:   mouvs.append((dt,"Virement CAF",None,v))
        for dt,mode,v in bloc['reglements']:
            lib=f"Règlement locataire {mode}".strip()
            if v>0: mouvs.append((dt,lib,v,None))
            else:   mouvs.append((dt,lib,None,v))
        mouvs.append((f"01/{mois:02d}/{annee}","Total de l'avis ci-dessus",total_avis,None))
    ym=y-7
    tot_deb =sum(m[2] for m in mouvs if m[2] is not None)
    tot_cred=sum(abs(m[3]) for m in mouvs if m[3] is not None)
    for i,(dt,lib,deb,cred) in enumerate(mouvs):
        if i%2==0:
            c.setFillColor(LIGHT); c.rect(20*mm,(ym-4)*mm,170*mm,6.5*mm,fill=1,stroke=0)
        c.setFillColor(BLACK); c.setFont("Helvetica",8.5)
        c.drawString(22*mm,ym*mm,dt)
        c.drawString(50*mm,ym*mm,lib[:55])
        if deb is not None:  c.drawRightString(155*mm,ym*mm,_fmt(deb))
        if cred is not None: c.drawRightString(188*mm,ym*mm,_fmt(cred))
        ym-=7
    _gold(c,ym+3)
    c.setFont("Helvetica-Bold",9)
    c.drawString(22*mm,(ym-3)*mm,"TOTAL")
    c.drawRightString(155*mm,(ym-3)*mm,_fmt(tot_deb))
    c.drawRightString(188*mm,(ym-3)*mm,_fmt(tot_cred))
    _band(c,ym-16,hh=9)
    c.setFillColor(WHITE); c.setFont("Helvetica-Bold",12)
    c.drawString(22*mm,(ym-11)*mm,"NET À PAYER")
    c.drawRightString(188*mm,(ym-11)*mm,_fmt(bloc['net_a_payer']))
    c.setFillColor(BLACK); c.setFont("Helvetica-Oblique",8)
    c.drawCentredString(w/2,(ym-22)*mm,"( Montants en Euros )")
    c.save(); buf.seek(0)
    return buf

TEXTE_LEGAL=(
    "Cette QUITTANCE annule tous les recus qui auraient pu etre donnes "
    "pour acompte verse sur le present terme meme si ces recus portent une date "
    "posterieure a la date ci-dessus, sans prejudice du courant ou d'instance "
    "judiciaire en cours et sous reserve du droit du proprietaire."
)
MENTIONS_BAS=[
    "LE PAIEMENT DE LA PRESENTE QUITTANCE N'EMPORTE PAS DE PRESOMPTION DE PAIEMENT DES TERMES ANTERIEURS.",
    "SI LE PAIEMENT EST EFFECTUE PAR CHEQUE, LA PRESENTE QUITTANCE NE PRENDRA SON CARACTERE LIBERATOIRE",
    "QU'APRES ENCAISSEMENT DU CHEQUE.",
    "EN CAS DE CONGE PRECEDEMMENT DONNE, CETTE QUITTANCE REPRESENTERAIT L'INDEMNITE D'OCCUPATION ET NE",
    "SAURAIT ETRE CONSIDEREE COMME UN TITRE DE LOCATION.",
]

def generer_quittance(bloc,cfg,mois,annee):
    fin=_fin_mois(mois,annee)
    today=date.today().strftime("%d/%m/%Y")
    debut_s=f"01/{mois:02d}/{annee}"; fin_s=f"{fin:02d}/{mois:02d}/{annee}"
    buf=BytesIO()
    c=rl_canvas.Canvas(buf,pagesize=A4)
    # Societe gauche
    c.setFont("Helvetica-Bold",10); c.setFillColor(BLACK)
    c.drawString(20*mm,272*mm,cfg['nom'])
    c.setFont("Helvetica",9)
    y_s=267
    for part in cfg['adresse'].split(' - '):
        c.drawString(20*mm,y_s*mm,part.strip()); y_s-=5
    if cfg.get('adresse2'): c.drawString(20*mm,y_s*mm,cfg['adresse2']); y_s-=5
    if cfg.get('adresse3'):
        for part in cfg['adresse3'].split(' - '):
            c.drawString(20*mm,y_s*mm,part.strip()); y_s-=5
    c.drawString(20*mm,y_s*mm,f"Tel : {cfg['tel']}"); y_s-=5
    c.drawString(20*mm,y_s*mm,f"Mob : {cfg['portable']}")
    # Titre droite
    c.setFont("Helvetica-Bold",18); c.setFillColor(NAVY)
    c.drawString(88*mm,272*mm,"QUITTANCE DE LOYER")
    c.setFont("Helvetica",9); c.setFillColor(BLACK)
    c.drawString(88*mm,265*mm,f"Période du  {debut_s}   Au  {fin_s}")
    # Locataire droite
    c.setFont("Helvetica",9); c.drawString(88*mm,255*mm,bloc['civilite'])
    c.setFont("Helvetica-Bold",10); c.drawString(88*mm,250*mm,bloc['nom'])
    c.setFont("Helvetica",9)
    c.drawString(88*mm,245*mm,bloc['adresse_loc'])
    c.drawString(88*mm,240*mm,bloc['cp_ville_loc'])
    _hline(c,230,lw=0.3)
    # Identification du bien — mêmes 3 lignes que l'avis
    c.setFont("Helvetica-Bold",9); c.setFillColor(NAVY)
    c.drawString(20*mm,225*mm,"Identification du bien :")
    c.setFont("Helvetica",9); c.setFillColor(BLACK)
    c.drawString(20*mm,220*mm,bloc['appartement'])
    c.drawString(20*mm,215*mm,bloc['local'])
    c.drawString(20*mm,210*mm,bloc['adresse_bien'])
    if bloc.get('cp_bien'): c.drawString(20*mm,205*mm,bloc['cp_bien'])
    c.setFont("Helvetica",9)
    c.drawString(30*mm,195*mm,bloc['civilite'])
    c.drawString(30*mm,190*mm,"Ci-dessous votre QUITTANCE de loyer.")
    # Tableau designation
    _hline(c,178,lw=0.3)
    _band(c,170,hh=7)
    c.setFillColor(WHITE); c.setFont("Helvetica-Bold",9)
    c.drawCentredString(80*mm,173.5*mm,"D  É  S  I  G  N  A  T  I  O  N")
    c.drawRightString(188*mm,173.5*mm,"Montants")
    c.setFillColor(BLACK)
    lignes=[]
    # Ordre naturel Excel : tout dans charges_supp, comme l'avis
    lignes=[(lbl,abs(v)) for lbl,v in bloc['charges_supp'] if v!=0]
    # Fallback si charges_supp vide
    if not lignes:
        if bloc['loyer_hab']!=0:    lignes.append(("Loyer d'habitation",bloc['loyer_hab']))
        if bloc['prov_charges']!=0: lignes.append(("Provision pour charges",bloc['prov_charges']))
    y=165
    for i,(lbl,mt) in enumerate(lignes):
        if i%2==0:
            c.setFillColor(LIGHT); c.rect(20*mm,(y-4)*mm,170*mm,6.5*mm,fill=1,stroke=0)
        c.setFillColor(BLACK); c.setFont("Helvetica",9)
        c.drawString(22*mm,y*mm,lbl)
        c.drawRightString(188*mm,y*mm,f"{abs(mt):,.2f}".replace(',',' '))
        y-=7
    y=min(y,125)
    total=sum(abs(mt) for _,mt in lignes)
    _hline(c,y+2,lw=0.3)
    c.setFont("Helvetica-Bold",10); c.setFillColor(NAVY)
    c.drawRightString(155*mm,(y-5)*mm,"TOTAL")
    c.setFillColor(BLACK); c.setFont("Helvetica-Bold",11)
    c.drawRightString(188*mm,(y-5)*mm,f"{total:,.2f}".replace(',',' '))
    _hline(c,y-9,lw=0.3)
    c.setFont("Helvetica-Bold",9)
    c.drawCentredString(w/2,(y-18)*mm,"( Montants en Euros )")
    # Encadre legal
    y_leg=y-25
    c.setStrokeColor(NAVY); c.setLineWidth(1)
    c.rect(88*mm,(y_leg-32)*mm,100*mm,32*mm,fill=0,stroke=1)
    c.setFont("Helvetica",7); c.setFillColor(BLACK)
    mots=TEXTE_LEGAL.split(); ligne_c=''; lignes_leg=[]
    for mot in mots:
        test=ligne_c+(' ' if ligne_c else '')+mot
        if c.stringWidth(test,"Helvetica",7)>93*mm:
            lignes_leg.append(ligne_c); ligne_c=mot
        else: ligne_c=test
    if ligne_c: lignes_leg.append(ligne_c)
    y_t=y_leg-5
    for ll in lignes_leg[:10]:
        c.drawString(90*mm,y_t*mm,ll); y_t-=4
    # Mentions bas
    c.setFont("Helvetica",6)
    y_bas=25
    for ligne in MENTIONS_BAS:
        c.drawString(20*mm,y_bas*mm,ligne); y_bas-=4
    c.save(); buf.seek(0)
    return buf

def run_ra_rq(loyers_io, societe_key, mois, annee=2026):
    cfg=SOCIETES.get(societe_key)
    if not cfg: return None,None,f"Societe inconnue : {societe_key}"
    if not cfg['avis']: return None,None,f"{societe_key} : pas d'avis d'echeance configure"
    locataires,err=lire_locataires(loyers_io,societe_key,mois,annee)
    if err: return None,None,err
    if not locataires: return None,None,"Aucun locataire trouve"
    mois_nm=MOIS_NOMS[mois]
    nom_soc=cfg['nom'].replace(' ','_').replace('/','_')
    # ZIP avis
    buf_avis=BytesIO()
    with zipfile.ZipFile(buf_avis,'w',zipfile.ZIP_DEFLATED) as zf:
        for b in locataires:
            if not b['nom']: continue
            pdf=generer_avis(b,cfg,mois,annee)
            fname=f"AVIS_{b['nom'].replace(' ','_').replace('/','_')}.pdf"
            zf.writestr(fname,pdf.read())
    buf_avis.seek(0)
    nom_zip_avis=f"{nom_soc}_{mois_nm.upper()}_{annee}_AVIS_ECHEANCE.zip"
    # ZIP quittances
    buf_quit=BytesIO()
    nb_quit=0
    with zipfile.ZipFile(buf_quit,'w',zipfile.ZIP_DEFLATED) as zf:
        for b in locataires:
            if not b['nom'] or not b['a_paye']: continue
            pdf=generer_quittance(b,cfg,mois,annee)
            fname=f"QUITTANCE_{b['nom'].replace(' ','_').replace('/','_')}.pdf"
            zf.writestr(fname,pdf.read())
            nb_quit+=1
    buf_quit.seek(0)
    nom_zip_quit=f"{nom_soc}_{mois_nm.upper()}_{annee}_QUITTANCES.zip"
    stats={
        'nb_locataires':len(locataires),
        'nb_avis':len(locataires),
        'nb_quittances':nb_quit,
        'nom_zip_avis':nom_zip_avis,
        'nom_zip_quit':nom_zip_quit,
    }
    return (buf_avis,nom_zip_avis),(buf_quit,nom_zip_quit),stats
