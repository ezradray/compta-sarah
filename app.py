import streamlit as st
import pandas as pd
from robot_ac_gl import run, MOIS_NOMS
from robot_ra_rq import run_ra_rq
from robot_gfs import run_gfs

st.set_page_config(page_title="Compta Sarah", page_icon="💼", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@600;700&family=Inter:wght@300;400;500;600&display=swap');
  html,body,[class*="css"]{font-family:'Inter',sans-serif;}
  .stApp{background:#F7F8FC;}
  [data-testid="stSidebar"]{background:linear-gradient(180deg,#0D1B2A 0%,#1a2f4a 100%);border-right:1px solid #ffffff10;}
  [data-testid="stSidebar"] *{color:#e8edf5 !important;}
  .logo-block{text-align:center;padding:28px 0 20px 0;border-bottom:1px solid #ffffff15;margin-bottom:24px;}
  .logo-title{font-family:'Playfair Display',serif;font-size:26px;font-weight:700;color:#ffffff !important;}
  .logo-sub{font-size:11px;color:#5b7fa6 !important;letter-spacing:.15em;text-transform:uppercase;margin-top:4px;}
  .stat-card{background:white;border-radius:14px;padding:20px 24px;box-shadow:0 2px 12px rgba(0,0,0,.06);border-left:4px solid;margin-bottom:8px;}
  .stat-ok{border-color:#22c55e;}.stat-alert{border-color:#f59e0b;}.stat-skip{border-color:#94a3b8;}.stat-total{border-color:#3b82f6;}
  .stat-num{font-size:32px;font-weight:700;line-height:1;}.stat-lbl{font-size:12px;color:#6b7280;margin-top:4px;text-transform:uppercase;letter-spacing:.05em;}
  .section-title{font-family:'Playfair Display',serif;font-size:20px;color:#0D1B2A;margin:28px 0 16px 0;padding-bottom:8px;border-bottom:2px solid #e5e7eb;}
  .result-table{width:100%;border-collapse:collapse;font-size:13px;}
  .result-table th{background:#0D1B2A;color:white;padding:10px 12px;text-align:left;font-weight:500;font-size:12px;}
  .result-table td{padding:9px 12px;border-bottom:1px solid #f1f5f9;}
  .result-table tr:hover td{background:#f8faff;}
  .badge-ok{background:#dcfce7;color:#166534;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:600;}
  .badge-alerte{background:#fef3c7;color:#92400e;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:600;}
  .badge-ignore{background:#f1f5f9;color:#475569;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:600;}
  .montant-pos{color:#16a34a;font-weight:600;}.montant-neg{color:#dc2626;font-weight:600;}
  .alert-box{background:#fffbeb;border:1px solid #fcd34d;border-radius:12px;padding:16px 20px;margin:16px 0;}
  .alert-box h4{color:#92400e;margin:0 0 10px 0;font-size:14px;}
  .alert-item{font-size:13px;color:#78350f;padding:4px 0;border-bottom:1px solid #fde68a;}
  .dl-card{background:white;border-radius:14px;padding:20px;box-shadow:0 2px 10px rgba(0,0,0,.06);text-align:center;margin-bottom:8px;}
</style>
""", unsafe_allow_html=True)

SOCIETES = ["SCI EMJJ","SCI EMJJ MIRABEAU","SCI MMA","SCI AZM","SCI MAZ","SOGEPA","M.A LA GARENNE"]

if "robot" not in st.session_state:
    st.session_state.robot = "AC_GL"

# ── SIDEBAR ──
with st.sidebar:
    st.markdown('<div class="logo-block"><div class="logo-title">💼 Compta Sarah</div><div class="logo-sub">La belle vie, occupe toi de ton mari</div></div>', unsafe_allow_html=True)

    st.markdown("<span style='font-size:11px;text-transform:uppercase;letter-spacing:.1em;color:#5b7fa6'>🤖 Robots</span>", unsafe_allow_html=True)
    c1,c2,c3 = st.columns(3)
    with c1:
        if st.button("🏦 Rapprochement & Attribution", use_container_width=True, type="primary" if st.session_state.robot=="AC_GL" else "secondary"):
            st.session_state.robot="AC_GL"; st.rerun()
    with c2:
        if st.button("📄 Avis & Quittances", use_container_width=True, type="primary" if st.session_state.robot=="RA_RQ" else "secondary"):
            st.session_state.robot="RA_RQ"; st.rerun()
    with c3:
        if st.button("🧾 Factures Salle", use_container_width=True, type="primary" if st.session_state.robot=="GFS" else "secondary"):
            st.session_state.robot="GFS"; st.rerun()

    st.markdown("---")
    st.markdown("<span style='font-size:11px;text-transform:uppercase;letter-spacing:.1em;color:#5b7fa6'>🏢 Société</span>", unsafe_allow_html=True)
    societe_sel = st.selectbox("Société", SOCIETES, key="societe")
    st.markdown("<span style='font-size:11px;text-transform:uppercase;letter-spacing:.1em;color:#5b7fa6'>📅 Période</span>", unsafe_allow_html=True)
    mois_options = list(MOIS_NOMS.values())
    mois_sel  = st.selectbox("Mois", mois_options, index=2)
    mois_num  = list(MOIS_NOMS.keys())[list(MOIS_NOMS.values()).index(mois_sel)]
    annee_sel = st.number_input("Année", min_value=2020, max_value=2035, value=2026, step=1)
    st.markdown("---")

    if st.session_state.robot == "AC_GL":
        st.markdown("<span style='font-size:11px;text-transform:uppercase;letter-spacing:.1em;color:#5b7fa6'>📂 Fichiers</span>", unsafe_allow_html=True)
        releve_file = st.file_uploader("Relevé bancaire", type=["xlsx"], key="releve")
        pivot_file  = st.file_uploader("Tableau Pivot",   type=["xlsx"], key="pivot")
        loyers_file = st.file_uploader("Tableau Loyers",  type=["xlsx"], key="loyers")
        lancer = st.button("▶  Lancer AC+GL", use_container_width=True)
    else:
        st.markdown("<span style='font-size:11px;text-transform:uppercase;letter-spacing:.1em;color:#5b7fa6'>📂 Fichier</span>", unsafe_allow_html=True)
        loyers_file_rq = st.file_uploader("Tableau Loyers", type=["xlsx"], key="loyers_rq")
        lancer = st.button("▶  Générer Avis & Quittances", use_container_width=True)

# ── SIDEBAR GFS ──
if st.session_state.robot == "GFS":
    st.markdown("<span style='font-size:11px;text-transform:uppercase;letter-spacing:.1em;color:#5b7fa6'>📂 Fichier</span>", unsafe_allow_html=True)
    recap_file = st.file_uploader("Récap soirées à facturer", type=["xlsx"], key="recap_gfs")
    annee_gfs  = st.selectbox("Année à traiter", [2022,2023,2024,2025,2026], index=2, key="annee_gfs")
    lancer = st.button("▶  Générer les Factures", use_container_width=True)

    st.markdown("<div style='padding-top:30px;text-align:center'><small style='color:#2d4a6a;font-size:10px'>Actifs : AC+GL · RA · RQ<br>À venir : RAA · VS · GFS · LF · LT</small></div>", unsafe_allow_html=True)

# ══ ROBOT AC+GL ══
if st.session_state.robot == "AC_GL":
    st.markdown("<h1 style='font-family:Playfair Display,serif;font-size:32px;color:#0D1B2A'>🏦 Rapprochement Bancaire + Attribution de Compte</h1><hr>", unsafe_allow_html=True)
    if not lancer:
        st.markdown("""<div style='text-align:center;padding:60px 20px;background:white;border-radius:20px;box-shadow:0 2px 12px rgba(0,0,0,.05);margin-top:20px'>
          <div style='font-size:56px'>🤖</div>
          <h3 style='font-family:Playfair Display,serif;color:#0D1B2A'>Prêt à traiter vos données</h3>
          <p style='color:#9ca3af;max-width:400px;margin:0 auto;font-size:14px'>Chargez vos 3 fichiers, sélectionnez le mois et cliquez sur <b>Lancer AC+GL</b>.</p>
        </div>""", unsafe_allow_html=True)
    else:
        if not releve_file or not pivot_file or not loyers_file:
            st.error("⚠️ Veuillez charger les 3 fichiers avant de lancer le robot."); st.stop()
        with st.spinner(f"🤖 Traitement — {mois_sel} {annee_sel}..."):
            df_res, buffers, maj, err = run(releve_file, pivot_file, loyers_file, mois=mois_num, annee=int(annee_sel))
        if err: st.error(f"❌ {err}"); st.stop()
        buf1,buf2,buf3 = buffers
        nb_ok=len(df_res[df_res.STATUT=='OK']); nb_al=len(df_res[df_res.STATUT=='ALERTE']); nb_ig=len(df_res[df_res.STATUT=='IGNORE'])
        c1,c2,c3,c4=st.columns(4)
        with c1: st.markdown(f'<div class="stat-card stat-total"><div class="stat-num" style="color:#3b82f6">{len(df_res)}</div><div class="stat-lbl">Transactions</div></div>',unsafe_allow_html=True)
        with c2: st.markdown(f'<div class="stat-card stat-ok"><div class="stat-num" style="color:#22c55e">{nb_ok}</div><div class="stat-lbl">Identifiées ✓</div></div>',unsafe_allow_html=True)
        with c3: st.markdown(f'<div class="stat-card stat-alert"><div class="stat-num" style="color:#f59e0b">{nb_al}</div><div class="stat-lbl">Alertes ⚠</div></div>',unsafe_allow_html=True)
        with c4: st.markdown(f'<div class="stat-card stat-skip"><div class="stat-num" style="color:#94a3b8">{nb_ig}</div><div class="stat-lbl">Ignorées ⏸</div></div>',unsafe_allow_html=True)
        if nb_al>0:
            items="".join([f"<div class='alert-item'>→ {row['DATE'].strftime('%d/%m/%Y') if pd.notna(row['DATE']) else '?'} | {str(row['LIBELLE'])[:60]} | <b>{row['MONTANT']:.2f}€</b></div>" for _,row in df_res[df_res.STATUT=='ALERTE'].iterrows()])
            st.markdown(f'<div class="alert-box"><h4>⚠️ {nb_al} ligne(s) à saisir manuellement</h4>{items}</div>',unsafe_allow_html=True)
        st.markdown('<div class="section-title">📋 Détail des transactions</div>',unsafe_allow_html=True)
        rows_html=""
        for _,r in df_res.iterrows():
            d=r['DATE'].strftime('%d/%m/%y') if pd.notna(r['DATE']) else '?'
            sgn='+' if r['SENS']=='CREDIT' else '-'
            cls_m='montant-pos' if r['SENS']=='CREDIT' else 'montant-neg'
            badge=f"<span class='badge-{r['STATUT'].lower()}'>{r['STATUT']}</span>"
            rows_html+=f"<tr><td>{d}</td><td style='max-width:260px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap'>{str(r['LIBELLE'])[:60]}</td><td class='{cls_m}'>{sgn}{abs(r['MONTANT']):.2f}€</td><td><code style='font-size:11px;background:#f1f5f9;padding:2px 6px;border-radius:4px'>{r['TYPE']}</code></td><td><code style='font-size:11px;background:#f1f5f9;padding:2px 6px;border-radius:4px'>{r['COMPTE']}</code></td><td style='font-size:12px'>{str(r['NOM'])[:38]}</td><td>{badge}</td></tr>"
        st.markdown(f"<div style='background:white;border-radius:16px;padding:20px;box-shadow:0 2px 12px rgba(0,0,0,.06);overflow-x:auto'><table class='result-table'><thead><tr><th>Date</th><th>Libellé</th><th>Montant</th><th>Type</th><th>Compte</th><th>Identifié comme</th><th>Statut</th></tr></thead><tbody>{rows_html}</tbody></table></div>",unsafe_allow_html=True)
        if maj:
            st.markdown('<div class="section-title">🏠 Loyers reportés</div>',unsafe_allow_html=True)
            cols=st.columns(min(len(maj),4))
            for i,m in enumerate(maj):
                with cols[i%4]: st.markdown(f"<div style='background:white;border-radius:12px;padding:14px 16px;box-shadow:0 2px 8px rgba(0,0,0,.05);border-left:3px solid #22c55e;margin-bottom:8px'><div style='font-size:12px;color:#6b7280'>✅ Mis à jour</div><div style='font-size:13px;font-weight:600;color:#0D1B2A'>{m}</div></div>",unsafe_allow_html=True)
        st.markdown('<div class="section-title">📥 Télécharger</div>',unsafe_allow_html=True)
        dl1,dl2,dl3=st.columns(3)
        with dl1:
            st.markdown("<div class='dl-card'><div style='font-size:32px'>📋</div><div style='font-weight:600;color:#0D1B2A;margin:8px 0 4px'>Intégration Logiciel</div><div style='font-size:12px;color:#9ca3af;margin-bottom:12px'>Écritures comptables</div></div>",unsafe_allow_html=True)
            st.download_button("⬇ Télécharger",data=buf1.getvalue(),file_name=f"INTEGRATION_LOGICIEL_{mois_sel.upper()}_{annee_sel}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",key="dl1")
        with dl2:
            st.markdown("<div class='dl-card'><div style='font-size:32px'>🏠</div><div style='font-weight:600;color:#0D1B2A;margin:8px 0 4px'>Tableau Loyers MAJ</div><div style='font-size:12px;color:#9ca3af;margin-bottom:12px'>Loyers avec dates et modes</div></div>",unsafe_allow_html=True)
            st.download_button("⬇ Télécharger",data=buf2.getvalue(),file_name=f"TABLEAU_LOYERS_{mois_sel.upper()}_{annee_sel}_MAJ.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",key="dl2")
        with dl3:
            st.markdown(f"<div class='dl-card'><div style='font-size:32px'>{'⚠️' if nb_al>0 else '✅'}</div><div style='font-weight:600;color:#0D1B2A;margin:8px 0 4px'>Rapport Alertes</div><div style='font-size:12px;color:#9ca3af;margin-bottom:12px'>{nb_al} ligne(s) manuelles</div></div>",unsafe_allow_html=True)
            st.download_button("⬇ Télécharger",data=buf3.getvalue(),file_name=f"ALERTES_{mois_sel.upper()}_{annee_sel}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",key="dl3")

# ══ ROBOT RA+RQ ══
elif st.session_state.robot == "RA_RQ":
    st.markdown(f"<h1 style='font-family:Playfair Display,serif;font-size:32px;color:#0D1B2A'>📄 Avis d'Échéance & Quittances — {societe_sel} · {mois_sel} {int(annee_sel)}</h1><hr>", unsafe_allow_html=True)
    if not lancer:
        st.markdown("""<div style='text-align:center;padding:60px 20px;background:white;border-radius:20px;box-shadow:0 2px 12px rgba(0,0,0,.05);margin-top:20px'>
          <div style='font-size:56px'>📄</div>
          <h3 style='font-family:Playfair Display,serif;color:#0D1B2A'>Prêt à générer vos documents</h3>
          <p style='color:#9ca3af;max-width:400px;margin:0 auto;font-size:14px'>Chargez le tableau des loyers, sélectionnez la société et le mois, puis cliquez sur <b>Générer Avis & Quittances</b>.</p>
        </div>""", unsafe_allow_html=True)
    else:
        if not loyers_file_rq: st.error("⚠️ Veuillez charger le tableau des loyers."); st.stop()
        if societe_sel=="M.A LA GARENNE": st.warning("⚠️ M.A LA GARENNE : pas d'avis configuré."); st.stop()
        from io import BytesIO as _BytesIO
        loyers_bytes = _BytesIO(loyers_file_rq.read()); loyers_bytes.seek(0)
        with st.spinner(f"📄 Génération — {societe_sel} · {mois_sel} {int(annee_sel)}..."):
            result_avis, result_quit, stats = run_ra_rq(loyers_bytes, societe_sel, mois_num, int(annee_sel))
        if isinstance(stats, str): st.error(f"❌ {stats}"); st.stop()
        c1,c2,c3=st.columns(3)
        with c1: st.markdown(f'<div class="stat-card stat-total"><div class="stat-num" style="color:#3b82f6">{stats["nb_locataires"]}</div><div class="stat-lbl">Locataires</div></div>',unsafe_allow_html=True)
        with c2: st.markdown(f'<div class="stat-card stat-ok"><div class="stat-num" style="color:#22c55e">{stats["nb_avis"]}</div><div class="stat-lbl">Avis générés</div></div>',unsafe_allow_html=True)
        with c3: st.markdown(f'<div class="stat-card stat-alert"><div class="stat-num" style="color:#f59e0b">{stats["nb_quittances"]}</div><div class="stat-lbl">Quittances</div></div>',unsafe_allow_html=True)
        nb_sans = stats['nb_locataires'] - stats['nb_quittances']
        if nb_sans>0: st.info(f"ℹ️ {nb_sans} locataire(s) sans paiement — quittance non générée.")
        st.markdown('<div class="section-title">📥 Télécharger les fichiers</div>',unsafe_allow_html=True)
        dl1,dl2=st.columns(2)
        with dl1:
            st.markdown(f"<div class='dl-card'><div style='font-size:40px'>📋</div><div style='font-weight:600;color:#0D1B2A;margin:10px 0 4px;font-size:16px'>Avis d'Échéance</div><div style='font-size:12px;color:#9ca3af;margin-bottom:4px'>{stats['nb_avis']} avis — {societe_sel}</div><div style='font-size:11px;color:#d1d5db;margin-bottom:14px'>{stats['nom_zip_avis']}</div></div>",unsafe_allow_html=True)
            st.download_button("⬇ Télécharger le ZIP",data=result_avis[0].getvalue(),file_name=stats['nom_zip_avis'],mime="application/zip",key="dl_avis",use_container_width=True)
        with dl2:
            st.markdown(f"<div class='dl-card'><div style='font-size:40px'>✅</div><div style='font-weight:600;color:#0D1B2A;margin:10px 0 4px;font-size:16px'>Quittances de Loyer</div><div style='font-size:12px;color:#9ca3af;margin-bottom:4px'>{stats['nb_quittances']} quittances — locataires ayant payé</div><div style='font-size:11px;color:#d1d5db;margin-bottom:14px'>{stats['nom_zip_quit']}</div></div>",unsafe_allow_html=True)
            st.download_button("⬇ Télécharger le ZIP",data=result_quit[0].getvalue(),file_name=stats['nom_zip_quit'],mime="application/zip",key="dl_quit",use_container_width=True)


# ══ ROBOT GFS ══
elif st.session_state.robot == "GFS":
    st.markdown(f"<h1 style='font-family:Playfair Display,serif;font-size:32px;color:#0D1B2A'>🧾 Génération Factures Salle — {annee_gfs}</h1><hr>", unsafe_allow_html=True)
    if not lancer:
        st.markdown("""<div style='text-align:center;padding:60px 20px;background:white;border-radius:20px;box-shadow:0 2px 12px rgba(0,0,0,.05);margin-top:20px'>
          <div style='font-size:56px'>🧾</div>
          <h3 style='font-family:Playfair Display,serif;color:#0D1B2A'>Prêt à générer vos factures</h3>
          <p style='color:#9ca3af;max-width:400px;margin:0 auto;font-size:14px'>Chargez le fichier récap soirées, sélectionnez l'année et cliquez sur <b>Générer les Factures</b>.</p>
          <div style='margin-top:32px;display:flex;justify-content:center;gap:32px'>
            <div style='text-align:center'><div style='font-size:28px'>📊</div><div style='font-size:12px;color:#9ca3af;margin-top:4px'>Récap soirées</div></div>
            <div style='font-size:28px;color:#d1d5db;padding-top:8px'>→</div>
            <div style='text-align:center'><div style='font-size:28px'>🧾</div><div style='font-size:12px;color:#9ca3af;margin-top:4px'>Factures PDF</div></div>
            <div style='font-size:28px;color:#d1d5db;padding-top:8px'>→</div>
            <div style='text-align:center'><div style='font-size:28px'>📦</div><div style='font-size:12px;color:#9ca3af;margin-top:4px'>ZIP facturier</div></div>
          </div>
        </div>""", unsafe_allow_html=True)
    else:
        if not recap_file:
            st.error("⚠️ Veuillez charger le fichier récap soirées.")
            st.stop()
        from io import BytesIO as _BytesIO
        recap_bytes = _BytesIO(recap_file.read())
        recap_bytes.seek(0)
        with st.spinner(f"🧾 Génération factures {annee_gfs}..."):
            buf_zip, nom_zip, stats = run_gfs(recap_bytes, int(annee_gfs))
        if isinstance(stats, str):
            st.error(f"❌ {stats}")
            st.stop()
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(f'<div class="stat-card stat-total"><div class="stat-num" style="color:#3b82f6">{stats["nb"]}</div><div class="stat-lbl">Factures générées</div></div>', unsafe_allow_html=True)
        with c2:
            st.markdown(f'<div class="stat-card stat-ok"><div class="stat-num" style="color:#22c55e;font-size:18px">{stats["premiere"]}</div><div class="stat-lbl">Première facture</div></div>', unsafe_allow_html=True)
        with c3:
            st.markdown(f'<div class="stat-card stat-alert"><div class="stat-num" style="color:#f59e0b;font-size:18px">{stats["derniere"]}</div><div class="stat-lbl">Dernière facture</div></div>', unsafe_allow_html=True)
        st.markdown('<div class="section-title">📥 Télécharger le facturier</div>', unsafe_allow_html=True)
        st.markdown(f"""<div class='dl-card'>
          <div style='font-size:48px'>📦</div>
          <div style='font-weight:600;color:#0D1B2A;margin:10px 0 4px;font-size:18px'>Facturier {annee_gfs}</div>
          <div style='font-size:13px;color:#6b7280;margin-bottom:4px'>{stats['nb']} factures en ordre chronologique</div>
          <div style='font-size:11px;color:#d1d5db;margin-bottom:16px'>{nom_zip}</div>
        </div>""", unsafe_allow_html=True)
        st.download_button(
            "⬇ Télécharger le ZIP",
            data=buf_zip.getvalue(),
            file_name=nom_zip,
            mime="application/zip",
            key="dl_gfs",
            use_container_width=True
        )
