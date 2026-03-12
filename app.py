import streamlit as st
import pandas as pd
from robot_ac_gl import run, MOIS_NOMS

# ── CONFIG PAGE ───────────────────────────────────────────────
st.set_page_config(
    page_title="Compta Sarah",
    page_icon="💼",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── CSS DESIGN ────────────────────────────────────────────────
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@600;700&family=Inter:wght@300;400;500;600&display=swap');

  html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

  /* Fond général */
  .stApp { background: #F7F8FC; }

  /* Sidebar */
  [data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0D1B2A 0%, #1a2f4a 100%);
    border-right: 1px solid #ffffff10;
  }
  [data-testid="stSidebar"] * { color: #e8edf5 !important; }
  [data-testid="stSidebar"] .stSelectbox label,
  [data-testid="stSidebar"] .stFileUploader label { color: #94a8c4 !important; font-size:12px; letter-spacing:0.05em; text-transform:uppercase; }

  /* Header logo */
  .logo-block {
    text-align: center;
    padding: 28px 0 20px 0;
    border-bottom: 1px solid #ffffff15;
    margin-bottom: 24px;
  }
  .logo-title {
    font-family: 'Playfair Display', serif;
    font-size: 26px;
    font-weight: 700;
    color: #ffffff !important;
    letter-spacing: -0.5px;
  }
  .logo-sub {
    font-size: 11px;
    color: #5b7fa6 !important;
    letter-spacing: 0.15em;
    text-transform: uppercase;
    margin-top: 4px;
  }

  /* Cartes stats */
  .stat-card {
    background: white;
    border-radius: 14px;
    padding: 20px 24px;
    box-shadow: 0 2px 12px rgba(0,0,0,0.06);
    border-left: 4px solid;
    margin-bottom: 8px;
  }
  .stat-ok    { border-color: #22c55e; }
  .stat-alert { border-color: #f59e0b; }
  .stat-skip  { border-color: #94a3b8; }
  .stat-total { border-color: #3b82f6; }
  .stat-num { font-size: 32px; font-weight: 700; line-height: 1; }
  .stat-lbl { font-size: 12px; color: #6b7280; margin-top: 4px; text-transform: uppercase; letter-spacing: 0.05em; }

  /* Titres sections */
  .section-title {
    font-family: 'Playfair Display', serif;
    font-size: 20px;
    color: #0D1B2A;
    margin: 28px 0 16px 0;
    padding-bottom: 8px;
    border-bottom: 2px solid #e5e7eb;
  }

  /* Tableau résultats */
  .result-table { width:100%; border-collapse:collapse; font-size:13px; }
  .result-table th {
    background: #0D1B2A; color: white;
    padding: 10px 12px; text-align:left;
    font-weight: 500; font-size: 11px;
    text-transform: uppercase; letter-spacing: 0.06em;
  }
  .result-table td { padding: 9px 12px; border-bottom: 1px solid #f1f5f9; }
  .result-table tr:hover td { background: #f8faff; }
  .badge-ok     { background:#dcfce7; color:#166534; padding:2px 10px; border-radius:20px; font-size:11px; font-weight:600; }
  .badge-alerte { background:#fef3c7; color:#92400e; padding:2px 10px; border-radius:20px; font-size:11px; font-weight:600; }
  .badge-ignore { background:#f1f5f9; color:#475569; padding:2px 10px; border-radius:20px; font-size:11px; font-weight:600; }
  .montant-pos  { color:#16a34a; font-weight:600; }
  .montant-neg  { color:#dc2626; font-weight:600; }

  /* Boutons download */
  .stDownloadButton > button {
    background: linear-gradient(135deg, #0D1B2A, #1a3a5c) !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    padding: 10px 20px !important;
    font-weight: 500 !important;
    width: 100% !important;
    transition: all 0.2s !important;
  }
  .stDownloadButton > button:hover {
    transform: translateY(-1px) !important;
    box-shadow: 0 4px 15px rgba(13,27,42,0.3) !important;
  }

  /* Upload zone */
  [data-testid="stFileUploader"] {
    background: #ffffff08;
    border: 1px dashed #ffffff30;
    border-radius: 10px;
    padding: 6px;
  }

  /* Bouton lancer */
  .stButton > button {
    background: linear-gradient(135deg, #16a34a, #15803d) !important;
    color: white !important;
    border: none !important;
    border-radius: 12px !important;
    padding: 14px 32px !important;
    font-size: 15px !important;
    font-weight: 600 !important;
    width: 100% !important;
    margin-top: 8px !important;
    letter-spacing: 0.02em !important;
    transition: all 0.2s !important;
  }
  .stButton > button:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 6px 20px rgba(22,163,74,0.4) !important;
  }

  /* Alert box */
  .alert-box {
    background: #fffbeb;
    border: 1px solid #f59e0b;
    border-radius: 12px;
    padding: 16px 20px;
    margin-top: 16px;
  }
  .alert-box h4 { color: #92400e; margin:0 0 10px 0; font-size:14px; }
  .alert-item { color:#78350f; font-size:13px; padding: 3px 0; }

  /* Onglets navigation robots */
  .robot-nav {
    display: flex; gap: 8px; margin-bottom: 24px;
    padding: 6px; background: white;
    border-radius: 14px; box-shadow: 0 2px 8px rgba(0,0,0,0.06);
  }
  .robot-btn-active {
    background: #0D1B2A; color: white;
    padding: 10px 18px; border-radius: 10px;
    font-size: 13px; font-weight: 600;
  }
  .robot-btn {
    background: transparent; color: #6b7280;
    padding: 10px 18px; border-radius: 10px;
    font-size: 13px; font-weight: 500;
  }

  /* Divider */
  hr { border: none; border-top: 1px solid #e5e7eb; margin: 20px 0; }

  /* Selectbox style */
  .stSelectbox > div > div { border-radius: 8px !important; }
</style>
""", unsafe_allow_html=True)

# ── SIDEBAR ───────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div class="logo-block">
      <div class="logo-title">💼 Compta Sarah</div>
      <div class="logo-sub">La belle vie, occupe toi de ton mari</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("#### 🏦 Rapprochement Bancaire + Attribution de Compte")
    st.markdown("<small style='color:#5b7fa6'>Attribution des comptes & intégration loyers</small>", unsafe_allow_html=True)
    st.markdown("---")

    st.markdown("<span style='font-size:11px;text-transform:uppercase;letter-spacing:0.1em;color:#5b7fa6'>🏢 Société</span>", unsafe_allow_html=True)
    SOCIETES = [
        "SCI EMJJ",
        "SCI EMJJ MIRABEAU",
        "SCI MMA",
        "SCI AZM",
        "SCI MAZ",
        "SOGEPA",
        "M.A LA GARENNE",
    ]
    societe_sel = st.selectbox("Société", SOCIETES, key="societe")

    st.markdown("<br><span style='font-size:11px;text-transform:uppercase;letter-spacing:0.1em;color:#5b7fa6'>📂 Fichiers</span>", unsafe_allow_html=True)
    releve_file = st.file_uploader("Relevé bancaire", type=["xlsx"], key="releve")
    pivot_file  = st.file_uploader("Tableau Pivot",   type=["xlsx"], key="pivot")
    loyers_file = st.file_uploader("Tableau Loyers",  type=["xlsx"], key="loyers")

    st.markdown("---")
    st.markdown("<span style='font-size:11px;text-transform:uppercase;letter-spacing:0.1em;color:#5b7fa6'>📅 Période</span>", unsafe_allow_html=True)

    mois_options = list(MOIS_NOMS.values())
    mois_sel = st.selectbox("Mois", mois_options, index=2)
    mois_num = list(MOIS_NOMS.keys())[list(MOIS_NOMS.values()).index(mois_sel)]
    annee_sel = st.number_input("Année", min_value=2020, max_value=2035, value=2026, step=1)

    st.markdown("---")
    lancer = st.button("▶  Lancer le Robot")

    st.markdown("""
    <div style='margin-top:auto;padding-top:40px;text-align:center'>
      <small style='color:#2d4a6a;font-size:10px'>Prochains robots :<br>
      RA · RQ · RAA · VS · GFS · LF · LT</small>
    </div>
    """, unsafe_allow_html=True)

# ── CONTENU PRINCIPAL ─────────────────────────────────────────
st.markdown("""
<div style='margin-bottom:8px'>
  <h1 style='font-family:Playfair Display,serif;font-size:32px;color:#0D1B2A;margin:0'>
    Rapprochement Bancaire + Attribution de Compte
  </h1>
  <p style='color:#6b7280;margin:4px 0 0 0;font-size:15px'>
    Attribution des comptes comptables &amp; mise à jour du tableau des loyers
  </p>
</div>
<hr>
""", unsafe_allow_html=True)

# État initial
if not lancer:
    st.markdown("""
    <div style='text-align:center;padding:60px 20px;background:white;border-radius:20px;
                box-shadow:0 2px 12px rgba(0,0,0,0.05);margin-top:20px'>
      <div style='font-size:56px;margin-bottom:16px'>🤖</div>
      <h3 style='font-family:Playfair Display,serif;color:#0D1B2A;margin:0 0 8px 0'>
        Prêt à traiter vos données
      </h3>
      <p style='color:#9ca3af;max-width:400px;margin:0 auto;font-size:14px'>
        Chargez vos 3 fichiers dans le panneau de gauche,
        sélectionnez le mois et cliquez sur <b>Lancer le Robot</b>.
      </p>
      <div style='margin-top:32px;display:flex;justify-content:center;gap:32px'>
        <div style='text-align:center'>
          <div style='font-size:28px'>📊</div>
          <div style='font-size:12px;color:#9ca3af;margin-top:4px'>Relevé bancaire</div>
        </div>
        <div style='font-size:28px;color:#d1d5db;padding-top:8px'>→</div>
        <div style='text-align:center'>
          <div style='font-size:28px'>🔍</div>
          <div style='font-size:12px;color:#9ca3af;margin-top:4px'>Matching comptes</div>
        </div>
        <div style='font-size:28px;color:#d1d5db;padding-top:8px'>→</div>
        <div style='text-align:center'>
          <div style='font-size:28px'>📥</div>
          <div style='font-size:12px;color:#9ca3af;margin-top:4px'>Exports Excel</div>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)

else:
    # Vérification fichiers
    if not releve_file or not pivot_file or not loyers_file:
        st.error("⚠️ Veuillez charger les 3 fichiers avant de lancer le robot.")
        st.stop()

    # Lancement
    with st.spinner(f"🤖 Traitement en cours — {mois_sel} {annee_sel}..."):
        df_res, buffers, maj, err = run(
            releve_file, pivot_file, loyers_file,
            mois=mois_num, annee=int(annee_sel)
        )

    if err:
        st.error(f"❌ {err}")
        st.stop()

    buf1, buf2, buf3 = buffers
    nb_ok   = len(df_res[df_res.STATUT=='OK'])
    nb_al   = len(df_res[df_res.STATUT=='ALERTE'])
    nb_ig   = len(df_res[df_res.STATUT=='IGNORE'])
    alertes = df_res[df_res.STATUT=='ALERTE']

    # ── STATS ─────────────────────────────────────────────────
    c1,c2,c3,c4 = st.columns(4)
    with c1:
        st.markdown(f"""<div class="stat-card stat-total">
          <div class="stat-num" style="color:#3b82f6">{len(df_res)}</div>
          <div class="stat-lbl">Transactions</div></div>""", unsafe_allow_html=True)
    with c2:
        st.markdown(f"""<div class="stat-card stat-ok">
          <div class="stat-num" style="color:#22c55e">{nb_ok}</div>
          <div class="stat-lbl">Identifiées ✓</div></div>""", unsafe_allow_html=True)
    with c3:
        st.markdown(f"""<div class="stat-card stat-alert">
          <div class="stat-num" style="color:#f59e0b">{nb_al}</div>
          <div class="stat-lbl">Alertes ⚠</div></div>""", unsafe_allow_html=True)
    with c4:
        st.markdown(f"""<div class="stat-card stat-skip">
          <div class="stat-num" style="color:#94a3b8">{nb_ig}</div>
          <div class="stat-lbl">Ignorées ⏸</div></div>""", unsafe_allow_html=True)

    # ── ALERTES ───────────────────────────────────────────────
    if nb_al > 0:
        items = "".join([
            f"<div class='alert-item'>→ {row['DATE'].strftime('%d/%m/%Y') if pd.notna(row['DATE']) else '?'} &nbsp;|&nbsp; "
            f"{str(row['LIBELLE'])[:60]} &nbsp;|&nbsp; <b>{row['MONTANT']:.2f}€</b></div>"
            for _,row in alertes.iterrows()
        ])
        st.markdown(f"""<div class="alert-box">
          <h4>⚠️ {nb_al} ligne(s) à saisir manuellement</h4>{items}</div>""",
          unsafe_allow_html=True)

    # ── TABLEAU RÉSULTATS ─────────────────────────────────────
    st.markdown('<div class="section-title">📋 Détail des transactions</div>', unsafe_allow_html=True)

    rows_html = ""
    for _,r in df_res.iterrows():
        d   = r['DATE'].strftime('%d/%m/%y') if pd.notna(r['DATE']) else '?'
        sgn = '+' if r['SENS']=='CREDIT' else '-'
        cls_m = 'montant-pos' if r['SENS']=='CREDIT' else 'montant-neg'
        badge = f"<span class='badge-{r['STATUT'].lower()}'>{r['STATUT']}</span>"
        rows_html += f"""<tr>
          <td>{d}</td>
          <td style='max-width:260px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap'>{str(r['LIBELLE'])[:60]}</td>
          <td class='{cls_m}'>{sgn}{abs(r['MONTANT']):.2f}€</td>
          <td><code style='font-size:11px;background:#f1f5f9;padding:2px 6px;border-radius:4px'>{r['TYPE']}</code></td>
          <td><code style='font-size:11px;background:#f1f5f9;padding:2px 6px;border-radius:4px'>{r['COMPTE']}</code></td>
          <td style='color:#374151;font-size:12px'>{str(r['NOM'])[:38]}</td>
          <td>{badge}</td>
        </tr>"""

    st.markdown(f"""
    <div style='background:white;border-radius:16px;padding:20px;box-shadow:0 2px 12px rgba(0,0,0,0.06);overflow-x:auto'>
      <table class='result-table'>
        <thead><tr>
          <th>Date</th><th>Libellé</th><th>Montant</th>
          <th>Type</th><th>Compte</th><th>Identifié comme</th><th>Statut</th>
        </tr></thead>
        <tbody>{rows_html}</tbody>
      </table>
    </div>""", unsafe_allow_html=True)

    # ── LOYERS MAJ ────────────────────────────────────────────
    if maj:
        st.markdown('<div class="section-title">🏠 Loyers reportés dans le tableau</div>', unsafe_allow_html=True)
        cols = st.columns(min(len(maj), 4))
        for i, m in enumerate(maj):
            with cols[i % 4]:
                st.markdown(f"""
                <div style='background:white;border-radius:12px;padding:14px 16px;
                            box-shadow:0 2px 8px rgba(0,0,0,0.05);border-left:3px solid #22c55e;
                            margin-bottom:8px'>
                  <div style='font-size:12px;color:#6b7280'>✅ Mis à jour</div>
                  <div style='font-size:13px;font-weight:600;color:#0D1B2A;margin-top:2px'>{m}</div>
                </div>""", unsafe_allow_html=True)

    # ── TÉLÉCHARGEMENTS ───────────────────────────────────────
    st.markdown('<div class="section-title">📥 Télécharger les fichiers</div>', unsafe_allow_html=True)
    dl1, dl2, dl3 = st.columns(3)

    with dl1:
        st.markdown("""<div style='background:white;border-radius:14px;padding:20px;
                        box-shadow:0 2px 10px rgba(0,0,0,0.06);text-align:center;margin-bottom:8px'>
          <div style='font-size:32px'>📋</div>
          <div style='font-weight:600;color:#0D1B2A;margin:8px 0 4px'>Intégration Logiciel</div>
          <div style='font-size:12px;color:#9ca3af;margin-bottom:12px'>Écritures comptables prêtes à importer</div>
        </div>""", unsafe_allow_html=True)
        st.download_button(
            "⬇ Télécharger",
            data=buf1.getvalue(),
            file_name=f"INTEGRATION_LOGICIEL_{mois_sel.upper()}_{annee_sel}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl1"
        )

    with dl2:
        st.markdown("""<div style='background:white;border-radius:14px;padding:20px;
                        box-shadow:0 2px 10px rgba(0,0,0,0.06);text-align:center;margin-bottom:8px'>
          <div style='font-size:32px'>🏠</div>
          <div style='font-weight:600;color:#0D1B2A;margin:8px 0 4px'>Tableau Loyers MAJ</div>
          <div style='font-size:12px;color:#9ca3af;margin-bottom:12px'>Loyers reportés avec dates et modes</div>
        </div>""", unsafe_allow_html=True)
        st.download_button(
            "⬇ Télécharger",
            data=buf2.getvalue(),
            file_name=f"TABLEAU_LOYERS_{mois_sel.upper()}_{annee_sel}_MAJ.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl2"
        )

    with dl3:
        st.markdown(f"""<div style='background:white;border-radius:14px;padding:20px;
                        box-shadow:0 2px 10px rgba(0,0,0,0.06);text-align:center;margin-bottom:8px'>
          <div style='font-size:32px'>{'⚠️' if nb_al>0 else '✅'}</div>
          <div style='font-weight:600;color:#0D1B2A;margin:8px 0 4px'>Rapport Alertes</div>
          <div style='font-size:12px;color:#9ca3af;margin-bottom:12px'>
            {nb_al} ligne(s) à saisir manuellement
          </div>
        </div>""", unsafe_allow_html=True)
        st.download_button(
            "⬇ Télécharger",
            data=buf3.getvalue(),
            file_name=f"ALERTES_{mois_sel.upper()}_{annee_sel}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl3"
        )
