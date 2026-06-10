import streamlit as st
import pandas as pd
import datetime
import os

# ─────────────────────────────────────────────
# CONFIGURAZIONE PAGINA
# ─────────────────────────────────────────────
st.set_page_config(page_title="Pianificazione Turni - COF", layout="wide")
st.title("📦 Pianificazione Mensile Reparto E-commerce")

# ─────────────────────────────────────────────
# COSTANTI
# ─────────────────────────────────────────────
FILE_ANAGRAFICA = "anagrafica_salvata.csv"

GIORNI_LABELS = ["Dom", "Lun", "Mar", "Mer", "Gio", "Ven", "Sab", "Dom"]
GIORNI_CHIAVI = ["Dom_P", "Lun", "Mar", "Mer", "Gio", "Ven", "Sab", "Dom_S"]
GIORNI_BASE   = ["Lunedì", "Martedì", "Mercoledì", "Giovedì", "Venerdì", "Sabato"]
OFFSETS       = [-1, 0, 1, 2, 3, 4, 5, 6]

OPZIONI_TURNO  = ["06:00-13:00", "12:30-19:30", "13:00-20:00", "RIPOSO", "MALATTIA", "FERIE", "PERMESSO"]
TARGET_DEFAULT = {"Dom_P": 45, "Lun": 90, "Mar": 75, "Mer": 75, "Gio": 75, "Ven": 90, "Sab": 90, "Dom_S": 45}
TARGET_DOM     = 10

MATRICE_TURNI = {
    0: {1: "06:00-13:00", 2: "06:00-13:00", 3: "13:00-20:00",  4: "12:30-19:30"},
    1: {1: "12:30-19:30", 2: "13:00-20:00", 3: "06:00-13:00",  4: "06:00-13:00"},
    2: {1: "06:00-13:00", 2: "06:00-13:00", 3: "12:30-19:30",  4: "13:00-20:00"},
    3: {1: "13:00-20:00", 2: "12:30-19:30", 3: "06:00-13:00",  4: "06:00-13:00"},
}

ASSENTE = {"RIPOSO", "MALATTIA", "FERIE", "PERMESSO"}

# ─────────────────────────────────────────────
# UTILITY
# ─────────────────────────────────────────────
def pulisci_riposi(val):
    if pd.isna(val) or val is None or str(val).strip() == "":
        return "Nessuno"
    return str(val).strip()

def turno_base(squadra, numero_settimana):
    ciclo = numero_settimana % 4
    return MATRICE_TURNI[ciclo][int(squadra)]

def colora_celle(valore):
    v = str(valore)
    if v == "MALATTIA":  return "background-color:#ffcccc;color:#cc0000;font-weight:bold;"
    if v == "FERIE":     return "background-color:#ffe6cc;color:#cc6600;font-weight:bold;"
    if v == "PERMESSO":  return "background-color:#e6f2ff;color:#0066cc;"
    if v == "RIPOSO":    return "background-color:#f2f2f2;color:#7f7f7f;"
    if "06:00" in v:     return "background-color:#e6ffed;color:#1a7f37;"
    if "12:30" in v or "13:00" in v: return "background-color:#fbefff;color:#8250df;"
    return ""

def file_settimana(anno, week):
    return f"Turni_W{week:02d}_{anno}.csv"

def is_definitiva(anno, week):
    fname = file_settimana(anno, week)
    if not os.path.exists(fname):
        return False
    try:
        df = pd.read_csv(fname, nrows=1)
        val = df.get("_definitiva", pd.Series([False]))[0]
        return bool(val) == True
    except Exception:
        return False

def salva_settimana(df, anno, week, definitiva):
    df = df.copy()
    df["_definitiva"] = definitiva
    df.to_csv(file_settimana(anno, week), index=False)

def carica_settimana(anno, week):
    fname = file_settimana(anno, week)
    if not os.path.exists(fname):
        return None
    df = pd.read_csv(fname)
    if "_definitiva" in df.columns:
        df = df.drop(columns=["_definitiva"])
    return df

def parse_data_malattia(val):
    """Converte qualsiasi valore in datetime.date o None, senza mai restituire NaT."""
    if val is None:
        return None
    try:
        parsed = pd.to_datetime(val)
        if pd.isnull(parsed):
            return None
        return parsed.date()
    except Exception:
        return None

# ─────────────────────────────────────────────
# INIT ANAGRAFICA
# ─────────────────────────────────────────────
def init_anagrafica():
    if os.path.exists(FILE_ANAGRAFICA):
        df = pd.read_csv(FILE_ANAGRAFICA)
        df = df.where(pd.notnull(df), None)
        for col in ["Riposo 1", "Riposo 2"]:
            if col in df.columns:
                df[col] = df[col].apply(pulisci_riposi)
        for col in ["Malattia Fino Al", "Ferie W1", "Ferie W2", "Ferie W3"]:
            if col not in df.columns:
                df[col] = None
        for col in ["Dom Scorsa"]:
            if col in df.columns:
                df = df.drop(columns=[col])
        return df

    nomi_base = [
        ("MARVIN MENDOZA","FT",1),("MANUEL MENDOZA","FT",2),
        ("ELAINE ALVES BRAGA","FT",3),("CRISTOPHER BULOSAN","FT",4),
        ("CATERINA CAVALLARI","FT",1),("MICHELA GAVAZZENI","FT",2),
        ("FRANCESCO CUARESMA","FT",3),("BARBARA PERALTA","FT",4),
        ("SIMONA RICCARDI","FT",1),("MARIA MABANTA","FT",2),
        ("ARVIN PASCUA","FT",3),("CAMILLA TOCCHETTI","FT",4),
        ("DAPHNI SPEERING","FT",1),("EUGENE TANADA","FT",2),
        ("RUSSEL CIPRIANO","FT",3),("CARBUNGCAL FELIX","FT",4),
        ("SARA GHITTI","FT",1),("DANIELA BONO","FT",2),
        ("JONIMAE TAN","FT",3),("ROSARIO PETILUNA","FT",4),
        ("LEA MAGTIBAY","FT",1),("BABYJANE MAGTIBAY","FT",2),
        ("NIKITA DONGHI","FT",3),
    ]
    rows = []
    for nome, contratto, sq in nomi_base:
        rows.append({
            "Nome": nome, "Contratto": contratto, "Squadra": sq,
            "Riposo 1": "Nessuno", "Riposo 2": "Nessuno",
            "Malattia Fino Al": None,
            "Ferie W1": None, "Ferie W2": None, "Ferie W3": None
        })
    return pd.DataFrame(rows)

if "df_anagrafica" not in st.session_state:
    st.session_state.df_anagrafica = init_anagrafica()

def salva_anagrafica(df):
    st.session_state.df_anagrafica = df.copy()
    df.to_csv(FILE_ANAGRAFICA, index=False)

# ─────────────────────────────────────────────
# LOGICA DOMENICHE PRECEDENTI
# ─────────────────────────────────────────────
def calcola_domeniche_precedenti(week_target, anno_target):
    week_prec = week_target - 1
    anno_prec = anno_target
    if week_prec < 1:
        week_prec = 52
        anno_prec -= 1
    df_prec = carica_settimana(anno_prec, week_prec)
    risultato = {}
    if df_prec is not None and "Dom_S" in df_prec.columns:
        for _, row in df_prec.iterrows():
            nome = row.get("Dipendente")
            if nome:
                risultato[nome] = str(row.get("Dom_S", "RIPOSO")) not in ASSENTE
    return risultato

# ─────────────────────────────────────────────
# GENERAZIONE TABELLONE
# ─────────────────────────────────────────────
def genera_tabellone(week_num, anno, lunedi, dom_precedenti, target_pct):
    dati = st.session_state.df_anagrafica.to_dict("records")
    dati = [d for d in dati if d.get("Nome") and str(d.get("Nome")).strip()]
    n_totale = len(dati)

    data_dom_p = lunedi - datetime.timedelta(days=1)
    data_dom_s = lunedi + datetime.timedelta(days=6)
    dip_map = {d["Nome"]: d for d in dati}

    # ── STEP 1: struttura base ──
    rows = []
    for dip in dati:
        nome = dip["Nome"]
        ferie_settimane = set()
        for fk in ["Ferie W1", "Ferie W2", "Ferie W3"]:
            v = dip.get(fk)
            if v is not None:
                try:
                    ferie_settimane.add(int(v))
                except Exception:
                    pass
        in_ferie = week_num in ferie_settimane
        data_mal = parse_data_malattia(dip.get("Malattia Fino Al"))

        rows.append({
            "Dipendente": nome,
            "Contratto": dip["Contratto"],
            "Squadra": dip["Squadra"],
            "_in_ferie": in_ferie,
            "_data_mal": data_mal,
            "Dom_P": None, "Lun": None, "Mar": None, "Mer": None,
            "Gio": None, "Ven": None, "Sab": None, "Dom_S": None,
        })

    df = pd.DataFrame(rows)

    # ── STEP 2: turni infrasettimanali Lun-Sab ──
    for idx, row in df.iterrows():
        dip = dip_map[row["Dipendente"]]
        in_ferie = row["_in_ferie"]
        data_mal = row["_data_mal"]
        t_base = turno_base(dip["Squadra"], week_num)

        for chiave, offset in zip(GIORNI_CHIAVI[1:7], OFFSETS[1:7]):
            data_g = lunedi + datetime.timedelta(days=offset)
            in_malattia = (data_mal is not None) and (data_g <= data_mal)
            if in_malattia:
                df.at[idx, chiave] = "MALATTIA"
            elif in_ferie:
                df.at[idx, chiave] = "FERIE"
            else:
                df.at[idx, chiave] = t_base

    # ── STEP 3: riposi PT ──
    for idx, row in df.iterrows():
        if row["Contratto"] != "PT" or row["_in_ferie"]:
            continue
        dip = dip_map[row["Dipendente"]]
        riposi = [r for r in [dip.get("Riposo 1"), dip.get("Riposo 2")]
                  if r and r != "Nessuno"]
        if not riposi:
            continue

        ha_lavorato_dom_p = dom_precedenti.get(row["Dipendente"], False)

        def target_di(giorno_nome):
            for k, off in zip(GIORNI_CHIAVI[1:7], OFFSETS[1:7]):
                if GIORNI_BASE[off - 1] == giorno_nome:
                    return target_pct.get(k, 0.75)
            return 0.75

        if len(riposi) == 2:
            t1 = target_di(riposi[0])
            t2 = target_di(riposi[1])
            riposo_primario   = riposi[0] if t1 <= t2 else riposi[1]
            riposo_secondario = riposi[1] if t1 <= t2 else riposi[0]
            riposi_da_applicare = [riposo_primario] if ha_lavorato_dom_p else [riposo_primario, riposo_secondario]
        else:
            riposi_da_applicare = riposi

        for giorno_nome in riposi_da_applicare:
            for chiave, offset in zip(GIORNI_CHIAVI[1:7], OFFSETS[1:7]):
                if GIORNI_BASE[offset - 1] == giorno_nome:
                    if df.at[idx, chiave] not in {"MALATTIA", "FERIE"}:
                        df.at[idx, chiave] = "RIPOSO"
                    break

    # ── STEP 4: riposo FT ──
    for idx, row in df.iterrows():
        if row["Contratto"] != "FT" or row["_in_ferie"]:
            continue
        tutti_malattia = all(df.at[idx, k] == "MALATTIA" for k in GIORNI_CHIAVI[1:7])
        if tutti_malattia:
            continue

        miglior_chiave = None
        max_surplus = -9999
        for chiave in GIORNI_CHIAVI[1:7]:
            if df.at[idx, chiave] in ASSENTE:
                continue
            lavoratori = (~df[chiave].isin(ASSENTE)).sum()
            target_n = n_totale * target_pct.get(chiave, 0.75)
            surplus = lavoratori - target_n
            if surplus > max_surplus:
                max_surplus = surplus
                miglior_chiave = chiave

        if miglior_chiave:
            df.at[idx, miglior_chiave] = "RIPOSO"

    # ── STEP 5: domeniche ──
    for idx, row in df.iterrows():
        dip = dip_map[row["Dipendente"]]
        in_ferie = row["_in_ferie"]
        data_mal = row["_data_mal"]
        in_malattia_dom_p = (data_mal is not None) and (data_dom_p <= data_mal)

        if in_malattia_dom_p:
            df.at[idx, "Dom_P"] = "MALATTIA"
        elif in_ferie:
            df.at[idx, "Dom_P"] = turno_base(dip["Squadra"], week_num)
        else:
            ha_lav = dom_precedenti.get(row["Dipendente"], None)
            if ha_lav is None:
                df.at[idx, "Dom_P"] = "06:00-13:00"
            elif ha_lav:
                df.at[idx, "Dom_P"] = "RIPOSO"
            else:
                df.at[idx, "Dom_P"] = turno_base(dip["Squadra"], week_num)

    for idx, row in df.iterrows():
        dip = dip_map[row["Dipendente"]]
        in_ferie = row["_in_ferie"]
        data_mal = row["_data_mal"]
        in_malattia_dom_s = (data_mal is not None) and (data_dom_s <= data_mal)

        if in_malattia_dom_s:
            df.at[idx, "Dom_S"] = "MALATTIA"
        elif in_ferie:
            df.at[idx, "Dom_S"] = "RIPOSO"
        else:
            ha_lav_dom_p = str(df.at[idx, "Dom_P"]) not in ASSENTE
            if ha_lav_dom_p:
                df.at[idx, "Dom_S"] = "RIPOSO"
            else:
                df.at[idx, "Dom_S"] = turno_base(dip["Squadra"], week_num)

    # ── STEP 6: garantisci TARGET_DOM domenica ──
    lav_dom_s = (~df["Dom_S"].isin(ASSENTE)).sum()
    da_aggiungere = TARGET_DOM - lav_dom_s

    if da_aggiungere > 0:
        candidati = df[df["Dom_S"] == "RIPOSO"].index.tolist()
        for idx in candidati:
            if da_aggiungere <= 0:
                break
            nome_dip = df.at[idx, "Dipendente"]
            df.at[idx, "Dom_S"] = turno_base(dip_map[nome_dip]["Squadra"], week_num)
            da_aggiungere -= 1

    df = df.drop(columns=["_in_ferie", "_data_mal"])
    return df[["Dipendente", "Contratto", "Squadra"] + GIORNI_CHIAVI]

# ─────────────────────────────────────────────
# RIFERIMENTI TEMPORALI
# ─────────────────────────────────────────────
oggi = datetime.date.today()
lunedi_corrente = oggi - datetime.timedelta(days=oggi.weekday())
iso_corrente = lunedi_corrente.isocalendar()
week_corrente = iso_corrente[1]
anno_corrente = iso_corrente[0]

settimane = []
for delta in range(-1, 14):
    lun = lunedi_corrente + datetime.timedelta(weeks=delta)
    iso = lun.isocalendar()
    settimane.append((iso[0], iso[1], lun))

# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
st.sidebar.header("⚙️ Parametri Operativi")
pieces_ora = st.sidebar.number_input("Pezzi/Ora (Default):", value=100, min_value=1)
st.sidebar.divider()
st.sidebar.subheader("🎯 Forza Lavoro Richiesta (%)")
st.sidebar.caption("Percentuale di dipendenti attivi per ogni giorno")

target_pct = {}
for chiave in GIORNI_CHIAVI:
    label = {"Dom_P": "Domenica (inizio)", "Dom_S": "Domenica (fine)"}.get(chiave, chiave)
    default = TARGET_DEFAULT[chiave]
    target_pct[chiave] = st.sidebar.slider(label, 0, 100, default, key=f"sl_{chiave}") / 100

# ─────────────────────────────────────────────
# TABS PRINCIPALI
# ─────────────────────────────────────────────
tab_anagrafica, tab_turni = st.tabs(["📋 Gestione Anagrafica", "📅 Turni Settimanali"])

# ══════════════════════════════════════════════
# SCHEDA 1 — ANAGRAFICA
# ══════════════════════════════════════════════
with tab_anagrafica:
    st.subheader("👥 Lista Personale e Assenze Programmate")

    df_show = st.session_state.df_anagrafica.copy()
    if "Malattia Fino Al" in df_show.columns:
        df_show["Malattia Fino Al"] = pd.to_datetime(df_show["Malattia Fino Al"], errors="coerce").dt.date

    config_anagrafica = {
        "Contratto": st.column_config.SelectboxColumn("Contratto", options=["FT", "PT"], required=True),
        "Squadra": st.column_config.NumberColumn("Squadra", min_value=1, max_value=4, step=1, required=True),
        "Riposo 1": st.column_config.SelectboxColumn("Riposo 1 (PT)", options=["Nessuno"] + GIORNI_BASE),
        "Riposo 2": st.column_config.SelectboxColumn("Riposo 2 (PT)", options=["Nessuno"] + GIORNI_BASE),
        "Malattia Fino Al": st.column_config.DateColumn("Malattia Fino Al", format="DD/MM/YYYY"),
        "Ferie W1": st.column_config.NumberColumn("Ferie W1 (N. Sett. ISO)", min_value=1, max_value=53),
        "Ferie W2": st.column_config.NumberColumn("Ferie W2 (N. Sett. ISO)", min_value=1, max_value=53),
        "Ferie W3": st.column_config.NumberColumn("Ferie W3 (N. Sett. ISO)", min_value=1, max_value=53),
    }

    df_editato = st.data_editor(
        df_show,
        column_config=config_anagrafica,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        key="anagrafica_editor"
    )

    if st.button("💾 Salva Modifiche Anagrafica", type="primary", use_container_width=True):
        salva_anagrafica(df_editato)
        st.success("✅ Anagrafica aggiornata!")

    st.divider()
    col_add, col_del = st.columns(2)

    with col_add:
        st.subheader("➕ Aggiungi Dipendente")
        with st.container(border=True):
            nuovo_nome      = st.text_input("Nome e Cognome")
            nuovo_contratto = st.selectbox("Contratto", ["FT", "PT"])
            nuova_squadra   = st.selectbox("Squadra", [1, 2, 3, 4])
            nuovo_r1 = st.selectbox("Riposo Fisso 1 (PT)", ["Nessuno"] + GIORNI_BASE)
            nuovo_r2 = st.selectbox("Riposo Fisso 2 (PT)", ["Nessuno"] + GIORNI_BASE)
            if st.button("Aggiungi", use_container_width=True):
                if not nuovo_nome.strip():
                    st.error("Inserisci un nome valido!")
                else:
                    nuova_riga = {
                        "Nome": nuovo_nome.upper(),
                        "Contratto": nuovo_contratto,
                        "Squadra": nuova_squadra,
                        "Riposo 1": nuovo_r1 if nuovo_contratto == "PT" else "Nessuno",
                        "Riposo 2": nuovo_r2 if nuovo_contratto == "PT" else "Nessuno",
                        "Malattia Fino Al": None,
                        "Ferie W1": None, "Ferie W2": None, "Ferie W3": None,
                    }
                    nuovo_df = pd.concat(
                        [st.session_state.df_anagrafica, pd.DataFrame([nuova_riga])],
                        ignore_index=True
                    )
                    salva_anagrafica(nuovo_df)
                    st.success(f"✅ {nuovo_nome.upper()} aggiunto!")
                    st.rerun()

    with col_del:
        st.subheader("🗑️ Rimuovi Dipendente")
        with st.container(border=True):
            lista_nomi = st.session_state.df_anagrafica["Nome"].tolist()
            nome_da_el = st.selectbox("Scegli chi eliminare:", ["Nessuno..."] + lista_nomi)
            if st.button("Elimina", type="primary", use_container_width=True):
                if nome_da_el != "Nessuno...":
                    nuovo_df = st.session_state.df_anagrafica[
                        st.session_state.df_anagrafica["Nome"] != nome_da_el
                    ]
                    salva_anagrafica(nuovo_df)
                    st.success(f"❌ {nome_da_el} rimosso.")
                    st.rerun()

# ══════════════════════════════════════════════
# SCHEDA 2 — TURNI
# ══════════════════════════════════════════════
with tab_turni:

    if st.session_state.df_anagrafica.empty:
        st.warning("⚠️ Aggiungi prima i dipendenti nell'Anagrafica.")
    else:
        labels_week = []
        for anno_w, week_w, lun_w in settimane:
            dom_p = lun_w - datetime.timedelta(days=1)
            dom_s = lun_w + datetime.timedelta(days=6)
            flag_def = "🔒" if is_definitiva(anno_w, week_w) else "📝"
            is_cur = (anno_w == anno_corrente and week_w == week_corrente)
            marker = " ◀" if is_cur else ""
            labels_week.append(f"{flag_def} W{week_w} ({dom_p.day}/{dom_p.month}–{dom_s.day}/{dom_s.month}){marker}")

        tabs_week = st.tabs(labels_week)

        for i, (t_week, (anno_w, week_w, lun_w)) in enumerate(zip(tabs_week, settimane)):
            with t_week:
                col_labels = {}
                rinomina_exp = {}
                for chiave, label, offset in zip(GIORNI_CHIAVI, GIORNI_LABELS, OFFSETS):
                    data_g = lun_w + datetime.timedelta(days=offset)
                    lbl = f"{label} {data_g.day}/{data_g.month}"
                    col_labels[chiave] = lbl
                    rinomina_exp[chiave] = lbl

                config_turni = {
                    chiave: st.column_config.SelectboxColumn(
                        col_labels[chiave],
                        options=OPZIONI_TURNO,
                        disabled=(chiave == "Dom_P" and i > 0)
                    )
                    for chiave in GIORNI_CHIAVI
                }

                definitiva = is_definitiva(anno_w, week_w)
                df_salvato = carica_settimana(anno_w, week_w)

                if df_salvato is not None:
                    df_calcolato = df_salvato.copy()
                    if i > 0:
                        anno_pw, week_pw, _ = settimane[i - 1]
                        df_prec_w = carica_settimana(anno_pw, week_pw)
                        if df_prec_w is not None and "Dom_S" in df_prec_w.columns:
                            dom_s_map = df_prec_w.set_index("Dipendente")["Dom_S"].to_dict()
                            for idx, row in df_calcolato.iterrows():
                                nome = row.get("Dipendente")
                                if nome and nome in dom_s_map:
 
