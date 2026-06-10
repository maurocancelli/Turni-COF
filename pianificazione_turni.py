import streamlit as st
import pandas as pd
import datetime
import os

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

# Domenica: solo turno mattino
TURNO_DOMENICA = "06:00-13:00"

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

def turno_infrasettimanale(squadra, numero_settimana):
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
        return bool(df.get("_definitiva", pd.Series([False]))[0])
    except Exception:
        return False

def salva_settimana(df, anno, week, definitiva):
    d = df.copy()
    d["_definitiva"] = definitiva
    d.to_csv(file_settimana(anno, week), index=False)

def carica_settimana(anno, week):
    fname = file_settimana(anno, week)
    if not os.path.exists(fname):
        return None
    df = pd.read_csv(fname)
    if "_definitiva" in df.columns:
        df = df.drop(columns=["_definitiva"])
    return df

def parse_data_malattia(val):
    if val is None:
        return None
    try:
        parsed = pd.to_datetime(val)
        return None if pd.isnull(parsed) else parsed.date()
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
# LEGGI DOM_S DA FILE STORICO (settimana prima della finestra)
# ─────────────────────────────────────────────
def leggi_dom_s_precedente(week_target, anno_target):
    """
    Restituisce dict nome -> valore_dom_s (stringa) dalla settimana
    immediatamente prima di week_target. Se non esiste ritorna {}.
    """
    w = week_target - 1
    a = anno_target
    if w < 1:
        w = 52
        a -= 1
    df = carica_settimana(a, w)
    if df is None or "Dom_S" not in df.columns:
        return {}
    return {
        row["Dipendente"]: str(row["Dom_S"])
        for _, row in df.iterrows()
        if row.get("Dipendente")
    }

# ─────────────────────────────────────────────
# GENERAZIONE TABELLONE
# dom_s_prec: dict nome -> valore stringa Dom_S settimana precedente
# ─────────────────────────────────────────────
def genera_tabellone(week_num, anno, lunedi, dom_s_prec, target_pct):
    dati = st.session_state.df_anagrafica.to_dict("records")
    dati = [d for d in dati if d.get("Nome") and str(d.get("Nome")).strip()]
    n_totale = len(dati)

    data_dom_p = lunedi - datetime.timedelta(days=1)
    data_dom_s = lunedi + datetime.timedelta(days=6)
    dip_map = {d["Nome"]: d for d in dati}

    rows = []
    for dip in dati:
        nome = dip["Nome"]
        ferie_set = set()
        for fk in ["Ferie W1", "Ferie W2", "Ferie W3"]:
            v = dip.get(fk)
            if v is not None:
                try:
                    ferie_set.add(int(v))
                except Exception:
                    pass
        in_ferie = week_num in ferie_set
        data_mal = parse_data_malattia(dip.get("Malattia Fino Al"))
        rows.append({
            "Dipendente": nome, "Contratto": dip["Contratto"], "Squadra": dip["Squadra"],
            "_in_ferie": in_ferie, "_data_mal": data_mal,
            "Dom_P": None, "Lun": None, "Mar": None, "Mer": None,
            "Gio": None, "Ven": None, "Sab": None, "Dom_S": None,
        })

    df = pd.DataFrame(rows)

    # ── Turni infrasettimanali Lun-Sab ──
    for idx, row in df.iterrows():
        dip = dip_map[row["Dipendente"]]
        t_base = turno_infrasettimanale(dip["Squadra"], week_num)
        for chiave, offset in zip(GIORNI_CHIAVI[1:7], OFFSETS[1:7]):
            data_g = lunedi + datetime.timedelta(days=offset)
            in_mal = (row["_data_mal"] is not None) and (data_g <= row["_data_mal"])
            if in_mal:
                df.at[idx, chiave] = "MALATTIA"
            elif row["_in_ferie"]:
                df.at[idx, chiave] = "FERIE"
            else:
                df.at[idx, chiave] = t_base

    # ── Riposi PT ──
    for idx, row in df.iterrows():
        if row["Contratto"] != "PT" or row["_in_ferie"]:
            continue
        dip = dip_map[row["Dipendente"]]
        riposi = [r for r in [dip.get("Riposo 1"), dip.get("Riposo 2")]
                  if r and r != "Nessuno"]
        if not riposi:
            continue
        val_dom_p = dom_s_prec.get(row["Dipendente"], None)
        ha_lavorato_dom = val_dom_p is not None and val_dom_p not in ASSENTE

        def target_di(g):
            for k, off in zip(GIORNI_CHIAVI[1:7], OFFSETS[1:7]):
                if GIORNI_BASE[off - 1] == g:
                    return target_pct.get(k, 0.75)
            return 0.75

        if len(riposi) == 2:
            t1, t2 = target_di(riposi[0]), target_di(riposi[1])
            r_prim = riposi[0] if t1 <= t2 else riposi[1]
            r_sec  = riposi[1] if t1 <= t2 else riposi[0]
            da_app = [r_prim] if ha_lavorato_dom else [r_prim, r_sec]
        else:
            da_app = riposi

        for g in da_app:
            for chiave, offset in zip(GIORNI_CHIAVI[1:7], OFFSETS[1:7]):
                if GIORNI_BASE[offset - 1] == g:
                    if df.at[idx, chiave] not in {"MALATTIA", "FERIE"}:
                        df.at[idx, chiave] = "RIPOSO"
                    break

    # ── Riposo FT (un giorno a settimana) ──
    for idx, row in df.iterrows():
        if row["Contratto"] != "FT" or row["_in_ferie"]:
            continue
        if all(df.at[idx, k] == "MALATTIA" for k in GIORNI_CHIAVI[1:7]):
            continue
        miglior = None
        max_sur = -9999
        for chiave in GIORNI_CHIAVI[1:7]:
            if df.at[idx, chiave] in ASSENTE:
                continue
            lav = (~df[chiave].isin(ASSENTE)).sum()
            sur = lav - n_totale * target_pct.get(chiave, 0.75)
            if sur > max_sur:
                max_sur = sur
                miglior = chiave
        if miglior:
            df.at[idx, miglior] = "RIPOSO"

    # ── Dom_P: copia esatta della Dom_S della settimana precedente ──
    for idx, row in df.iterrows():
        data_mal = row["_data_mal"]
        in_mal_dom_p = (data_mal is not None) and (data_dom_p <= data_mal)
        if in_mal_dom_p:
            df.at[idx, "Dom_P"] = "MALATTIA"
        else:
            val_prec = dom_s_prec.get(row["Dipendente"], None)
            if val_prec is None:
                # Nessuna info storica: default mattino
                df.at[idx, "Dom_P"] = TURNO_DOMENICA
            else:
                # Copia esatta — stesso valore della Dom_S settimana scorsa
                df.at[idx, "Dom_P"] = val_prec

    # ── Dom_S: rotazione rispetto a Dom_P della STESSA settimana ──
    for idx, row in df.iterrows():
        data_mal = row["_data_mal"]
        in_mal_dom_s = (data_mal is not None) and (data_dom_s <= data_mal)
        if in_mal_dom_s:
            df.at[idx, "Dom_S"] = "MALATTIA"
        elif row["_in_ferie"]:
            # Chi è in ferie: lavora Dom_P (già assegnata sopra), riposa Dom_S
            df.at[idx, "Dom_S"] = "RIPOSO"
        else:
            dom_p_val = str(df.at[idx, "Dom_P"])
            if dom_p_val in ASSENTE:
                # Non ha lavorato Dom_P → lavora Dom_S
                df.at[idx, "Dom_S"] = TURNO_DOMENICA
            else:
                # Ha lavorato Dom_P → riposa Dom_S
                df.at[idx, "Dom_S"] = "RIPOSO"

    # ── Garantisci TARGET_DOM lavoratori in Dom_S ──
    lav = (~df["Dom_S"].isin(ASSENTE)).sum()
    mancanti = TARGET_DOM - lav
    if mancanti > 0:
        for idx in df[df["Dom_S"] == "RIPOSO"].index:
            if mancanti <= 0:
                break
            df.at[idx, "Dom_S"] = TURNO_DOMENICA
            mancanti -= 1

    df = df.drop(columns=["_in_ferie", "_data_mal"])
    return df[["Dipendente", "Contratto", "Squadra"] + GIORNI_CHIAVI]

# ─────────────────────────────────────────────
# RIFERIMENTI TEMPORALI
# ─────────────────────────────────────────────
oggi = datetime.date.today()
lunedi_corrente = oggi - datetime.timedelta(days=oggi.weekday())
iso_c = lunedi_corrente.isocalendar()
week_corrente = iso_c[1]
anno_corrente = iso_c[0]

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
    target_pct[chiave] = st.sidebar.slider(label, 0, 100, TARGET_DEFAULT[chiave], key=f"sl_{chiave}") / 100

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
        df_show, column_config=config_anagrafica,
        num_rows="dynamic", use_container_width=True,
        hide_index=True, key="anagrafica_editor"
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
                        "Nome": nuovo_nome.upper(), "Contratto": nuovo_contratto,
                        "Squadra": nuova_squadra,
                        "Riposo 1": nuovo_r1 if nuovo_contratto == "PT" else "Nessuno",
                        "Riposo 2": nuovo_r2 if nuovo_contratto == "PT" else "Nessuno",
                        "Malattia Fino Al": None,
                        "Ferie W1": None, "Ferie W2": None, "Ferie W3": None,
                    }
                    nuovo_df = pd.concat([st.session_state.df_anagrafica, pd.DataFrame([nuova_riga])], ignore_index=True)
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
                    nuovo_df = st.session_state.df_anagrafica[st.session_state.df_anagrafica["Nome"] != nome_da_el]
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
        # ── Costruisci labels ──
        labels_week = []
        for anno_w, week_w, lun_w in settimane:
            dom_p = lun_w - datetime.timedelta(days=1)
            dom_s = lun_w + datetime.timedelta(days=6)
            flag  = "🔒" if is_definitiva(anno_w, week_w) else "📝"
            cur   = " ◀" if (anno_w == anno_corrente and week_w == week_corrente) else ""
            labels_week.append(f"{flag} W{week_w} ({dom_p.day}/{dom_p.month}–{dom_s.day}/{dom_s.month}){cur}")

        # ── Pre-genera la catena completa ──
        # dom_s_map: nome -> valore stringa Dom_S della settimana precedente
        tabelloni = {}

        for j, (anno_w, week_w, lun_w) in enumerate(settimane):
            # Ricava dom_s_map dalla settimana precedente nella catena
            if j == 0:
                dom_s_map = leggi_dom_s_precedente(week_w, anno_w)
            else:
                anno_pw, week_pw, _ = settimane[j - 1]
                df_prec = tabelloni.get((anno_pw, week_pw))
                if df_prec is not None and "Dom_S" in df_prec.columns:
                    dom_s_map = {
                        rrow["Dipendente"]: str(rrow["Dom_S"])
                        for _, rrow in df_prec.iterrows()
                        if rrow.get("Dipendente")
                    }
                else:
                    dom_s_map = leggi_dom_s_precedente(week_w, anno_w)

            df_salvato = carica_settimana(anno_w, week_w)
            if df_salvato is not None:
                df_chain = df_salvato.copy()
                # Forza Dom_P = copia esatta Dom_S settimana precedente
                for ridx, rrow in df_chain.iterrows():
                    nome = rrow.get("Dipendente")
                    if nome:
                        val = dom_s_map.get(nome, None)
                        if val is not None:
                            df_chain.at[ridx, "Dom_P"] = val
            else:
                df_chain = genera_tabellone(week_w, anno_w, lun_w, dom_s_map, target_pct)

            tabelloni[(anno_w, week_w)] = df_chain

        # ── Renderizza i tab ──
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
                df_calcolato = tabelloni[(anno_w, week_w)].copy()

                if definitiva:
                    st.success("🔒 **Settimana DEFINITIVA**")
                else:
                    st.info("📝 **Settimana PROVVISORIA** — generata automaticamente")

                df_modificato = st.data_editor(
                    df_calcolato,
                    column_config=config_turni,
                    use_container_width=True,
                    hide_index=True,
                    key=f"editor_{anno_w}_{week_w}"
                )

                col1, col2, col3 = st.columns(3)
                with col1:
                    if not definitiva:
                        if st.button("🔒 Blocca come Definitiva", type="primary",
                                     use_container_width=True, key=f"blocca_{anno_w}_{week_w}"):
                            salva_settimana(df_modificato, anno_w, week_w, definitiva=True)
                            st.success("✅ Bloccata!")
                            st.rerun()
                    else:
                        if st.button("🔓 Sblocca", type="secondary",
                                     use_container_width=True, key=f"sblocca_{anno_w}_{week_w}"):
                            salva_settimana(df_modificato, anno_w, week_w, definitiva=False)
                            st.success("↩️ Sbloccata!")
                            st.rerun()
                with col2:
                    if st.button("💾 Salva Modifiche", use_container_width=True,
                                 key=f"salva_{anno_w}_{week_w}"):
                        salva_settimana(df_modificato, anno_w, week_w, definitiva=definitiva)
                        st.success("✅ Salvato!")
                        st.rerun()
                with col3:
                    df_exp = df_modificato.copy()
                    df_exp.rename(columns=rinomina_exp, inplace=True)
                    xlsx_cols = ["Dipendente", "Contratto", "Squadra"] + list(rinomina_exp.values())
                    df_exp = df_exp[[c for c in xlsx_cols if c in df_exp.columns]]
                    csv_data = df_exp.to_csv(index=False, sep=";").encode("utf-8-sig")
                    st.download_button(
                        label="📥 Esporta CSV",
                        data=csv_data,
                        file_name=f"Turni_W{week_w}_{anno_w}.csv",
                        mime="text/csv",
                        use_container_width=True,
                        key=f"down_{anno_w}_{week_w}"
                    )

                st.write("**Vista Colorata:**")
                st.dataframe(df_modificato.style.map(colora_celle), use_container_width=True)

                st.write("**Stima Volumi Giornalieri:**")
                report = []
                for chiave in GIORNI_CHIAVI:
                    op_m = (df_modificato[chiave] == "06:00-13:00").sum()
                    op_p = (df_modificato[chiave].isin(["12:30-19:30", "13:00-20:00"])).sum()
                    report.append({
                        "Giorno": col_labels[chiave],
                        "Mattina (06-13)": int(op_m),
                        "Pomeriggio": int(op_p),
                        "Tot. Operatori": int(op_m + op_p),
                        "Pezzi Stimati": int((op_m + op_p) * 7 * pieces_ora),
                    })
                st.dataframe(pd.DataFrame(report).set_index("Giorno").T, use_container_width=True)
