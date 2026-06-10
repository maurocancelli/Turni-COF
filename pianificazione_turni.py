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

GIORNI_LABELS   = ["Dom", "Lun", "Mar", "Mer", "Gio", "Ven", "Sab", "Dom"]
GIORNI_CHIAVI   = ["Dom_P", "Lun", "Mar", "Mer", "Gio", "Ven", "Sab", "Dom_S"]
GIORNI_BASE     = ["Lunedì", "Martedì", "Mercoledì", "Giovedì", "Venerdì", "Sabato"]
# offset rispetto al lunedì della settimana corrente
OFFSETS         = [-1, 0, 1, 2, 3, 4, 5, 6]

OPZIONI_TURNO   = ["06:00-13:00", "12:30-19:30", "13:00-20:00", "RIPOSO", "MALATTIA", "FERIE", "PERMESSO"]
TARGET_DEFAULT  = {"Dom_P": 45, "Lun": 90, "Mar": 75, "Mer": 75, "Gio": 75, "Ven": 90, "Sab": 90, "Dom_S": 45}
TARGET_DOM      = 10   # numero fisso di lavoratori la domenica

# Matrice rotazione turni (week_iso % 4 → squadra)
# ciclo == 0: sq1=M, sq2=M, sq3=P13, sq4=P12:30
# ciclo == 1: sq1=P12:30, sq2=P13, sq3=M, sq4=M
# ciclo == 2: sq1=M, sq2=M, sq3=P12:30, sq4=P13
# ciclo == 3: sq1=P13, sq2=P12:30, sq3=M, sq4=M
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

def turno_base(squadra: int, numero_settimana: int) -> str:
    ciclo = numero_settimana % 4
    return MATRICE_TURNI[ciclo][squadra]

def colora_celle(valore):
    v = str(valore)
    if v == "MALATTIA":   return "background-color:#ffcccc;color:#cc0000;font-weight:bold;"
    if v == "FERIE":      return "background-color:#ffe6cc;color:#cc6600;font-weight:bold;"
    if v == "PERMESSO":   return "background-color:#e6f2ff;color:#0066cc;"
    if v == "RIPOSO":     return "background-color:#f2f2f2;color:#7f7f7f;"
    if "06:00" in v:      return "background-color:#e6ffed;color:#1a7f37;"
    if "12:30" in v or "13:00" in v: return "background-color:#fbefff;color:#8250df;"
    return ""

def file_settimana(anno: int, week: int) -> str:
    return f"Turni_W{week:02d}_{anno}.csv"

def is_definitiva(anno: int, week: int) -> bool:
    fname = file_settimana(anno, week)
    if not os.path.exists(fname):
        return False
    try:
        df = pd.read_csv(fname, nrows=1)
        return df.get("_definitiva", pd.Series([False]))[0] == True
    except Exception:
        return False

def salva_settimana(df: pd.DataFrame, anno: int, week: int, definitiva: bool):
    df = df.copy()
    df["_definitiva"] = definitiva
    df.to_csv(file_settimana(anno, week), index=False)

def carica_settimana(anno: int, week: int) -> pd.DataFrame | None:
    fname = file_settimana(anno, week)
    if not os.path.exists(fname):
        return None
    df = pd.read_csv(fname)
    if "_definitiva" in df.columns:
        df = df.drop(columns=["_definitiva"])
    return df

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
        # rimuovi colonne legacy
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
        rows.append({"Nome": nome, "Contratto": contratto, "Squadra": sq,
                     "Riposo 1": "Nessuno", "Riposo 2": "Nessuno",
                     "Malattia Fino Al": None,
                     "Ferie W1": None, "Ferie W2": None, "Ferie W3": None})
    return pd.DataFrame(rows)

if "df_anagrafica" not in st.session_state:
    st.session_state.df_anagrafica = init_anagrafica()

def salva_anagrafica(df):
    st.session_state.df_anagrafica = df.copy()
    df.to_csv(FILE_ANAGRAFICA, index=False)

# ─────────────────────────────────────────────
# LOGICA DOMENICHE
# ─────────────────────────────────────────────
def calcola_domeniche_precedenti(week_target: int, anno_target: int) -> dict[str, bool]:
    """
    Per ogni dipendente, calcola se ha lavorato la domenica della settimana
    precedente a week_target (utile per la rotazione "una sì una no").
    Restituisce dict nome -> ha_lavorato_domenica_scorsa (bool)
    """
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
# GENERAZIONE TABELLONE SETTIMANA
# ─────────────────────────────────────────────
def genera_tabellone(week_num: int, anno: int, lunedi: datetime.date,
                     dom_precedenti: dict[str, bool],
                     target_pct: dict[str, float]) -> pd.DataFrame:
    """
    Genera il tabellone turni per una settimana.
    dom_precedenti: dict nome -> bool (True = ha lavorato Dom_S settimana precedente)
    target_pct: dict chiave -> float (0-1)
    """
    dati = st.session_state.df_anagrafica.to_dict("records")
    dati = [d for d in dati if d.get("Nome") and str(d.get("Nome")).strip()]

    n_totale = len(dati)
    data_dom_p = lunedi - datetime.timedelta(days=1)
    data_dom_s = lunedi + datetime.timedelta(days=6)

    # Dizionario rapido per lookup
    dip_map = {d["Nome"]: d for d in dati}

    # ── STEP 1: costruisci struttura base ──
    rows = []
    for dip in dati:
        nome = dip["Nome"]
        settimane_ferie = {dip.get("Ferie W1"), dip.get("Ferie W2"), dip.get("Ferie W3")} - {None}
        in_ferie = week_num in settimane_ferie

        data_mal = dip.get("Malattia Fino Al")
        try:
            parsed = pd.to_datetime(data_mal)
            data_mal = None if pd.isnull(parsed) else parsed.date()
        except Exception:
            data_mal = None

        rows.append({
            "Dipendente": nome,
            "Contratto": dip["Contratto"],
            "Squadra": dip["Squadra"],
            "_in_ferie": in_ferie,
            "_data_mal": data_mal,
            "Dom_P": None,
            "Lun": None, "Mar": None, "Mer": None,
            "Gio": None, "Ven": None, "Sab": None,
            "Dom_S": None,
        })

    df = pd.DataFrame(rows)

    # ── STEP 2: assegna turni infrasettimanali (Lun-Sab) ──
    for idx, row in df.iterrows():
        dip = dip_map[row["Dipendente"]]
        in_ferie = row["_in_ferie"]
        data_mal = row["_data_mal"]
        t_base = turno_base(int(dip["Squadra"]), week_num)

        for chiave, offset in zip(GIORNI_CHIAVI[1:7], OFFSETS[1:7]):
            data_g = lunedi + datetime.timedelta(days=offset)
            in_malattia = data_mal is not None and data_g <= data_mal

            if in_malattia:
                df.at[idx, chiave] = "MALATTIA"
            elif in_ferie:
                df.at[idx, chiave] = "FERIE"
            else:
                df.at[idx, chiave] = t_base

    # ── STEP 3: gestione riposi PT (Lun-Sab) ──
    for idx, row in df.iterrows():
        if row["Contratto"] != "PT":
            continue
        if row["_in_ferie"]:
            continue
        dip = dip_map[row["Dipendente"]]
        riposi = [r for r in [dip.get("Riposo 1"), dip.get("Riposo 2")]
                  if r and r != "Nessuno"]
        if not riposi:
            continue

        # Determina se lavora la domenica (Dom_S) — prima stima: usa rotazione
        # La domenica verrà assegnata nello step successivo,
        # per ora usiamo la domenica precedente come proxy
        ha_lavorato_dom_p = dom_precedenti.get(row["Dipendente"], False)

        if len(riposi) == 2:
            # Scegli quale dei due sospendere in base alle esigenze di copertura
            # (il giorno con target più alto ha più bisogno → tenere quel giorno lavorativo)
            # → mettiamo a riposo il giorno con target MINORE tra i due
            def target_di(giorno_nome):
                for k, off in zip(GIORNI_CHIAVI[1:7], OFFSETS[1:7]):
                    if GIORNI_BASE[off - 1] == giorno_nome:
                        return target_pct.get(k, 0.75)
                return 0.75

            t1 = target_di(riposi[0])
            t2 = target_di(riposi[1])
            # giorno con target minore → candidato a diventare riposo
            riposo_primario = riposi[0] if t1 <= t2 else riposi[1]
            riposo_secondario = riposi[1] if t1 <= t2 else riposi[0]

            # Se ha lavorato domenica precedente → usa solo 1 riposo infrasettimanale
            # scegliamo quello col target minore (riposo_primario)
            # Se non ha lavorato domenica → usa entrambi i riposi
            riposi_da_applicare = [riposo_primario] if ha_lavorato_dom_p else [riposo_primario, riposo_secondario]
        else:
            riposi_da_applicare = riposi

        for giorno_nome in riposi_da_applicare:
            # Trova la chiave corrispondente
            for chiave, offset in zip(GIORNI_CHIAVI[1:7], OFFSETS[1:7]):
                if GIORNI_BASE[offset - 1] == giorno_nome:
                    if df.at[idx, chiave] not in {"MALATTIA", "FERIE"}:
                        df.at[idx, chiave] = "RIPOSO"
                    break

    # ── STEP 4: riposo FT (Lun-Sab) ──
    # Per ogni FT che non ha assenze, assegna 1 giorno di riposo
    # scegliendo il giorno con il maggiore surplus di copertura
    for idx, row in df.iterrows():
        if row["Contratto"] != "FT":
            continue
        if row["_in_ferie"]:
            continue
        data_mal = row["_data_mal"]
        # se è in malattia tutta la settimana, salta
        tutti_malattia = all(
            df.at[idx, k] == "MALATTIA" for k in GIORNI_CHIAVI[1:7]
        )
        if tutti_malattia:
            continue

        # calcola il surplus per ogni giorno candidato
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
    # Dom_P
    for idx, row in df.iterrows():
        dip = dip_map[row["Dipendente"]]
        in_ferie = row["_in_ferie"]
        data_mal = row["_data_mal"]
        in_malattia_dom_p = data_mal is not None and data_dom_p <= data_mal

        if in_malattia_dom_p:
            df.at[idx, "Dom_P"] = "MALATTIA"
        elif in_ferie:
            # chi è in ferie lavora la domenica PRIMA di partire
            # (la domenica precedente all'inizio ferie = Dom_P di questa settimana)
            df.at[idx, "Dom_P"] = turno_base(int(dip["Squadra"]), week_num)
        else:
            # ripristina il turno dalla domenica precedente se disponibile
            ha_lav = dom_precedenti.get(row["Dipendente"], None)
            if ha_lav is None:
                # prima settimana: nessuna info → mattino di default
                df.at[idx, "Dom_P"] = "06:00-13:00"
            elif ha_lav:
                # ha lavorato domenica scorsa → oggi riposo (salvo necessità)
                df.at[idx, "Dom_P"] = "RIPOSO"
            else:
                df.at[idx, "Dom_P"] = turno_base(int(dip["Squadra"]), week_num)

    # Dom_S
    for idx, row in df.iterrows():
        dip = dip_map[row["Dipendente"]]
        in_ferie = row["_in_ferie"]
        data_mal = row["_data_mal"]
        in_malattia_dom_s = data_mal is not None and data_dom_s <= data_mal

        if in_malattia_dom_s:
            df.at[idx, "Dom_S"] = "MALATTIA"
        elif in_ferie:
            # rientra sempre di lunedì → domenica di rientro è libera
            df.at[idx, "Dom_S"] = "RIPOSO"
        else:
            ha_lav_dom_p = str(df.at[idx, "Dom_P"]) not in ASSENTE
            if ha_lav_dom_p:
                df.at[idx, "Dom_S"] = "RIPOSO"
            else:
                df.at[idx, "Dom_S"] = turno_base(int(dip["Squadra"]), week_num)

    # ── STEP 6: garantisci TARGET_DOM lavoratori domenica (Dom_S) ──
    lav_dom_s = (~df["Dom_S"].isin(ASSENTE)).sum()
    da_aggiungere = TARGET_DOM - lav_dom_s

    if da_aggiungere > 0:
        # priorità: chi non ha lavorato Dom_P e non ha assenze
        candidati = df[
            (df["Dom_S"] == "RIPOSO") &
            (~df["Dom_P"].isin(ASSENTE)) == False
        ].index.tolist()
        # fallback: chiunque sia in RIPOSO
        candidati_fallback = df[df["Dom_S"] == "RIPOSO"].index.tolist()

        for idx in candidati + [c for c in candidati_fallback if c not in candidati]:
            if da_aggiungere <= 0:
                break
            df.at[idx, "Dom_S"] = turno_base(int(dip_map[df.at[idx, "Dipendente"]]["Squadra"]), week_num)
            da_aggiungere -= 1

    # ── CLEANUP colonne interne ──
    df = df.drop(columns=["_in_ferie", "_data_mal"])
    colonne = ["Dipendente", "Contratto", "Squadra"] + GIORNI_CHIAVI
    return df[colonne]

# ─────────────────────────────────────────────
# CALCOLO RIFERIMENTI TEMPORALI
# ─────────────────────────────────────────────
oggi = datetime.date.today()
lunedi_corrente = oggi - datetime.timedelta(days=oggi.weekday())
iso_corrente = lunedi_corrente.isocalendar()
week_corrente = iso_corrente[1]
anno_corrente = iso_corrente[0]

# Settimana precedente
lunedi_prec = lunedi_corrente - datetime.timedelta(weeks=1)
iso_prec = lunedi_prec.isocalendar()
week_prec = iso_prec[1]
anno_prec = iso_prec[0]

# Lista settimane da mostrare: 1 precedente + corrente + 13 avanti = 15 totali
settimane = []
for delta in range(-1, 14):
    lun = lunedi_corrente + datetime.timedelta(weeks=delta)
    iso = lun.isocalendar()
    settimane.append((iso[0], iso[1], lun))  # (anno, week, lunedi)

# ─────────────────────────────────────────────
# SIDEBAR — parametri
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

    # Converti date per visualizzazione
    df_show = st.session_state.df_anagrafica.copy()
    if "Malattia Fino Al" in df_show.columns:
        df_show["Malattia Fino Al"] = pd.to_datetime(df_show["Malattia Fino Al"], errors="coerce").dt.date

    config_anagrafica = {
        "Contratto": st.column_config.SelectboxColumn("Contratto", options=["FT", "PT"], required=True),
        "Squadra": st.column_config.NumberColumn("Squadra", min_value=1, max_value=4, step=1, required=True),
        "Riposo 1": st.column_config.SelectboxColumn("Riposo 1 (PT)", options=["Nessuno"] + GIORNI_BASE),
        "Riposo 2": st.column_config.SelectboxColumn("Riposo 2 (PT)", options=["Nessuno"] + GIORNI_BASE),
        "Malattia Fino Al": st.column_config.DateColumn("Malattia Fino Al", format="DD/MM/YYYY"),
        "Ferie W1": st.column_config.NumberColumn("Ferie W1 (N. Settimana ISO)", min_value=1, max_value=53),
        "Ferie W2": st.column_config.NumberColumn("Ferie W2 (N. Settimana ISO)", min_value=1, max_value=53),
        "Ferie W3": st.column_config.NumberColumn("Ferie W3 (N. Settimana ISO)", min_value=1, max_value=53),
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
            nuovo_nome       = st.text_input("Nome e Cognome")
            nuovo_contratto  = st.selectbox("Contratto", ["FT", "PT"])
            nuova_squadra    = st.selectbox("Squadra", [1, 2, 3, 4])
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
                    st.s
