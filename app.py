"""Lista cose da fare per clienti e stagioni (Streamlit + Google Sheets).

Avvio locale:  streamlit run app.py   (serve .streamlit/secrets.toml)

Dati salvati in un Google Sheet con 3 fogli: todo_clienti, todo_stagioni, todo_task
(creati in automatico al primo avvio). Se lo sheet è quello di topsheet, gli elenchi dei
clienti e delle stagioni vengono letti (sola lettura) dai suoi fogli Clienti e Stagioni.
Secrets: gli stessi di topsheet (sheet_id + gcp_json).

Predisposta per i topsheet: l'importatore deve solo chiamare
`ensure_cliente`, `ensure_stagione` e `add_task(..., chiave_esterna=...)`.
La chiave_esterna (es. "cliente|stagione|ordine|tipo") evita i duplicati
quando lo stesso topsheet viene reimportato.
"""
import html
import json
import re
from datetime import date, timedelta
from io import BytesIO

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# Prefisso dei fogli creati da questa app: così si può usare lo stesso Google Sheet
# di altre app (es. TOPSHEETDB) senza toccare i fogli esistenti (Clienti, Stagioni, Ordini...).
PREFISSO_TAB = "todo_"

TIPI = [
    "Telefonare",
    "Scadenza",
    "Controllo ordine",
    "Ordine evaso?",
    "Pagamento ricevuto?",
    "Altro",
]

# Colonne di ogni foglio (l'ordine conta: A, B, C, ...)
SCHEMA = {
    "clienti": ["nome", "contatti"],
    "stagioni": ["nome"],
    "task": ["id", "cliente", "stagione", "tipo", "descrizione", "scadenza",
             "stato", "origine", "chiave_esterna", "note", "creato_il", "completato_il"],
}
COL_ID = 1               # A
COL_CHIAVE = 9           # I  (chiave_esterna)
COL_STATO = "G"
COL_COMPLETATO = "L"


# ------------------------------------------------------------ Google Sheets
@st.cache_resource(show_spinner="Connessione a Google Sheets…")
def get_book():
    # Stessi Secrets dell'app topsheet: sheet_id + gcp_json (oppure gcp_service_account)
    if "gcp_json" in st.secrets:
        info = json.loads(st.secrets["gcp_json"])
    else:
        info = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return gspread.authorize(creds).open_by_key(st.secrets["sheet_id"])


@st.cache_resource(show_spinner=False)
def get_ws():
    sh = get_book()
    esistenti = {w.title.lower(): w for w in sh.worksheets()}
    fogli = {}
    for nome, colonne in SCHEMA.items():
        tab = PREFISSO_TAB + nome
        ws = esistenti.get(tab.lower()) or sh.add_worksheet(tab, rows=1000, cols=len(colonne))
        if not ws.row_values(1):
            ws.append_row(colonne, value_input_option="RAW")
        fogli[nome] = ws
    return fogli


@st.cache_data(ttl=60, show_spinner=False)
def read(nome):
    righe = get_ws()[nome].get_all_records(numericise_ignore=["all"])
    return pd.DataFrame(righe, columns=SCHEMA[nome]).fillna("").astype(str)


def season_key(code):
    """SS26 < FW26 < SS27 ...: ordina per anno, poi primavera/estate prima di autunno/inverno."""
    c = str(code).upper().strip()
    m = re.search(r"(\d{2})\s*$", c)
    year = int(m.group(1)) if m else 0
    half = 0 if c.startswith("PS") else 1 if c.startswith(("SS", "S", "PE", "RE")) else 2
    return (year, half)


@st.cache_data(ttl=120, show_spinner=False)
def esterno(foglio, colonna):
    """Legge (solo lettura) una colonna di un foglio di topsheet, es. Clienti/name. Se manca: lista vuota."""
    try:
        rec = get_book().worksheet(foglio).get_all_records(numericise_ignore=["all"], default_blank="")
    except Exception:
        return []
    return [str(r.get(colonna, "")) for r in rec if str(r.get(colonna, "")).strip()]


def _unici(nomi):
    visti = {}
    for n in nomi:
        n = " ".join(str(n).split())
        if n:
            visti.setdefault(n.casefold(), n)
    return list(visti.values())


def elenco_clienti():
    """Clienti di topsheet (foglio Clienti) + quelli aggiunti da questa app."""
    return sorted(_unici(esterno("Clienti", "name") + list(read("clienti")["nome"])), key=str.casefold)


def elenco_stagioni():
    """Stagioni di topsheet (foglio Stagioni) + quelle aggiunte da questa app, dalla più recente."""
    return sorted(_unici(esterno("Stagioni", "code") + list(read("stagioni")["nome"])),
                  key=lambda c: (season_key(c), c.casefold()), reverse=True)


def _riga(ws, task_id):
    """Numero di riga (1-based) del task, oppure None."""
    try:
        return ws.col_values(COL_ID).index(str(task_id)) + 1
    except ValueError:
        return None


def _ensure(foglio, nome, esistenti, extra=()):
    """Ritorna il nome già presente (anche in topsheet, senza badare alle maiuscole) o lo aggiunge."""
    nome = " ".join((nome or "").split())
    if not nome:
        return None
    for e in esistenti:
        if e.casefold() == nome.casefold():
            return e
    get_ws()[foglio].append_row([nome, *extra], value_input_option="RAW")
    read.clear()
    return nome


def ensure_cliente(nome):
    return _ensure("clienti", nome, elenco_clienti(), ("",))


def ensure_stagione(nome):
    return _ensure("stagioni", nome, elenco_stagioni())


def add_task(cliente, stagione, tipo, descrizione="", scadenza=None,
             origine="manuale", note="", chiave_esterna=None):
    """Inserisce un task. Se `chiave_esterna` esiste già, non duplica (ritorna None)."""
    ws = get_ws()["task"]
    if chiave_esterna and chiave_esterna in ws.col_values(COL_CHIAVE):
        return None
    ids = [int(x) for x in ws.col_values(COL_ID)[1:] if str(x).isdigit()]
    nuovo_id = max(ids) + 1 if ids else 1
    ws.append_row(
        [nuovo_id, cliente, stagione or "", tipo, descrizione,
         scadenza.isoformat() if isinstance(scadenza, date) else (scadenza or ""),
         "da fare", origine, chiave_esterna or "", note,
         date.today().isoformat(), ""],
        value_input_option="RAW",
    )
    read.clear()
    return nuovo_id


def set_stato(task_id, fatto):
    ws = get_ws()["task"]
    r = _riga(ws, task_id)
    if r is None:
        return
    ws.batch_update(
        [
            {"range": f"{COL_STATO}{r}", "values": [["fatto" if fatto else "da fare"]]},
            {"range": f"{COL_COMPLETATO}{r}", "values": [[date.today().isoformat() if fatto else ""]]},
        ],
        value_input_option="RAW",
    )
    read.clear()


def delete_task(task_id):
    ws = get_ws()["task"]
    r = _riga(ws, task_id)
    if r is not None:
        ws.delete_rows(r)
    read.clear()


# ----------------------------------------------------------------- callbacks
def on_toggle(task_id):
    set_stato(task_id, st.session_state[f"chk_{task_id}"])


def on_delete(task_id):
    delete_task(task_id)


# ----------------------------------------------------------------- pagine
def esc(s):
    return html.escape(str(s))


def pagina_da_fare():
    st.header("Da fare")

    clienti = elenco_clienti()
    stagioni = elenco_stagioni()

    c1, c2, c3 = st.columns(3)
    f_cli = c1.multiselect("Cliente", clienti)
    f_sta = c2.multiselect("Stagione", stagioni)
    f_tipo = c3.multiselect("Tipo", TIPI)
    c4, c5 = st.columns(2)
    mostra_fatti = c4.toggle("Mostra completati", value=False)
    solo_urgenti = c5.toggle("Solo scaduti / entro 7 giorni", value=False)

    df = read("task").copy()
    if not mostra_fatti:
        df = df[df["stato"] == "da fare"]
    if f_cli:
        df = df[df["cliente"].isin(f_cli)]
    if f_sta:
        df = df[df["stagione"].isin(f_sta)]
    if f_tipo:
        df = df[df["tipo"].isin(f_tipo)]
    df["_fatto"] = df["stato"] == "fatto"
    df["_senza"] = df["scadenza"] == ""
    df = df.sort_values(["_fatto", "_senza", "scadenza", "cliente"])

    oggi = date.today()
    tra7 = oggi + timedelta(days=7)

    visibili = 0
    for r in df.to_dict("records"):
        scad = date.fromisoformat(r["scadenza"]) if r["scadenza"] else None
        fatto = r["stato"] == "fatto"

        if scad and not fatto:
            semaforo = "🔴" if scad < oggi else "🟡" if scad <= tra7 else "⚪"
        else:
            semaforo = "✅" if fatto else "⚪"

        if solo_urgenti and not (scad and scad <= tra7 and not fatto):
            continue
        visibili += 1

        testo = f"**{esc(r['cliente'])}**"
        if r["stagione"]:
            testo += f" · {esc(r['stagione'])}"
        testo += f" · {esc(r['tipo'])}"
        if r["descrizione"]:
            testo += f" — {esc(r['descrizione'])}"
        if fatto:
            testo = f"~~{testo}~~"
        dettaglio = []
        if scad:
            dettaglio.append(f"📅 {scad.strftime('%d/%m/%Y')}")
        if r["origine"] == "topsheet":
            dettaglio.append("📄 topsheet")
        if r["note"]:
            dettaglio.append(f"📝 {esc(r['note'])}")

        col_chk, col_txt, col_del = st.columns([0.08, 0.80, 0.12], vertical_alignment="center")
        col_chk.checkbox(
            "fatto", value=fatto, key=f"chk_{r['id']}",
            on_change=on_toggle, args=(r["id"],), label_visibility="collapsed",
        )
        col_txt.markdown(
            f"{semaforo} {testo}" + (f"  \n<small>{' · '.join(dettaglio)}</small>" if dettaglio else ""),
            unsafe_allow_html=True,
        )
        with col_del.popover("🗑"):
            st.write("Eliminare definitivamente?")
            st.button("Sì, elimina", key=f"del_{r['id']}", type="primary",
                      on_click=on_delete, args=(r["id"],))

    if visibili == 0:
        st.info("Niente da fare con questi filtri. 🎉")
    else:
        st.caption(f"{visibili} task · 🔴 scaduto · 🟡 entro 7 giorni")


def pagina_aggiungi():
    st.header("Aggiungi task")

    clienti = elenco_clienti()
    stagioni = elenco_stagioni()
    NUOVO = "➕ Nuovo…"

    with st.form("nuovo_task", clear_on_submit=True):
        c1, c2 = st.columns(2)
        sel_cli = c1.selectbox("Cliente", clienti + [NUOVO])
        nuovo_cli = c1.text_input("Nome nuovo cliente (se hai scelto «Nuovo»)")
        sel_sta = c2.selectbox("Stagione", ["(nessuna)"] + stagioni + [NUOVO])
        nuova_sta = c2.text_input("Nome nuova stagione (se hai scelto «Nuovo»)", placeholder="es. FW27")

        c3, c4 = st.columns(2)
        tipo = c3.selectbox("Tipo", TIPI)
        senza_scad = c4.checkbox("Senza scadenza (ignora la data qui sotto)")
        scad = c4.date_input("Scadenza", value=date.today() + timedelta(days=7), format="DD/MM/YYYY")

        descrizione = st.text_input("Descrizione")
        note = st.text_area("Note", height=80)

        if st.form_submit_button("Aggiungi", type="primary"):
            nome_cli = nuovo_cli if sel_cli == NUOVO else sel_cli
            if not (nome_cli or "").strip():
                st.error("Indica il cliente.")
            else:
                cliente = ensure_cliente(nome_cli)
                if sel_sta == "(nessuna)":
                    stagione = None
                elif sel_sta == NUOVO:
                    stagione = ensure_stagione(nuova_sta)
                else:
                    stagione = sel_sta
                add_task(cliente, stagione, tipo, descrizione,
                         None if senza_scad else scad, note=note)
                st.success("Task aggiunto.")


def _con_conteggi(nomi, colonna):
    task = read("task")
    aperti = task[task["stato"] == "da fare"].groupby(colonna).size()
    fatti = task[task["stato"] == "fatto"].groupby(colonna).size()
    out = pd.DataFrame({"nome": nomi})
    out["aperti"] = out["nome"].map(aperti).fillna(0).astype(int)
    out["fatti"] = out["nome"].map(fatti).fillna(0).astype(int)
    return out


def pagina_anagrafica():
    st.header("Clienti e stagioni")
    st.caption("Clienti e stagioni di topsheet compaiono qui in automatico; "
               "quelli aggiunti da qui restano solo in questa app.")
    c1, c2 = st.columns(2)

    with c1:
        st.subheader("Clienti")
        with st.form("f_cli", clear_on_submit=True):
            nome = st.text_input("Nome")
            if st.form_submit_button("Aggiungi cliente") and nome.strip():
                ensure_cliente(nome)
                st.rerun()
        st.dataframe(_con_conteggi(elenco_clienti(), "cliente"), hide_index=True, width="stretch")

    with c2:
        st.subheader("Stagioni")
        with st.form("f_sta", clear_on_submit=True):
            nome = st.text_input("Nome stagione", placeholder="es. SS27")
            if st.form_submit_button("Aggiungi stagione") and nome.strip():
                ensure_stagione(nome)
                st.rerun()
        st.dataframe(_con_conteggi(elenco_stagioni(), "stagione"), hide_index=True, width="stretch")


def pagina_esporta():
    st.header("Esporta")
    df = read("task").drop(columns=["id", "chiave_esterna", "creato_il"])
    st.dataframe(df, hide_index=True, width="stretch")
    buf = BytesIO()
    df.to_excel(buf, index=False)
    st.download_button("Scarica Excel", buf.getvalue(), file_name="task.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ----------------------------------------------------------------- main
def main():
    st.set_page_config(page_title="Cose da fare", page_icon="✅", layout="wide")
    try:
        get_ws()
    except Exception as e:
        st.error("Impossibile collegarsi al Google Sheet. Controlla i Secrets (sheet_id e gcp_json, "
                 "gli stessi di topsheet) e che il foglio sia condiviso con il service account (Editor).")
        st.exception(e)
        st.stop()

    pagina = st.sidebar.radio(
        "Menu", ["Da fare", "Aggiungi task", "Clienti e stagioni", "Esporta"]
    )
    if st.sidebar.button("🔄 Aggiorna dati"):
        st.cache_data.clear()
        st.rerun()
    {
        "Da fare": pagina_da_fare,
        "Aggiungi task": pagina_aggiungi,
        "Clienti e stagioni": pagina_anagrafica,
        "Esporta": pagina_esporta,
    }[pagina]()


main()
