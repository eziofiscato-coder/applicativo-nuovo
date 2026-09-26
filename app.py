
import streamlit as st
import pandas as pd
from datetime import date, datetime, time
import io

# PDF export - usa reportlab se disponibile, altrimenti fallback
try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet
    REPORTLAB_OK = True
except:
    REPORTLAB_OK = False

st.set_page_config(page_title="ANA Varese - 5 Form Completi", page_icon="🛡️", layout="wide")

# Inizializza sessione
for k in ["volontari", "radio_db", "consegna_radio", "alias_radio", "brogliaccio"]:
    if k not in st.session_state:
        st.session_state[k] = []

# Mock radio_db per test se vuoto
if not st.session_state.radio_db:
    st.session_state.radio_db = [
        {"Matricola": "PD785-001", "Modello": "Hytera PD785", "Tipo": "DMR Regionali", "Stato": "Operativa"},
        {"Matricola": "ANY-878-002", "Modello": "Anytone 878", "Tipo": "Radio Amatoriale", "Stato": "Operativa"}
    ]

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

def to_pdf(df, title):
    """Genera PDF da DataFrame - per tutti i form"""
    output = io.BytesIO()
    if REPORTLAB_OK and not df.empty:
        try:
            doc = SimpleDocTemplate(output, pagesize=landscape(A4), leftMargin=20, rightMargin=20, topMargin=30, bottomMargin=20)
            styles = getSampleStyleSheet()
            story = []
            story.append(Paragraph(f"<b>{title} - ANA Varese</b> - {datetime.now().strftime('%d/%m/%Y %H:%M')}", styles['Title']))
            story.append(Spacer(1, 20))
            # Prepara dati tabella
            cols = list(df.columns)[:10]  # max 10 colonne per stare in pagina
            data = [cols] + df[cols].astype(str).values.tolist()[:100]  # max 100 righe
            table = Table(data, repeatRows=1)
            table.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1A5D1A')),
                ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                ('ALIGN', (0,0), (-1,-1), 'LEFT'),
                ('FONTSIZE', (0,0), (-1,0), 8),
                ('FONTSIZE', (0,1), (-1,-1), 6),
                ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#e8f5e9')]),
            ]))
            story.append(table)
            doc.build(story)
            return output.getvalue()
        except Exception as e:
            st.error(f"Errore PDF: {e}")
            return None
    else:
        # Fallback: PDF testuale semplice se reportlab non c'è
        try:
            from reportlab.pdfgen import canvas
            from reportlab.lib.pagesizes import landscape, A4
            c = canvas.Canvas(output, pagesize=landscape(A4))
            c.setFont("Helvetica-Bold", 14)
            c.drawString(30, 550, f"{title} - ANA Varese - {datetime.now().strftime('%d/%m/%Y %H:%M')}")
            c.setFont("Helvetica", 8)
            y = 520
            cols = " | ".join(list(df.columns)[:8])
            c.drawString(30, y, cols)
            y -= 15
            for _, row in df.head(80).iterrows():
                txt = " | ".join([str(row[c])[:20] for c in list(df.columns)[:8]])
                c.drawString(30, y, txt)
                y -= 10
                if y < 30:
                    c.showPage()
                    y = 550
            c.save()
            return output.getvalue()
        except:
            return None

TIPI_RADIO = ["DMR Regionali", "Tetra", "PMR 446", "Nautiche", "Radio Amatoriale", "DMR Commerciale", "Analogica VHF", "Analogica UHF", "CB 27 MHz", "LPD 433", "Altro"]

st.title("🛡️ ANA Varese - Gestionale 5 Form")
st.caption("Completo - Volontari / DB Radio / Consegna Radio / Alias / Brogliaccio - con PDF")

menu = st.sidebar.radio("Seleziona Form", ["Volontari (con foto)", "DB Radio", "Consegna Radio", "Alias Radio", "Brogliaccio"], index=0)

# ================= VOLONTARI =================
if menu == "Volontari (con foto)":
    st.header("👤 Volontari (con foto) - Campi Originali")
    tab1, tab2, tab3, tab4 = st.tabs(["Anagrafica + Capo ODV", "Contatti", "Ruolo e Squadra", "Foto e Documenti"])
    with st.form("form_volontario_completo", clear_on_submit=False):
        with tab1:
            c1, c2 = st.columns(2)
            with c1:
                nome = st.text_input("Nome *", key="v_nome")
                cognome = st.text_input("Cognome *", key="v_cognome")
                comune = st.text_input("Comune Residenza", key="v_comune")
                via = st.text_input("Via", key="v_via")
                data_nascita = st.date_input("Data Nascita", value=date(1990,1,1), format="DD/MM/YYYY", key="v_data_nasc")
                cod_fisc = st.text_input("Codice Fiscale", key="v_cf")
            with c2:
                capo_odv = st.text_input("Capo ODV *", key="v_capo_odv", help="Nome del Capo ODV di riferimento")
                odv_app = st.selectbox("ODV Associazione di Appartenenza *", ["ANA Varese", "ANA Milano", "Protezione Civile Varese", "Croce Rossa Italiana - Varese", "Misericordia", "ANPAS", "Altro"], key="v_odv")
                if odv_app == "Altro":
                    odv_app = st.text_input("Specifica ODV", key="v_odv_altro") or odv_app
        with tab2:
            c1, c2 = st.columns(2)
            with c1:
                cellulare = st.text_input("Cellulare *", key="v_cell")
                email = st.text_input("Email", key="v_email")
            with c2:
                tel_emerg = st.text_input("Telefono Emergenza", key="v_tel_em")
                note_cont = st.text_area("Note Contatti", key="v_note_cont")
        with tab3:
            c1, c2 = st.columns(2)
            with c1:
                ruolo = st.selectbox("Ruolo *", ["Volontario", "Capo Squadra", "Coordinatore", "Autista", "Radio Operatore", "Capo ODV", "Vice Capo ODV"], key="v_ruolo")
                squadra = st.selectbox("Squadra *", ["Squadra A", "Squadra B", "Squadra C", "Logistica", "Segreteria", "ODV Centrale"], key="v_squadra")
                radio_id = st.text_input("ID Radio / Matricola", key="v_radio_id")
                modello_radio = st.selectbox("Modello Radio", ["Hytera PD785", "Anytone 878", "Motorola", "Altro"], key="v_mod_radio")
            with c2:
                data_iscriz = st.date_input("Data Iscrizione ODV", value=date.today(), format="DD/MM/YYYY", key="v_data_iscr")
                stato_vol = st.selectbox("Stato", ["Attivo", "Inattivo", "In Formazione", "Sospeso"], key="v_stato")
                note_dot = st.text_area("Note Dotazione", key="v_note_dot")
        with tab4:
            doc_tipo = st.text_input("Tipo Documento", key="v_doc_tipo")
            doc_num = st.text_input("Numero Documento", key="v_doc_num")
            doc_scad = st.date_input("Scadenza Documento", value=date.today(), format="DD/MM/YYYY", key="v_doc_scad")
            foto_file = st.file_uploader("Foto Volontario", type=["jpg","jpeg","png"], key="v_foto")
        
        submitted = st.form_submit_button("💾 SALVA VOLONTARIO", type="primary", use_container_width=True)
        if submitted:
            if nome and cognome and cellulare and capo_odv:
                nuovo = {
                    "Nome": nome, "Cognome": cognome, "Comune": comune, "Via": via,
                    "CapoODV": capo_odv, "ODVAppartenenza": odv_app,
                    "DataNascita": str(data_nascita), "CodFisc": cod_fisc,
                    "Cellulare": cellulare, "Email": email, "TelEmergenza": tel_emerg,
                    "Ruolo": ruolo, "Squadra": squadra,
                    "RadioID": radio_id, "ModelloRadio": modello_radio,
                    "Documento": doc_tipo, "DocNum": doc_num, "ScadDoc": str(doc_scad),
                    "Stato": stato_vol,
                    "DataAgg": datetime.now().strftime("%d/%m/%Y %H:%M")
                }
                st.session_state.volontari.append(nuovo)
                st.success(f"Volontario {cognome} {nome} salvato!")
            else:
                st.error("Compila Nome, Cognome, Cellulare e Capo ODV *")

    if st.session_state.volontari:
        df = pd.DataFrame(st.session_state.volontari)
        st.dataframe(df, use_container_width=True)
        c1, c2 = st.columns(2)
        with c1:
            st.download_button("⬇️ Excel Volontari", to_excel(df), "volontari.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        with c2:
            pdf = to_pdf(df, "VOLONTARI")
            if pdf:
                st.download_button("📄 PDF Volontari", pdf, "volontari.pdf", mime="application/pdf", use_container_width=True)

# ================= DB RADIO =================
elif menu == "DB Radio":
    st.header("📡 DB Radio - Anagrafica Apparati")
    st.caption("Con campo Tipo: DMR Regionali, Tetra, PMR 446, Nautiche, Radio Amatoriale, ecc.")
    
    with st.form("form_db_radio", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            matricola = st.text_input("Matricola *", placeholder="Es: PD785-001", key="db_mat")
            modello = st.selectbox("Modello *", ["Hytera PD785", "Hytera PD785G", "Anytone 878", "Anytone 578", "Motorola DP4400", "Motorola DP4800", "Baofeng UV-5R", "Altro"], key="db_mod")
            if modello == "Altro":
                modello = st.text_input("Specifica Modello", key="db_mod_altro") or modello
            num_serie = st.text_input("Numero Serie", key="db_serie")
        with c2:
            tipo_radio = st.selectbox("Tipo Radio *", TIPI_RADIO, index=0, key="db_tipo", help="Seleziona: DMR Regionali, Tetra, PMR 446, Nautiche, Radio Amatoriale, ecc.")
            if tipo_radio == "Altro":
                tipo_radio = st.text_input("Specifica Tipo", placeholder="Es: DMR Tier III", key="db_tipo_altro") or tipo_radio
            frequenza = st.text_input("Frequenza / Canale", placeholder="Es: 172.350", key="db_freq")
            proprietario = st.selectbox("Proprietario ODV", ["ANA Varese", "ANA Milano", "Protezione Civile", "Altro"], key="db_prop")
        with c3:
            stato_radio = st.selectbox("Stato", ["Operativa", "In riparazione", "Fuori servizio", "Riserva", "Dispersa"], key="db_stato")
            data_acq = st.date_input("Data Acquisto", key="db_data")
            note_radio = st.text_area("Note", key="db_note")
        
        submitted_db = st.form_submit_button("💾 SALVA RADIO IN DB", type="primary", use_container_width=True)
        if submitted_db:
            if matricola and modello:
                st.session_state.radio_db.append({
                    "Matricola": matricola,
                    "Modello": modello,
                    "Tipo": tipo_radio,
                    "NumSerie": num_serie,
                    "Frequenza": frequenza,
                    "Proprietario": proprietario,
                    "Stato": stato_radio,
                    "DataAcquisto": str(data_acq),
                    "Note": note_radio
                })
                st.success(f"Radio {matricola} - {modello} [{tipo_radio}] salvata!")
                st.rerun()
            else:
                st.error("Compila Matricola e Modello *")

    if st.session_state.radio_db:
        st.subheader(f"Elenco Radio ({len(st.session_state.radio_db)})")
        df = pd.DataFrame(st.session_state.radio_db)
        # Filtro per Tipo
        if "Tipo" in df.columns:
            tipi_presenti = ["Tutti"] + sorted(df["Tipo"].dropna().unique().tolist())
            filtro_tipo = st.selectbox("Filtra per Tipo", tipi_presenti, key="filtro_tipo")
            if filtro_tipo != "Tutti":
                df_view = df[df["Tipo"] == filtro_tipo]
            else:
                df_view = df
        else:
            df_view = df
        st.dataframe(df_view, use_container_width=True)
        col1, col2, col3 = st.columns(3)
        with col1:
            st.download_button("⬇️ Excel DB Radio", to_excel(df_view), "db_radio.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        with col2:
            pdf = to_pdf(df_view, "DB RADIO")
            if pdf:
                st.download_button("📄 PDF DB Radio", pdf, "db_radio.pdf", mime="application/pdf", use_container_width=True)
        with col3:
            if st.button("🗑️ Svuota DB Radio", use_container_width=True):
                st.session_state.radio_db = []
                st.rerun()
    else:
        st.info("Nessuna radio in archivio - aggiungi la prima sopra")

# ================= CONSEGNA RADIO =================
elif menu == "Consegna Radio":
    st.header("📻 Consegna Radio - Campi Originali")
    c1, c2 = st.columns(2)
    with c1:
        vol_list = [f"{v.get('Cognome','')} {v.get('Nome','')}" for v in st.session_state.volontari]
        sel_vol = st.selectbox("Volontario", vol_list, key="c_vol") if vol_list else st.text_input("Volontario (manuale)", key="c_vol_man")
        radio_list = [f"{r.get('Matricola','')} - {r.get('Modello','')} [{r.get('Tipo','')}]" for r in st.session_state.radio_db]
        sel_radio = st.selectbox("Radio", radio_list, key="c_radio") if radio_list else st.text_input("Radio manuale", key="c_radio_man")
    with c2:
        data_cons = st.date_input("Data Consegna", value=date.today(), format="DD/MM/YYYY", key="c_data")
        ora_cons = st.time_input("Ora", value=datetime.now().time(), key="c_ora")
        motivo = st.text_input("Motivo / Evento", key="c_motivo")

    if st.button("Registra Consegna", type="primary", use_container_width=True):
        st.session_state.consegna_radio.append({
            "Volontario": sel_vol,
            "Radio": sel_radio,
            "Data": str(data_cons),
            "Ora": str(ora_cons),
            "Motivo": motivo,
            "Stato": "Consegnata"
        })
        st.success("Consegna registrata")
        st.rerun()

    if st.session_state.consegna_radio:
        df = pd.DataFrame(st.session_state.consegna_radio)
        st.dataframe(df, use_container_width=True)
        c1, c2 = st.columns(2)
        with c1:
            st.download_button("⬇️ Excel Consegne", to_excel(df), "consegne.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        with c2:
            pdf = to_pdf(df, "CONSEGNA RADIO")
            if pdf:
                st.download_button("📄 PDF Consegne", pdf, "consegna_radio.pdf", mime="application/pdf", use_container_width=True)

# ================= ALIAS RADIO =================
elif menu == "Alias Radio":
    st.header("🏷️ Alias Radio - Campi Originali")
    c1, c2 = st.columns(2)
    with c1:
        alias_n = st.text_input("Alias *", key="a_alias")
        id_r = st.text_input("ID Radio *", key="a_id")
    with c2:
        gruppo = st.selectbox("Gruppo", ["Squadra A", "Squadra B", "Squadra C", "Coordinamento", "Logistica"], key="a_gruppo")
        desc = st.text_input("Descrizione", key="a_desc")

    if st.button("Salva Alias", type="primary", use_container_width=True):
        if alias_n and id_r:
            st.session_state.alias_radio.append({
                "Alias": alias_n,
                "ID Radio": id_r,
                "Gruppo": gruppo,
                "Descrizione": desc
            })
            st.success("Alias salvato")
            st.rerun()
        else:
            st.error("Compila Alias e ID Radio")

    if st.session_state.alias_radio:
        df = pd.DataFrame(st.session_state.alias_radio)
        st.dataframe(df, use_container_width=True)
        c1, c2 = st.columns(2)
        with c1:
            st.download_button("⬇️ Excel Alias", to_excel(df), "alias.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        with c2:
            pdf = to_pdf(df, "ALIAS RADIO")
            if pdf:
                st.download_button("📄 PDF Alias", pdf, "alias_radio.pdf", mime="application/pdf", use_container_width=True)

# ================= BROGLIACCIO =================
elif menu == "Brogliaccio":
    st.header("📝 Brogliaccio Radio - Campi Originali")
    st.info("Chiamate e Ricevente da Alias Radio | Operatore da Volontari Nome Cognome")
    
    alias_list = list(set([a.get("Alias","") for a in st.session_state.alias_radio if a.get("Alias")])) or ["Centrale Operativa", "Squadra A", "Squadra B"]
    vol_list = [f"{v.get('Nome','')} {v.get('Cognome','')}" for v in st.session_state.volontari] or ["Ezio Fiscato"]

    c1, c2 = st.columns(2)
    with c1:
        data_b = st.date_input("Data", value=date.today(), format="DD/MM/YYYY", key="b_data")
        ora_b = st.time_input("Ora", value=datetime.now().time(), key="b_ora")
        operatore = st.selectbox("Operatore (da Volontari)", vol_list, key="b_op")
        operatore_custom = st.text_input("Oppure Operatore manuale", key="b_op_man")
        if operatore_custom.strip(): operatore = operatore_custom.strip()
        chiamate = st.selectbox("Chiamate (da Alias Radio)", alias_list, key="b_chiam")
        chiamate_custom = st.text_input("Oppure Chiamate manuale", key="b_chiam_man")
        if chiamate_custom.strip(): chiamate = chiamate_custom.strip()
    with c2:
        evento_b = st.text_input("Evento Riferimento", key="b_evento")
        emerg_b = st.text_input("Emergenza Riferimento", key="b_emerg")
        ricevente = st.selectbox("Ricevente (da Alias Radio)", alias_list, key="b_ricev")
        ricevente_custom = st.text_input("Oppure Ricevente manuale", key="b_ricev_man")
        if ricevente_custom.strip(): ricevente = ricevente_custom.strip()
        blindato = st.checkbox("Blinda Evento/Emergenza", key="b_blind")
    
    testo_b = st.text_area("Testo Brogliaccio *", height=150, key="b_testo")

    if st.button("Salva Brogliaccio", type="primary", use_container_width=True):
        if testo_b:
            st.session_state.brogliaccio.append({
                "Data": str(data_b), "Ora": str(ora_b),
                "Operatore": operatore, "Chiamate": chiamate, "Ricevente": ricevente,
                "Evento": evento_b, "Emergenza": emerg_b,
                "Testo": testo_b, "Blindato": blindato
            })
            st.success(f"Brogliaccio salvato - {operatore} - {chiamate} -> {ricevente}")
            st.rerun()
        else:
            st.error("Compila Testo Brogliaccio *")

    if st.session_state.brogliaccio:
        df = pd.DataFrame(st.session_state.brogliaccio)
        st.dataframe(df, use_container_width=True)
        cols_show = [c for c in ["Data","Ora","Operatore","Chiamate","Ricevente","Testo","Evento"] if c in df.columns]
        if cols_show:
            st.markdown("**Vista sintetica:**")
            st.dataframe(df[cols_show], use_container_width=True)
        c1, c2 = st.columns(2)
        with c1:
            st.download_button("⬇️ Excel Brogliaccio", to_excel(df), "brogliaccio.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        with c2:
            pdf = to_pdf(df, "BROGLIACCIO RADIO")
            if pdf:
                st.download_button("📄 PDF Brogliaccio", pdf, "brogliaccio.pdf", mime="application/pdf", use_container_width=True)
