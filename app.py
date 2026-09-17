import streamlit as st
from datetime import date, datetime
from decimal import Decimal, InvalidOperation

st.set_page_config(page_title="Generatore RDA Olivetti", page_icon="📄", layout="wide")

SOGLIA_NUOVA_PROCEDURA = date(2026, 4, 15)
SOGLIA_3A = Decimal("20000.00")

CASISTICHE_TD = [
    "1. Particolarità tecnologiche/infrastrutturali",
    "2. Acquisti in condizioni di emergenza adeguatamente motivati",
    "3. Consulenze/prestazioni professionali di particolare specializzazione, integrazione o complementarità",
    "4. Acquisti d'opportunità a condizioni estremamente vantaggiose e non ripetibili",
    "5. Gara andata deserta",
    "6. Prezzi imposti da autorità di settore o accordi che impediscono la competizione",
    "7. Estensione di condizioni/prezzi definiti in competizione",
    "8. Progetti business/tecnologici con fabbisogni puntuali e fornitori identificati",
    "9. Commercializzazione di prodotti/soluzioni non standard per clienti privati",
    "10. Opportunità di business per commercializzazione verso la PA",
]

def euro_to_decimal(value: str) -> Decimal:
    s = (value or "").strip().replace("€", "").replace(" ", "")
    if not s:
        return Decimal("0")
    # Accetta 23.200,50 oppure 23200.50
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return Decimal(s)
    except InvalidOperation:
        return Decimal("0")

def calcola_documenti(data_rda: date, totale: Decimal, deroga: bool):
    nuova = data_rda >= SOGLIA_NUOVA_PROCEDURA
    docs = []
    if nuova:
        docs += [
            ("Allegato 6 - Trattamento dati", True, "Sempre previsto"),
            ("Allegato 4 - Requisiti sicurezza", False, "Sempre previsto"),
            ("Allegato 7 - Razionali RDA", True, "Sempre previsto"),
        ]
        if deroga:
            docs.append(("Offerta", False, "Prevista in deroga"))
        if totale > SOGLIA_3A:
            docs.append(("Allegato 3A - Trattativa Diretta", True, "Vademecum: totale > 20.000 €"))
    else:
        docs += [
            ("Allegato 6 - Trattamento dati", True, "Sempre previsto"),
            ("Allegato 1 - Requisiti sicurezza (vecchia procedura)", False, "Sempre previsto"),
            ("Allegato 5 - Scheda motivazionale", True, "Sempre previsto"),
            ("Offerta", False, "Sempre prevista"),
        ]
        if totale > SOGLIA_3A:
            docs.append(("Allegato 3A - Trattativa Diretta", True, "Vademecum: totale > 20.000 €"))
    return nuova, docs

def stato_campo(nome, valore):
    return None if valore not in (None, "", []) else nome

st.title("Generatore documentazione RDA Olivetti")
st.caption("Versione 0.1 — motore regole + raccolta dati + anteprima. Nessun dato viene inventato.")

with st.sidebar:
    st.header("Modelli")
    st.info(
        "Nella prossima fase collegheremo i modelli PDF/DOCX reali. "
        "Questa versione decide quali documenti servono e raccoglie i dati necessari."
    )

st.subheader("1. Dati della RDA")
c1, c2, c3 = st.columns(3)
with c1:
    fornitore = st.text_input("Intestatario / Ragione sociale")
    codice_sap = st.text_input("Codice SAP fornitore")
    numero_rda = st.text_input("Numero RDA")
with c2:
    data_rda = st.date_input("Data RDA", value=date.today())
    commessa = st.text_input("Rif. Commessa")
    cliente = st.text_input("Cliente / Gara")
with c3:
    totale_testo = st.text_input("Totale RDA", placeholder="es. 23.200,00")
    oggetto = st.text_area("Oggetto", height=115)

totale = euro_to_decimal(totale_testo)

st.subheader("2. Informazioni decisionali")
c1, c2, c3 = st.columns(3)
with c1:
    deroga = st.radio("RDA in deroga / Trattativa Diretta?", ["No", "Sì"], horizontal=True) == "Sì"
with c2:
    tratta_dati = st.radio("La fornitura tratta dati personali?", ["No", "Sì", "Da verificare"], horizontal=True)
with c3:
    funzione_richiedente = st.text_input("Funzione richiedente", placeholder="es. P&D ...")

nuova, documenti = calcola_documenti(data_rda, totale, deroga)
serve_3a = any("3A" in d[0] for d in documenti)

st.subheader("3. Dati aggiuntivi richiesti solo quando servono")
piva = ""
periodo_inizio = None
periodo_fine = None
casistica_td = ""
descrizione_esigenza = ""
tecnologia = ""
continuita = "Da verificare"
fornitore_impegnato = "Da verificare"
stato_fornitura = "Da verificare"

if serve_3a:
    st.warning("Il vademecum richiede l'Allegato 3A perché il totale supera 20.000 €.")
    a, b = st.columns(2)
    with a:
        piva = st.text_input("P.IVA fornitore *")
        casistica_td = st.selectbox("Casistica Trattativa Diretta *", [""] + CASISTICHE_TD)
        periodo_inizio = st.date_input("Inizio fornitura", value=data_rda)
        continuita = st.selectbox("Attività in continuità con lo stesso fornitore?", ["Da verificare", "No", "Sì"])
    with b:
        periodo_fine = st.date_input("Fine fornitura", value=data_rda)
        tecnologia = st.text_input("Tecnologia (se applicabile)")
        fornitore_impegnato = st.selectbox("Fornitore già impegnato?", ["Da verificare", "No", "Sì"])
        stato_fornitura = st.selectbox("Stato della fornitura", ["Da verificare", "Non iniziata", "In corso", "Completata"])
    descrizione_esigenza = st.text_area(
        "Descrizione dettagliata dell'esigenza / motivazione TD *",
        placeholder="Descrivere in modo chiaro e circostanziato. Il programma non inventerà questa motivazione.",
        height=120,
    )
else:
    periodo_inizio = st.date_input("Inizio fornitura (se noto)", value=data_rda)
    periodo_fine = st.date_input("Fine fornitura (se noto)", value=data_rda)

if tratta_dati == "Sì":
    st.markdown("**Dettagli Allegato 6**")
    tipo_dati = st.multiselect(
        "Dati personali trattati",
        ["Dati comuni", "Dati particolari", "Videosorveglianza", "Dati di traffico", "Geolocalizzazione", "Altro"],
    )
    interessati = st.multiselect(
        "Soggetti interessati",
        ["Clienti TIM/SdG", "Dipendenti TIM/SdG", "Fornitori/consulenti", "Visitatori", "Clienti/Dipendenti di Clienti", "Altro"],
    )
    tipo_trattamento = st.multiselect(
        "Tipi di trattamento",
        ["Prestazioni professionali/consulenziali", "Servizi amministrativi", "Gestione documentale",
         "Logistica/reception/vigilanza", "Servizi IT", "Servizi tecnici", "Altro"],
    )
else:
    tipo_dati, interessati, tipo_trattamento = [], [], []

st.divider()
st.subheader("4. Documentazione calcolata")
st.write(
    f"Procedura applicata: **{'NUOVA (dal 15/04/2026)' if nuova else 'PRECEDENTE al 15/04/2026'}** "
    f"— Totale interpretato: **€ {totale:,.2f}**".replace(",", "X").replace(".", ",").replace("X", ".")
)

for nome, firma, motivo in documenti:
    st.success(f"✓ {nome} — {'DA FIRMARE' if firma else 'NO FIRMA'} — {motivo}")

# Controlli bloccanti / warning
mancanti = []
for nome, valore in [
    ("Ragione sociale", fornitore),
    ("Codice SAP", codice_sap),
    ("Numero RDA", numero_rda),
    ("Cliente/Gara", cliente),
    ("Oggetto", oggetto),
    ("Funzione richiedente", funzione_richiedente),
]:
    x = stato_campo(nome, valore)
    if x:
        mancanti.append(x)

if totale <= 0:
    mancanti.append("Totale RDA valido")

if serve_3a:
    for nome, valore in [
        ("P.IVA fornitore", piva),
        ("Casistica TD", casistica_td),
        ("Descrizione/motivazione TD", descrizione_esigenza),
    ]:
        x = stato_campo(nome, valore)
        if x:
            mancanti.append(x)

if tratta_dati == "Da verificare":
    mancanti.append("Verifica trattamento dati personali")
elif tratta_dati == "Sì" and (not tipo_dati or not interessati or not tipo_trattamento):
    mancanti.append("Dettagli trattamento dati personali")

st.subheader("5. Controllo prima della generazione")
if mancanti:
    st.error("Campi da completare prima della generazione: " + " • ".join(mancanti))
else:
    st.success("Controlli principali superati. I dati sono pronti per il motore di compilazione.")

with st.expander("Anteprima dati strutturati"):
    st.json({
        "fornitore": fornitore,
        "codice_sap": codice_sap,
        "numero_rda": numero_rda,
        "data_rda": str(data_rda),
        "commessa": commessa,
        "cliente_gara": cliente,
        "totale": str(totale),
        "oggetto": oggetto,
        "deroga_td": deroga,
        "tratta_dati_personali": tratta_dati,
        "funzione_richiedente": funzione_richiedente,
        "piva": piva,
        "periodo_inizio": str(periodo_inizio) if periodo_inizio else "",
        "periodo_fine": str(periodo_fine) if periodo_fine else "",
        "casistica_td": casistica_td,
        "descrizione_esigenza": descrizione_esigenza,
        "tecnologia": tecnologia,
        "continuita": continuita,
        "fornitore_impegnato": fornitore_impegnato,
        "stato_fornitura": stato_fornitura,
        "documenti_da_generare": [d[0] for d in documenti],
    })

if st.button("GENERA DOCUMENTI", type="primary", disabled=bool(mancanti)):
    st.info(
        "Motore di compilazione documenti non ancora collegato nella v0.1. "
        "La logica e i controlli sono pronti; il prossimo modulo compilerà i modelli PDF/DOCX."
    )
