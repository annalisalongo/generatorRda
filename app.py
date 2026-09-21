from pathlib import Path
import io, re, shutil, tempfile, zipfile
from datetime import datetime
import streamlit as st
from docx import Document
from pypdf import PdfReader
from openpyxl import load_workbook

BASE = Path(__file__).parent
T = BASE / 'templates'

ALL_DOCS = ['1','2','3A','4','6','7','10','11','12','13','14','15','16','OFFERTA']

CASI_TD = [
'1 - Particolarità tecnologiche/infrastrutturali / continuità',
'2 - Emergenza adeguatamente motivata',
'3 - Consulenze/prestazioni professionali specialistiche o complementari',
'4 - Acquisto di opportunità','5 - Gara andata deserta',
'6 - Prezzi imposti / cessione azienda / impossibilità oggettiva di competizione',
'7 - Estensione condizioni/prezzi di competizione',
'8 - Progetto business/tecnologico identificato con fabbisogni e fornitori puntuali',
'9 - Commercializzazione prodotti/soluzioni non standard a clienti privati',
'10 - Opportunità business per commercializzazione verso PA']

SPECIALI = [
'1 - Scelte tecnologiche/business di Gruppo da unico fornitore',
'2 - Diritti/sponsorizzazioni/collaboratori artistici unici',
'3 - Manutenzione applicativi proprietari / rete non TPM','4 - Partnership commerciali',
'5 - Canali del Costruttore certificati / prezzi Procurement',
'6 - Vendor tecnologico/distributore per gare PA con tecnologia esplicita ed esclusività',
'7 - Nuove tecnologie / primo impiego / sperimentazioni','8 - Postazioni alta frequenza',
'9 - Prodotti di rete per ampliamento reti esistenti','10 - Adesione ad Accordi Quadro di Gruppo',
'11 - Networking ICT con accordi/listini PR >=80% commessa','12 - Leasing individuato da Enterprise/CF.F']


def euro(s):
    if isinstance(s, (int,float)): return float(s)
    s = str(s or '').strip().replace('EUR','').replace('€','').replace('\xa0','').replace(' ','')
    if not s: return 0.0
    if ',' in s:
        s = s.replace('.','').replace(',','.')
    try: return float(s)
    except: return 0.0


def fmt_eur(x):
    return f"€ {float(x):,.2f}".replace(',', 'X').replace('.', ',').replace('X','.')


def pdf_text(upload):
    if not upload: return ''
    try:
        reader = PdfReader(io.BytesIO(upload.getvalue()))
        return '\n'.join((p.extract_text() or '') for p in reader.pages)
    except Exception as e:
        st.warning(f'Non riesco a leggere automaticamente {upload.name}: {e}')
        return ''


def first(patterns, text, flags=re.I|re.S):
    for pat in patterns:
        m = re.search(pat, text, flags)
        if m:
            return re.sub(r'\s+', ' ', m.group(1)).strip(' :-\n\t')
    return ''


def normalize_date(s):
    s = (s or '').strip()
    for fmt in ('%d/%m/%Y','%d-%m-%Y','%d.%m.%Y','%d %B %Y','%d %b %Y'):
        try: return datetime.strptime(s,fmt).strftime('%d/%m/%Y')
        except: pass
    # English month names from supplier quotations
    for fmt in ('%d %B %Y','%d %b %Y'):
        try: return datetime.strptime(s,fmt).strftime('%d/%m/%Y')
        except: pass
    return s


def extract_offer(text):
    d = {}
    d['offerta_numero'] = first([r'(?:Numero offerta|RIFERIMENTO OFFERTA)\s*[:\-]?\s*([A-Z0-9_./-]+)'], text)
    d['offerta_data'] = normalize_date(first([r'\b(\d{1,2}\s+[A-Za-z]+\s+20\d{2})\b', r'\b(\d{1,2}/\d{1,2}/20\d{2})\b'], text))
    d['fornitore'] = first([r'OFFERTA\s*\n\s*([^\n]+?)(?=\n)', r'\n([A-Z][A-Z0-9 .&\'-]{3,})\s*\n[^\n]*(?:Belgium|Italia|Italy)'], text)
    d['piva'] = first([r'\b(IT\d{11})\b'], text)
    d['oggetto'] = first([r'Nome scheda\s+([^\n]+)', r'Servizio\s+([^\n]+)'], text)
    d['progetto'] = first([r'Riferimento progetto\s+([^\n]+)'], text)
    d['richiedente'] = first([r'Riferimento d.acquisto\s+([^\n]+)'], text)
    d['totale'] = euro(first([r'PREZZO LORDO\s*(?:EUR|€)?\s*([\d.,]+)', r'PREZZO NETTO\s*(?:EUR|€)?\s*([\d.,]+)'], text))
    d['consegna'] = first([r'TEMPO DI CONSEGNA\s*([0-9]+\s*(?:WD|giorni[^\n]*))'], text)
    return d


def extract_rda(text):
    d = {}
    d['data'] = normalize_date(first([r'Data Richiesta\s+Numero Commessa\s+(\d{1,2}/\d{1,2}/20\d{2})', r'Data Richiesta\s+(\d{1,2}/\d{1,2}/20\d{2})'], text))
    d['commessa'] = first([r'Data Richiesta\s+Numero Commessa\s+\d{1,2}/\d{1,2}/20\d{2}\s+([^\n]+)', r'Numero Commessa\s+([^\n]+)'], text)
    d['richiedente'] = first([r'Richiedente.*?\n\s*([^\n]+?)\s*\nRiferimento', r'Richiedente\s+Destinazione merce.*?\n\s*([^\n]+)'], text)
    d['offerta_ref'] = first([r'(Offerta\s+[A-Z0-9_./-]+\s+del\s+\d{1,2}/\d{1,2}/20\d{2})'], text)
    # Riga acquisto: posizione, quantità, fornitore, descrizione, costo. Heuristics + fallback.
    m = re.search(r'\n\s*1\s+(\d+)\s+([^\n]+?)\s+([A-Za-z0-9_./-]{4,})\s+([\d.,]+)\s*(?:\n|$)', text)
    if m:
        d['quantita'] = m.group(1); d['fornitore'] = m.group(2).strip(); d['oggetto'] = m.group(3); d['totale'] = euro(m.group(4))
    else:
        d['oggetto'] = first([r'\b(HW_[A-Za-z0-9_./-]+)\b'], text)
        d['totale'] = euro(first([r'\b(\d{3,}[.,]?\d*)\s*\n2\s*\n3\s*\n4'], text))
    return d


def extract_excel(upload, row_number=None):
    if not upload: return {}, []
    try:
        wb = load_workbook(io.BytesIO(upload.getvalue()), data_only=True, read_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        if not rows: return {}, []
        headers = [str(x or '').strip() for x in rows[0]]
        data_rows = rows[1:]
        if not data_rows: return {}, headers
        idx = max(0, min((row_number or 2)-2, len(data_rows)-1))
        vals = data_rows[idx]
        raw = {headers[i]: vals[i] for i in range(min(len(headers),len(vals))) if headers[i]}
        return raw, headers
    except Exception as e:
        st.warning(f'Non riesco a leggere il file Excel: {e}')
        return {}, []


def excel_map(raw):
    out = {}
    for k,v in raw.items():
        lk = k.lower()
        if 'rda' in lk and ('n.' in lk or 'numero' in lk or lk.strip()=='n. rda'): out['rda'] = str(v or '').strip()
        elif 'sap' in lk and 'forn' in lk: out['sap'] = str(v or '').strip()
        elif 'ragione' in lk or ('fornitore' in lk and 'cod' not in lk): out['fornitore'] = str(v or '').strip()
        elif 'totale' in lk: out['totale'] = euro(v)
        elif 'commessa' in lk: out['commessa'] = str(v or '').strip()
        elif 'cliente' in lk or 'gara' in lk: out['cliente'] = str(v or '').strip()
        elif 'oggetto' in lk: out['oggetto'] = str(v or '').strip()
        elif 'data rda' in lk: out['data'] = str(v or '').strip()
    return out


def merge_detected(offer, rda, xls):
    # Precedenza: Excel per dati gestionali, RDA per richiesta interna, offerta per dati economici/fornitore precisi.
    d = {}
    for src in (offer, rda, xls):
        for k,v in src.items():
            if v not in ('',None,0,0.0): d[k]=v
    # eccezioni utili
    if offer.get('totale'): d['totale'] = offer['totale']
    if offer.get('fornitore'): d['fornitore'] = offer['fornitore']
    if rda.get('data'): d['data'] = rda['data']
    if xls.get('rda'): d['rda'] = xls['rda']
    return d


def decide(d):
    docs=['4','6']
    reasons=['Allegato 4: requisiti sicurezza/compliance','Allegato 6: trattamento dati']
    if not d.get('nuova_tranche_senza_variazione') and not d.get('precontrattuale'):
        docs.append('7'); reasons.append('Allegato 7: razionali fabbisogno/prezzo')
    if d.get('deroga') or d.get('side_letter'):
        docs.append('OFFERTA'); reasons.append('Offerta: deroga/Side Letter')
    if float(d.get('totale') or 0) > 20000:
        docs.append('3A'); reasons.append('Allegato 3A: importo > €20.000 (regola operativa impostata)')
    if d.get('gara_prest_prof'): docs.append('13')
    if d.get('saas'): docs.append('14')
    if d.get('cliente_tipo')=='Privato' and d.get('trattamento_cliente'): docs.append('11')
    if d.get('cliente_tipo')=='PA' and d.get('trattamento_cliente'): docs.append('12')
    if d.get('amministratore_sistema'): docs.append('15')
    if d.get('trattamento_dati'): docs.append('10')
    if d.get('attestazione_conformita'): docs.append('16')
    return list(dict.fromkeys(docs)), reasons


def xml_replace(src,dst,repls,occurrence_repls=None):
    from lxml import etree
    occurrence_repls=occurrence_repls or []; counts={}
    with zipfile.ZipFile(src,'r') as zin, zipfile.ZipFile(dst,'w',zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data=zin.read(item.filename)
            if item.filename.endswith('.xml'):
                try:
                    root=etree.fromstring(data)
                    for node in root.xpath('//*[local-name()="t"]'):
                        if node.text is None: continue
                        text=node.text
                        for a,b in repls.items(): text=text.replace(a,str(b)) if a in text else text
                        for old,new,n in occurrence_repls:
                            if old in text:
                                counts[old]=counts.get(old,0)+1
                                if counts[old]==n: text=text.replace(old,str(new),1)
                        node.text=text
                    data=etree.tostring(root,xml_declaration=True,encoding='UTF-8',standalone=True)
                except Exception: pass
            zout.writestr(item,data)


def make7(d,out):
    repl={
      '21/09/2026':d.get('data',''), 'OLIVETTI - P&DI CPP':d.get('funzione',''),
      'ANTONIO SCUCCIMARRA':d.get('fornitore',''),
      'Nota Spese Giugno, CB Centro Sud Puglia - Acque del Sud':d.get('oggetto',''),
      'supporto per attività commerciale propedeutica alla riuscita del Progetto Water Management System gestito da Portfolio IoT':d.get('motivazione',''),
      '216,40':fmt_eur(d.get('totale',0)).replace('€ ','')}
    xml_replace(T/'allegato7.docx',out,repl)


def make6(d,out):
    repl={'2101416996':d.get('rda','')}
    occ=[('Scrivere qui',d.get('riferimenti',''),1),('Scrivere qui',d.get('fornitore',''),2)]
    xml_replace(T/'allegato6.docx',out,repl,occ)


def make3a(d,out):
    shutil.copy2(T/'allegato3.docx',out)
    doc=Document(out); t=doc.tables[1]
    for p in doc.paragraphs:
        if p.text.strip()=='Luogo, gg/mm/aaaa': p.text='Luogo Roma, '+d.get('data',''); break
    t.cell(2,0).text='Fornitore proposto:\n'+d.get('fornitore','')
    t.cell(2,1).text='Cod. Sap: '+d.get('sap',''); t.cell(2,3).text='P.IVA: '+d.get('piva','')
    t.cell(3,0).text='Importo TD: '+fmt_eur(d.get('totale',0)); t.cell(3,3).text='N° RDA (se emessa): '+d.get('rda','')
    if d.get('casistica_td'): t.cell(6,0).text="Casistica Allegato 1: "+d['casistica_td']
    if d.get('motivazione'): t.cell(7,0).text='Descrizione dettagliata:\n'+d['motivazione']
    if d.get('cliente'): t.cell(17,0).text='Cliente destinatario della fornitura: '+d['cliente']
    doc.save(out)


def package(d,docs,offer_bytes=None,offer_name=None):
    td=Path(tempfile.mkdtemp()); produced=[]
    def cp(src,name):
        p=td/name; shutil.copy2(src,p); produced.append(p)
    if '4' in docs: cp(T/'allegato4.docx',f"RDA {d.get('rda','')} - Allegato 4.docx")
    if '6' in docs:
        p=td/f"RDA {d.get('rda','')} - Allegato 6.docx"; make6(d,p); produced.append(p)
    if '7' in docs:
        p=td/f"RDA {d.get('rda','')} - Allegato 7.docx"; make7(d,p); produced.append(p)
    if '3A' in docs:
        p=td/f"RDA {d.get('rda','')} - Allegato 3A - DA VERIFICARE.docx"; make3a(d,p); produced.append(p)
    for n in ['1','2','10','11','12','13','15','16']:
        if n in docs: cp(T/f'allegato{n}.docx',f"RDA {d.get('rda','')} - Allegato {n}.docx")
    if '14' in docs: cp(T/'allegato14.xlsx',f"RDA {d.get('rda','')} - Allegato 14.xlsx")
    if 'OFFERTA' in docs:
        if offer_bytes:
            p=td/(offer_name or 'Offerta.pdf'); p.write_bytes(offer_bytes); produced.append(p)
        else:
            p=td/'OFFERTA_MANCANTE.txt'; p.write_text('Offerta non caricata.',encoding='utf-8'); produced.append(p)
    z=io.BytesIO()
    with zipfile.ZipFile(z,'w',zipfile.ZIP_DEFLATED) as zz:
        for p in produced: zz.write(p,p.name)
    return z.getvalue()


st.set_page_config(page_title='Generatore RDA Olivetti v0.5',layout='wide')
st.title('Generatore RDA Olivetti — v0.5')
st.caption('Carica Offerta + Richiesta RDA + riga Excel. L’app estrae i dati, te li fa verificare e poi genera gli allegati scelti.')

st.subheader('1. Carica i documenti di partenza')
a,b,c=st.columns(3)
with a: offer_file=st.file_uploader('OFFERTA fornitore',type=['pdf'],key='offer')
with b: rda_file=st.file_uploader('RICHIESTA RDA',type=['pdf'],key='rda_pdf')
with c:
    excel_file=st.file_uploader('FILE EXCEL (opzionale)',type=['xlsx','xlsm'],key='xls')
    excel_row=st.number_input('Numero riga Excel da leggere',min_value=2,value=2,step=1,disabled=not bool(excel_file))

offer_data=extract_offer(pdf_text(offer_file)) if offer_file else {}
rda_data=extract_rda(pdf_text(rda_file)) if rda_file else {}
raw_xls,_=extract_excel(excel_file,int(excel_row)) if excel_file else ({},[])
xls_data=excel_map(raw_xls)
detected=merge_detected(offer_data,rda_data,xls_data)

if offer_file or rda_file or excel_file:
    st.subheader('2. Dati trovati automaticamente — controlla/correggi')
    st.caption('I campi vuoti non sono stati trovati con sufficiente affidabilità: compilali tu. Il programma non inventa dati mancanti.')
    c1,c2,c3=st.columns(3)
    with c1:
        rda=st.text_input('N. RDA',value=str(detected.get('rda','')))
        fornitore=st.text_input('Fornitore',value=str(detected.get('fornitore','')))
        sap=st.text_input('Codice SAP fornitore',value=str(detected.get('sap','')))
        piva=st.text_input('P.IVA fornitore',value=str(detected.get('piva','')))
        totale=euro(st.text_input('Totale RDA / Offerta',value=fmt_eur(detected.get('totale',0)).replace('€ ','')))
    with c2:
        data=st.text_input('Data richiesta',value=str(detected.get('data','')))
        commessa=st.text_input('Commessa',value=str(detected.get('commessa','')))
        cliente=st.text_input('Cliente / Gara / Progetto',value=str(detected.get('cliente') or detected.get('progetto','')))
        richiedente=st.text_input('Richiedente',value=str(detected.get('richiedente','')))
        funzione=st.text_input('Funzione richiedente',value='OLIVETTI - P&DI CPP')
    with c3:
        oggetto_default=detected.get('oggetto','')
        if rda_data.get('quantita') and oggetto_default: oggetto_default=f"n. {rda_data['quantita']} {oggetto_default}"
        oggetto=st.text_area('Oggetto',value=str(oggetto_default))
        motivazione=st.text_area('Motivazione / descrizione esigenza',value='')
        riferimenti=st.text_input('Riferimenti (SDW/SH/PCD, se applicabile)')
        consegna=st.text_input('Consegna rilevata',value=str(detected.get('consegna','')))

    with st.expander('Mostra cosa è stato letto dai tre input'):
        st.write('**Offerta:**', offer_data or 'nessun dato')
        st.write('**Richiesta RDA:**', rda_data or 'nessun dato')
        st.write('**Riga Excel:**', raw_xls or 'nessun dato')

    st.subheader('3. Scegli come determinare gli allegati')
    mode=st.radio('Modalità', ['Li indico io','Determina automaticamente'],horizontal=True)

    # defaults shared by both modes
    flags={k:False for k in ['nuova_tranche_senza_variazione','precontrattuale','deroga','side_letter','gara_prest_prof','saas','trattamento_cliente','amministratore_sistema','trattamento_dati','attestazione_conformita']}
    cliente_tipo='n.a.'; casistica_td=''

    if mode=='Li indico io':
        default_docs=['4','6','7'] + (['3A'] if totale>20000 else [])
        selected_docs=st.multiselect('Allegati da generare',ALL_DOCS,default=default_docs)
        docs=selected_docs
        st.caption('In questa modalità la tua scelta prevale sul motore decisionale.')
        if '3A' in docs:
            casistica_td=st.selectbox('Casistica TD per il 3A (se applicabile)',['']+CASI_TD)
    else:
        q1,q2,q3=st.columns(3)
        with q1:
            flags['deroga']=st.checkbox('Acquisto in deroga')
            flags['side_letter']=st.checkbox('Side Letter')
            flags['nuova_tranche_senza_variazione']=st.checkbox('Nuova tranche senza variazione valore')
            flags['precontrattuale']=st.checkbox('Ingaggio precontrattuale')
        with q2:
            td=st.checkbox('Trattativa Diretta')
            speciale=st.checkbox('Acquisto Speciale')
            if td and not speciale: casistica_td=st.selectbox('Casistica TD',['']+CASI_TD)
            if speciale: st.selectbox('Casistica Acquisto Speciale',['']+SPECIALI)
            flags['gara_prest_prof']=st.checkbox('Gara prestazioni professionali')
        with q3:
            flags['trattamento_dati']=st.checkbox('Tratta dati personali')
            flags['trattamento_cliente']=st.checkbox('Tratta dati del Cliente')
            cliente_tipo=st.selectbox('Tipo cliente',['n.a.','Privato','PA'])
            flags['saas']=st.checkbox('SaaS')
            flags['amministratore_sistema']=st.checkbox('Amministratore di Sistema')
            flags['attestazione_conformita']=st.checkbox('Attestazione conformità privacy')
        temp={'totale':totale,'cliente_tipo':cliente_tipo,**flags}
        docs,reasons=decide(temp)
        st.write('**Proposta automatica:** '+', '.join(docs))
        for reason in reasons: st.caption('• '+reason)

    st.subheader('4. Controllo finale')
    missing=[]
    for label,val in [('N. RDA',rda),('Fornitore',fornitore),('Oggetto',oggetto)]:
        if not val: missing.append(label)
    if '3A' in docs and totale<=20000:
        st.warning('Hai selezionato manualmente il 3A con importo non superiore a €20.000: verrà comunque generato perché in modalità manuale la tua scelta prevale.')
    if missing: st.warning('Da completare prima della generazione: '+', '.join(missing))
    st.write('**Pacchetto:** '+(', '.join(docs) if docs else 'nessun allegato'))

    d=dict(rda=rda,fornitore=fornitore,sap=sap,piva=piva,totale=totale,data=data,commessa=commessa,
           cliente=cliente,richiedente=richiedente,funzione=funzione,oggetto=oggetto,motivazione=motivazione,
           riferimenti=riferimenti,consegna=consegna,casistica_td=casistica_td,cliente_tipo=cliente_tipo,**flags)

    if st.button('GENERA PACCHETTO RDA',type='primary',disabled=bool(missing or not docs)):
        z=package(d,docs,offer_file.getvalue() if offer_file else None,offer_file.name if offer_file else None)
        st.download_button('Scarica ZIP RDA',z,file_name=f"RDA_{rda}_allegati.zip",mime='application/zip')
else:
    st.info('Carica almeno uno dei documenti per iniziare. Il flusso consigliato è: Offerta + Richiesta RDA + file Excel.')
