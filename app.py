from pathlib import Path
import io, re, shutil, tempfile, zipfile
from datetime import datetime
import streamlit as st
from docx import Document
from pypdf import PdfReader

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


def parse_pasted_excel_row(text):
    """Legge una singola riga copiata da Excel (celle separate da TAB).
    Supporta la Bibbia RDA sia con sia senza la colonna 'RDA Gestita da F / C'.
    """
    text = str(text or '').strip('\r\n')
    if not text.strip():
        return {}, []
    # Excel copia le celle con TAB; conserva anche le celle vuote.
    lines = [ln for ln in text.splitlines() if ln.strip()]
    if not lines:
        return {}, []
    # Se l'utente ha incollato intestazione + riga, usa l'intestazione reale.
    if len(lines) >= 2 and ('ragione sociale' in lines[0].lower() or 'n. rda' in lines[0].lower()):
        headers = lines[0].split('\t')
        vals = lines[1].split('\t')
    else:
        vals = lines[0].split('\t')
        headers14 = ['Intestatario - Ragione sociale','COD. SAP FORNITORE','Allegati alla RDA','n. RDA','DATA RDA','n. ODA','DATA ODA','EM','DATA EM','Economics OK','Rif. Commessa','Cliente / Gara','Totale','OGGETTO']
        headers15 = ['Intestatario - Ragione sociale','COD. SAP FORNITORE','Allegati alla RDA','n. RDA','DATA RDA','RDA Gestita da F / C','n. ODA','DATA ODA','EM','DATA EM','Economics OK','Rif. Commessa','Cliente / Gara','Totale','OGGETTO']
        headers = headers15 if len(vals) >= 15 else headers14
    # Non comprimere celle vuote: la posizione delle colonne è significativa.
    raw = {str(headers[i]).strip(): vals[i].strip() if i < len(vals) else '' for i in range(min(len(headers), len(vals)))}
    extras = vals[len(headers):] if len(vals) > len(headers) else []
    return raw, extras


def pasted_row_map(raw):
    out = {}
    for k,v in raw.items():
        lk = k.lower().strip()
        sv = str(v or '').strip()
        if 'allegati' in lk and 'rda' in lk:
            out['allegati_rda'] = sv
        elif ('n. rda' in lk or 'numero rda' in lk) and 'data' not in lk:
            out['rda'] = sv
        elif 'sap' in lk and 'forn' in lk:
            out['sap'] = sv
        elif 'ragione' in lk or (lk == 'fornitore'):
            out['fornitore'] = sv
        elif 'totale' in lk:
            out['totale'] = euro(sv)
        elif 'commessa' in lk:
            out['commessa'] = sv
        elif 'cliente' in lk or 'gara' in lk:
            out['cliente'] = sv
        elif 'oggetto' in lk:
            out['oggetto'] = sv
        elif 'data rda' in lk:
            out['data'] = sv
    return out


def docs_from_attachment_text(text):
    """Converte la dicitura della colonna 'Allegati alla RDA' nei codici dell'app."""
    original = str(text or '').strip()
    t = original.lower()
    docs=[]
    # 3A prima del generico 3.
    if re.search(r'\b(?:all(?:egato)?\.?\s*)?3\s*a\b', t): docs.append('3A')
    for n in ['1','2','4','6','7','10','11','12','13','14','15','16']:
        if re.search(rf'\b(?:all(?:egato)?\.?\s*){n}\b', t): docs.append(n)
    # Dicitura operativa usata nella Bibbia.
    if re.search(r'req(?:uisiti)?\.?\s*(?:di\s*)?sicurezza|requisiti\s+sicurezza|sicurezza\s*(?:ict)?', t):
        if '4' not in docs: docs.append('4')
    if re.search(r'\bofferta\b', t): docs.append('OFFERTA')
    return list(dict.fromkeys(docs))

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



def _clean_item(oggetto):
    return re.sub(r'^\s*n\.\s*\d+\s+', '', str(oggetto or ''), flags=re.I).strip()


def _qty_from_object(oggetto, fallback=''):
    m = re.match(r'^\s*n\.\s*(\d+)\s+', str(oggetto or ''), flags=re.I)
    return m.group(1) if m else str(fallback or '').strip()


def classify_purchase(oggetto):
    t = str(oggetto or '').lower()
    if any(k in t for k in ['licenz', 'sql server', 'windows', 'autodesk', 'software', 'subscription', 'cal']):
        return 'licenze'
    if any(k in t for k in ['servizio', 'consulenz', 'supporto', 'prestaz', 'giornata uomo', 'nota spese', 'attività profession']):
        return 'servizi'
    return 'hardware'


def clean_all7_object(oggetto, qty=''):
    """Pulisce l'oggetto per l'All.7: una sola quantità e niente riferimenti economici/offerta."""
    t = re.sub(r'\s+', ' ', str(oggetto or '')).strip()
    # Rimuove quantità ripetute iniziali: n. 3 N. 3 ... -> contenuto
    while re.match(r'^\s*n\.\s*\d+\s+', t, re.I):
        t = re.sub(r'^\s*n\.\s*\d+\s+', '', t, count=1, flags=re.I).strip()
    # Rimuove riferimenti all'offerta dall'oggetto: vanno nel razionale economico.
    t = re.sub(r'\s*[-–—]?\s*Offerta\s+[A-Z0-9_./-]+(?:\s+del\s+\d{1,2}/\d{1,2}/20\d{2})?\s*$', '', t, flags=re.I).strip(' -–—')
    q = str(qty or '').strip()
    return (f"n. {q} {t}" if q and t else t)


def generate_all7_sections(d, offer_data=None, rda_data=None):
    """Genera separatamente fabbisogno e razionale prezzo per l'Allegato 7."""
    offer_data = offer_data or {}
    rda_data = rda_data or {}
    raw = str(d.get('oggetto') or '').strip()
    qty = _qty_from_object(raw, rda_data.get('quantita',''))
    item = _clean_item(raw)
    # Elimina eventuale seconda quantità e riferimento offerta anche dalla descrizione.
    item = re.sub(r'^\s*n\.\s*\d+\s+', '', item, flags=re.I).strip()
    item = re.sub(r'\s*[-–—]?\s*Offerta\s+[A-Z0-9_./-]+(?:\s+del\s+\d{1,2}/\d{1,2}/20\d{2})?\s*$', '', item, flags=re.I).strip(' -–—')
    item = item or 'la fornitura indicata nella richiesta di acquisto'
    commessa = str(d.get('commessa') or '').strip()
    cliente = str(d.get('cliente') or '').strip()
    tipo = classify_purchase(raw)
    low = raw.lower()

    if commessa and cliente:
        scope = f" nell'ambito della commessa {commessa}, relativa al progetto {cliente}"
    elif commessa:
        scope = f" nell'ambito della commessa {commessa}"
    elif cliente:
        scope = f" nell'ambito delle attività previste per {cliente}"
    else:
        scope = ''

    qtxt = f"n. {qty} " if qty else ''
    if tipo == 'licenze':
        action = 'rinnovare' if any(k in low for k in ['rinnovo','renew','subscription']) else 'acquisire'
        bisogno = f"Necessità di {action} {qtxt}{item}{scope}, al fine di garantire la disponibilità delle licenze necessarie alle attività progettuali e la relativa continuità operativa."
    elif tipo == 'servizi':
        bisogno = f"Necessità di acquisire {item}{scope}, a supporto delle attività previste e della corretta esecuzione del progetto."
        if qty:
            bisogno += f" Il quantitativo richiesto ({qty}) corrisponde al fabbisogno indicato nella richiesta di acquisto."
    else:
        bisogno = f"Necessità di acquisire {qtxt}{item}{scope}."
        if qty:
            bisogno += " Il quantitativo richiesto è stato determinato sulla base del fabbisogno progettuale e delle specifiche tecniche previste per la commessa."
        else:
            bisogno += " Il fabbisogno è quello indicato nella richiesta di acquisto e nella documentazione tecnica disponibile."

    off_num = str(offer_data.get('offerta_numero') or '').strip()
    off_date = str(offer_data.get('offerta_data') or '').strip()
    supplier = str(d.get('fornitore') or '').strip()
    total = float(d.get('totale') or 0)
    if off_num:
        ref = f"offerta {supplier + ' ' if supplier else ''}{off_num}"
        if off_date:
            ref += f" del {off_date}"
        prezzo = f"Valore definito sulla base dell'{ref}."
    elif supplier:
        prezzo = f"Valore definito sulla base della documentazione economica presentata da {supplier}."
    else:
        prezzo = "Valore definito sulla base della documentazione economica disponibile."
    return bisogno, prezzo


def generate_all7_rationale(d, offer_data=None, rda_data=None):
    """Compatibilità: restituisce le due sezioni in un unico testo."""
    bisogno, prezzo = generate_all7_sections(d, offer_data, rda_data)
    return bisogno + "\n\n" + prezzo

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


def xml_replace(src,dst,repls,occurrence_repls=None,black_replacements=None):
    from lxml import etree
    occurrence_repls=occurrence_repls or []; counts={}
    black_replacements=set(black_replacements or [])
    W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    with zipfile.ZipFile(src,'r') as zin, zipfile.ZipFile(dst,'w',zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data=zin.read(item.filename)
            if item.filename.endswith('.xml'):
                try:
                    root=etree.fromstring(data)
                    if black_replacements:
                        for color_node in root.xpath('//*[local-name()="color" and translate(@*[local-name()="val"], "abcdef", "ABCDEF")="808080"]'):
                            color_node.set(W+'val','000000')
                    for node in root.xpath('//*[local-name()="t"]'):
                        if node.text is None: continue
                        text=node.text
                        touched=False
                        for a,b in repls.items():
                            if a in text:
                                text=text.replace(a,str(b)); touched = touched or a in black_replacements
                        for old,new,n in occurrence_repls:
                            if old in text:
                                counts[old]=counts.get(old,0)+1
                                if counts[old]==n: text=text.replace(old,str(new),1)
                        node.text=text
                        if touched:
                            run=node.getparent()
                            rpr=run.find(W+'rPr')
                            if rpr is None:
                                rpr=etree.Element(W+'rPr'); run.insert(0,rpr)
                            color=rpr.find(W+'color')
                            if color is None:
                                color=etree.SubElement(rpr,W+'color')
                            color.set(W+'val','000000')
                            # Mantiene il corpo compatto del modello (9 pt) per evitare overflow.
                            sz=rpr.find(W+'sz')
                            if sz is None: sz=etree.SubElement(rpr,W+'sz')
                            sz.set(W+'val','18')
                            szcs=rpr.find(W+'szCs')
                            if szcs is None: szcs=etree.SubElement(rpr,W+'szCs')
                            szcs.set(W+'val','18')
                    data=etree.tostring(root,xml_declaration=True,encoding='UTF-8',standalone=True)
                except Exception: pass
            zout.writestr(item,data)


def make7(d,out):
    qty = _qty_from_object(d.get('oggetto',''), '')
    oggetto7 = d.get('oggetto_all7') or clean_all7_object(d.get('oggetto',''), qty)
    repl={
      '21/09/2026':d.get('data',''), 'OLIVETTI - P&DI CPP':d.get('funzione',''),
      'ANTONIO SCUCCIMARRA':d.get('fornitore',''),
      'Nota Spese Giugno, CB Centro Sud Puglia - Acque del Sud':oggetto7,
      'supporto per attività commerciale propedeutica alla riuscita del Progetto Water Management System gestito da Portfolio IoT':d.get('motivazione',''),
      '216,40':fmt_eur(d.get('totale',0)).replace('€ ',''),
      'Valore definito':d.get('prezzo_razionale','Valore definito sulla base dell’offerta').rstrip('.'),
      'sulla base dell’offerta':''
    }
    black={
      'OLIVETTI - P&DI CPP','ANTONIO SCUCCIMARRA',
      'Nota Spese Giugno, CB Centro Sud Puglia - Acque del Sud',
      'supporto per attività commerciale propedeutica alla riuscita del Progetto Water Management System gestito da Portfolio IoT',
      '216,40','Valore definito','sulla base dell’offerta'
    }
    xml_replace(T/'allegato7.docx',out,repl,black_replacements=black)


def _replace_text_preserve_runs(paragraph, old, new, occurrence=1):
    """Sostituisce testo in un paragrafo senza ricreare i run e senza perdere la formattazione."""
    full = ''.join(r.text for r in paragraph.runs)
    pos = -1
    cursor = 0
    for _ in range(occurrence):
        pos = full.find(old, cursor)
        if pos < 0:
            return False
        cursor = pos + len(old)
    end = pos + len(old)

    spans=[]
    cur=0
    for i,r in enumerate(paragraph.runs):
        nxt=cur+len(r.text)
        spans.append((i,cur,nxt))
        cur=nxt

    touched=[x for x in spans if x[1] < end and x[2] > pos]
    if not touched:
        return False
    first_i=touched[0][0]
    first_start=touched[0][1]
    last_i=touched[-1][0]
    last_end=touched[-1][2]
    prefix=paragraph.runs[first_i].text[:max(0,pos-first_start)]
    suffix=paragraph.runs[last_i].text[max(0,end-spans[last_i][1]):]
    paragraph.runs[first_i].text = prefix + new + (suffix if first_i==last_i else '')
    for i,_,_ in touched[1:]:
        paragraph.runs[i].text=''
    if first_i != last_i:
        paragraph.runs[last_i].text=suffix
    return True


def make6(d,out):
    # Conserva il layout e la formattazione del modello originale.
    repl={'2101416996': d.get('rda','')}
    xml_replace(T/'allegato6.docx',out,repl)

    doc=Document(out)
    for p in doc.paragraphs:
        if p.text.startswith('Riferimenti (*):') and 'Fornitore:' in p.text:
            _replace_text_preserve_runs(p, 'Scrivere qui', d.get('riferimenti',''), 1)
            _replace_text_preserve_runs(p, 'Scrivere qui', d.get('fornitore',''), 1)
    doc.save(out)


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


ATTACHMENT_NAMES = {
    '1': 'ALLEGATO 1 - Elenco tipologie di Trattativa Diretta_lista casistiche ammesse per gli acquisti di trattativa diretta',
    '2': 'ALLEGATO 2 - Casistiche di Acquisti Speciali',
    '3A': 'ALLEGATO 3A - Richiesta preventiva di autorizzazione per il ricorso a Trattativa Diretta',
    '4': 'ALLEGATO 4 - Requisiti di Sicurezza e di Compliance ICT per i fornitori',
    '5': 'ALLEGATO 5 - Scheda motivazionale per le richieste d’acquisto',
    '6': 'ALLEGATO 6 - Trattamento dei Dati personali',
    '7': 'ALLEGATO 7 - Razionali per RDA_modulo razionali dimensionanti dei fabbisogni e prezzi di riferimento di definizione budget RdA',
    '10': 'ALLEGATO 10 - Adempimenti previsti in caso di violazioni di dati personali - Data Breach',
    '11': 'ALLEGATO 11 - Allegato Tecnico di Compliance e Sicurezza - Settore Privato',
    '12': 'ALLEGATO 12 - Allegato Tecnico di Compliance e Sicurezza - Pubblica Amministrazione',
    '13': 'ALLEGATO 13 - Allegato tecnico per l’indizione di una gara per prestazioni professionali',
    '14': 'ALLEGATO 14 - Check List Requisiti di Sicurezza SaaS',
    '15': 'ALLEGATO 15 - Misure e accorgimenti relativi alle attribuzioni delle funzioni di Amministratore di Sistema',
    '16': 'ALLEGATO 16 - Attestazione di conformità ai requisiti di compliance relativi al trattamento dei dati personali',
}

def attachment_filename(rda, code, ext):
    """Nome file standard: RDA <numero> ALLEGATO <n> - <nome ufficiale>.<ext>"""
    title = ATTACHMENT_NAMES[code]
    # Evita caratteri non validi nei nomi file Windows mantenendo il titolo leggibile.
    title = re.sub(r'[\\/:*?"<>|]+', '-', title).strip()
    return f"RDA {rda} {title}.{ext}"

def package(d,docs,offer_bytes=None,offer_name=None):
    td=Path(tempfile.mkdtemp()); produced=[]
    rda=str(d.get('rda','')).strip() or 'SENZA_NUMERO'

    def cp(src,code,ext='docx'):
        p=td/attachment_filename(rda,code,ext)
        shutil.copy2(src,p); produced.append(p)

    if '4' in docs: cp(T/'allegato4.docx','4')
    if '6' in docs:
        p=td/attachment_filename(rda,'6','docx'); make6(d,p); produced.append(p)
    if '7' in docs:
        p=td/attachment_filename(rda,'7','docx'); make7(d,p); produced.append(p)
    if '3A' in docs:
        p=td/attachment_filename(rda,'3A','docx'); make3a(d,p); produced.append(p)
    for n in ['1','2','10','11','12','13','15','16']:
        if n in docs: cp(T/f'allegato{n}.docx',n)
    if '14' in docs: cp(T/'allegato14.xlsx','14','xlsx')
    if 'OFFERTA' in docs:
        if offer_bytes:
            # L'offerta mantiene il proprio nome originale: non è un Allegato numerato.
            p=td/(offer_name or 'Offerta.pdf'); p.write_bytes(offer_bytes); produced.append(p)
        else:
            p=td/'OFFERTA_MANCANTE.txt'; p.write_text('Offerta non caricata.',encoding='utf-8'); produced.append(p)
    z=io.BytesIO()
    with zipfile.ZipFile(z,'w',zipfile.ZIP_DEFLATED) as zz:
        for p in produced: zz.write(p,p.name)
    return z.getvalue()


st.set_page_config(page_title='Generatore RDA Olivetti v0.5.4',layout='wide')
st.title('Generatore RDA Olivetti — v0.5.4')
st.caption('Carica Offerta + Richiesta RDA e incolla una riga copiata dalla Bibbia Excel. L’app estrae i dati e legge anche gli allegati da generare.')

st.subheader('1. Documenti e riga della Bibbia Excel')
a,b=st.columns(2)
with a: offer_file=st.file_uploader('OFFERTA fornitore',type=['pdf'],key='offer')
with b: rda_file=st.file_uploader('RICHIESTA RDA',type=['pdf'],key='rda_pdf')
excel_paste=st.text_area('📋 Incolla qui UNA RIGA copiata da Excel',height=105,placeholder='Seleziona l’intera riga nella Bibbia Excel → Copia → Incolla qui. Le celle restano separate automaticamente.')

offer_data=extract_offer(pdf_text(offer_file)) if offer_file else {}
rda_data=extract_rda(pdf_text(rda_file)) if rda_file else {}
raw_xls,extra_xls=parse_pasted_excel_row(excel_paste)
xls_data=pasted_row_map(raw_xls)
docs_excel=docs_from_attachment_text(xls_data.get('allegati_rda',''))
detected=merge_detected(offer_data,rda_data,xls_data)

if offer_file or rda_file or excel_paste.strip():
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
        nota_razionale=st.text_area('Indicazioni aggiuntive per Allegato 7 (opzionale)',value='',help='Scrivi solo ciò che non emerge dai documenti: es. continuità, precedente acquisto, motivazione specifica.')
        riferimenti=st.text_input('Riferimenti (SDW/SH/PCD, se applicabile)')
        consegna=st.text_input('Consegna rilevata',value=str(detected.get('consegna','')))

    with st.expander('Mostra cosa è stato letto dai tre input'):
        st.write('**Offerta:**', offer_data or 'nessun dato')
        st.write('**Richiesta RDA:**', rda_data or 'nessun dato')
        st.write('**Riga Excel incollata:**', raw_xls or 'nessun dato')
        st.write('**Allegati letti dalla riga:**', docs_excel or 'nessun allegato riconosciuto')

    st.subheader('3. Allegati da generare')
    allegati_testo=xls_data.get('allegati_rda','')
    if allegati_testo:
        st.success('Dalla riga Excel: ' + allegati_testo)
    if docs_excel:
        st.write('**Riconosciuti automaticamente:** ' + ', '.join(docs_excel))
        mode=st.radio('Modalità', ['Usa gli allegati della riga Excel','Modifica manualmente','Determina automaticamente'],horizontal=True)
    else:
        st.warning('Nella riga incollata non ho riconosciuto la colonna allegati: puoi indicarli manualmente o usare il controllo automatico.')
        mode=st.radio('Modalità', ['Modifica manualmente','Determina automaticamente'],horizontal=True)

    # defaults shared by both modes
    flags={k:False for k in ['nuova_tranche_senza_variazione','precontrattuale','deroga','side_letter','gara_prest_prof','saas','trattamento_cliente','amministratore_sistema','trattamento_dati','attestazione_conformita']}
    cliente_tipo='n.a.'; casistica_td=''

    if mode=='Usa gli allegati della riga Excel':
        docs=docs_excel
        st.caption('La colonna “Allegati alla RDA” della Bibbia è la fonte principale. Il motore non aggiunge allegati da solo.')
        if '3A' in docs:
            casistica_td=st.selectbox('Casistica TD per il 3A (se applicabile)',['']+CASI_TD)
    elif mode=='Modifica manualmente':
        default_docs=docs_excel or (['4','6','7'] + (['3A'] if totale>20000 else []))
        selected_docs=st.multiselect('Allegati da generare',ALL_DOCS,default=default_docs)
        docs=selected_docs
        st.caption('Puoi correggere la selezione letta dalla riga Excel.')
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

    # Allegato 7: razionale automatico, sempre verificabile e modificabile
    base_for_7=dict(rda=rda,fornitore=fornitore,sap=sap,piva=piva,totale=totale,data=data,commessa=commessa,
                    cliente=cliente,richiedente=richiedente,funzione=funzione,oggetto=oggetto)
    bisogno7, prezzo7 = generate_all7_sections(base_for_7, offer_data, rda_data)
    if nota_razionale.strip():
        bisogno7 = bisogno7.rstrip() + ' ' + nota_razionale.strip()
    qty7 = _qty_from_object(oggetto, rda_data.get('quantita',''))
    oggetto7_default = clean_all7_object(oggetto, qty7)
    if '7' in docs:
        st.subheader('4. Allegato 7 — anteprima contenuti')
        st.caption('Le tre sezioni restano separate per mantenere pulita la formattazione del modello Word.')
        oggetto_all7=st.text_area('Oggetto Allegato 7',value=oggetto7_default,height=70)
        motivazione=st.text_area('Descrizione esigenza / dimensionamento fabbisogno',value=bisogno7,height=150)
        prezzo_razionale=st.text_area('Razionale del prezzo',value=prezzo7,height=90)
    else:
        oggetto_all7=oggetto7_default
        motivazione=bisogno7
        prezzo_razionale=prezzo7

    st.subheader('5. Controllo finale')
    missing=[]
    for label,val in [('N. RDA',rda),('Fornitore',fornitore),('Oggetto',oggetto)]:
        if not val: missing.append(label)
    if '3A' in docs and totale<=20000:
        st.warning('Il pacchetto include il 3A anche se l’importo non supera €20.000. Verrà generato perché è indicato nella riga Excel o nella selezione manuale: verifica che sia corretto per questa pratica.')
    if missing: st.warning('Da completare prima della generazione: '+', '.join(missing))
    st.write('**Pacchetto:** '+(', '.join(docs) if docs else 'nessun allegato'))

    d=dict(rda=rda,fornitore=fornitore,sap=sap,piva=piva,totale=totale,data=data,commessa=commessa,
           cliente=cliente,richiedente=richiedente,funzione=funzione,oggetto=oggetto,oggetto_all7=oggetto_all7,motivazione=motivazione,prezzo_razionale=prezzo_razionale,
           riferimenti=riferimenti,consegna=consegna,casistica_td=casistica_td,cliente_tipo=cliente_tipo,**flags)

    if st.button('GENERA PACCHETTO RDA',type='primary',disabled=bool(missing or not docs)):
        z=package(d,docs,offer_file.getvalue() if offer_file else None,offer_file.name if offer_file else None)
        st.download_button('Scarica ZIP RDA',z,file_name=f"RDA_{rda}_allegati.zip",mime='application/zip')
else:
    st.info('Carica almeno un documento oppure incolla una riga Excel. Flusso consigliato: Offerta + Richiesta RDA + riga copiata dalla Bibbia.')
