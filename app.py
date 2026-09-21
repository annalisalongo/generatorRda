from pathlib import Path
import io, re, shutil, tempfile, zipfile
import streamlit as st
from docx import Document

BASE=Path(__file__).parent
T=BASE/'templates'

CASI_TD=[
'1 - Particolarità tecnologiche/infrastrutturali / continuità',
'2 - Emergenza adeguatamente motivata','3 - Consulenze/prestazioni professionali specialistiche o complementari',
'4 - Acquisto di opportunità','5 - Gara andata deserta','6 - Prezzi imposti / cessione azienda / impossibilità oggettiva di competizione',
'7 - Estensione condizioni/prezzi di competizione','8 - Progetto business/tecnologico identificato con fabbisogni e fornitori puntuali',
'9 - Commercializzazione prodotti/soluzioni non standard a clienti privati','10 - Opportunità business per commercializzazione verso PA']

SPECIALI=[
'1 - Scelte tecnologiche/business di Gruppo da unico fornitore','2 - Diritti/sponsorizzazioni/collaboratori artistici unici',
'3 - Manutenzione applicativi proprietari / rete non TPM','4 - Partnership commerciali','5 - Canali del Costruttore certificati / prezzi Procurement',
'6 - Vendor tecnologico/distributore per gare PA con tecnologia esplicita ed esclusività','7 - Nuove tecnologie / primo impiego / sperimentazioni',
'8 - Postazioni alta frequenza','9 - Prodotti di rete per ampliamento reti esistenti','10 - Adesione ad Accordi Quadro di Gruppo',
'11 - Networking ICT con accordi/listini PR >=80% commessa','12 - Leasing individuato da Enterprise/CF.F']

def euro(s):
    if isinstance(s,(int,float)): return float(s)
    s=str(s).strip().replace('€','').replace(' ','')
    if ',' in s: s=s.replace('.','').replace(',','.')
    return float(s or 0)

def fmt_eur(x):
    return f"€ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X','.')

def decide(d):
    docs=[]; reasons=[]; warnings=[]
    # Nuova procedura: il motore evita automatismi non supportati e chiede le condizioni rilevanti.
    docs += ['4','6']
    reasons += ['Allegato 4: requisiti sicurezza/compliance','Allegato 6: valutazione trattamento dati e firma del modulo']
    if not d['nuova_tranche_senza_variazione'] and not d['precontrattuale']:
        docs.append('7'); reasons.append('Allegato 7: razionali fabbisogno/prezzo')
    if d['deroga'] or d['side_letter']:
        docs.append('OFFERTA'); reasons.append('Offerta: acquisto in deroga o gruppo merce Side Letter')
    if d['trattativa_diretta'] and d['totale']>20000 and not d['acquisto_speciale'] and not d['infragruppo'] and not d['accordo_esistente']:
        docs.append('3A'); reasons.append('Allegato 3A: TD >20k e nessuna esclusione dichiarata')
    if d['gara_prest_prof']:
        docs.append('13'); reasons.append('Allegato 13: gara per prestazioni professionali')
    if d['saas']:
        docs.append('14'); reasons.append('Allegato 14: checklist sicurezza SaaS')
    if d['cliente_tipo']=='Privato' and d['trattamento_cliente']:
        docs.append('11'); reasons.append('Allegato 11: ATCS cliente privato')
    if d['cliente_tipo']=='PA' and d['trattamento_cliente']:
        docs.append('12'); reasons.append('Allegato 12: ATCS cliente PA')
    if d['amministratore_sistema']:
        docs.append('15'); reasons.append('Allegato 15: attività di Amministratore di Sistema')
    if d['trattamento_dati']:
        docs.append('10'); reasons.append('Allegato 10: istruzioni data breach per responsabili terzi')
    if d['attestazione_conformita']:
        docs.append('16'); reasons.append('Allegato 16: attestazione conformità privacy')
    if d['acquisto_speciale']:
        warnings.append('Acquisto Speciale selezionato: non genera 3A automaticamente; si applica per il resto il processo standard.')
    if d['trattativa_diretta'] and not d['casistica_td'] and not d['acquisto_speciale']:
        warnings.append('Selezionare la casistica di Trattativa Diretta (Allegato 1).')
    return list(dict.fromkeys(docs)), reasons, warnings

def xml_replace(src,dst,repls,occurrence_repls=None):
    occurrence_repls=occurrence_repls or []
    with zipfile.ZipFile(src,'r') as zin, zipfile.ZipFile(dst,'w',zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data=zin.read(item.filename)
            if item.filename.endswith('.xml'):
                text=data.decode('utf-8')
                for a,b in repls.items(): text=text.replace(a,str(b))
                for old,new,n in occurrence_repls:
                    parts=text.split(old)
                    if len(parts)>n:
                        text=old.join(parts[:n])+str(new)+old.join(parts[n:])
                data=text.encode('utf-8')
            zout.writestr(item,data)

def make7(d,out):
    repl={
      '21/09/2026':d['data'], 'OLIVETTI - P&amp;DI CPP':d['funzione'], 'OLIVETTI - P&DI CPP':d['funzione'],
      'ANTONIO SCUCCIMARRA':d['fornitore'],
      'Nota Spese Giugno, CB Centro Sud Puglia - Acque del Sud':d['oggetto'],
      'supporto per attività commerciale propedeutica alla riuscita del Progetto Water Management System gestito da Portfolio IoT':d['motivazione'],
      '216,40':fmt_eur(d['totale']).replace('€ ','')
    }
    xml_replace(T/'allegato7.docx',out,repl)

def make6(d,out):
    repl={'2101416996':d['rda']}
    # Il modello caricato contiene due placeholder consecutivi nella riga Riferimenti/Fornitore.
    occ=[('Scrivere qui',d['riferimenti'],1),('Scrivere qui',d['fornitore'],2)]
    xml_replace(T/'allegato6.docx',out,repl,occ)

def make3a(d,out):
    repl={
      'Scrivere qui il nome del fornitore':d['fornitore'], '00000':d['sap'], '00000000000':d['piva'],
      '€ 000.000.000,00':fmt_eur(d['totale']), '0000000000':d['rda'],
      'Scegliere un elemento.': d['casistica_td'] or 'DA COMPILARE',
      'Scrivere qui la risposta':d['motivazione'], 'Scrivere qui il cliente':d['cliente']
    }
    xml_replace(T/'allegato3.docx',out,repl)

def package(d,docs,offer_bytes=None,offer_name=None):
    td=Path(tempfile.mkdtemp())
    produced=[]
    if '4' in docs: shutil.copy2(T/'allegato4.docx',td/f"RDA {d['rda']} - Allegato 4.docx"); produced.append(td/f"RDA {d['rda']} - Allegato 4.docx")
    if '6' in docs: make6(d,td/f"RDA {d['rda']} - Allegato 6.docx"); produced.append(td/f"RDA {d['rda']} - Allegato 6.docx")
    if '7' in docs: make7(d,td/f"RDA {d['rda']} - Allegato 7.docx"); produced.append(td/f"RDA {d['rda']} - Allegato 7.docx")
    if '3A' in docs: make3a(d,td/f"RDA {d['rda']} - Allegato 3A - DA VERIFICARE.docx"); produced.append(td/f"RDA {d['rda']} - Allegato 3A - DA VERIFICARE.docx")
    for n in ['10','11','12','13','15','16']:
        if n in docs:
            shutil.copy2(T/f'allegato{n}.docx',td/f"RDA {d['rda']} - Allegato {n}.docx"); produced.append(td/f"RDA {d['rda']} - Allegato {n}.docx")
    if '14' in docs:
        shutil.copy2(T/'allegato14.xlsx',td/f"RDA {d['rda']} - Allegato 14.xlsx"); produced.append(td/f"RDA {d['rda']} - Allegato 14.xlsx")
    if 'OFFERTA' in docs:
        if offer_bytes:
            p=td/(offer_name or 'Offerta.pdf'); p.write_bytes(offer_bytes); produced.append(p)
        else:
            p=td/'OFFERTA_MANCANTE.txt'; p.write_text('Caricare l’offerta del fornitore prima di chiudere la RDA.',encoding='utf-8'); produced.append(p)
    z=io.BytesIO()
    with zipfile.ZipFile(z,'w',zipfile.ZIP_DEFLATED) as zz:
        for p in produced: zz.write(p,p.name)
    return z.getvalue()

st.set_page_config(page_title='Generatore RDA Olivetti v0.4',layout='wide')
st.title('Generatore RDA Olivetti — v0.4')
st.caption('Nuova procedura: motore decisionale + compilazione sui DOCX originali. I campi non supportati non vengono inventati.')

c1,c2,c3=st.columns(3)
with c1:
    rda=st.text_input('N. RDA')
    fornitore=st.text_input('Fornitore')
    sap=st.text_input('Codice SAP fornitore')
    piva=st.text_input('P. IVA fornitore')
    totale=euro(st.text_input('Totale RDA','0,00'))
with c2:
    data=st.text_input('Data','21/09/2026')
    funzione=st.text_input('Funzione richiedente','OLIVETTI - P&DI CPP')
    cliente=st.text_input('Cliente / Gara')
    cliente_tipo=st.selectbox('Tipo cliente',['n.a.','Privato','PA'])
    riferimenti=st.text_input('Riferimenti (SDW/SH/PCD, se applicabile)')
with c3:
    oggetto=st.text_area('Oggetto')
    motivazione=st.text_area('Motivazione / descrizione esigenza')

st.subheader('Domande decisionali')
a,b,c=st.columns(3)
with a:
    pre_sales=st.checkbox('RDA Pre Sales')
    trattativa_diretta=st.checkbox('Trattativa Diretta')
    acquisto_speciale=st.checkbox('Acquisto Speciale (Allegato 2)')
    infragruppo=st.checkbox('Acquisto infragruppo')
with b:
    accordo_esistente=st.checkbox('Contratto/AQ/listino già in essere')
    nuova_tranche_senza_variazione=st.checkbox('Nuova tranche senza variazione valore complessivo')
    precontrattuale=st.checkbox('Ingaggio fornitore in fase precontrattuale')
    deroga=st.checkbox('Acquisto in deroga')
    side_letter=st.checkbox('Gruppo merce soggetto a Side Letter')
with c:
    trattamento_dati=st.checkbox('La fornitura tratta dati personali')
    trattamento_cliente=st.checkbox('Trattamento dati di cui il cliente è Titolare/Responsabile')
    saas=st.checkbox('Servizio SaaS')
    amministratore_sistema=st.checkbox('Attività di Amministratore di Sistema')
    gara_prest_prof=st.checkbox('Gara per prestazioni professionali')
    attestazione_conformita=st.checkbox('Richiesta attestazione conformità privacy')

casistica_td=''
if trattativa_diretta and not acquisto_speciale:
    casistica_td=st.selectbox('Casistica Trattativa Diretta — Allegato 1',['']+CASI_TD)
if acquisto_speciale:
    st.selectbox('Casistica Acquisto Speciale — Allegato 2',['']+SPECIALI)

d=locals().copy()
docs,reasons,warnings=decide(d)
st.subheader('Anteprima allegati')
st.write('**Da predisporre:** '+(', '.join(docs) if docs else 'nessuno'))
for x in reasons: st.write('• '+x)
for w in warnings: st.warning(w)

offer=st.file_uploader('Offerta del fornitore (se richiesta)',type=['pdf','docx','xlsx'])
if st.button('GENERA PACCHETTO RDA',type='primary',disabled=not bool(rda and fornitore and oggetto)):
    z=package(d,docs,offer.getvalue() if offer else None,offer.name if offer else None)
    st.download_button('Scarica ZIP RDA',z,file_name=f'RDA_{rda}_allegati.zip',mime='application/zip')
    if '3A' in docs: st.info('Il 3A viene precompilato nei campi sicuri e marcato DA VERIFICARE: contiene campi autorizzativi/contrattuali che non vanno inventati.')
