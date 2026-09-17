import streamlit as st
from datetime import date
from decimal import Decimal, InvalidOperation
from pathlib import Path
from io import BytesIO
import zipfile, shutil, tempfile
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from pypdf import PdfReader, PdfWriter

ROOT=Path(__file__).parent
T=ROOT/'templates'
CUT=date(2026,4,15); TH=Decimal('20000')
CASI=[
'1. Particolarità tecnologiche/infrastrutturali','2. Acquisti in condizioni di emergenza','3. Consulenze/prestazioni professionali di particolare specializzazione','4. Acquisti d’opportunità','5. Gara andata deserta','6. Prezzi imposti / impossibilità oggettiva di competizione','7. Estensione condizioni/prezzi di competizione','8. Progetti business/tecnologici con fabbisogni puntuali','9. Commercializzazione non standard per clienti privati','10. Opportunità business verso PA']

def money(s):
    s=(s or '').replace('€','').replace(' ','')
    if ',' in s:s=s.replace('.','').replace(',','.')
    try:return Decimal(s)
    except:return Decimal('0')

def fmt(v):
    return f"{v:,.2f}".replace(',','X').replace('.',',').replace('X','.')

def overlay_pdf(src, pages_draw):
    r=PdfReader(str(src)); w=PdfWriter()
    for i,p in enumerate(r.pages):
        if i in pages_draw:
            mb=p.mediabox; width=float(mb.width); height=float(mb.height)
            b=BytesIO(); c=canvas.Canvas(b,pagesize=(width,height))
            pages_draw[i](c,width,height); c.save(); b.seek(0)
            op=PdfReader(b).pages[0]; p.merge_page(op)
        w.add_page(p)
    out=BytesIO(); w.write(out); return out.getvalue()

def text(c,x,y,s,size=8,maxw=85):
    c.setFont('Helvetica',size)
    words=str(s or '').split(); line=''; yy=y
    for wd in words:
        test=(line+' '+wd).strip()
        if c.stringWidth(test,'Helvetica',size)>maxw and line:
            c.drawString(x,yy,line); yy-=size+2; line=wd
        else: line=test
    if line:c.drawString(x,yy,line)

def gen6(d):
    def p0(c,w,h):
        c.setFillColorRGB(1,1,1); c.rect(210,h-160,220,16,fill=1,stroke=0); c.setFillColorRGB(0,0,0)
        text(c,220,h-154,f"RdA N° {d['rda']}",9,190)
        text(c,365,h-272,d['fornitore'],8,170)
        # segno risposta trattamento
        text(c,470,h-187,'X' if d['tratta']=='Sì' else '',10,20)
        text(c,515,h-187,'X' if d['tratta']=='No' else '',10,20)
    return overlay_pdf(T/'allegato6.pdf',{0:p0})

def gen7(d):
    def p0(c,w,h):
        text(c,420,h-173,d['data'].strftime('%d/%m/%Y'),8,120)
        text(c,65,h-220,d['funzione'],9,580)
        text(c,65,h-267,d['fornitore'],9,580)
        text(c,65,h-330,d['oggetto'],8,580)
        desc=d['descrizione'] or f"Fabbisogno relativo alla commessa {d['commessa']} - cliente/gara {d['cliente']}."
        text(c,65,h-455,desc,8,580)
        text(c,80,120,'€ '+fmt(d['totale']),10,250)
    return overlay_pdf(T/'allegato7.pdf',{0:p0})

def gen3a(d):
    # Il modello 3A è pagina 2 del PDF complessivo: estraiamo solo quella pagina e compiliamo i campi principali.
    r=PdfReader(str(T/'allegato3.pdf')); base=PdfWriter(); base.add_page(r.pages[1]); tmp=BytesIO(); base.write(tmp); tmp.seek(0)
    rr=PdfReader(tmp); p=rr.pages[0]; w=float(p.mediabox.width); h=float(p.mediabox.height)
    b=BytesIO(); c=canvas.Canvas(b,pagesize=(w,h))
    text(c,72,h-152,d['fornitore'],8,150); text(c,270,h-152,d['sap'],8,90); text(c,445,h-152,d['piva'],8,100)
    text(c,75,h-194,'€ '+fmt(d['totale']),8,120); text(c,450,h-194,d['rda'],8,100)
    text(c,75,h-250,d['casistica'],7,460)
    text(c,75,h-300,d['descrizione'],7,460)
    text(c,75,h-375,d['tecnologia'],7,200)
    text(c,75,h-650,d['cliente'],8,400)
    c.save(); b.seek(0); p.merge_page(PdfReader(b).pages[0]); out=BytesIO(); ww=PdfWriter(); ww.add_page(p); ww.write(out); return out.getvalue()

def make_zip(d,docs):
    zbuf=BytesIO()
    with zipfile.ZipFile(zbuf,'w',zipfile.ZIP_DEFLATED) as z:
        if '6' in docs:z.writestr(f"RDA_{d['rda']}_Allegato_6.pdf",gen6(d))
        if '7' in docs:z.writestr(f"RDA_{d['rda']}_Allegato_7.pdf",gen7(d))
        if '3A' in docs:z.writestr(f"RDA_{d['rda']}_Allegato_3A.pdf",gen3a(d))
        if '4' in docs:z.write(T/'allegato4.docx',f"RDA_{d['rda']}_Allegato_4.docx")
        if '5' in docs:z.write(T/'allegato5.docx',f"RDA_{d['rda']}_Allegato_5.docx")
        if '1' in docs:z.write(T/'allegato1_casistiche.pdf',f"RDA_{d['rda']}_Allegato_1.pdf")
        if 'OFFERTA' in docs:z.writestr('INSERIRE_OFFERTA.txt','Inserire qui il file di offerta della RDA. Il programma non inventa né sostituisce l’offerta del fornitore.')
    return zbuf.getvalue()

st.set_page_config(page_title='Generatore RDA Olivetti',page_icon='📄',layout='wide')
st.title('Generatore documentazione RDA Olivetti')
st.caption('v0.2 - genera materialmente il pacchetto di allegati dai modelli reali caricati.')

c1,c2,c3=st.columns(3)
with c1:
    fornitore=st.text_input('Intestatario / Ragione sociale'); sap=st.text_input('Codice SAP fornitore'); rda=st.text_input('Numero RDA')
with c2:
    data_rda=st.date_input('Data RDA'); commessa=st.text_input('Rif. Commessa'); cliente=st.text_input('Cliente / Gara')
with c3:
    totale_s=st.text_input('Totale RDA',placeholder='23.200,00'); oggetto=st.text_area('Oggetto',height=110)
tot=money(totale_s)

c1,c2,c3=st.columns(3)
with c1: deroga=st.radio('In deroga / TD?', ['No','Sì'],horizontal=True)=='Sì'
with c2: tratta=st.radio('Tratta dati personali?', ['No','Sì','Da verificare'],horizontal=True)
with c3: funzione=st.text_input('Funzione richiedente')

nuova=data_rda>=CUT
docs=['6','4','7'] if nuova else ['6','1','5','OFFERTA']
if nuova and deroga:docs.append('OFFERTA')
if tot>TH:docs.append('3A')

piva=casistica=descrizione=tecnologia=''
if '3A' in docs:
    st.subheader('Dati Allegato 3A')
    a,b=st.columns(2)
    with a:piva=st.text_input('P.IVA fornitore'); casistica=st.selectbox('Casistica TD',['']+CASI)
    with b:tecnologia=st.text_input('Tecnologia (se applicabile)')
    descrizione=st.text_area('Descrizione dettagliata / motivazione TD',height=110)
else:
    descrizione=st.text_area('Descrizione esigenza per Allegato 7 (facoltativa)',height=80)

st.subheader('Documenti previsti')
labels={'6':'Allegato 6 - trattamento dati (firma)','4':'Allegato 4 - sicurezza (no firma)','7':'Allegato 7 - razionali (firma)','3A':'Allegato 3A - TD (firma)','1':'Allegato 1 - vecchia procedura','5':'Allegato 5 - scheda motivazionale (firma)','OFFERTA':'Offerta'}
for x in docs:st.write('✓ '+labels[x])

missing=[]
for n,v in [('Ragione sociale',fornitore),('Codice SAP',sap),('Numero RDA',rda),('Cliente/Gara',cliente),('Oggetto',oggetto),('Funzione richiedente',funzione)]:
    if not v:missing.append(n)
if tot<=0:missing.append('Totale valido')
if tratta=='Da verificare':missing.append('Trattamento dati')
if '3A' in docs:
    if not piva:missing.append('P.IVA')
    if not casistica:missing.append('Casistica TD')
    if not descrizione:missing.append('Descrizione/motivazione TD')

if missing:st.error('Da completare: '+' • '.join(missing))
else:
    d={'fornitore':fornitore,'sap':sap,'rda':rda,'data':data_rda,'commessa':commessa,'cliente':cliente,'totale':tot,'oggetto':oggetto,'funzione':funzione,'tratta':tratta,'piva':piva,'casistica':casistica,'descrizione':descrizione,'tecnologia':tecnologia}
    payload=make_zip(d,docs)
    st.success('Pacchetto pronto.')
    st.download_button('GENERA E SCARICA DOCUMENTI',payload,file_name=f'RDA_{rda}_documentazione.zip',mime='application/zip',type='primary')

st.info('Nota v0.2: Allegati 6, 7 e 3A vengono compilati automaticamente sui PDF originali. Gli allegati statici di sicurezza/casistiche vengono copiati dal modello. La scheda motivazionale vecchia è inclusa come modello nella procedura precedente e sarà resa compilabile campo-per-campo nella revisione successiva.')
