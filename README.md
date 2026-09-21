# Generatore RDA Olivetti v0.4

Avvio:
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Cosa cambia rispetto alla v0.3
- usa i DOCX originali caricati, non overlay PDF a coordinate;
- separa Trattativa Diretta da Acquisto Speciale;
- non applica più la regola semplicistica `>20k = 3A` senza verificare le esclusioni;
- Allegato 7 non viene proposto per nuova tranche senza variazione del valore complessivo o ingaggio precontrattuale;
- gestisce condizioni per ATCS privato/PA, SaaS, AdS, data breach e gara prestazioni professionali;
- consente di allegare la vera offerta del fornitore.

## Stato compilazione
- Allegato 6: precompila RDA, riferimenti e fornitore sul DOCX originale.
- Allegato 7: precompila i campi del modello-esempio fornito (data, funzione, fornitore, oggetto, motivazione, importo).
- Allegato 3A: precompila solo i campi sicuri e lascia da verificare i campi autorizzativi/contrattuali.
- Allegati 4/10/11/12/13/14/15/16: inclusi come template originali quando la condizione li richiede.

La v0.4 è volutamente conservativa: se un dato non è disponibile, non lo inventa.
