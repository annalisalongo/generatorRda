# Generatore RDA Olivetti v0.5

## Novità
La v0.5 parte dai documenti reali:
1. carica l'**Offerta PDF**;
2. carica la **Richiesta RDA PDF**;
3. opzionalmente carica il **file Excel** e indica la riga;
4. l'app estrae i dati e li mostra per controllo/correzione;
5. scegli **Li indico io** per indicare direttamente gli allegati oppure **Determina automaticamente**;
6. genera il pacchetto ZIP.

L'app non inventa i campi che non trova: li lascia vuoti e richiede verifica manuale.

## Installazione
```bash
pip install -r requirements.txt
```

## Avvio
```bash
streamlit run app.py
```

## Nota
Il parser è euristico: funziona sui modelli RDA/offerta simili agli esempi usati nello sviluppo, ma i dati estratti vanno sempre controllati nella schermata di revisione prima della generazione.
