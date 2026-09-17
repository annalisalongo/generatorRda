# Generatore RDA Olivetti v0.2

Questa versione genera un file ZIP con gli allegati richiesti dal vademecum.

## Avvio

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Funzioni
- selezione automatica procedura prima/dopo 15/04/2026;
- soglia > 20.000 euro per Allegato 3A secondo vademecum;
- Offerta solo in deroga nella nuova procedura;
- compilazione automatica dei campi principali di Allegato 6, 7 e 3A sui PDF originali;
- inclusione degli allegati statici previsti;
- blocco se mancano dati obbligatori.

## Nota
La v0.2 è una prima calibrazione delle coordinate dei PDF. Prima dell'uso produttivo è consigliato fare una prova con una RDA reale e verificare visivamente i documenti generati.
