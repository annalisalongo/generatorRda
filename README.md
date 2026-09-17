# Generatore RDA Olivetti — v0.1

Prima versione del programma.

## Cosa fa già
- raccoglie i dati della RDA;
- applica automaticamente la procedura precedente / successiva al 15/04/2026;
- applica la soglia > 20.000 € per l'Allegato 3A secondo il vademecum operativo;
- gestisce la presenza dell'Offerta in base alla deroga nella nuova procedura;
- propone le 10 casistiche ufficiali di Trattativa Diretta;
- apre i campi aggiuntivi solo quando servono;
- gestisce il primo livello di domande per l'Allegato 6;
- segnala i dati mancanti e impedisce la generazione finché non sono completi;
- mostra un'anteprima strutturata dei dati.

## Avvio

1. Installa Python 3.11+.
2. Apri il terminale nella cartella del progetto.
3. Esegui:

    pip install -r requirements.txt
    streamlit run app.py

Si aprirà l'interfaccia nel browser.

## Prossimo modulo
La v0.2 collegherà i modelli reali e produrrà i file di output:
- Allegato 6;
- Allegato 4 / Allegato 1 vecchio;
- Allegato 7 / Allegato 5 vecchio;
- Allegato 3A quando richiesto.

I PDF originali richiedono una strategia di compilazione per coordinate/campi; i DOCX possono essere compilati preservando il più possibile il modello originale.
