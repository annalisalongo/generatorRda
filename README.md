# Generatore RDA Olivetti v0.5.5

Novità rispetto alla v0.5.3:

- Tutti gli allegati generati usano il formato nome completo:
  `RDA <numero> ALLEGATO <n> - <nome ufficiale>.<estensione>`
- Esempio Allegato 7:
  `RDA 2101417276 ALLEGATO 7 - Razionali per RDA_modulo razionali dimensionanti dei fabbisogni e prezzi di riferimento di definizione budget RdA.docx`
- L'Allegato 6 inserisce automaticamente nella sezione firme:
  `Paolo Sigismondi - 1° Riporto`
- Restano invariati il flusso Offerta + Richiesta RDA + riga Excel incollata, la lettura automatica degli allegati e il generatore argomentato dell'Allegato 7.

## Avvio

```bash
pip install -r requirements.txt
streamlit run app.py
```
