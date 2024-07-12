# GlycemyStats

GlycemyStats è uno script Python progettato per monitorare i livelli di glicemia giornalieri, organizzati per data, ora e livello di glicemia. I dati inseriti vengono salvati in un file CSV e successivamente esportati in un file Excel, che include una formattazione condizionale e un grafico professionale.

## Funzionalità

- Inserimento interattivo della data, ora e livello di glicemia.
- Supporto per più formati di data e ora.
- Salvataggio dei dati in un file CSV.
- Esportazione dei dati in un file Excel con formattazione condizionale e grafico.
- Evidenziazione dei livelli di glicemia in rosso se tra 101 e 129 e in verde se tra 70 e 99.

## Requisiti

- Python 3.x
- Moduli Python:
  - `csv`
  - `datetime`
  - `pyfiglet`
  - `openpyxl`
  - `re`

## Installazione

1. Clonare il repository o scaricare il file script Python.
2. Installare le dipendenze necessarie usando pip:
   ```bash
   pip install pyfiglet openpyxl
