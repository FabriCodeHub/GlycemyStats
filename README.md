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

## Esecuzione dello Script

1. Eseguire lo script Python:
   ```bash
   python glycemystats.py
   ```
2. Seguire le istruzioni per l'inserimento dei dati:
   - Inserire la data nel formato richiesto (YYYY-MM-DD, DD/MM/YYYY o DD\MM\YYYY).
   - Inserire l'ora nel formato richiesto (HH:MM o HH.MM).
   - Inserire il livello di glicemia come numero di 2 o 3 cifre.
3. Ripetere l'inserimento dei dati per più voci, se necessario.
4. Al termine, il file Excel con i dati registrati e il grafico verrà creato automaticamente.

## Dettagli del Codice

### `stampa_titolo()`

Questa funzione stampa il titolo "GlycemyStats" in grande usando la libreria `pyfiglet`.

### `inserisci_dati_glicemia()`

Questa funzione gestisce l'inserimento interattivo dei dati da parte dell'utente. Verifica i formati della data e dell'ora e assicura che il livello di glicemia sia un numero di 2 o 3 cifre.

### `crea_file_excel()`

Questa funzione crea un file Excel a partire dai dati salvati nel file CSV. Aggiunge la formattazione condizionale per evidenziare i livelli di glicemia e crea un grafico professionale.

### `main()`

La funzione principale che coordina le operazioni:
- Stampa il titolo.
- Esegue un ciclo per l'inserimento dei dati.
- Crea il file Excel al termine dell'inserimento dei dati.

## Esempio di Utilizzo

```
Inserisci la data (YYYY-MM-DD o DD/MM/YYYY o DD\MM\YYYY): 11\07\2024
Inserisci l'ora (HH:MM o HH.MM): 19:15
Inserisci il livello di glicemia (2 o 3 cifre): 109
Dati inseriti correttamente.
Vuoi inserire un altro dato? (s/n): n
File Excel creato correttamente: livelli_glicemia.xlsx
```

## Formattazione Condizionale

- Livelli di glicemia tra 101 e 129: Rosso
- Livelli di glicemia tra 70 e 99: Verde

## Grafico

Il grafico presente nel file Excel visualizza i livelli di glicemia nel tempo, con un asse per la data e l'ora e un asse per i livelli di glicemia.

## Contatti

Per domande o suggerimenti, contattare l'autore dello script.


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
