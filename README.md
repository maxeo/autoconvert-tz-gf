# Script di Aggiornamento del Google Form

Lo script di aggiornamento del Google Form è progettato per gestire le operazioni del form in base ai dati recuperati da un foglio di lavoro in Google Spreadsheet. Questo script è scritto in ECMAScript 6 e utilizza le Google Apps Script Services per interagire con Google Form e Google Spreadsheet.

## Funzionalità Principali

1. **Aggiornamento delle domande del Google Form:** Lo script elimina tutte le domande esistenti nel form e crea nuove domande basate sui dati recuperati dal foglio di lavoro.

2. **Gestione delle Timezone:** Lo script gestisce le domande basate sulla timezone degli utenti. Per ogni timezone, crea una nuova sezione nel form.

3. **Preparazione dei Dati:** Lo script ha un metodo per preparare i dati recuperati dal foglio di lavoro, formattando e riorganizzando i dati come richiesto.

## Come utilizzare lo script

1. **Struttura del Google Spreadsheet:** Lo script richiede che il tuo foglio di lavoro di Google Spreadsheet sia strutturato in un modo specifico per funzionare correttamente. Nella prima pagina del tuo Google Spreadsheet, devi avere una tabella con le seguenti colonne e tipo di dati:

   | Data disponibilità | Ora disponibilità from | Persone           |
       |------------------- |----------------------  |------------------ |
   | (Data)             | (Orario nel formato HH:MM) | (Stringa separata da virgole di nomi) |

   Esempio:

   | Data disponibilità | Ora disponibilità from | Persone                  |
       |------------------- |----------------------  |-----------------------|
   | 2024-07-01        | 16:00                 | Mario, Luigi, Peach |
   | 2024-07-01        | 17:00                 | Mario, Luigi, Peach |
   | 2024-07-01        | 18:00                 | Mario, Luigi, Peach |

2. **Recuperare `formId` e `sheetId`:** Questi sono gli ID univoci del tuo Google Form e Google Spreadsheet. Dovresti sostituire questi valori con i tuoi ID univoci. Puoi trovare l'ID del tuo Google Form nell'URL mentre stai modificando il form, e l'ID del Google Spreadsheet nell'URL mentre stai visualizzando il foglio di lavoro.

3. **Esecuzione dello script:** Una volta impostato l'ID del form e del foglio di lavoro, puoi eseguire lo script. Lo script aggiornerà il Google Form sulla base dei dati presenti nel foglio di lavoro.

## Contribuire

Le segnalazioni di bug, le richieste di funzionalità e i contributi di codice sono i benvenuti. Se stai contribuendo con del codice, assicurati di seguire le linee guida di stile del codice in uso.

## Licenza

Questo progetto è con licenza sotto i termini della licenza MIT.
