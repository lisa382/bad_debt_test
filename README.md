# Bad Debt Reserve Dashboard

Lokal Streamlit-dashboard som beräknar bad debt-reserv utifrån ett NetSuite-utdrag
med öppna kundfakturor och credit memos.

## Kom igång

1. Lägg ett NetSuite-utdrag (`.xlsx`) i samma mapp som `app.py`. Filen ska ha kolumnerna:
   `Internal ID`, `Date`, `Type`, `Document Number`, `Name`, `Due Date/Receive By`,
   `Amount`, `Amount Remaining`, `Currency`.
2. Dubbelklicka på `Starta Dashboard.command` (skapar en Python-miljö automatiskt
   första gången) — eller kör manuellt:

   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   pip install -r requirements.txt
   streamlit run app.py
   ```

3. Dashboarden öppnas i webbläsaren och läser automatiskt in den senast ändrade
   relevanta `.xlsx`-filen i mappen. Byt ut filen och ladda om sidan för att
   räkna om med ny data — ingen kod behöver ändras.

## Beräkningslogik

- **Signed Open Amount**: `Amount Remaining` för Invoice, `-Amount Remaining` för
  Credit Memo (credit memos minskar exponeringen).
- **Days Open**: `Beräkningsdatum - Date` (fakturadatum, inte förfallodatum).
- **Reserve %** enligt aging-trappan nedan, tillämpad på Days Open.
- **Bad Debt Reserve** = `Signed Open Amount × Reserve %`.

| Days Open | Reserve % |
|---|---:|
| 0–30 | 0,5 % |
| 31–60 | 1,0 % |
| 61–90 | 4,5 % |
| 91–180 | 10,0 % |
| 181–300 | 25,0 % |
| 301–500 | 50,0 % |
| 501–800 | 60,0 % |
| 801+ | 70,0 % |

## Bokföring

Bad debt-justeringen konteras:

- **Debet 6350** / **Kredit 1515** vid ökad reserv
- **Debet 1515** / **Kredit 6350** vid minskad reserv

Department och Location används inte.

## Innehåll i repot

Endast kod och konfiguration ingår — inga NetSuite-utdrag, interna Word-dokument
eller skärmdumpar. Lägg din egen exempel- eller produktionsfil lokalt enligt
"Kom igång" ovan.
