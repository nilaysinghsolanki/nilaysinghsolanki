# netsearch

Offline search for people, customers and accounts across your Excel and CSV lists.
It runs entirely on your own PC: no internet, no API calls, no Claude credits.

Your spreadsheets are read once into a local SQLite full-text index
(`netsearch.db`). After that, each search takes a few milliseconds, even with
hundreds of thousands of rows. Only files that changed get re-read.

## Setup (once)

1. Install Python 3.9+ from python.org. On Windows, tick "Add to PATH".
2. In this folder, run:
   ```
   pip install -r requirements.txt
   ```
   `rapidfuzz` is optional. Without it, typo matching uses Python's built-in `difflib`.

## Use

Put your lists (`.xlsx`, `.xlsm`, `.csv`) in a `data` folder here, or point at any folder:

```
python netsearch.py index data "C:\Users\me\OneDrive\Network"
python netsearch.py                     # interactive search prompt
python netsearch.py serve               # search page at http://127.0.0.1:8765
python netsearch.py search titan capital --csv out.csv
python netsearch.py stats               # what is indexed
python netsearch.py prune               # forget deleted files
```

On Windows you can double-click these instead: `index-data.bat` indexes the
`data` folder, `search.bat` opens the prompt, and `search-ui.bat` opens the browser page.

## Query syntax

| Query | Finds |
|---|---|
| `rahul` | anything starting with "rahul" (names, companies, notes...) |
| `"venture capital"` | the exact phrase |
| `name:abhi org:meddo` | narrows to one field: `name`, `org`/`company`, `role`/`title`, `email`, `phone` |
| `batch:2010..2015` | any other column, matched by part of its header, as a number range |
| `sector:vc help:funding` | a substring of that column |
| `(989) 986-2689` | phone numbers, however they're formatted |
| `endya` | typos: if exact search finds little, close name/company matches are added (marked `~`) |

All terms must match (AND). Results are ranked with name matches first,
then company, role, and everything else.

## Columns

Sheets with a header row (`Name`, `Company`, `Account Name`, `Mail Id`,
`Mobile`, `Designation`...) are understood automatically. For sheets without
a header row, columns are guessed from their content (emails, phones,
LinkedIn links, years, names). For the best results, add the column names to
`schemas.json` by filename pattern. `*DTU-Network*` is already there.
