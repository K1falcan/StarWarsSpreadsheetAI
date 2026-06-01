# Star Wars SpreadsheetAI

Generates a formatted Excel workbook of Star Wars data, pulling everything live
from the **[swapi.info](https://swapi.info)** API. No data is hard-coded — every
cell comes from `https://swapi.info/api`.

## What it builds

A single `star_wars.xlsx` workbook with one sheet per category:

| Sheet | Contents |
| --- | --- |
| **Characters** | People, with homeworld, species, ships, vehicles & films resolved to names |
| **Planets** | Climate, terrain, population, residents, films |
| **Species** | Classification, language, homeworld, members, films |
| **Starships** | Model, manufacturer, specs, pilots, films |
| **Vehicles** | Model, manufacturer, specs, pilots, films |
| **Films (Chronological)** | The movies ordered by in-universe episode number |
| **Films (Release Order)** | The movies ordered by real-world release date |

The API links resources to each other by URL (e.g. a character's homeworld is a
`.../planets/1` link); the generator resolves those links to readable names so
the spreadsheet is human-friendly. Each sheet has a styled header, frozen top
row, auto-filter, and auto-sized columns.

## Setup

Requires Python 3.9+.

```bash
python3 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Run

```bash
python main.py                     # writes ./star_wars.xlsx
python main.py --output out/sw.xlsx # custom location
python main.py --verbose           # log every API request
```

Run `python main.py --help` for all options.

## Debugging in VS Code

1. Install the **Python** extension.
2. Open this folder. The interpreter is preset to `.venv` via `.vscode/settings.json`.
3. Open the **Run & Debug** panel (⇧⌘D / Ctrl+Shift+D), choose
   **Build Star Wars Spreadsheet** (or the verbose variant), and press **F5**.
4. Set breakpoints anywhere in `src/` — `justMyCode` is off, so you can step
   into the request and openpyxl internals too.

## Project layout

```
main.py                 # CLI entry point
src/
  swapi_client.py       # fetches collections + resolves URL cross-references
  spreadsheet.py        # column definitions + workbook/styling
.vscode/                # launch + interpreter settings for debugging
requirements.txt
```

## Data source

All data: [swapi.info](https://swapi.info) — `https://swapi.info/api`.
