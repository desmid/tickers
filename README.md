### Introduction

LibreOffice/OpenOffice macros and associated Python packages for scraping
Yahoo prices for: stock tickers, currency pairs, indices. A work in progress -
the code works in live spreadsheets, but this is essentially a demo repository
for amusement.

### Run the demo spreadsheet

To try a pre-built spreadsheet with a selection of tickers, load:

  `Templates/test.ods`

into LibreOffice (OpenOffice should also work but was not tested) and enable
macros if necessary, then click some buttons.

Creates a debugging logfile `out.log` in the session startup directory of
LO/OO.

### Functional requirements

- Update a LO/OO spreadsheet containing columns of Yahoo tickers with a button
  that populates specified columns with market price and currency.
- Data columns will be in register with the ticker column, typically adjacent,
  but may be anywhere else in the sheet and in any order.
- Recognise a ticker in a spreadsheet cell as a distinct uppercased word
  of form:
  <PRE>
    Stock            FX pair              Index
    -------------------------------------------
    BP               EURUSD=X             ^FTSE
    BP.              EUR/USD              FTSE
    BP.L             EUR:USD
                     EURUSD
  </PRE>
- Ignore other text not matching these types of pattern.
- Web fetches should be robust and produce an informative message on failure.
  
### Design requirements

- Macros must be embedded within the spreadsheet.
- Macros must use human-readable spreadsheet coordinates like `A1`,
  `B10:B100`, `C`.
- Spreadsheet operations should be abstracted behind a facade.
- Code should be structured and reusable, not a monolithic blob.
- Code should be extendable to other data sources.
- Code should use a logging facility.

### Design issues

LO/OO Python macro bindings in a spreadsheet can reference an external system
directory called `pythonpath`. Spreadsheet embedded macros are implemented in
an `ods` spreadsheet in a different way: the spreadsheet file is a zip archive
with Python macros typically placed under `Scripts/python` in a single file
with appropriate entries in `META-INF/manifest.xml`.

Placing multiple files under `Scripts/python` presents a problem as they are
not added to the Python package search path. This is of course relative to the
root of the zip archive, and not to any place on the host file system. It
turns out that a relative path inside the zipfile can be inserted dynamically
into the Python path, allowing a sensible embedded package hierarchy.

The solution was found in a technical report: "Umstieg von Basic auf Python
mit AOO unter Linux, Ein Erfahrungsbericht mit Makros, die in einem
Calc-Dokument eingebettet sind"
[pdf](https://www.uni-due.de/~abi070/files/OOo/Erfahrungsbericht.pdf "") by
Volker Lenhardt at Universität Duisburg-Essen, Germany.

### External dependencies

- Installation script for UNIX/Linux systems uses `sh`.
- LO/OO running Python 2 or 3.
- Tested in LO 5.0.6.2 using Python 3.4.3.

### Python package dependencies

- Uses `uno`, `logging` in the run-time and `unittest` in the tests.
- Probably should use `json` and `requests` but doesn't yet...

### Layout

<PRE>
     .
     ├── LoadPrices.py
     ├── README.md
     ├── Templates
     │   ├── empty.ods
     │   └── test.ods
     ├── manifest.xml
     ├── nosetests
     ├── pack
     ├── pythonpath
     │   ├── libreoffice
     │   │   ├── __init__.py
     │   │   ├── cell.py
     │   │   ├── cellrange.py
     │   │   ├── controls.py
     │   │   └── datasheet.py
     │   ├── sites
     │   │   ├── __init__.py
     │   │   └── yahoo.py
     │   └── web
     │       ├── __init__.py
     │       └── httpagent.py
     └── tests
         ├── __init__.py
         ├── libreoffice
         │   ├── __init__.py
         │   ├── test_Cell.py
         │   └── test_CellRange.py
         └── sites
             ├── __init__.py
             ├── data_yahoo_json.py
             └── test_yahoo_PriceDict.py
</PRE>

### Tests

In the top-level directory, run:

  `python3 -m unittest discover -v`

or, if installed, use `nose2-3` (the Python 3 version of `nose` - see
`./nosetests` for usage).

### Install

The demo includes a `pack` script which should be run in the *parent
directory* above this top-level. It takes the vanilla spreadsheet
`Templates/empty.ods`, inserts the code and `manifest.xml`, and writes a
functioning spreadsheet into a new `test.ods` in the now current directory.

### Acknowledgements

The tool was inspired by `simpleYahooPriceScrape.ods` on the
[LemonFool UK investor site](https://www.lemonfool.co.uk "LemonFool"). The LO/OO
messagebox functionality was adapted from that tool, but otherwise the design
is different.

END
