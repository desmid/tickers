### Introduction

This code provides LibreOffice/OpenOffice (LO/OO) macros and associated Python
packages for scraping Yahoo prices for stock tickers, currency pairs, or
indices.

It is a work in progress - the code works in live spreadsheets, but this is
essentially a demo.

### Run the demo spreadsheet

To try a pre-built spreadsheet with a selection of tickers, load:

  `./test.ods`

into LibreOffice (OpenOffice should also work but was not tested) and enable
macros if necessary, then click some buttons.

Creates a debugging logfile `out.log` in the session startup directory of
LO/OO.

### Functional requirements

- Populate a LO/OO spreadsheet containing columns of Yahoo tickers with their
  market price and currency.
  
- The data values will be fetched from Yahoo when the appropriate button is
  clicked.

- The output data columns will usually lie alongside the ticker column, but
  may be anywhere else in the sheet and the columns can be in any order.

- A ticker in a spreadsheet cell should be recognised as a distinct uppercased
  word of form:
  
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

- The application should be self-contained, i.e., macros must be embedded
  within the spreadsheet not stored in external files.

- Macros must use human-readable spreadsheet coordinates like `A1`,
  `B10:B100`, `C`.

- Spreadsheet operations should be abstracted for code portability to other
  systems.

- Code should be structured and extendable to other data sources.

- Code shoud have an optional logging facility for debugging.

### Design issues

LO/OO Python macro bindings in a spreadsheet can reference an external system
directory called `pythonpath`, but this would break the design requirement of
being self-contained.

Alternatively, spreadsheet embedded macros can be implemented in an `ods`
spreadsheet in a different way: the spreadsheet file is a zip archive with
Python macros typically placed under `Scripts/python` in a single file with
appropriate entries in `META-INF/manifest.xml`.

Problem:

Placing multiple files under `Scripts/python` presents a problem as they are
not added to the Python package search path. This is of course relative to the
root of the zip archive, and not to any place on the host file system. It
turns out that a relative path inside the zipfile can be inserted dynamically
into the Python path, allowing a sensible embedded package hierarchy [1].

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
      ├── LoadPrices.py            #macros
      ├── README.md
      ├── Templates
      │   ├── empty.ods
      │   └── manifest.xml
      ├── nosetests
      ├── out.log
      ├── pack
      ├── pythonpath               #packages
      │   ├── sites
      │   │   ├── __init__.py
      │   │   └── yahoo.py
      │   ├── spreadsheet
      │   │   ├── __init__.py
      │   │   ├── api
      │   │   │   ├── __init__.py
      │   │   │   ├── factory.py
      │   │   │   └── libreoffice.py
      │   │   ├── cell.py
      │   │   ├── cellrange.py
      │   │   └── datasheet.py
      │   └── web
      │       ├── __init__.py
      │       └── httpagent.py
      ├── rpurge
      ├── test.ods                 #test spreadsheet
      └── tests                    #unit tests
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

### Unit tests

In the top-level directory, run:

  `python3 -m unittest discover -v`

or, if installed, use `nose2-3` (the Python 3 version of `nose` - see
`./nosetests` for usage).

### Install

The demo includes a `pack` script which should be run in the top-level
directory. It takes the vanilla spreadsheet `Templates/empty.ods`, unpacks it
into a directory `t`, injects the code and `manifest.xml`, and packs it all
back up again into a functioning spreadsheet `test.ods` in the current
directory.

### Acknowledgements

The tool was inspired by `simpleYahooPriceScrape.ods` on the
[LemonFool UK investor site](https://www.lemonfool.co.uk "LemonFool"). The LO/OO
messagebox functionality was adapted from that tool, but otherwise the design
is different.

[1] The solution to the pythonpath embedding problem was found in a technical
report:

"Umstieg von Basic auf Python mit AOO unter Linux, Ein Erfahrungsbericht mit
Makros, die in einem Calc-Dokument eingebettet sind." February, 2016. Volker
Lenhardt, Universität Duisburg-Essen, Germany.
\[[pdf](https://www.uni-due.de/~abi070/files/OOo/Erfahrungsbericht.pdf "Umstieg von Basic auf Python mit AOO unter Linux")\]

END
