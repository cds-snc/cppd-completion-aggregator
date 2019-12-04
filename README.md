# cppd-completion-aggregator

Very quick and dirty script to parse CPPD completion tracking Excel sheets. Not really productionized, and probably never will have to be. Delete or archive this repo if it's here a year without modification.


# Setup

Requires Python3 (Mac: `brew install python`)

First time only:

1. Clone this repo and navigate in
1. Setup virtual environment: `python3 -m venv .venv`
1. Activate virtual env: `source .venv/bin/activate`
1. Install dependencies: `pip install -r requirements.txt`

All other times:

1. Activate virtual env: `source .venv/bin/activate`
1. Drop all the spreadsheets into a folder, without any other kinds of files going into that folder
1. Run the script: `python parser.py (folder path) > data.csv`
1. Checkout your aggregated data in Excel / Google Spreadsheets with `data.csv`