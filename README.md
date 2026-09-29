# Keystone Partslink Crawler

This project crawls the Keystone crash catalog, reads Partslink numbers from an Excel file, and exports matched result rows into an Excel report.

## Requirements

- Python 3.10+
- Windows/macOS/Linux with a local browser available to Playwright

## Setup

From the project folder, install dependencies:

```bash
python -m pip install -r requirements.txt
python -m playwright install chromium
```

Create a `.env` file in the project root if you want to avoid entering credentials each run:

```env
KEYSTONE_USERNAME=your_email_or_username
KEYSTONE_PASSWORD=your_password
```

## Prepare the input file

Create a template Excel file:

```bash
python keystone_crawler.py --make-template
```

This creates `parts.xlsx` with a sample `Partslink Number` column. Replace the example values with your actual Partslink numbers.

## Run the crawler

Run a small visible test with one part:

```bash
python keystone_crawler.py --show --limit 1
```

Run the full crawl:

```bash
python keystone_crawler.py
```

Optional filters:

```bash
python keystone_crawler.py --limit 5
python keystone_crawler.py --input my_parts.xlsx --column "Partslink Number"
python keystone_crawler.py --fresh
python keystone_crawler.py --manual-login
```

## Output

The script writes results to:

- `keystone_results.xlsx`

The output includes:

- `Partslink Number`
- `Oldest Year`
- `Newest Year`
- `Brand`
- `Model`
- `Type`
- `Interchange Number`
- `Interchange 1` to `Interchange 5`
- `OEM Number`
- `OEM 1` to `OEM 5`
- `Status`

## Troubleshooting

If Playwright says the browser is missing:

```bash
python -m playwright install chromium
```

If the script cannot import a package:

```bash
python -m pip install -r requirements.txt
```

If the input file is missing:

```bash
python keystone_crawler.py --make-template
```

## Keystone Model Crawler

`keystone_model_crawler.py` searches the Crash catalog by vehicle and model-part category, then opens each matching Partslink result and exports its fitment, description, and OEM/alternate numbers. It uses the same Keystone login credentials and saved browser session as `keystone_crawler.py`.

### Prepare the input file

Create the model-search template:

```bash
python keystone_model_crawler.py --make-template
```

This creates `find_model.xlsx` with these required columns:

- `Year`: vehicle model year
- `Brand`: vehicle make
- `Model`: vehicle model
- `Parts`: model-parts category shown in the Crash search form, such as `All Parts` or `BUMPERS (PLASTIC)`
- `Category`: one or more catalog categories, such as `CAPA`, `NSF`, or `OE`; separate multiple values with commas, for example `CAPA,NSF,OE`

Enter one search per spreadsheet row. The vehicle selections are dependent, so the Brand, Model, and Parts values must be available for the selected Year and preceding selection.

### Run the model crawler

Run one row with the browser visible first:

```bash
python keystone_model_crawler.py --show --limit 1
```

Run all rows:

```bash
python keystone_model_crawler.py
```

Optional arguments:

```bash
python keystone_model_crawler.py --input another_model_file.xlsx
python keystone_model_crawler.py --output another_results_file.xlsx
python keystone_model_crawler.py --limit 5
python keystone_model_crawler.py --fresh
python keystone_model_crawler.py --manual-login
```

`--fresh` ignores existing results. `--manual-login` opens the site for manual sign-in when automatic login is unavailable.

### Model crawler output

The default output file is `keystone_model_results.xlsx`. It contains:

- `Partslink Number`, `Oldest Year`, `Newest Year`, `Brand`, and `Model`
- `Interchange Number` and `OEM Number`, plus numbered slots 1 through 5 for each
- `Part Type`: the first three digits of the first interchange number
- `Number Values` and `Description`

The crawler scrolls through the matching results and saves each result as a separate output row.
