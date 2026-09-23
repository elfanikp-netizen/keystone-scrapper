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
