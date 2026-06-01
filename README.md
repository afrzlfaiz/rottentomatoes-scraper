# Rotten Tomatoes Scraper

A Python CLI tool to scrape Rotten Tomatoes reviews (audience and critic) and export results to JSON, CSV, and Excel.

## Features

- Fetch reviews by **movie/TV URL** or **movie ID**
- Supports review type:
  - `audience`
  - `critic`
  - `both`
- **Batch scraping** — scrape multiple films/shows in one run, results appended to the same output file
  - Interactive: enter first film with full settings, then loop "Add another?" with URL + count only
  - CLI: `--batch-file` with a text file of URLs (one per line)
- Interactive mode (no arguments)
- Export formats:
  - `json`
  - `csv`
  - `excel`
  - `all`
- Optional filters:
  - `--verified` (audience only)
  - `--top-only`

## Requirements

- Python 3.10+
- Dependencies listed in `requirements.txt`

## Setup

### Windows

```bat
setup.bat
```

### Linux/macOS/Git Bash

```bash
sh setup.sh
```

## Run

### Windows

```bat
start.bat
```

`start.bat` will wait for a key press before closing the terminal window.

### Linux/macOS/Git Bash

```bash
sh start.sh
```

## CLI Usage

```bash
python scrape_rottentomatoes.py [movie_id] [options]
```

### Options

- `--url` Rotten Tomatoes URL (movie or TV)
- `--batch-file <path>` text file with one RT URL per line for batch scraping (all films use the same `--type`, `--count`, `--format`)
- `--type {audience,critic,both}` (default: `both`)
- `--count <int>` number of reviews per type (default: `20`)
- `--verified` audience only: fetch verified reviews only
- `--top-only` fetch top reviews only if available
- `--output <path>` base output file path/name
- `--format {json,csv,excel,all}` (default: `json`)

## Examples

Fetch both review types from URL and save all formats:

```bash
python scrape_rottentomatoes.py --url "https://www.rottentomatoes.com/m/lee_cronins_the_mummy" --type both --count 50 --output output/reviews --format all
```

Fetch critic reviews by movie ID and save JSON:

```bash
python scrape_rottentomatoes.py 771306662 --type critic --count 30 --output output/critic_reviews --format json
```

Batch scrape multiple films from a file (all use same `--type`, `--count`, `--format`; results appended row-by-row):

```bash
python scrape_rottentomatoes.py --batch-file urls.txt --type critic --count 30 --output output/batch_critics --format csv
```

Run without arguments (interactive mode with batch support):

```bash
python scrape_rottentomatoes.py
```

In interactive mode you'll enter the first film's URL, choose review type/format/output, then optionally add more films (URL + count each) whose results are appended to the same file.

## Output Notes

- `json`: single JSON file
- `csv`: one CSV file per review type (e.g. `reviews_audience.csv`, `reviews_critic.csv`)
- `excel`: one `.xlsx` file with separate sheets per review type
- **Batch mode**: the first film creates the output file with headers; subsequent films append rows without re-adding headers. All films in a batch share the same review type to keep column format consistent.

## Disclaimer

This tool uses Rotten Tomatoes public endpoints/pages. Availability and response format may change at any time.
