# Netflix Content Analysis

A polished, dependency-free MVP of the Netflix Content Analysis Platform described in the supplied PRD.

## Included

- Catalogue overview with title, movie, TV, and release KPIs
- Cross-dashboard filters for type, region, genre, and release year
- Release-year trends, runtime distribution, genre momentum, and an illustrative five-year outlook
- Geographic content footprint, regional mix, and genre concentration
- Cast/director leaderboards and collaboration intelligence
- Searchable content library and filtered CSV export
- CSV/JSON dataset import with drag-and-drop, validation, safe text rendering, and a downloadable template
- Responsive desktop, tablet, and mobile layouts
- Real dataset with 5,850 titles and 77,801 cast/director credits

## Run locally

Open `index.html` directly in a browser, or serve the directory:

```powershell
python -m http.server 8000
```

Then visit `http://localhost:8000`.

## Import another dataset

Select **Import data** in the header. You can download the included CSV template or upload a CSV/JSON file with:

- Required: `title` and `release_year`/`year`
- Recommended: `type`, `genres`/`listed_in`, `country`, `rating`, `runtime`, `director`, and `cast`
- Limits: CSV or JSON, up to 15 MB and 25,000 rows

Imported data is used for the current browser session. Select **Restore bundled data** to return to the included catalogue.

## Project structure

- `index.html` - application shell and accessible dashboard markup
- `styles.css` - Netflix-inspired responsive design system
- `app.js` - demo data model, filtering, analytics, charts, views, and export
- `data/titles.csv` - source title metadata
- `data/credits.csv` - source cast and director credits
- `data/catalogue.js` - normalized browser-ready catalogue generated from the CSV files
- `scripts/build_catalogue.py` - repeatable data preparation script
- `Netflix_Content_Analysis_PRD.docx` - source product requirements

## Production boundaries

This is a frontend product MVP, not the full 20-week production architecture. The supplied dataset ends in 2022 and does not contain Netflix catalogue-addition dates or regional availability, so the interface uses release-year and production-country analysis. The forecast is an illustrative frontend baseline rather than a trained Prophet/ARIMA model. Production delivery would additionally require the PRD's Airflow/dbt pipeline, warehouse, FastAPI service, authentication, trained forecasting models, scheduled reporting, observability, and deployment infrastructure.
