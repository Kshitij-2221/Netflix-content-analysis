import ast
import csv
import json
from collections import defaultdict
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data"

COUNTRY_NAMES = {
    "AR": "Argentina", "AT": "Austria", "AU": "Australia", "BE": "Belgium",
    "BR": "Brazil", "CA": "Canada", "CH": "Switzerland", "CL": "Chile",
    "CN": "China", "CO": "Colombia", "CZ": "Czech Republic", "DE": "Germany",
    "DK": "Denmark", "EG": "Egypt", "ES": "Spain", "FI": "Finland",
    "FR": "France", "GB": "United Kingdom", "HK": "Hong Kong", "HU": "Hungary",
    "ID": "Indonesia", "IE": "Ireland", "IL": "Israel", "IN": "India",
    "IS": "Iceland", "IT": "Italy", "JP": "Japan", "KR": "South Korea",
    "MX": "Mexico", "MY": "Malaysia", "NG": "Nigeria", "NL": "Netherlands",
    "NO": "Norway", "NZ": "New Zealand", "PH": "Philippines", "PL": "Poland",
    "PT": "Portugal", "RO": "Romania", "RU": "Russia", "SE": "Sweden",
    "SG": "Singapore", "TH": "Thailand", "TR": "Turkey", "TW": "Taiwan",
    "UA": "Ukraine", "US": "United States", "ZA": "South Africa",
}

REGIONS = {
    "North America": {"US", "CA", "MX"},
    "LATAM": {"AR", "BO", "BR", "CL", "CO", "CR", "CU", "DO", "EC", "GT", "PE", "PR", "PY", "UY", "VE"},
    "APAC": {"AU", "BD", "CN", "HK", "ID", "IN", "JP", "KR", "MY", "NZ", "PH", "PK", "SG", "TH", "TW", "VN"},
    "Africa": {"DZ", "EG", "ET", "GH", "KE", "MA", "NG", "SN", "TN", "TZ", "UG", "ZA", "ZW"},
}


def parse_list(value):
    if not value:
        return []
    try:
        parsed = ast.literal_eval(value)
        return parsed if isinstance(parsed, list) else []
    except (ValueError, SyntaxError):
        return []


def region_for(code):
    for region, codes in REGIONS.items():
        if code in codes:
            return region
    return "Europe / Other"


def title_case_genre(value):
    return value.replace("_", " ").title()


credits = defaultdict(lambda: {"directors": [], "cast": []})
with (DATA_DIR / "credits.csv").open(encoding="utf-8-sig", newline="") as handle:
    for row in csv.DictReader(handle):
        target = credits[row["id"]]
        if row["role"] == "DIRECTOR" and len(target["directors"]) < 3:
            target["directors"].append(row["name"])
        elif row["role"] == "ACTOR" and len(target["cast"]) < 8:
            target["cast"].append(row["name"])

catalogue = []
with (DATA_DIR / "titles.csv").open(encoding="utf-8-sig", newline="") as handle:
    for row in csv.DictReader(handle):
        genre_values = [title_case_genre(item) for item in parse_list(row["genres"])]
        country_codes = parse_list(row["production_countries"])
        primary_code = country_codes[0] if country_codes else ""
        contributors = credits[row["id"]]
        title_type = "Movie" if row["type"] == "MOVIE" else "TV Show"

        catalogue.append({
            "id": row["id"],
            "title": row["title"],
            "type": title_type,
            "genre": genre_values[0] if genre_values else "Unclassified",
            "genres": genre_values,
            "country": COUNTRY_NAMES.get(primary_code, primary_code or "Unknown"),
            "countryCode": primary_code,
            "region": region_for(primary_code) if primary_code else "Unknown",
            "year": int(row["release_year"]) if row["release_year"] else 0,
            "rating": row["age_certification"] or "Not rated",
            "duration": int(float(row["runtime"])) if row["runtime"] else 0,
            "seasons": int(float(row["seasons"])) if row["seasons"] else 0,
            "director": ", ".join(contributors["directors"]) or "Unknown",
            "cast": contributors["cast"],
            "imdbScore": float(row["imdb_score"]) if row["imdb_score"] else None,
            "imdbVotes": int(float(row["imdb_votes"])) if row["imdb_votes"] else 0,
            "tmdbPopularity": float(row["tmdb_popularity"]) if row["tmdb_popularity"] else None,
            "description": row["description"] or "",
        })

output = DATA_DIR / "catalogue.js"
with output.open("w", encoding="utf-8", newline="\n") as handle:
    handle.write("window.NETFLIX_CATALOGUE = ")
    json.dump(catalogue, handle, ensure_ascii=False, separators=(",", ":"))
    handle.write(";\n")

print(f"Wrote {len(catalogue):,} titles to {output}")
