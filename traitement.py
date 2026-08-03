import pandas as pd
import requests
import re
import time
from tqdm import tqdm

# =====================================================
# SETTINGS
# =====================================================

INPUT_FILE = "Données Excel 26-200167.xlsx"
OUTPUT_FILE = "Données Excel 26-200167_traitées.xlsx"

API_KEY = "YOUR_GOOGLE_API_KEY"

DRY_RUN = True       # <- True = don't call Google
PAUSE = 0.05         # seconds between requests

# =====================================================
# Street name normalization
# =====================================================

TYPE_MAP = {
    "RUE": "Rue",
    "AVEN": "Avenue",
    "BOUL": "Boulevard",
    "PLACE": "Place",
    "CHEM": "Chemin",
}

DIRECTION_MAP = {
    " E ": " Est ",
    " O ": " Ouest ",
    " N ": " Nord ",
    " S ": " Sud ",
}


def normalize_street(s):

    if pd.isna(s):
        return ""

    s = s.strip().upper()

    # Replace directions
    for k, v in DIRECTION_MAP.items():
        s = s.replace(k, v.upper())

    tokens = s.split()

    if len(tokens) == 0:
        return ""

    street_type = ""

    if tokens[-1] in TYPE_MAP:
        street_type = TYPE_MAP[tokens[-1]]
        tokens = tokens[:-1]

    name = " ".join(tokens)

    # Est/Ouest at end
    name = name.replace(" EST", " Est")
    name = name.replace(" OUEST", " Ouest")
    name = name.replace(" NORD", " Nord")
    name = name.replace(" SUD", " Sud")

    # Fix capitalization
    name = name.title()

    # Common apostrophes
    name = name.replace(" D'", " d'")
    name = name.replace(" L'", " l'")
    name = name.replace(" De ", " de ")
    name = name.replace(" Des ", " des ")
    name = name.replace(" Du ", " du ")
    name = name.replace(" La ", " la ")
    name = name.replace(" Le ", " le ")

    return f"{street_type} {name}".strip()


# =====================================================
# Build Google query
# =====================================================

def make_query(row):

    street = normalize_street(row["Rue - Infraction"])

    inter = row["Intersection"]

    civ = row.get("#Civ.", "")

    if pd.notna(civ):
        try:
            civ = str(int(float(civ)))
        except (ValueError, TypeError):
            civ = str(civ).strip()
    else:
        civ = ""

    if pd.notna(inter) and str(inter).strip() != "":

        inter = normalize_street(inter)

        return f"{street} & {inter}, Montréal, Québec"

    if pd.notna(civ) and str(civ).strip() != "":

        return f"{civ} {street}, Montréal, Québec"

    return f"{street}, Montréal, Québec"


# =====================================================
# Google Geocoder
# =====================================================

def geocode(query):

    url = "https://maps.googleapis.com/maps/api/geocode/json"

    r = requests.get(
        url,
        params={
            "address": query,
            "key": API_KEY
        }
    )

    data = r.json()

    if data["status"] != "OK":
        return None, None, data["status"]

    loc = data["results"][0]["geometry"]["location"]

    return loc["lat"], loc["lng"], "OK"


# =====================================================
# MAIN
# =====================================================

df = pd.read_excel(INPUT_FILE)

df["google_query"] = df.apply(make_query, axis=1)

unique = (
    df[["google_query"]]
    .drop_duplicates()
    .reset_index(drop=True)
)

print(f"{len(df)} rows")
print(f"{len(unique)} unique locations")

if DRY_RUN:

    print("\nQueries that would be sent:\n")

    for q in unique.google_query:
        print(q)

    unique.to_csv("google_queries_preview.csv", index=False)

    print("\nSaved preview to google_queries_preview.csv")

    quit()

# -----------------------------------------------------

cache = {}

for q in tqdm(unique.google_query):

    lat, lon, status = geocode(q)

    cache[q] = (lat, lon, status)

    time.sleep(PAUSE)

df["lat"] = df.google_query.map(lambda x: cache[x][0])

df["lon"] = df.google_query.map(lambda x: cache[x][1])

df["status"] = df.google_query.map(lambda x: cache[x][2])

df.to_excel(OUTPUT_FILE, index=False)

print("Done.")