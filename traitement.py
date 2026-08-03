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
SPECIAL_NAMES = {

    # Existing
    "ESPLANADE DE L'": "de l'Esplanade",
    "LORIMIER DE": "De Lorimier",
    "CASTELNAU DE": "De Castelnau",
    "FONTAINE LA": "La Fontaine",
    "IBERVILLE D'": "D'Iberville",
    "ARLES D'": "D'Arles",
    "GALETS DES": "Des Galets",
    "LAMARTINE DE": "De Lamartine",
    "LIEGE DE": "De Liège",

    # Saint...
    "ST-HUBERT": "Saint-Hubert",
    "ST-DENIS": "Saint-Denis",
    "ST-LAURENT": "Saint-Laurent",
    "ST-MICHEL": "Saint-Michel",
    "ST-MARC": "Saint-Marc",
    "ST-PATRICK": "Saint-Patrick",

    # Sainte...
    "STE-CATHERINE": "Sainte-Catherine",
    "STE-CLAIRE": "Sainte-Claire",

    # De...
    "ACADIE DE L'": "de l'Acadie",
    "BELLECHASSE DE": "De Bellechasse",
    "BRETAGNE DE": "De Bretagne",
    "CADILLAC DE": "De Cadillac",
    "CASTILLE DE": "De Castille",
    "CHATEAUBRIAND DE": "De Chateaubriand",
    "COMPIEGNE DE": "De Compiègne",
    "COTE-DE-LIESSE DE LA": "de la Côte-de-Liesse",
    "COTE-ST-LUC DE LA": "de la Côte-Saint-Luc",
    "COTE-STE-CATHERINE DE LA": "de la Côte-Sainte-Catherine",
    "DORCHESTER": "Dorchester",
    "ECORES DES": "Des Écores",
    "HONORE-BEAUGRAND": "Honoré-Beaugrand",
    "INSPECTEUR DE L'": "de l'Inspecteur",
    "LACHENAIE DE": "De Lachenaie",
    "MAISONNEUVE DE": "De Maisonneuve",
    "MARSEILLE DE": "De Marseille",
    "MONTAGNE DE LA": "de la Montagne",
    "NOTRE-DAME": "Notre-Dame",
    "ORMEAUX DES": "Des Ormeaux",
    "PAIMPOL DE": "De Paimpol",
    "PIERRE-DE-COUBERTIN": "Pierre-De-Coubertin",
    "RECOLLETS DES": "Des Récollets",
    "REIMS DE": "De Reims",
    "RENTY DE": "De Renty",
    "ROSAIRE DU": "Du Rosaire",
    "ROSEMONT": "Rosemont",
    "ROUEN DE": "De Rouen",
    "SALABERRY DE": "De Salaberry",
    "SHANNON DU": "Du Shannon",
    "THOMAS-KEEFER": "Thomas-Keefer",
    "TOULOUSE DE": "De Toulouse",
    "VIGER": "Viger",
    "VITERBE DE": "De Viterbe",
    "ST-REAL DE": "De Saint-Real",
    "ARTAGNAN D'": "D'Artagnan",
    "CENTRE DU": "Centre",

    "MACQUEEN PLAC": "Place Macqueen",

    # Hyphenated names
    "CHRISTOPHE-COLOMB": "Christophe-Colomb",
    "JEAN-TALON": "Jean-Talon",
    "LOUIS-H-LA-FONTAINE": "Louis-H.-La Fontaine",
    "MADELEINE-HUGUENIN": "Madeleine-Huguenin",
    "PIE-IX": "Pie-IX",
    "RENE-LEVESQUE": "René-Lévesque",
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

    s = str(s).strip().upper()

    # Replace direction abbreviations before splitting
    for k, v in DIRECTION_MAP.items():
        s = s.replace(k, v.upper())

    tokens = s.split()

    if not tokens:
        return ""

    street_type = ""

    # Street type is the last token
    if tokens[-1] in TYPE_MAP:
        street_type = TYPE_MAP[tokens[-1]]
        tokens = tokens[:-1]

    name = " ".join(tokens)

    # Move direction from the middle to the end if needed
    # e.g. CASTELNAU DE EST -> CASTELNAU DE + Est
    direction = ""

    for d in (" EST", " OUEST", " NORD", " SUD"):
        if name.endswith(d):
            direction = d.title().strip()
            name = name[:-len(d)].strip()
            break

    # Use exact exception if available
    if name in SPECIAL_NAMES:
        name = SPECIAL_NAMES[name]
    else:
        name = name.title()

    if direction:
        name += f" {direction}"

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