import sys
import json
import os
import unicodedata
from collections import defaultdict
from difflib import get_close_matches

import pandas as pd
BOM_FILE = None
EPMS_FILE = None

CONFIG_FILE = "config/MainConfig.xlsx"

REFERENCE_FILE = "config/ReferenceSubcon&Region.xlsx"


NORMALIZATION_SHEET = "Equipment_Normalization"


# =========================================================
# CLEAN TEXT
# =========================================================

def clean_text(value):

    if pd.isna(value):
        return ""

    text = str(value)

    text = unicodedata.normalize("NFKC", text)

    text = text.replace("\n", " ")
    text = text.replace("\r", " ")

    text = " ".join(text.split())

    return text.strip()


# =========================================================
# NORMALIZATION MAP
# =========================================================

def build_normalization_map():

    df = pd.read_excel(
        CONFIG_FILE,
        sheet_name=NORMALIZATION_SHEET
    )

    normalization_map = {}

    for _, row in df.iterrows():

        item_list = clean_text(row["Item List"])

        item_key = clean_text(row["ItemKey"])

        if item_list and item_key:

            normalization_map[item_list] = item_key

    return normalization_map


# =========================================================
# REGION + CONTRACT REFERENCE
# =========================================================

def load_reference_tables():

    df = pd.read_excel(
        REFERENCE_FILE,
        header=None
    )

    # =====================================================
    # REGION -> PURCHASING AREA
    # =====================================================

    region_map = {}

    for _, row in df.iterrows():

        region = clean_text(row[1])

        purchasing_area = clean_text(row[2])

        if not region:
            continue

        if not purchasing_area:
            continue

        lower_region = region.lower()
        lower_area = purchasing_area.lower()

        # skip headers
        if "region" in lower_region:
            continue

        if "purchasing area" in lower_area:
            continue

        region_map[region] = purchasing_area

    # =====================================================
    # OFFICIAL SUBCON -> CONTRACT NUMBER
    # =====================================================

    subcon_contract_map = {}

    for _, row in df.iterrows():

        official_subcon = clean_text(row[4])

        contract_number = clean_text(row[5])

        if not official_subcon:
            continue

        if not contract_number:
            continue

        lower_subcon = official_subcon.lower()
        lower_contract = contract_number.lower()

        # skip headers
        if "subcontractor" in lower_subcon:
            continue

        if "contract number" in lower_contract:
            continue

        subcon_contract_map[
            official_subcon.upper()
        ] = {
            "official_name": official_subcon,
            "contract_number": contract_number,
        }

    print(f"Loaded region mappings: {len(region_map)}")

    print(f"Loaded subcon contracts: {len(subcon_contract_map)}")

    return (
        region_map,
        subcon_contract_map
    )


# =========================================================
# LOAD EPMS LOOKUP
# =========================================================

def load_epms_lookup():

    print("Loading EPMS lookup...")

    if not os.path.exists(EPMS_FILE):

        print(
            "WARNING: EPMS file not found. "
            "Using empty lookup."
        )

        return {}


    excel = pd.ExcelFile(EPMS_FILE)

    sheet_name = excel.sheet_names[0]

    df = pd.read_excel(
        EPMS_FILE,
        sheet_name=sheet_name,
        header=3
    )

    print(f"Loaded EPMS rows: {len(df)}")

    epms_lookup = {}

    for _, row in df.iterrows():

        site_code = clean_text(
            row.get("customer site code", "")
        ).upper()

        du_code = clean_text(
            row.get("du code", "")
        ).upper()

        subcon_ti = clean_text(
            row.get("SubCon - TI", "")
        )

        if not site_code:
            continue

        if not du_code:
            continue

        if not subcon_ti:
            continue

        key = (
            site_code,
            du_code
        )

        epms_lookup[key] = subcon_ti

    print(f"Loaded EPMS lookup keys: {len(epms_lookup)}")

    return epms_lookup


# =========================================================
# DETECT BOM COLUMNS
# =========================================================

def detect_site_columns(df):

    columns = list(df.columns)

    site_code_col = None
    site_name_col = None
    region_col = None
    du_col = None

    for col in columns:

        lower = clean_text(col).lower()

        if "site code" in lower:

            site_code_col = col

        elif "site name" in lower:

            site_name_col = col

        elif lower == "region":

            region_col = col

        elif "du code" in lower:

            du_col = col

    return (
        site_code_col,
        site_name_col,
        region_col,
        du_col,
    )
# =========================================================
# BOM LAST UPDATED DATE
# =========================================================

def get_bom_last_updated_date(df):

    for col in df.columns:

        col_name = clean_text(col).lower()

        if "lastupdated" in col_name:

            series = pd.to_datetime(
                df[col],
                errors="coerce"
            )

            max_date = series.max()

            if pd.notna(max_date):

                return max_date.strftime(
                    "%Y-%m-%d %H:%M:%S"
                )

    return ""

# =========================================================
# MAIN
# =========================================================

def main():

    global BOM_FILE
    global EPMS_FILE

    BOM_FILE = os.environ.get("BOM_FILE_PATH", "input/BOM.xlsx")

    EPMS_FILE = os.environ.get("EPMS_FILE_PATH", "input/EPMS.xlsx")

    print("Loading normalization config...")

    normalization_map = build_normalization_map()

    print(f"Loaded {len(normalization_map)} normalization rules")

    print("Loading reference tables...")

    (
        region_map,
        subcon_contract_map
    ) = load_reference_tables()

    print("Loading EPMS lookup...")

    epms_lookup = load_epms_lookup()

    print("Loading BOM...")
    print(f"BOM FILE = {BOM_FILE}")

    excel = pd.ExcelFile(
    BOM_FILE,
    engine="openpyxl"
    )

    sheet_name = excel.sheet_names[0]

    print(f"Using BOM sheet: {sheet_name}")

    df = pd.read_excel(
    BOM_FILE,
    sheet_name=sheet_name,
    header=2,
    engine="openpyxl"
    )

    print(f"Loaded BOM rows: {len(df)}")

    bom_last_updated_date = (
        get_bom_last_updated_date(df)
    )

    print(
        f"BOM Last Updated Date: "
        f"{bom_last_updated_date}"
    )

    (
        site_code_col,
        site_name_col,
        region_col,
        du_col,
    ) = detect_site_columns(df)
    
    results = {}

    equipment_columns = []

    print("\nChecking BOM columns against normalization map...\n")

    for col in df.columns:

        cleaned_col = clean_text(col)

        if cleaned_col in normalization_map:

            equipment_columns.append(col)

            print(
                f"{cleaned_col} "
                f"-> "
                f"{normalization_map[cleaned_col]}"
            )

    print(f"\nMatched equipment columns: {len(equipment_columns)}")

    # =====================================================
    # PROCESS BOM
    # =====================================================

    for _, row in df.iterrows():

        site_code = clean_text(
            row.get(site_code_col, "")
        ).upper()

        if not site_code:
            continue

        site_name = clean_text(
            row.get(site_name_col, "")
        )

        region = clean_text(
            row.get(region_col, "")
        )

        du_code = clean_text(
            row.get(du_col, "")
        ).upper()

        # -------------------------------------------------
        # EPMS LOOKUP
        # -------------------------------------------------

        lookup_key = (
            site_code,
            du_code
        )

        epms_subcon = epms_lookup.get(
            lookup_key,
            ""
        )

        # -------------------------------------------------
        # USE EPMS SUBCON DIRECTLY
        # -------------------------------------------------

        subcon_ti = epms_subcon

        # -------------------------------------------------
        # PURCHASING AREA
        # -------------------------------------------------

        purchasing_area = region_map.get(
            region,
            ""
        )

        # -------------------------------------------------
        # CONTRACT NUMBER
        # -------------------------------------------------

        contract_number = ""

        if subcon_ti:

            exact_key = subcon_ti.upper()

            # exact match
            if exact_key in subcon_contract_map:

                subcon_info = subcon_contract_map[
                    exact_key
                ]

                subcon_ti = subcon_info[
                    "official_name"
                ]

                contract_number = subcon_info[
                    "contract_number"
                ]

            else:

                available = list(
                    subcon_contract_map.keys()
                )

                matches = get_close_matches(
                    exact_key,
                    available,
                    n=1,
                    cutoff=0.6
                )

                if matches:

                    matched = matches[0]

                    subcon_info = subcon_contract_map[
                        matched
                    ]

                    subcon_ti = subcon_info[
                        "official_name"
                    ]

                    contract_number = subcon_info[
                        "contract_number"
                    ]

                    print(
                        f"Fuzzy subcon match: "
                        f"{exact_key} -> {matched}"
                    )

        # -------------------------------------------------
        # CREATE SITE OBJECT
        # -------------------------------------------------

        if site_code not in results:

            results[site_code] = {

                "metadata": {

                    "site_code":
                        site_code,

                    "site_name":
                        site_name,

                    "region":
                        region,

                    "du_code":
                        du_code,

                    "subcon_ti":
                        subcon_ti,

                    "contract_number":
                        contract_number,

                    "purchasing_area":
                        purchasing_area,

                    "bom_last_updated_date":
                        bom_last_updated_date,
                },

                "quantities":
                    defaultdict(float)
            }

        # -------------------------------------------------
        # QUANTITY NORMALIZATION
        # -------------------------------------------------

        for col in equipment_columns:

            value = row.get(col)

            if pd.isna(value):
                continue

            try:
                qty = float(value)

            except:
                continue

            if qty == 0:
                continue

            cleaned_col = clean_text(col)

            item_key = normalization_map[
                cleaned_col
            ]

            results[site_code]["quantities"][
                item_key
            ] += qty

    # =====================================================
    # CONVERT DEFAULTDICT
    # =====================================================

    for site_code in results:

        results[site_code]["quantities"] = dict(
            results[site_code]["quantities"]
        )

    # =====================================================
    # EXPORT
    # =====================================================

    OUTPUT_FILE = os.environ.get(
        "NORMALIZED_OUTPUT_FILE",
        "output/simple_normalized.json"
    )

    with open(
        OUTPUT_FILE,
        "w",
        encoding="utf-8"
    ) as f:

        json.dump(
            results,
            f,
            indent=2
        )

    print("\nDONE")

    print(
        f"Exported: "
        f"{OUTPUT_FILE}"
    )

if __name__ == "__main__":

    main()