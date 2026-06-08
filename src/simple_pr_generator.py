import argparse
import json
import os

import pandas as pd


CONFIG_FILE = "config/MainConfig.xlsx"

CALCULATED_FILE = "output/simple_calculated.json"

OUTPUT_FILE = "output/simple_pr_output.json"

GENERAL_ITEM_CONFIG_FILE = "config/GENERAL ITEM FOR ALL DU PROJECT Overall.xlsx"
GENERAL_ITEM_OUTPUT_FILE = "output/general_pr_output.json"
COMBINED_PR_OUTPUT_FILE = "output/simple_pr_output_with_general_items.json"

MAIN_RULE_SHEET = "MainRuleTable"


def load_pr_rules():

    df = pd.read_excel(
        CONFIG_FILE,
        sheet_name=MAIN_RULE_SHEET
    )

    print("Detected MainRuleTable columns:")
    print(df.columns.tolist())

    rules = []

    for _, row in df.iterrows():

        enabled = row.get("Enabled", True)

        if pd.isna(enabled):
            enabled = True

        if enabled in [False, 0, "FALSE", "False"]:
            continue

        material_code = row.get(
            "newContractcode",
            ""
        )

        if pd.isna(material_code):
            material_code = ""
        elif isinstance(material_code, float):
            if material_code.is_integer():
                material_code = int(material_code)

        line_item_text = row.get(
            "LineItemText",
            ""
        )

        if pd.isna(line_item_text):
            line_item_text = ""

        unit = row.get(
            "Unit*",
            ""
        )

        if pd.isna(unit):
            unit = ""

        rules.append({

            "ItemKey": str(
                row.get("ItemKey", "")
            ).strip(),

            "RuleType": str(
                row.get("RuleType", "")
            ).strip(),

            "MinQty": float(
                row.get("MinQty", 0)
            ),

            "MaxQty": float(
                row.get("MaxQty", 999999)
            ),

            "LineItemText": str(
                line_item_text
            ).strip(),

            "MaterialCode": str(
                material_code
            ).strip(),

            "Enabled": enabled,

            "PR_Mode": str(
                row.get("PR_Mode", "")
            ).strip(),

            "SourceType": str(
                row.get("SourceType", "")
            ).strip(),

            "Unit": str(
                unit
            ).strip(),
        })

    return rules


def get_selected_project():
    parser = argparse.ArgumentParser(
        description="Generate PR output and optionally add general DU project mandatory line items."
    )
    parser.add_argument(
        "--selected-project",
        dest="selected_project",
        help="Selected project name to use for general DU item inclusion.",
        default=None,
    )

    args = parser.parse_args()
    selected_project = args.selected_project

    if not selected_project:
        selected_project = os.environ.get("SELECTED_PROJECT")
        if not selected_project:
            selected_project = os.environ.get("GENERAL_ITEM_PROJECT")

    if selected_project:
        selected_project = selected_project.strip()

    return selected_project


def load_general_item_config():
    df_by_region = {}

    for region_sheet in ["Central", "Sabah", "Sarawak", "Northern"]:
        df = pd.read_excel(
            GENERAL_ITEM_CONFIG_FILE,
            sheet_name=region_sheet,
            header=[0, 1],
        )

        project_columns = [
            column for column in df.columns
            if column[0] != "General Item"
        ]

        rows = []

        for _, row in df.iterrows():
            item_code = row.get(("General Item", "Item Code"))
            if pd.isna(item_code):
                continue

            item_name = row.get(("General Item", "Item Name"), "")
            item_unit = row.get(("General Item", "Unit"), "")
            location_field = None

            if region_sheet in ["Sabah", "Sarawak"]:
                location_field = row.get(("General Item", "City"), "")
            elif region_sheet == "Northern":
                location_field = row.get(("General Item", "Province/State"), "")

            project_flags = {}
            for project_col in project_columns:
                project_name = str(project_col[0]).strip()
                project_flags[project_name] = row.get(project_col)

            rows.append({
                "item_code": str(item_code).strip(),
                "item_name": str(item_name).strip(),
                "unit": str(item_unit).strip(),
                "location": str(location_field).strip(),
                "project_flags": project_flags,
            })

        df_by_region[region_sheet] = rows

    em_df = pd.read_excel(
        GENERAL_ITEM_CONFIG_FILE,
        sheet_name="EM Transportation Model",
        header=0,
    )

    em_transport_map = {}
    for _, row in em_df.iterrows():
        epms_region = str(row.get("EPMS Region", "")).strip()
        current_area = str(row.get("Current BOQ Area", "")).strip()
        if epms_region and current_area:
            em_transport_map[epms_region.lower()] = current_area
        if current_area:
            em_transport_map[current_area.lower()] = current_area

    return df_by_region, em_transport_map


def load_epms_map(path):
    df = pd.read_excel(path, sheet_name="data", header=3, engine="openpyxl")
    site_map = {}

    for _, row in df.iterrows():
        site_code = row.get("customer site code")
        if pd.isna(site_code):
            continue
        site_code = str(site_code).strip()
        site_map[site_code] = {
            "Region": str(row.get("region", "")).strip(),
            "ProvinceState": str(row.get("Province/State", "")).strip(),
            "City": str(row.get("City", "")).strip(),
        }

    return site_map


def _is_truthy_project_flag(value):
    if pd.isna(value):
        return False

    if isinstance(value, (int, float)):
        return value == 1

    value = str(value).strip().lower()
    return value in {"1", "1.0", "yes", "y", "true"}


def _normalize_text(value):
    if value is None:
        return ""
    # Treat pandas NA, NaN and literal 'nan'/'none' as empty
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass

    s = str(value).strip().lower()
    if s in {"nan", "none"}:
        return ""
    return s


def _matches_location(site_code, location, config_region, province_state, city, em_transport_map):
    location_value = _normalize_text(location)
    if not location_value:
        match = True
        print(
            f"Site={site_code} "
            f"Region={config_region} "
            f"Province={province_state} "
            f"City={city} "
            f"Location={location} "
            f"Decision={match}"
        )
        return True
    province_value = _normalize_text(province_state)
    city_value = _normalize_text(city)
    match = False

    if config_region == "Northern":
        match = province_value == location_value
    elif config_region in {"Sabah", "Sarawak"}:
        mapped_area = em_transport_map.get(city_value, "")
        match = _normalize_text(mapped_area) == location_value

    print(
        f"Site={site_code} "
        f"Region={config_region} "
        f"Province={province_state} "
        f"City={city} "
        f"Location={location} "
        f"Decision={match}"
    )
    return match


def generate_general_pr_for_site(
    site_code,
    site_data,
    general_item_config,
    selected_project,
    em_transport_map,
    epms_map,
):
    metadata = site_data.get("metadata", {})
    region = _normalize_text(metadata.get("region", ""))

    if not selected_project:
        return []

    config_region = None
    if region == "central":
        config_region = "Central"
    elif region == "sabah":
        config_region = "Sabah"
    elif region == "sarawak":
        config_region = "Sarawak"
    elif region in {"northern", "north"}:
        config_region = "Northern"

    if not config_region:
        return []

    site_name = metadata.get("site_name", "")
    site_code_text = metadata.get("site_code", site_code)
    epms_site = epms_map.get(site_code_text, {})
    province_state = epms_site.get("ProvinceState", "")
    city = epms_site.get("City", "")

    OPTIONAL_THRESHOLD_KEYS = {
        "AAU",
        "Antenna",
        "Battery",
        "BBU",
        "Bracket",
        "Cabinet",
        "Power Module",
        "Combiner",
        "PadPower",
        "Post",
        "RRU",
        "RRU Cage",
        "Security Bar",
        "Feeder 1/2",
        "Feeder 1 5/8",
        "Feeder 7/8",
    }

    general_rows = []
    seen_material_codes = set()

    for row in general_item_config.get(config_region, []):
        project_value = row["project_flags"].get(selected_project)
        item_name = row.get("item_name")
        total_qty = sum(
            site_data.get("quantities", {}).get(item_key, 0)
            for item_key in OPTIONAL_THRESHOLD_KEYS
        )

        print(
            f"Site={site_code} | "
            f"Item={item_name} | "
            f"ProjectValue={project_value} | "
            f"ThresholdQty={total_qty}"
        )

        if str(project_value).strip().lower() == "optional":
            if total_qty <= 6:
                continue
        elif not _is_truthy_project_flag(project_value):
            continue

        if config_region == "Central":
            match = True
        else:
            match = _matches_location(
                site_code_text,
                row["location"],
                config_region,
                province_state,
                city,
                em_transport_map,
            )

        if not match:
            continue

        material_code = row["item_code"]
        if material_code in seen_material_codes:
            continue

        seen_material_codes.add(material_code)

        general_rows.append({
            "MaterialCode": material_code,
            "LineItemText": row["item_name"],
            "ItemKey": row["item_name"],
            "PR_Mode": "GENERAL",
            "CalculatedQty": 1,
            "FinalPRQty": 1,
            "Unit": row["unit"],
            "Remarks": "General Item TI",
        })

    return general_rows


def evaluate_rule(quantity, rule):

    min_qty = rule["MinQty"]

    max_qty = rule["MaxQty"]

    if quantity < min_qty:
        return False

    if quantity > max_qty:
        return False

    return True


def generate_pr_for_site(site_code, site_data, rules):

    quantities = site_data["quantities"]

    pr_lines = []

    print()
    print("=" * 60)
    print(f"GENERATING PR: {site_code}")
    print("=" * 60)

    for rule in rules:

        item_key = rule["ItemKey"]

        site_qty = quantities.get(
            item_key,
            0
        )

        if site_qty == 0:
            continue

        print(
            f"Checking {item_key} "
            f"= {site_qty}"
        )

        if evaluate_rule(site_qty, rule):

            final_qty = 1

            pr_line = {

                "MaterialCode":
                    rule["MaterialCode"],

                "LineItemText":
                    rule["LineItemText"],

                "ItemKey":
                    item_key,

                "PR_Mode":
                    rule["PR_Mode"],

                "CalculatedQty":
                    site_qty,

                "FinalPRQty":
                    final_qty,

                "Unit":
                    rule["Unit"],

                "Remarks":
                    "BOM TI",
            }

            pr_lines.append(pr_line)

            print(
                f"  -> GENERATED "
                f"{rule['MaterialCode']} "
                f"qty={final_qty}"
            )

    return pr_lines


def main():

    selected_project = get_selected_project()

    if selected_project:
        print(f"Selected project for general items: {selected_project}")

    print("Loading calculated quantities...")

    with open(
        CALCULATED_FILE,
        "r",
        encoding="utf-8"
    ) as f:

        calculated_data = json.load(f)

    print("Loading PR rules...")

    rules = load_pr_rules()

    print(f"Loaded {len(rules)} PR rules")

    general_item_config = {}
    em_transport_map = {}
    epms_map = {}
    if selected_project:
        print("Loading general DU project configuration...")
        general_item_config, em_transport_map = load_general_item_config()
        epms_map = load_epms_map("input/EPMS.xlsx")

    final_output = {}
    general_output = {}

    for site_code, site_data in calculated_data.items():

        pr_lines = generate_pr_for_site(
            site_code,
            site_data,
            rules
        )

        general_pr_lines = []
        if selected_project:
            general_pr_lines = generate_general_pr_for_site(
                site_code,
                site_data,
                general_item_config,
                selected_project,
                em_transport_map,
                epms_map,
            )

        # =====================================================
        # PRESERVE METADATA
        # =====================================================

        final_output[site_code] = {

            "metadata":
                site_data.get(
                    "metadata",
                    {}
                ),

            "pr_lines":
                pr_lines
        }

        general_output[site_code] = {

            "metadata":
                site_data.get(
                    "metadata",
                    {}
                ),

            "pr_lines":
                general_pr_lines
        }

    with open(
        OUTPUT_FILE,
        "w",
        encoding="utf-8"
    ) as f:

        json.dump(
            final_output,
            f,
            indent=2
        )

    if selected_project:
        with open(
            GENERAL_ITEM_OUTPUT_FILE,
            "w",
            encoding="utf-8"
        ) as f:
            json.dump(
                general_output,
                f,
                indent=2
            )

        combined_output = {}
        for site_code, site_data in final_output.items():
            combined_pr_lines = list(site_data["pr_lines"])
            combined_pr_lines.extend(
                general_output.get(site_code, {}).get("pr_lines", [])
            )
            combined_output[site_code] = {
                "metadata": site_data["metadata"],
                "pr_lines": combined_pr_lines,
            }

        with open(
            COMBINED_PR_OUTPUT_FILE,
            "w",
            encoding="utf-8"
        ) as f:
            json.dump(
                combined_output,
                f,
                indent=2
            )

    print()
    print("DONE")
    print(f"Exported: {OUTPUT_FILE}")
    if selected_project:
        print(f"Exported: {GENERAL_ITEM_OUTPUT_FILE}")
        print(f"Exported: {COMBINED_PR_OUTPUT_FILE}")


if __name__ == "__main__":
    main()