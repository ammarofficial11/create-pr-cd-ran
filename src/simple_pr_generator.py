import json

import pandas as pd


CONFIG_FILE = "config/MainConfig.xlsx"

CALCULATED_FILE = "output/simple_calculated.json"

OUTPUT_FILE = "output/simple_pr_output.json"

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
            }

            pr_lines.append(pr_line)

            print(
                f"  -> GENERATED "
                f"{rule['MaterialCode']} "
                f"qty={final_qty}"
            )

    return pr_lines


def main():

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

    final_output = {}

    for site_code, site_data in calculated_data.items():

        pr_lines = generate_pr_for_site(
            site_code,
            site_data,
            rules
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

    print()
    print("DONE")
    print(f"Exported: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()