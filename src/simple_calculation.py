import json

import pandas as pd


CONFIG_FILE = "config/MainConfig.xlsx"

NORMALIZED_FILE = "output/simple_normalized.json"

OUTPUT_FILE = "output/simple_calculated.json"

CALCULATION_SHEET = "Calculation_Rules"


def load_rules():

    df = pd.read_excel(
        CONFIG_FILE,
        sheet_name=CALCULATION_SHEET
    )

    rules = []

    for _, row in df.iterrows():

        item_key = str(
            row.get("ItemKey", "")
        ).strip()

        calc_type = str(
            row.get("CalcType", "")
        ).strip()

        source_item_key = str(
            row.get("SourceItemKey", "")
        ).strip()

        condition_item_key = str(
            row.get("ConditionItemKey", "")
        ).strip()

        threshold = row.get(
            "ThresholdMin",
            None
        )

        divisor = row.get(
            "Divisor",
            None
        )

        condition_type = str(
            row.get("ConditionType", "")
        ).strip()

        condition_value = row.get(
            "ConditionValue",
            None
        )

        round_mode = str(
            row.get("RoundMode", "")
        ).strip()

        rules.append({

            "ItemKey": item_key,
            "CalcType": calc_type,
            "SourceItemKey": source_item_key,
            "ConditionItemKey": condition_item_key,
            "ThresholdMin": threshold,
            "Divisor": divisor,
            "ConditionType": condition_type,
            "ConditionValue": condition_value,
            "RoundMode": round_mode,
        })

    return rules


def apply_round(value, round_mode):

    if round_mode == "UP":

        return int(-(-value // 1))

    return value


def calculate_site(site_data, rules):

    source_quantities = site_data["quantities"]

    calculated = {}

    for rule in rules:

        item_key = rule["ItemKey"]

        calc_type = rule["CalcType"]

        source_item_key = rule["SourceItemKey"]

        condition_item_key = rule["ConditionItemKey"]

        threshold = rule["ThresholdMin"]

        divisor = rule["Divisor"]

        condition_type = rule["ConditionType"]

        condition_value = rule["ConditionValue"]

        round_mode = rule["RoundMode"]

        result = 0

        # =====================================================
        # SUM
        # =====================================================

        if calc_type == "SUM":

            result = source_quantities.get(
                item_key,
                0
            )

        # =====================================================
        # DIVIDE_BY_ITEMKEY
        # =====================================================

        elif calc_type == "DIVIDE_BY_ITEMKEY":

            source_qty = source_quantities.get(
                source_item_key,
                0
            )

            divisor_qty = source_quantities.get(
                condition_item_key,
                0
            )

            if divisor_qty > 0:

                result = source_qty / divisor_qty

            if threshold is not None:

                if result < float(threshold):

                    result = 0

            result = apply_round(
                result,
                round_mode
            )

        # =====================================================
        # CONDITIONAL_SUM_BY_ITEMKEY
        # =====================================================

        elif calc_type == "CONDITIONAL_SUM_BY_ITEMKEY":

            condition_qty = source_quantities.get(
                condition_item_key,
                0
            )

            if condition_type == "GT":

                if condition_qty > float(condition_value):

                    result = 0

                else:

                    result = source_quantities.get(
                        source_item_key,
                        0
                    )

        # =====================================================
        # CONDITIONAL_DIVIDE
        # =====================================================

        elif calc_type == "CONDITIONAL_DIVIDE":

            source_qty = source_quantities.get(
                source_item_key,
                0
            )

            condition_qty = source_quantities.get(
                condition_item_key,
                0
            )

            if condition_type == "SUM_GT":

                if source_qty + condition_qty > float(condition_value):

                    if divisor and float(divisor) > 0:

                        result = source_qty / float(divisor)

        calculated[item_key] = result

        print(
            f"{item_key} "
            f"({calc_type}) "
            f"= {result}"
        )

    return calculated


def main():

    print("Loading normalized quantities...")

    with open(
        NORMALIZED_FILE,
        "r",
        encoding="utf-8"
    ) as f:

        normalized_data = json.load(f)

    print("Loading calculation rules...")

    rules = load_rules()

    print(f"Loaded {len(rules)} rules")

    final_results = {}

    for site_code, site_data in normalized_data.items():

        print()
        print("=" * 60)
        print(f"SITE: {site_code}")
        print("=" * 60)

        calculated = calculate_site(
            site_data,
            rules
        )

        # =====================================================
        # PRESERVE METADATA
        # =====================================================

        final_results[site_code] = {

            "metadata": site_data.get(
                "metadata",
                {}
            ),

            "quantities": calculated
        }

    with open(
        OUTPUT_FILE,
        "w",
        encoding="utf-8"
    ) as f:

        json.dump(
            final_results,
            f,
            indent=2
        )

    print()
    print("Exported:")
    print(OUTPUT_FILE)


if __name__ == "__main__":
    main()