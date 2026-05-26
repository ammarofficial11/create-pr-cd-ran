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


def normalize_lookup_key(value):

    return (
        str(value)
        .strip()
        .replace("_", "")
        .replace("/", "")
        .replace(" ", "")
        .lower()
    )


def calculate_site(site_data, rules):

    source_quantities = site_data["quantities"]

    calculated = {}

    working_quantities = dict(source_quantities)

    # build a normalized lookup map to allow semantic-safe key matching
    normalized_working_lookup = {}

    for k, v in working_quantities.items():

        normalized_working_lookup[
            normalize_lookup_key(k)
        ] = v

    def get_quantity(key):

        normalized_key = normalize_lookup_key(key)

        return normalized_working_lookup.get(
            normalized_key,
            0
        )

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

            # source_qty must come from original normalized quantities
            # to avoid contamination from previously calculated run counts
            source_qty = source_quantities.get(
                source_item_key,
                0
            )

            divisor_qty = get_quantity(
                condition_item_key
            )

            result = 0

            # Safeguard: if there's no feeder length (source_qty) or no connectors (divisor_qty),
            # do not produce runs.
            if source_qty <= 0:

                result = 0

            elif divisor_qty <= 0:

                result = 0

            else:

                total_runs = divisor_qty / 2

                if total_runs > 0:

                    average = source_qty / total_runs

                    if threshold is not None:

                        if average >= float(threshold):

                            result = total_runs

                        else:

                            result = 0

                    else:

                        result = total_runs

            result = apply_round(
                result,
                round_mode
            )

        # =====================================================
        # CONDITIONAL_SUM_BY_ITEMKEY
        # =====================================================

        elif calc_type == "CONDITIONAL_SUM_BY_ITEMKEY":

            condition_qty = get_quantity(
                condition_item_key
            )

            if condition_type == "GT":

                if condition_qty > float(condition_value):

                    result = 0

                else:

                    result = get_quantity(
                        source_item_key
                    )

        # =====================================================
        # CONDITIONAL_DIVIDE
        # =====================================================

        elif calc_type == "CONDITIONAL_DIVIDE":

            source_qty = get_quantity(
                source_item_key
            )

            condition_qty = source_quantities.get(
                condition_item_key,
                0
            )

            result = 0

            # feeder quantity must come from original normalized quantities
            # and connector quantity must be present to generate runs
            if condition_qty > 0:

                if source_qty > 0:

                    if divisor and float(divisor) > 0:

                        result = source_qty / float(divisor)

            result = apply_round(
                result,
                round_mode
            )

        calculated[item_key] = result
        working_quantities[item_key] = result

        # keep normalized lookup in sync so downstream rules can find
        # this calculated value using semantic key matching
        normalized_working_lookup[
            normalize_lookup_key(item_key)
        ] = result

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