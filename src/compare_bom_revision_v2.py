import json
import os
from datetime import datetime
import pandas as pd

# ==========================================================
# FILES
# ==========================================================

PR_FILE = "output/simple_calculated.json"
COMPARE_FILE ="output/simple_calculated_bom_c.json"

BOM_A_FILE ="output/simple_calculated_bom_a.json"
BOM_B_FILE ="output/simple_calculated_bom_b.json"

OUTPUT_JSON = "output/bom_revision_report.json"
OUTPUT_XLSX = "output/bom_revision_report.xlsx"


def load_json(filepath):
    with open(filepath, "r", encoding="utf-8") as f:
        return json.load(f)


def determine_change_type(old_qty, new_qty):

    if old_qty == 0 and new_qty > 0:
        return "ADDED"

    if old_qty > 0 and new_qty == 0:
        return "REMOVED"

    if new_qty > old_qty:
        return "INCREASE"

    if new_qty < old_qty:
        return "DECREASE"

    return "UNCHANGED"


def build_remark(item_key, old_qty, new_qty, change_type):

    old_qty = int(old_qty) if float(old_qty).is_integer() else old_qty
    new_qty = int(new_qty) if float(new_qty).is_integer() else new_qty

    if change_type == "ADDED":
        return f"{item_key} added Qty - 0 to {new_qty}"

    if change_type == "REMOVED":
        return f"{item_key} removed Qty - {old_qty} to 0"

    if change_type == "INCREASE":
        return f"{item_key} increase Qty - {old_qty} to {new_qty}"

    if change_type == "DECREASE":
        return f"{item_key} decrease Qty - {old_qty} to {new_qty}"

    return ""


def build_site_index(data):

    index = {}

    for site_code, site_data in data.items():

        metadata = site_data.get("metadata", {})

        du_code = metadata.get("du_code", "")

        if not du_code:
            continue

        key = (
            site_code.upper(),
            du_code.upper()
        )

        index[key] = site_data

    return index


from datetime import datetime

def get_dataset_latest_date(data):

    dates = []

    for _, site_data in data.items():

        value = site_data.get(
            "metadata",
            {}
        ).get(
            "bom_last_updated_date",
            ""
        )

        if value:

            try:

                dates.append(
                    datetime.strptime(
                        value,
                        "%Y-%m-%d %H:%M:%S"
                    )
                )

            except:
                pass

    if not dates:
        return ""

    return max(dates).strftime(
        "%Y-%m-%d %H:%M:%S"
    )

def compare_datasets(
    legacy_data,
    latest_data,
    comparison_method,
    legacy_date,
    latest_date
):

    legacy_index = build_site_index(legacy_data)
    latest_index = build_site_index(latest_data)

    common_sites = (
        set(legacy_index.keys())
        &
        set(latest_index.keys())
    )

    print(f"Common Sites Found: {len(common_sites)}")

    results = []
    total_changes = 0

    for site_key in sorted(common_sites):

        legacy_site = legacy_index[site_key]
        latest_site = latest_index[site_key]

        site_code = legacy_site["metadata"]["site_code"]
        du_code = legacy_site["metadata"]["du_code"]

        legacy_qtys = legacy_site.get(
            "quantities",
            {}
        )

        latest_qtys = latest_site.get(
            "quantities",
            {}
        )

        all_item_keys = (
            set(legacy_qtys.keys())
            |
            set(latest_qtys.keys())
        )

        site_changes = []

        for item_key in sorted(all_item_keys):

            old_qty = legacy_qtys.get(item_key, 0)
            new_qty = latest_qtys.get(item_key, 0)

            if old_qty == new_qty:
                continue

            delta = new_qty - old_qty

            change_type = determine_change_type(
                old_qty,
                new_qty
            )

            remark = build_remark(
                item_key,
                old_qty,
                new_qty,
                change_type
            )

            site_changes.append({
                "item_key": item_key,
                "legacy_qty": old_qty,
                "latest_qty": new_qty,
                "delta": delta,
                "change_type": change_type,
                "remark": remark,
                "comparison_method": comparison_method,
                "legacy_date": legacy_date,
                "latest_date": latest_date
            })

            total_changes += 1

        if site_changes:

           results.append({

            "site_code":
                site_code,

            "site_name":
                legacy_site["metadata"].get(
                    "site_name",
                    ""
                ),

            "region":
                legacy_site["metadata"].get(
                    "region",
                    ""
                ),

            "du_code":
                du_code,

            "changes":
                site_changes
        })

    print(f"Mismatch Records: {total_changes}")

    return results, len(common_sites), total_changes


def export_json(results):

    with open(
        OUTPUT_JSON,
        "w",
        encoding="utf-8"
    ) as f:

        json.dump(
            results,
            f,
            indent=2
        )
    
    print(
    f"Exported JSON: "
    f"{OUTPUT_JSON}"
    )

def export_excel(
    results,
    comparison_method,
    legacy_date,
    latest_date,
    sites_compared,
    item_changes
):

    rows = []

    changed_sites = len(results)

    for site in results:

        for change in site["changes"]:

            rows.append({
                "Site Code": site["site_code"],
                "Site Name": site["site_name"],
                "Region": site["region"],
                "DU Code": site["du_code"],
                "Item Key": change["item_key"],
                "Legacy Qty": change["legacy_qty"],
                "Latest Qty": change["latest_qty"],
                "Delta": change["delta"],
                "Change Type": change["change_type"],
                "Remark": change["remark"],
                "Comparison Method": change["comparison_method"],
                "Legacy BOM Last Updated Date": change["legacy_date"],
                "Latest BOM Last Updated Date": change["latest_date"]
            })

    summary_df = pd.DataFrame([{
        "Metric": "Comparison Method",
        "Value": comparison_method
    },{
        "Metric": "Legacy BOM Last Updated Date",
        "Value": legacy_date
    },{
        "Metric": "Latest BOM Last Updated Date",
        "Value": latest_date
    },{
        "Metric": "Sites Compared",
        "Value": sites_compared
    },{
        "Metric": "Sites Changed",
        "Value": changed_sites
    },{
        "Metric": "Item Changes",
        "Value": item_changes
    }])

    details_df = pd.DataFrame(rows)

    with pd.ExcelWriter(OUTPUT_XLSX) as writer:
        summary_df.to_excel(
            writer,
            sheet_name="SUMMARY",
            index=False
        )

        details_df.to_excel(
            writer,
            sheet_name="DETAILS",
            index=False
        )

    print(
    f"Exported Excel: "
    f"{OUTPUT_XLSX}"
    )

def compare_pr_vs_bom():
    print()
    print("=" * 60)
    print("WORKFLOW 1")
    print("PR GENERATED BOM VS UPLOADED BOM")
    print("=" * 60)
    print()
    
    legacy_data = load_json(PR_FILE)
    latest_data = load_json(COMPARE_FILE)

    comparison_method = (
        "Workflow 1 - PR Generated BOM vs Uploaded BOM"
    )

    legacy_date = get_dataset_latest_date(
        legacy_data
    )

    latest_date = get_dataset_latest_date(
        latest_data
    )
    results, sites_compared, item_changes = compare_datasets(
        legacy_data,
        latest_data,
        comparison_method,
        legacy_date,
        latest_date
    )

    export_json(results)

    export_excel(
        results,
        comparison_method,
        legacy_date,
        latest_date,
        sites_compared,
        item_changes
    )

    print()
    print("=" * 60)
    print("BOM REVISION COMPARISON COMPLETE")
    print("=" * 60)

    print(
        f"Sites Compared : "
        f"{sites_compared}"
    )

    print(
        f"Sites Changed  : "
        f"{len(results)}"
    )

    print(
        f"Item Changes   : "
        f"{item_changes}"
    )



def compare_bom_vs_bom():
    
    print()
    print("=" * 60)
    print("WORKFLOW 2")
    print("LATEST BOM VS LEGACY BOM")
    print("=" * 60)
    print()

    data_a = load_json(BOM_A_FILE)
    data_b = load_json(BOM_B_FILE)

    date_a = get_dataset_latest_date(data_a)
    date_b = get_dataset_latest_date(data_b)

    if date_a >= date_b:

        latest_data = data_a
        legacy_data = data_b

        latest_date = date_a
        legacy_date = date_b

    else:

        latest_data = data_b
        legacy_data = data_a

        latest_date = date_b
        legacy_date = date_a

    comparison_method = (
        "Workflow 2 - Latest BOM vs Legacy BOM"
    )
    
    print()
    print("BOM A Date :", date_a)
    print("BOM B Date :", date_b)
    print()

    print("Legacy BOM :", legacy_date)
    print("Latest BOM :", latest_date)
    print()

    results, sites_compared, item_changes = compare_datasets(
        legacy_data,
        latest_data,
        comparison_method,
        legacy_date,
        latest_date
    )

    export_json(results)

    export_excel(
        results,
        comparison_method,
        legacy_date,
        latest_date,
        sites_compared,
        item_changes
    )
    print()
    print("=" * 60)
    print("BOM REVISION COMPARISON COMPLETE")
    print("=" * 60)

    print(
        f"Sites Compared : "
        f"{sites_compared}"
    )

    print(
        f"Sites Changed  : "
        f"{len(results)}"
    )

    print(
        f"Item Changes   : "
        f"{item_changes}"
    )


