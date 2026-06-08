import json
import os

import pandas as pd


PR_OUTPUT_FILE = "output/simple_pr_output.json"
GENERAL_PR_OUTPUT_FILE = "output/general_pr_output.json"
OUTPUT_FILE = "output/ECC_PR_Output.xlsx"
OUTPUT_FILE_WITH_GENERAL = "output/ECC_PR_Output_With_GeneralItems.xlsx"


def load_pr_data(path):
    if not os.path.exists(path):
        return {}

    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def compose_ecc_rows(pr_data):
    ecc_rows = []
    sn = 1

    for site_code, site_data in pr_data.items():
        metadata = site_data.get("metadata", {})
        line_items = site_data.get("pr_lines", [])

        for item in line_items:
            ecc_rows.append({
                "SN.": sn,
                "Purchasing Area*": metadata.get("purchasing_area", ""),
                "Region*": metadata.get("region", ""),
                "Site ID*": metadata.get("site_code", ""),
                "Site Name*": metadata.get("site_name", ""),
                "Delivery Unit Code*": metadata.get("du_code", ""),
                "Logical Site Name": "",
                "Contract Number*": metadata.get("contract_number", ""),
                "Subcontractor*": metadata.get("subcon_ti", ""),
                "PBOM Code*": item.get("MaterialCode", ""),
                "SOW*": item.get("LineItemText", ""),
                "Unit*": item.get("Unit", ""),
                "Quantity*": item.get("FinalPRQty", 0),
                "Remarks": item.get("LineItemText", "") if not item.get("MaterialCode", "") else "",
            })
            sn += 1

    return ecc_rows


def main():

    print("Loading PR output...")

    pr_data = load_pr_data(PR_OUTPUT_FILE)

    ecc_rows = compose_ecc_rows(pr_data)

    print(f"Generated ECC rows: {len(ecc_rows)}")

    print("Exporting ECC output...")

    df = pd.DataFrame(ecc_rows)

    with pd.ExcelWriter(
        OUTPUT_FILE,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="ECC_PR"
        )

    print("\nDONE")
    print(f"Exported: {OUTPUT_FILE}")

    general_pr_data = load_pr_data(GENERAL_PR_OUTPUT_FILE)
    merged_pr_data = {}
    merged_pr_data.update(pr_data)

    for site_code, site_data in general_pr_data.items():
        if site_code in merged_pr_data:
            merged_pr_data[site_code]["pr_lines"].extend(
                site_data.get("pr_lines", [])
            )
        else:
            merged_pr_data[site_code] = site_data

    merged_ecc_rows = compose_ecc_rows(merged_pr_data)

    print(f"Generated ECC rows with general items: {len(merged_ecc_rows)}")
    print("Exporting ECC output with general items...")

    df = pd.DataFrame(merged_ecc_rows)

    with pd.ExcelWriter(
        OUTPUT_FILE_WITH_GENERAL,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="ECC_PR"
        )

    print("\nDONE")
    print(f"Exported: {OUTPUT_FILE_WITH_GENERAL}")


if __name__ == "__main__":
    main()