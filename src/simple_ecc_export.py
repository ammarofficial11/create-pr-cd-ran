import json

import pandas as pd


PR_OUTPUT_FILE = "output/simple_pr_output.json"
OUTPUT_FILE = "output/ECC_PR_Output.xlsx"


def main():

    print("Loading PR output...")

    with open(PR_OUTPUT_FILE, "r", encoding="utf-8") as f:

        pr_data = json.load(f)

    ecc_rows = []

    sn = 1

    for site_code, site_data in pr_data.items():

        metadata = site_data["metadata"]

        line_items = site_data["pr_lines"]

        for item in line_items:

            ecc_rows.append({

                # ============================================
                # ECC TEMPLATE FIELDS
                # ============================================

                "SN.": sn,

                "Purchasing Area*":
                    metadata.get("purchasing_area", ""),

                "Region*":
                    metadata.get("region", ""),

                "Site ID*":
                    metadata.get("site_code", ""),

                "Site Name*":
                    metadata.get("site_name", ""),

                "Delivery Unit Code*":
                    metadata.get("du_code", ""),

                "Logical Site Name":
                    "",

                "Contract Number*":
                    metadata.get("contract_number", ""),

                "Subcontractor*":
                    metadata.get("subcon_ti", ""),

                "PBOM Code*":
                    item.get("MaterialCode", ""),

                "SOW*":
                    item.get("LineItemText", ""),

                "Unit*":
                    "",

                "Quantity*":
                    item.get("FinalPRQty", 0),

                "Remarks":
                    "",
            })

            sn += 1

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


if __name__ == "__main__":
    main()