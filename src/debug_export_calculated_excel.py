import json
from pathlib import Path

import pandas as pd


INPUT_FILE = Path("output/simple_calculated.json")
OUTPUT_FILE = Path("output/debug_calculated.xlsx")

METADATA_COLUMNS = [
    "Site Code",
    "Site Name",
    "Region",
    "DU Code",
    "Subcon TI",
    "Contract Number",
    "Purchasing Area",
]


def load_calculated_data(input_path: Path) -> dict:
    if not input_path.exists():
        raise FileNotFoundError(
            f"Calculated JSON input not found: {input_path}"
        )

    with input_path.open("r", encoding="utf-8") as reader:
        return json.load(reader)


def build_rows(calculated_data: dict) -> tuple[list[dict], list[str]]:
    all_item_keys = set()
    rows = []

    for site_code, site_data in calculated_data.items():
        metadata = site_data.get("metadata", {}) or {}
        quantities = site_data.get("quantities", {}) or {}

        all_item_keys.update(quantities.keys())

        row = {
            "Site Code": metadata.get("site_code", site_code),
            "Site Name": metadata.get("site_name", ""),
            "Region": metadata.get("region", ""),
            "DU Code": metadata.get("du_code", ""),
            "Subcon TI": metadata.get("subcon_ti", ""),
            "Contract Number": metadata.get("contract_number", ""),
            "Purchasing Area": metadata.get("purchasing_area", ""),
        }

        for item_key, item_value in quantities.items():
            row[item_key] = item_value

        rows.append(row)

    sorted_item_keys = sorted(all_item_keys)
    return rows, sorted_item_keys


def export_to_excel(rows: list[dict], item_keys: list[str], output_path: Path) -> None:
    columns = METADATA_COLUMNS + item_keys

    df = pd.DataFrame(rows, columns=columns)
    df = df.fillna(0)

    output_path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(
        output_path,
        engine="openpyxl",
    ) as writer:
        df.to_excel(
            writer,
            sheet_name="Calculated",
            index=False,
        )

        worksheet = writer.sheets["Calculated"]

        worksheet.auto_filter.ref = worksheet.dimensions
        worksheet.freeze_panes = worksheet["A2"]

        for column_cells in worksheet.columns:
            max_length = 0
            column = column_cells[0].column_letter
            for cell in column_cells:
                try:
                    cell_value = cell.value
                    cell_length = len(str(cell_value)) if cell_value is not None else 0
                    if cell_length > max_length:
                        max_length = cell_length
                except Exception:
                    continue
            adjusted_width = max(10, min(max_length + 2, 50))
            worksheet.column_dimensions[column].width = adjusted_width

    print(f"Exported calculated debug Excel: {output_path}")


def main() -> None:
    calculated_data = load_calculated_data(INPUT_FILE)
    rows, item_keys = build_rows(calculated_data)
    export_to_excel(rows, item_keys, OUTPUT_FILE)


if __name__ == "__main__":
    main()
