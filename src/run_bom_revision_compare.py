import os

import simple_normalize
import simple_calculation

from compare_bom_revision_v2 import (
    compare_pr_vs_bom,
    compare_bom_vs_bom
)

print()
print("=" * 60)
print("BOM REVISION COMPARISON TOOL")
print("=" * 60)

print()
print("1 - PR Generated BOM vs Uploaded BOM")
print("2 - BOM A vs BOM B")

choice = input(
    "\nSelect Workflow (1/2): "
).strip()

# =====================================================
# WORKFLOW 1
# =====================================================

if choice == "1":

    print()
    print("=" * 60)
    print("GENERATING BOM C DATASET")
    print("=" * 60)

    os.environ[
        "BOM_FILE_PATH"
    ] = "input/BOM_COMPARE_C.xlsx"

    os.environ[
        "NORMALIZED_OUTPUT_FILE"
    ] = (
        "output/simple_normalized_bom_c.json"
    )

    simple_normalize.main()

    os.environ[
        "NORMALIZED_INPUT_FILE"
    ] = (
        "output/simple_normalized_bom_c.json"
    )

    os.environ[
        "CALCULATED_OUTPUT_FILE"
    ] = (
        "output/simple_calculated_bom_c.json"
    )

    simple_calculation.main()

    compare_pr_vs_bom()

# =====================================================
# WORKFLOW 2
# =====================================================

elif choice == "2":

    print()
    print("=" * 60)
    print("GENERATING BOM A DATASET")
    print("=" * 60)

    os.environ[
        "BOM_FILE_PATH"
    ] = "input/BOM_COMPARE_A.xlsx"

    os.environ[
        "NORMALIZED_OUTPUT_FILE"
    ] = (
        "output/simple_normalized_bom_a.json"
    )

    simple_normalize.main()

    os.environ[
        "NORMALIZED_INPUT_FILE"
    ] = (
        "output/simple_normalized_bom_a.json"
    )

    os.environ[
        "CALCULATED_OUTPUT_FILE"
    ] = (
        "output/simple_calculated_bom_a.json"
    )

    simple_calculation.main()

    print()
    print("=" * 60)
    print("GENERATING BOM B DATASET")
    print("=" * 60)

    os.environ[
        "BOM_FILE_PATH"
    ] = "input/BOM_COMPARE_B.xlsx"

    os.environ[
        "NORMALIZED_OUTPUT_FILE"
    ] = (
        "output/simple_normalized_bom_b.json"
    )

    simple_normalize.main()

    os.environ[
        "NORMALIZED_INPUT_FILE"
    ] = (
        "output/simple_normalized_bom_b.json"
    )

    os.environ[
        "CALCULATED_OUTPUT_FILE"
    ] = (
        "output/simple_calculated_bom_b.json"
    )

    simple_calculation.main()

    compare_bom_vs_bom()

else:

    print("Invalid Selection")