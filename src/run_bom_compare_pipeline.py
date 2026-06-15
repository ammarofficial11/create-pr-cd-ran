import subprocess
import sys
from pathlib import Path


def run_bom_compare_pipeline(
    workflow
):

    repo_root = (
        Path(__file__)
        .resolve()
        .parents[1]
    )

    print()
    print("=" * 70)
    print(
        f"RUNNING BOM COMPARE WORKFLOW {workflow}"
    )
    print("=" * 70)

    print()
    print("SYS EXECUTABLE")
    print(sys.executable)

    print()
    print("WORKING DIRECTORY")
    print(repo_root)

    result = subprocess.run(

        [
            "py",
            "src/run_bom_revision_compare.py"
        ],

        input=f"{workflow}\n",

        text=True,

        cwd=repo_root
    )

    if result.returncode != 0:

        raise RuntimeError(
            "BOM Compare Failed"
        )

    print()
    print("=" * 70)
    print(
        "BOM COMPARE COMPLETE"
    )
    print("=" * 70)


if __name__ == "__main__":

    run_bom_compare_pipeline(
        workflow="2"
    )