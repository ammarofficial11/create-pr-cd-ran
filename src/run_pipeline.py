import subprocess
import sys


def run_script(script_name):

    print()
    print("=" * 70)
    print(f"RUNNING: {script_name}")
    print("=" * 70)

    result = subprocess.run(
        [sys.executable, script_name]
    )

    if result.returncode != 0:

        print()
        print(f"FAILED: {script_name}")

        sys.exit(1)

    print()
    print(f"SUCCESS: {script_name}")


def main():

    pipeline = [

        "src/simple_normalize.py",

        "src/simple_calculation.py",

        "src/simple_pr_generator.py",

        "src/simple_ecc_export.py",
    ]

    for script in pipeline:

        run_script(script)

    print()
    print("=" * 70)
    print("ALL PROCESS COMPLETED")
    print("=" * 70)

    print()
    print("Generated files:")

    print("1. output/simple_normalized.json")

    print("2. output/simple_calculated.json")

    print("3. output/simple_pr_output.json")

    print("4. output/ECC_PR_Output.xlsx")


if __name__ == "__main__":

    main()