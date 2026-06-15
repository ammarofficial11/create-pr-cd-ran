import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


GENERAL_PROJECTS = {
    "1": "CD consolidation 2023 (Swap/ Modernize)",
    "2": "CD consolidation 2023 (Decomm RAN)",
    "3": "CD consolidation 2023 (Decomm TX)",
    "4": "NIC congest upgrade",
    "5": "2023 Celcomdigi BAU",
    "6": "2024 Celcomdigi BAU",
    "7": "CR with Equipment",
    "8": "Celcomdigi USP",
}


def select_general_project():

    print("=" * 60)
    print("GENERAL DU PROJECT SELECTION")
    print("=" * 60)
    print()

    print("1. 2023 (Swap/Modernize)")
    print("2. 2023 (Decomm RAN)")
    print("3. 2023 (Decomm TX)")
    print("4. NIC congest upgrade")
    print("5. 2023 Celcomdigi BAU")
    print("6. 2024 Celcomdigi BAU")
    print("7. CR with Equipment")
    print("8. Celcomdigi USP")
    print()
    print("0. Skip General Item Processing")
    print()

    selection = input("Enter selection: ").strip()

    if selection == "" or selection == "0":
        return None

    project = GENERAL_PROJECTS.get(selection)
    if not project:
        print()
        print("Invalid selection. Skipping General Item Processing.")
        return None

    return project


def run_script(script_name, selected_project=None, cwd=None, env=None):

    print()
    print("=" * 70)
    print(f"RUNNING: {script_name}")
    print("=" * 70)
    print()
    print("SYS EXECUTABLE")
    print(sys.executable)

   

    command = ["python", script_name]
    if selected_project and script_name.endswith("simple_pr_generator.py"):
        command.extend(["--selected-project", selected_project])

    result = subprocess.run(command, cwd=cwd, env=env)

    if result.returncode != 0:

        print()
        print(f"FAILED: {script_name}")

        sys.exit(1)

    print()
    print(f"SUCCESS: {script_name}")


def build_pipeline_env(bom_file_path=None, epms_file_path=None, selected_project=None):
    env = os.environ.copy()
    if bom_file_path:
        env["BOM_FILE_PATH"] = str(bom_file_path)
    if epms_file_path:
        env["EPMS_FILE_PATH"] = str(epms_file_path)
    if selected_project:
        env["SELECTED_PROJECT"] = str(selected_project)
    return env


def run_pipeline(bom_file_path=None, epms_file_path=None, selected_project=None, cwd=None):
    if cwd is None:
        cwd = Path(__file__).resolve().parents[1]

    env = build_pipeline_env(
        bom_file_path=bom_file_path,
        epms_file_path=epms_file_path,
        selected_project=selected_project,
    )

    pipeline = [
        "src/simple_normalize.py",
        "src/simple_calculation.py",
        "src/simple_pr_generator.py",
        "src/simple_ecc_export.py",
    ]

    for script in pipeline:
        run_script(script, selected_project=selected_project, cwd=cwd, env=env)

    return {
        "selected_project": selected_project,
        "bom_file_path": bom_file_path,
        "epms_file_path": epms_file_path,
    }


def main():

    pipeline = [

        "src/simple_normalize.py",

        "src/simple_calculation.py",

        "src/simple_pr_generator.py",

        "src/simple_ecc_export.py",
    ]

    selected_project = os.getenv("SELECTED_PROJECT")

    if not selected_project:
        selected_project = select_general_project()

    pipeline_env = build_pipeline_env(
        selected_project=selected_project
    )

    repo_root = Path(__file__).resolve().parents[1]

    for script in pipeline:

        run_script(
            script,
            selected_project=selected_project,
            cwd=repo_root,
            env=pipeline_env
        )

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

    if selected_project:
        print("5. output/general_pr_output.json")
        print("6. output/simple_pr_output_with_general_items.json")
        print("7. output/ECC_PR_Output_With_GeneralItems.xlsx")
        
if __name__ == "__main__":

    main()