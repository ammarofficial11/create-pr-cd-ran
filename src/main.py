#!/usr/bin/env python3
"""AI Worker Platform contract entrypoint for the standalone RAN PR skill."""
from __future__ import annotations

import argparse
import hashlib
import json
import mimetypes
import os
import shutil
import signal
import subprocess
import sys
import threading
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


SKILL_ROOT = Path(__file__).resolve().parents[1]
CONTRACT_VERSION = "1.0"
SKILL_ID = "create-pr-cd-ran"
SKILL_VERSION = "1.1.0"
PIPELINE = (
    "simple_normalize.py",
    "simple_calculation.py",
    "simple_pr_generator.py",
    "simple_ecc_export.py",
)
APPROVED_OUTPUTS = (
    "simple_normalized.json",
    "simple_calculated.json",
    "simple_pr_output.json",
    "general_pr_output.json",
    "simple_pr_output_with_general_items.json",
    "ECC_PR_Output.xlsx",
    "ECC_PR_Output_With_GeneralItems.xlsx",
)


class ContractError(Exception):
    def __init__(self, code: str, message: str, category: str = "domain_input", details: dict | None = None):
        super().__init__(message)
        self.code = code
        self.category = category
        self.details = details or {}


class CancelledError(ContractError):
    def __init__(self):
        super().__init__("SKILL_CANCELLED", "Cancellation was requested.", "cancelled")


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="milliseconds").replace("+00:00", "Z")


def emit(event_type: str, phase: str, message: str, percent: int | None = None) -> None:
    event: dict[str, Any] = {"type": event_type, "timestamp": utc_now(), "phase": phase, "message": message}
    if percent is not None:
        event["percent"] = percent
    print(json.dumps(event, ensure_ascii=False), flush=True)


def parse_cli() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run create-pr-cd-ran through the AI Worker Platform skill contract.")
    parser.add_argument("--input-manifest", required=True, type=Path)
    return parser.parse_args()


def resolve_inside(root: Path, raw_value: Any, label: str) -> Path:
    raw = str(raw_value or "").strip()
    if not raw or Path(raw).is_absolute():
        raise ContractError("CONTRACT_PATH_INVALID", f"{label} must be a workspace-relative path.")
    resolved = (root / raw).resolve()
    try:
        resolved.relative_to(root.resolve())
    except ValueError as exc:
        raise ContractError("CONTRACT_PATH_INVALID", f"{label} escapes the workspace.") from exc
    return resolved


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def load_envelope(path: Path) -> tuple[dict[str, Any], Path, Path, Path, Path]:
    manifest_path = path.resolve()
    try:
        envelope = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ContractError("INPUT_MANIFEST_INVALID", "The input manifest is missing or invalid.") from exc
    if envelope.get("schemaVersion") != CONTRACT_VERSION:
        raise ContractError("CONTRACT_VERSION_UNSUPPORTED", "Unsupported input contract version.")
    identity = envelope.get("skill") or {}
    if identity.get("skillId") != SKILL_ID or identity.get("version") != SKILL_VERSION:
        raise ContractError("SKILL_IDENTITY_MISMATCH", "Input manifest skill identity does not match this package.")
    if not str(envelope.get("jobId") or "").strip():
        raise ContractError("JOB_ID_REQUIRED", "jobId is required.")
    paths = envelope.get("paths") or {}
    workspace = resolve_inside(manifest_path.parent, paths.get("workspace", "."), "paths.workspace")
    output = resolve_inside(workspace, paths.get("output", "output"), "paths.output")
    result = resolve_inside(workspace, paths.get("result", "result.json"), "paths.result")
    cancellation = resolve_inside(workspace, paths.get("cancellation", "temp/cancel.requested"), "paths.cancellation")
    output.mkdir(parents=True, exist_ok=True)
    result.parent.mkdir(parents=True, exist_ok=True)
    return envelope, workspace, output, result, cancellation


def declared_file(envelope: dict[str, Any], workspace: Path, name: str) -> Path:
    matches = [item for item in envelope.get("files", []) if item.get("name") == name]
    if len(matches) != 1:
        raise ContractError("INPUT_FILE_INVALID", f"Exactly one {name} file is required.")
    item = matches[0]
    path = resolve_inside(workspace, item.get("path"), f"files.{name}.path")
    if not path.is_file() or path.suffix.lower() != ".xlsx":
        raise ContractError("INPUT_FILE_INVALID", f"{name} must be an existing .xlsx file.")
    if item.get("size") is not None and int(item["size"]) != path.stat().st_size:
        raise ContractError("INPUT_FILE_SIZE_MISMATCH", f"{name} size does not match its declaration.")
    if item.get("sha256") and str(item["sha256"]).lower() != sha256(path):
        raise ContractError("INPUT_FILE_CHECKSUM_MISMATCH", f"{name} checksum does not match its declaration.")
    return path


def validate_parameters(parameters: dict[str, Any]) -> tuple[str, str | None]:
    unknown = sorted(set(parameters) - {"runMode", "selectedProject"})
    if unknown:
        raise ContractError("PARAMETERS_INVALID", "Unsupported parameters were supplied.", details={"fields": unknown})
    run_mode = parameters.get("runMode")
    if run_mode not in {"standard-pr", "general-item"}:
        raise ContractError("INVALID_RAN_RUN_MODE", "runMode must be standard-pr or general-item.")
    project = str(parameters.get("selectedProject") or "").strip() or None
    if run_mode == "standard-pr":
        return run_mode, None
    if not project:
        raise ContractError("INVALID_RAN_PROJECT", "selectedProject is required for general-item mode.")
    try:
        import pandas as pd
        sheets = pd.read_excel(SKILL_ROOT / "config" / "GENERAL ITEM FOR ALL DU PROJECT Overall.xlsx", sheet_name=None)
        approved = {
            str(column).strip()
            for frame in sheets.values()
            for column in list(frame.columns)[4:]
            if str(column).strip() and str(column).lower() != "nan" and not str(column).startswith("Unnamed")
        }
    except Exception as exc:
        raise ContractError("RAN_PROJECT_CATALOG_INVALID", "The RAN project catalog could not be read.", "domain_configuration") from exc
    if project not in approved:
        raise ContractError("INVALID_RAN_PROJECT", "selectedProject is not present in the approved RAN project catalog.")
    return run_mode, project


def check_cancel(path: Path) -> None:
    if path.exists():
        raise CancelledError()


def run_stage(script_name: str, workspace: Path, env: dict[str, str], cancellation: Path, percent: int) -> None:
    check_cancel(cancellation)
    emit("progress", "ran_pipeline", f"Running {script_name}.", percent)
    command = [sys.executable, str(SKILL_ROOT / "src" / script_name)]
    if script_name == "simple_pr_generator.py" and env.get("SELECTED_PROJECT"):
        command.extend(["--selected-project", env["SELECTED_PROJECT"]])
    process = subprocess.Popen(command, cwd=workspace, env=env, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE, text=True, encoding="utf-8", errors="replace")
    last_heartbeat = time.monotonic()
    while process.poll() is None:
        if cancellation.exists():
            process.terminate()
            try:
                process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                process.kill()
            raise CancelledError()
        if time.monotonic() - last_heartbeat >= 30:
            emit("progress", "ran_pipeline", f"{script_name} is still running.", percent)
            last_heartbeat = time.monotonic()
        time.sleep(0.2)
    stderr = process.stderr.read() if process.stderr else ""
    if process.returncode != 0:
        raise ContractError("RAN_PIPELINE_STAGE_FAILED", f"RAN pipeline stage failed: {script_name}.", "domain_processing", {"stage": script_name, "exitCode": process.returncode, "stderrTail": stderr[-1000:]})


def output_item(path: Path, workspace: Path) -> dict[str, Any]:
    return {
        "name": path.stem,
        "path": path.resolve().relative_to(workspace.resolve()).as_posix(),
        "mediaType": mimetypes.guess_type(path.name)[0] or "application/octet-stream",
        "displayName": path.name,
        "size": path.stat().st_size,
        "sha256": sha256(path),
    }


def write_result(path: Path, payload: dict[str, Any]) -> None:
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(temporary, path)


def site_count(output: Path) -> int:
    source = output / "simple_pr_output.json"
    try:
        payload = json.loads(source.read_text(encoding="utf-8"))
        return len(payload) if isinstance(payload, dict) else 0
    except (OSError, json.JSONDecodeError):
        return 0


def run(input_manifest: Path) -> int:
    envelope: dict[str, Any] = {}
    result_path = input_manifest.resolve().parent / "result.json"
    try:
        envelope, workspace, output, result_path, cancellation = load_envelope(input_manifest)
        signal.signal(signal.SIGTERM, lambda *_: (_ for _ in ()).throw(CancelledError()))
        bom = declared_file(envelope, workspace, "bom")
        epms = declared_file(envelope, workspace, "epms")
        run_mode, project = validate_parameters(envelope.get("parameters") or {})
        check_cancel(cancellation)
        emit("progress", "contract_validation", "Validated RAN inputs and parameters.", 5)
        config_target = workspace / "config"
        if config_target.exists():
            shutil.rmtree(config_target)
        shutil.copytree(SKILL_ROOT / "config", config_target)
        env = os.environ.copy()
        env.update({"BOM_FILE_PATH": str(bom), "EPMS_FILE_PATH": str(epms), "RAN_RUN_MODE": run_mode, "PYTHONUTF8": "1", "PYTHONIOENCODING": "utf-8"})
        if project:
            env.update({"SELECTED_PROJECT": project, "GENERAL_ITEM_PROJECT": project})
        for index, stage in enumerate(PIPELINE):
            run_stage(stage, workspace, env, cancellation, 15 + index * 20)
        check_cancel(cancellation)
        outputs = [output / name for name in APPROVED_OUTPUTS if (output / name).is_file()]
        if not outputs:
            raise ContractError("RAN_OUTPUT_MISSING", "The RAN pipeline produced no approved outputs.", "domain_processing")
        sites = site_count(output)
        payload = {
            "schemaVersion": CONTRACT_VERSION,
            "jobId": envelope["jobId"],
            "skillId": SKILL_ID,
            "skillVersion": SKILL_VERSION,
            "status": "succeeded",
            "summary": {"message": "RAN PR processing completed.", "metrics": {"runMode": run_mode, "selectedProject": project, "siteCount": sites, "outputFileCount": len(outputs)}},
            "outputs": [output_item(path, workspace) for path in outputs],
            "warnings": [],
            "error": None,
        }
        write_result(result_path, payload)
        emit("progress", "completed", "RAN PR processing completed.", 100)
        return 0
    except CancelledError as exc:
        status, exit_code, error = "cancelled", 130, exc
    except ContractError as exc:
        status, exit_code, error = "failed", 2, exc
    except Exception as exc:
        status, exit_code = "failed", 4
        error = ContractError(getattr(exc, "code", "RAN_PR_FAILED"), str(exc), "domain_processing")
    payload = {
        "schemaVersion": CONTRACT_VERSION,
        "jobId": str(envelope.get("jobId") or "unknown"),
        "skillId": SKILL_ID,
        "skillVersion": SKILL_VERSION,
        "status": status,
        "summary": {"message": str(error), "metrics": {}},
        "outputs": [],
        "warnings": [],
        "error": {"code": error.code, "category": error.category, "message": str(error), "retryable": False, "details": error.details},
    }
    result_path.parent.mkdir(parents=True, exist_ok=True)
    write_result(result_path, payload)
    emit("warning", status, str(error))
    return exit_code


def main() -> int:
    return run(parse_cli().input_manifest)


if __name__ == "__main__":
    raise SystemExit(main())
