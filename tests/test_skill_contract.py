import hashlib
import importlib.util
import json
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location("ran_contract", ROOT / "src" / "main.py")
MODULE = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(MODULE)


def test_manifest_identity_matches_entrypoint():
    manifest = json.loads((ROOT / "skill.json").read_text(encoding="utf-8"))
    assert manifest["skillId"] == MODULE.SKILL_ID
    assert manifest["version"] == MODULE.SKILL_VERSION
    assert manifest["resultContractVersion"] == MODULE.CONTRACT_VERSION


def test_declared_file_rejects_checksum_mismatch(tmp_path):
    input_dir = tmp_path / "input"
    input_dir.mkdir()
    source = input_dir / "bom.xlsx"
    source.write_bytes(b"workbook")
    envelope = {"files": [{"name": "bom", "path": "input/bom.xlsx", "size": 8, "sha256": "0" * 64}]}
    try:
        MODULE.declared_file(envelope, tmp_path, "bom")
    except MODULE.ContractError as exc:
        assert exc.code == "INPUT_FILE_CHECKSUM_MISMATCH"
    else:
        raise AssertionError("checksum mismatch should fail closed")


def test_output_item_contains_actual_checksum(tmp_path):
    output = tmp_path / "output"
    output.mkdir()
    artifact = output / "result.json"
    artifact.write_bytes(b"{}")
    item = MODULE.output_item(artifact, tmp_path)
    assert item["path"] == "output/result.json"
    assert item["sha256"] == hashlib.sha256(b"{}").hexdigest()
