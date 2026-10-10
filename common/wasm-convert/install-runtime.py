#!/usr/bin/env python3
"""Install the pinned full x2t release; generated artifacts remain outside Git."""
import hashlib
import json
import shutil
import sys
from pathlib import Path

root = Path(__file__).resolve().parent
release = Path(sys.argv[1])
pin = json.loads((root / "runtime-release.json").read_text())
raw = (release / "manifest.json").read_bytes()
if hashlib.sha256(raw).hexdigest() != pin["manifest_sha256"]:
    raise SystemExit("unexpected full conversion manifest")
manifest = json.loads(raw)
if manifest.get("build_flags", {}).get("module_kind") == "pdf":
    raise SystemExit("local conversion requires the full module")
if manifest["artifact_version"] != pin["artifact_version"]:
    raise SystemExit("unexpected conversion version")
names = {"x2t.js", "x2t.wasm", "x2t.js.gz", "x2t.js.br", "x2t.wasm.gz", "x2t.wasm.br"}
if set(manifest["artifacts"]) != names:
    raise SystemExit("incomplete conversion artifact set")
for name in names:
    data = (release / name).read_bytes()
    item = manifest["artifacts"][name]
    if len(data) != item["bytes"] or hashlib.sha256(data).hexdigest() != item["sha256"]:
        raise SystemExit("conversion artifact checksum mismatch: " + name)
destination = root / "runtime" / pin["artifact_version"]
destination.mkdir(parents=True, exist_ok=True)
for name in names | {"manifest.json"}:
    shutil.copy2(release / name, destination / name)
print("installed full conversion runtime", pin["artifact_version"])
