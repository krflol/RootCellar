#!/usr/bin/env python3
"""Assemble deterministic Excel interop corpus from fixtures + curated samples.

The assembled corpus contains:
1) Generated deterministic fixtures from `generate_corpus_fixtures.py`.
2) Optional curated real Excel-authored samples listed in
   `corpus/excel-authored/manifest.json`.
"""

from __future__ import annotations

import argparse
import json
import pathlib
import shutil
from datetime import datetime, timezone

import generate_corpus_fixtures


ALLOWED_LEGAL_CLEARANCE = {
    "approved",
    "restricted_internal",
    "restricted_partner",
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Assemble deterministic Excel interop corpus from generated fixtures "
            "plus optional curated Excel-authored samples."
        )
    )
    parser.add_argument(
        "--output-dir",
        default=".ci/excel-interop-corpus",
        help="Directory where assembled corpus files and manifest are written.",
    )
    parser.add_argument(
        "--repo-root",
        default=".",
        help="Repository root used to resolve curated sample paths.",
    )
    parser.add_argument(
        "--excel-authored-manifest",
        default="./corpus/excel-authored/manifest.json",
        help="Curated Excel-authored sample metadata manifest path.",
    )
    parser.add_argument(
        "--min-excel-authored-samples",
        type=int,
        default=0,
        help="Minimum curated Excel-authored samples required in assembled corpus.",
    )
    parser.add_argument(
        "--strict-manifest",
        action="store_true",
        help="Fail if the curated manifest path does not exist.",
    )
    parser.add_argument(
        "--required-curated-feature",
        action="append",
        default=[],
        help=(
            "Required curated feature tag that must be present in at least one "
            "curated sample. May be supplied multiple times."
        ),
    )
    return parser.parse_args()


def load_json(path: pathlib.Path) -> dict:
    payload = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        raise RuntimeError(f"expected JSON object at {path}")
    return payload


def normalize_sample_features(value: object, *, sample_id: str) -> list[str]:
    if not isinstance(value, list):
        raise RuntimeError(f"sample {sample_id} field 'features' must be a list")
    cleaned = []
    for item in value:
        if not isinstance(item, str):
            raise RuntimeError(
                f"sample {sample_id} field 'features' must contain only strings"
            )
        token = item.strip().lower()
        if token:
            cleaned.append(token)
    deduped = sorted(set(cleaned))
    if not deduped:
        raise RuntimeError(f"sample {sample_id} must include at least one feature tag")
    return deduped


def resolve_curated_path(
    *,
    sample_id: str,
    path_value: str,
    manifest_dir: pathlib.Path,
    repo_root: pathlib.Path,
) -> pathlib.Path:
    rel = pathlib.Path(path_value)
    if rel.is_absolute():
        raise RuntimeError(
            f"sample {sample_id} path must be repository-relative (not absolute): {path_value}"
        )
    candidate = (manifest_dir / rel).resolve()
    if not str(candidate).lower().startswith(str(repo_root.resolve()).lower()):
        raise RuntimeError(
            f"sample {sample_id} path escapes repository root: {path_value}"
        )
    if not candidate.is_file():
        raise RuntimeError(f"sample {sample_id} file does not exist: {candidate}")
    if candidate.suffix.lower() != ".xlsx":
        raise RuntimeError(
            f"sample {sample_id} must point to an .xlsx file: {candidate}"
        )
    return candidate


def load_curated_samples(
    *,
    repo_root: pathlib.Path,
    manifest_path: pathlib.Path,
    strict_manifest: bool,
) -> tuple[list[dict], dict]:
    if not manifest_path.is_file():
        if strict_manifest:
            raise RuntimeError(
                f"curated Excel-authored manifest does not exist: {manifest_path}"
            )
        return [], {"status": "missing", "path": str(manifest_path.resolve())}

    payload = load_json(manifest_path)
    samples_payload = payload.get("samples", [])
    if not isinstance(samples_payload, list):
        raise RuntimeError("curated manifest field 'samples' must be a list")

    manifest_dir = manifest_path.parent.resolve()
    seen_ids: set[str] = set()
    curated_samples: list[dict] = []

    for sample in samples_payload:
        if not isinstance(sample, dict):
            raise RuntimeError("curated manifest entries must be objects")

        sample_id = str(sample.get("id", "")).strip()
        if not sample_id:
            raise RuntimeError("curated sample missing required field: id")
        if sample_id in seen_ids:
            raise RuntimeError(f"duplicate curated sample id: {sample_id}")
        seen_ids.add(sample_id)

        path_value = str(sample.get("path", "")).strip()
        if not path_value:
            raise RuntimeError(f"sample {sample_id} missing required field: path")

        authoring_app = str(sample.get("authoring_app", "")).strip()
        if not authoring_app:
            raise RuntimeError(
                f"sample {sample_id} missing required field: authoring_app"
            )

        legal_clearance = str(sample.get("legal_clearance", "")).strip().lower()
        if legal_clearance not in ALLOWED_LEGAL_CLEARANCE:
            raise RuntimeError(
                f"sample {sample_id} legal_clearance must be one of "
                f"{sorted(ALLOWED_LEGAL_CLEARANCE)}"
            )

        source_category = str(sample.get("source_category", "")).strip()
        if not source_category:
            raise RuntimeError(
                f"sample {sample_id} missing required field: source_category"
            )

        features = normalize_sample_features(
            sample.get("features", []),
            sample_id=sample_id,
        )
        resolved_path = resolve_curated_path(
            sample_id=sample_id,
            path_value=path_value,
            manifest_dir=manifest_dir,
            repo_root=repo_root,
        )

        curated_samples.append(
            {
                "id": sample_id,
                "path": resolved_path,
                "path_value": path_value,
                "authoring_app": authoring_app,
                "legal_clearance": legal_clearance,
                "source_category": source_category,
                "features": features,
                "notes": str(sample.get("notes", "")).strip(),
            }
        )

    curated_samples.sort(key=lambda x: x["id"])
    meta = {
        "status": "loaded",
        "path": str(manifest_path.resolve()),
        "sample_count": len(curated_samples),
    }
    return curated_samples, meta


def assemble_corpus(
    *,
    output_dir: pathlib.Path,
    repo_root: pathlib.Path,
    curated_manifest_path: pathlib.Path,
    min_excel_authored_samples: int,
    strict_manifest: bool,
    required_curated_features: list[str],
) -> dict:
    if min_excel_authored_samples < 0:
        raise RuntimeError("--min-excel-authored-samples must be >= 0")

    if output_dir.exists():
        shutil.rmtree(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    generated_dir = output_dir / "generated-fixtures"
    curated_out_dir = output_dir / "excel-authored"
    generated_dir.mkdir(parents=True, exist_ok=True)
    curated_out_dir.mkdir(parents=True, exist_ok=True)

    fixture_descriptors = generate_corpus_fixtures.generate_descriptors(generated_dir)
    generate_corpus_fixtures.write_manifest(
        generated_dir,
        generated_dir / "manifest.json",
        fixture_descriptors,
    )

    curated_samples, curated_meta = load_curated_samples(
        repo_root=repo_root,
        manifest_path=curated_manifest_path,
        strict_manifest=strict_manifest,
    )
    if len(curated_samples) < min_excel_authored_samples:
        raise RuntimeError(
            "insufficient curated Excel-authored samples: "
            f"required={min_excel_authored_samples}, found={len(curated_samples)}"
        )

    normalized_required_features = sorted(
        {
            feature.strip().lower()
            for feature in required_curated_features
            if feature.strip()
        }
    )
    curated_feature_coverage = sorted(
        {feature for sample in curated_samples for feature in sample["features"]}
    )
    missing_required_curated_features = sorted(
        set(normalized_required_features) - set(curated_feature_coverage)
    )
    if missing_required_curated_features:
        raise RuntimeError(
            "required curated features missing from curated sample set: "
            f"{missing_required_curated_features}; "
            f"coverage={curated_feature_coverage}"
        )

    copied_curated: list[dict] = []
    for index, sample in enumerate(curated_samples, start=1):
        dest_name = f"{index:03d}-{sample['id']}.xlsx"
        dest = curated_out_dir / dest_name
        shutil.copy2(sample["path"], dest)
        copied_curated.append(
            {
                "id": sample["id"],
                "source_path": str(sample["path"]),
                "declared_path": sample["path_value"],
                "copied_path": str(dest),
                "relative_copied_path": str(dest.relative_to(output_dir)).replace("\\", "/"),
                "authoring_app": sample["authoring_app"],
                "legal_clearance": sample["legal_clearance"],
                "source_category": sample["source_category"],
                "features": sample["features"],
                "notes": sample["notes"],
            }
        )

    manifest = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "output_dir": str(output_dir.resolve()),
        "required_fixture_basenames": sorted([d.path.name for d in fixture_descriptors]),
        "generated_fixture_count": len(fixture_descriptors),
        "generated_fixture_dir": str(generated_dir.resolve()),
        "excel_authored_sample_count": len(copied_curated),
        "excel_authored_min_required": min_excel_authored_samples,
        "required_curated_features": normalized_required_features,
        "covered_curated_features": curated_feature_coverage,
        "missing_required_curated_features": missing_required_curated_features,
        "excel_authored_manifest": curated_meta,
        "excel_authored_samples": copied_curated,
    }

    manifest_path = output_dir / "manifest.json"
    manifest_path.write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    manifest["manifest_path"] = str(manifest_path.resolve())
    return manifest


def main() -> None:
    args = parse_args()
    output_dir = pathlib.Path(args.output_dir)
    repo_root = pathlib.Path(args.repo_root).resolve()
    curated_manifest_path = pathlib.Path(args.excel_authored_manifest)
    if not curated_manifest_path.is_absolute():
        curated_manifest_path = (repo_root / curated_manifest_path).resolve()

    manifest = assemble_corpus(
        output_dir=output_dir,
        repo_root=repo_root,
        curated_manifest_path=curated_manifest_path,
        min_excel_authored_samples=args.min_excel_authored_samples,
        strict_manifest=args.strict_manifest,
        required_curated_features=args.required_curated_feature,
    )

    print(
        "Assembled Excel interop corpus "
        f"generated={manifest['generated_fixture_count']} "
        f"excel_authored={manifest['excel_authored_sample_count']} "
        f"min_required={manifest['excel_authored_min_required']}"
    )
    print(f"Corpus directory: {output_dir}")
    print(f"Corpus manifest: {manifest['manifest_path']}")


if __name__ == "__main__":
    main()
