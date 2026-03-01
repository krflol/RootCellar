#!/usr/bin/env python3
"""Build a deterministic nightly batch corpus slice for CI throughput checks."""

from __future__ import annotations

import argparse
import json
import pathlib
import shutil
from datetime import datetime, timezone

import generate_corpus_fixtures


CURATED_WORKBOOKS = (
    "sample-formula.xlsx",
    "tx-source.xlsx",
    "preserve-source.xlsx",
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Build a deterministic nightly batch corpus from generated fixtures "
            "plus curated workbook samples."
        )
    )
    parser.add_argument(
        "--output-dir",
        default=".ci/batch-nightly-corpus",
        help="Directory where the assembled nightly corpus slice is written.",
    )
    parser.add_argument(
        "--target-files",
        type=int,
        default=32,
        help="Target number of corpus files to include after deterministic replication.",
    )
    parser.add_argument(
        "--repo-root",
        default=".",
        help="Repository root used to resolve curated workbook sample paths.",
    )
    return parser.parse_args()


def collect_seed_files(repo_root: pathlib.Path, fixtures_dir: pathlib.Path) -> list[pathlib.Path]:
    seed_files: list[pathlib.Path] = []

    generated = generate_corpus_fixtures.generate(fixtures_dir)
    seed_files.extend(sorted(generated, key=lambda p: p.name))

    for rel in CURATED_WORKBOOKS:
        candidate = repo_root / rel
        if candidate.is_file() and candidate.suffix.lower() == ".xlsx":
            seed_files.append(candidate)

    # Preserve deterministic ordering while removing duplicates.
    deduped = []
    seen = set()
    for path in sorted(seed_files, key=lambda p: str(p).replace("\\", "/")):
        key = str(path.resolve()).lower()
        if key in seen:
            continue
        seen.add(key)
        deduped.append(path)
    return deduped


def assemble_corpus(
    output_dir: pathlib.Path, repo_root: pathlib.Path, target_files: int
) -> dict:
    if target_files <= 0:
        raise ValueError("--target-files must be greater than zero")

    files_dir = output_dir / "files"
    fixtures_dir = output_dir / "fixtures"

    if output_dir.exists():
        shutil.rmtree(output_dir)
    files_dir.mkdir(parents=True, exist_ok=True)
    fixtures_dir.mkdir(parents=True, exist_ok=True)

    seeds = collect_seed_files(repo_root, fixtures_dir)
    if not seeds:
        raise RuntimeError("no seed workbooks available for nightly batch corpus assembly")

    copied_files: list[pathlib.Path] = []
    source_counts: dict[str, int] = {}

    # Copy each seed at least once.
    for idx, src in enumerate(seeds):
        dest = files_dir / f"seed-{idx:03d}-{src.name}"
        shutil.copy2(src, dest)
        copied_files.append(dest)
        key = src.name
        source_counts[key] = source_counts.get(key, 0) + 1

    # Deterministically replicate seeds until target size is reached.
    replica_idx = 0
    while len(copied_files) < target_files:
        src = seeds[replica_idx % len(seeds)]
        dest = files_dir / f"replica-{len(copied_files):03d}-{src.stem}.xlsx"
        shutil.copy2(src, dest)
        copied_files.append(dest)
        key = src.name
        source_counts[key] = source_counts.get(key, 0) + 1
        replica_idx += 1

    manifest = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "target_files": target_files,
        "actual_files": len(copied_files),
        "seed_file_count": len(seeds),
        "seed_files": [str(p).replace("\\", "/") for p in seeds],
        "files_dir": str(files_dir).replace("\\", "/"),
        "fixtures_dir": str(fixtures_dir).replace("\\", "/"),
        "source_distribution": dict(sorted(source_counts.items())),
    }
    manifest_path = output_dir / "manifest.json"
    manifest_path.write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    manifest["manifest_path"] = str(manifest_path).replace("\\", "/")
    return manifest


def main() -> None:
    args = parse_args()
    output_dir = pathlib.Path(args.output_dir)
    repo_root = pathlib.Path(args.repo_root)
    manifest = assemble_corpus(output_dir, repo_root, args.target_files)

    print(
        f"Built nightly batch corpus in {output_dir} with {manifest['actual_files']} files "
        f"from {manifest['seed_file_count']} seeds"
    )
    print(f"Corpus files directory: {manifest['files_dir']}")
    print(f"Corpus manifest: {manifest['manifest_path']}")
    for source_name, count in manifest["source_distribution"].items():
        print(f" - {source_name}: {count}")


if __name__ == "__main__":
    main()
