"""Audit helpers for resumable Microsoft Office COM baseline reports.

This intentionally uses Python's UTF-8 JSON parser instead of Windows
PowerShell 5.1 ConvertFrom-Json: Office COM errors can contain Unicode which
the latter has proven unreliable at decoding in long-running batch supervisors.
"""

from __future__ import annotations

import argparse
import json
import os
import time
from collections import Counter
from pathlib import Path


EXPECTED_BY_EXTENSION = {
    ".doc": 1000,
    ".docx": 1000,
    ".ppt": 1000,
    ".pptx": 1008,
    ".xls": 1000,
    ".xlsx": 1000,
}


def load(path: Path) -> dict:
    # The supervisor replaces checkpoints atomically, but Windows readers can
    # briefly receive ERROR_SHARING_VIOLATION while antivirus or the writer
    # holds the destination. A report command is diagnostic-only, so retry the
    # bounded transient instead of presenting a healthy active audit as a
    # failed/absent report. JSON decoding is retried too in case a non-local
    # filesystem exposes the new name before all metadata is visible.
    last_error: OSError | json.JSONDecodeError | None = None
    for attempt in range(12):
        try:
            with path.open(encoding="utf-8") as stream:
                return json.load(stream)
        except (OSError, json.JSONDecodeError) as exc:
            last_error = exc
            if attempt == 11:
                break
            time.sleep(0.15)
    assert last_error is not None
    raise last_error


def path_key(value: str | None) -> str:
    """Canonical path identity for mixing bulk-relative and retry-absolute paths."""
    if not value:
        return ""
    return os.path.normcase(os.path.abspath(value))


def classify_error(item: dict) -> str:
    if not item.get("error"):
        return ""
    existing = item.get("diagnosis")
    if existing:
        return existing
    message = item["error"].lower()
    if "file block" in message or "trust center" in message or "文件阻止" in item["error"] or "信任中心" in item["error"]:
        return "office-policy-blocked"
    if "80070520" in message or "specified logon session does not exist" in message or "指定的登录会话不存在" in item["error"]:
        return "office-session-unavailable"
    if "password-protected office package" in message:
        return "office-password-protected"
    return "office-baseline-unavailable"


def summary(report: dict, report_path: Path) -> int:
    files = report.get("files", [])
    source = report.get("summary", {})
    errors = Counter(classify_error(item) for item in files if item.get("error"))
    scopes = Counter(item.get("comparisonScope", "not-compared") or "not-compared" for item in files)
    diagnoses = Counter(item.get("diagnosis") or classify_error(item) or "unclassified" for item in files)
    print(f"report: {report_path}")
    print(f"paths: {len(files)}; compared: {source.get('compared', 0)}; errors: {source.get('errors', 0)}")
    print(f"token F1: {source.get('f1', 0):.9f}; Office images: {source.get('officeImages', 0)}; extracted images: {source.get('extractedImages', 0)}")
    print("errors: " + ", ".join(f"{key}={value}" for key, value in sorted(errors.items())))
    print("scopes: " + ", ".join(f"{key}={value}" for key, value in sorted(scopes.items())))
    print("diagnoses: " + ", ".join(f"{key}={value}" for key, value in sorted(diagnoses.items())))
    gate = report.get("qualityGate", {})
    if gate:
        print(
            "Office-visible quality gate: "
            f"compared={gate.get('contentCompared', 0)}; "
            f"exact-text={gate.get('contentTextMatches', 0)}; "
            f"image-count={gate.get('contentImageMatches', 0)}; "
            f"fully-aligned={gate.get('contentFullyAligned', 0)}; "
            f"fully-aligned-rate={gate.get('contentFullyAlignedRate', 0):.9f}; "
            f"Value2-excluded={len(gate.get('excludedScopeMismatchFiles', []))}"
        )
    by_ext = report.get("byExt", {})
    if by_ext:
        print("by extension:")
        for extension in sorted(by_ext):
            values = by_ext[extension]
            print(
                f"  {extension}: paths={values.get('total', 0)}; "
                f"compared={values.get('compared', 0)}; "
                f"errors={values.get('errors', 0)}; F1={values.get('f1', 0):.9f}; "
                f"Office-images={values.get('officeImages', 0)}; extracted-images={values.get('extractedImages', 0)}"
            )
    print(f"reused byte-identical paths: {sum(bool(item.get('reusedFrom')) for item in files)}")
    return 0


def write_errors(report: dict, output: Path, category: str | None, extensions: set[str] | None = None) -> int:
    paths = []
    for item in report.get("files", []):
        kind = classify_error(item)
        # A strict Office recovery pass must retry every failed COM baseline,
        # including Trust Center/File Block failures.  They are not content
        # mismatches and must never be left out merely because the diagnostic
        # is more specific than "office-baseline-unavailable".
        is_baseline_issue = kind in {
            "office-baseline-unavailable",
            "office-policy-blocked",
            "office-session-unavailable",
            "office-password-protected",
        }
        extension = (item.get("ext") or Path(item.get("path") or "").suffix).lower()
        extension_matches = extensions is None or extension in extensions
        if extension_matches and kind and (category is None or kind == category or (category == "office-baseline-issue" and is_baseline_issue)):
            paths.append(item["path"])
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text("\n".join(paths) + ("\n" if paths else ""), encoding="utf-8")
    print(f"wrote {len(paths)} paths to {output}")
    return 0


def write_diagnosis(report: dict, output: Path, diagnosis: str, visible_only: bool) -> int:
    """Export a deterministic path list for a focused content-quality rerun."""
    paths = []
    for item in report.get("files", []):
        current = item.get("diagnosis") or classify_error(item)
        if current != diagnosis:
            continue
        if visible_only and item.get("comparisonScope") != "office-visible":
            continue
        path = item.get("path")
        if path:
            paths.append(path)
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text("\n".join(sorted(paths)) + ("\n" if paths else ""), encoding="utf-8")
    print(f"wrote {len(paths)} {diagnosis} paths to {output}")
    return 0


def suite_summary(directory: Path, overrides: dict[str, Path] | None = None) -> int:
    """Print one path-level audit over the six independently resumable runs.

    Summaries embedded in old checkpoints cannot safely be added: a resumed
    report may have been produced by a previous binary.  Calculate coverage
    and the strict quality gate from each current file record instead.
    """
    reports = sorted(directory.glob("office-com-*-bulk.json"))
    # A recovery replaces records path-for-path.  Select one audit per
    # extension so stale provisional bulk timeouts can never leak into the
    # final 6008-sample suite merely because an override was added later.
    for extension, override in (overrides or {}).items():
        if not override.exists():
            raise SystemExit(f"override report does not exist for {extension}: {override}")
        reports = [path for path in reports if path.name != f"office-com-{extension.lstrip('.')}-bulk.json"]
        reports.append(override)
    if not reports:
        raise SystemExit(f"no bulk reports found in {directory}")
    by_extension: dict[str, list[dict]] = {}
    for path in reports:
        for item in load(path).get("files", []):
            extension = item.get("ext", "").lower()
            if extension:
                by_extension.setdefault(extension, []).append(item)

    print(f"suite reports: {directory}")
    total = compared = unavailable = visible = exact_text = image_count = fully_aligned = stored = 0
    for extension in sorted(EXPECTED_BY_EXTENSION):
        items = by_extension.get(extension, [])
        records = len(items)
        expected = EXPECTED_BY_EXTENSION[extension]
        current_compared = [x for x in items if x.get("baselineStatus") == "compared" and not x.get("error")]
        current_unavailable = [x for x in items if x.get("baselineStatus") == "baseline-unavailable"]
        visible_items = [x for x in current_compared if x.get("comparisonScope") == "office-visible"]
        stored_items = [x for x in current_compared if x.get("comparisonScope") == "office-stored-value"]
        aligned_items = [x for x in visible_items if x.get("f1") == 1 and x.get("imageMatch")]
        print(
            f"  {extension}: coverage={records}/{expected}; compared={len(current_compared)}; "
            f"unavailable={len(current_unavailable)}; visible={len(visible_items)}; "
            f"strict-aligned={len(aligned_items)}/{len(visible_items)}; Value2={len(stored_items)}"
        )
        total += records
        compared += len(current_compared)
        unavailable += len(current_unavailable)
        visible += len(visible_items)
        exact_text += sum(x.get("f1") == 1 for x in visible_items)
        image_count += sum(bool(x.get("imageMatch")) for x in visible_items)
        fully_aligned += len(aligned_items)
        stored += len(stored_items)
    print(
        f"suite coverage={total}/{sum(EXPECTED_BY_EXTENSION.values())}; compared={compared}; "
        f"unavailable={unavailable}; Office-visible={visible}; exact-text={exact_text}; "
        f"image-count={image_count}; fully-aligned={fully_aligned}; Value2-excluded={stored}"
    )
    return 0 if total == sum(EXPECTED_BY_EXTENSION.values()) else 1


def merge_xlsx_recovery(base_path: Path, recovery_path: Path, output_path: Path) -> int:
    """Replace provisional XLSX unavailable records with focused retry results.

    The bounded two-second run is retained separately as a path-coverage
    artifact. This merge produces the current quality report without ever
    treating its provisional timeout as an extraction mismatch. Recovery only
    replaces records for exactly the same absolute path and leaves unresolved
    entries explicit as baseline-unavailable.
    """
    base = load(base_path)
    recovery = load(recovery_path)
    replacements = {path_key(item.get("path")): item for item in recovery.get("files", []) if item.get("path")}
    files = []
    replaced = 0
    for item in base.get("files", []):
        candidate = replacements.get(path_key(item.get("path")))
        # A focused replay is intended to improve a provisional result, not to
        # erase a prior credible Office comparison merely because Word/Excel
        # happened to fail while opening that path on the replay.  Prefer the
        # new record when it compared successfully, or when the base itself
        # was unavailable (where the retry supplies the best evidence).
        candidate_compared = candidate and candidate.get("baselineStatus") == "compared" and not candidate.get("error")
        base_compared = item.get("baselineStatus") == "compared" and not item.get("error")
        if candidate is not None and (candidate_compared or not base_compared):
            files.append(candidate)
            replaced += 1
        else:
            files.append(item)
    # The aggregate fields embedded in a checkpoint are produced by whichever
    # officebaseline binary last wrote it.  After record-level replacement they
    # are necessarily stale, and consumers must never mistake them for a
    # re-evaluated audit.  Rebuild all derived aggregates from the merged path
    # records before writing the artifact.
    base["files"] = files
    rebuild_aggregates(base)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(base, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"merged {replaced} recovery records into {output_path}")
    return 0


def merge_recoveries(base_path: Path, output_path: Path, recovery_paths: list[Path]) -> int:
    """Merge any number of focused recovery checkpoints into one audit.

    Later successful recovery files win for the same path. This makes the
    normal workflow composable: a deliberately bounded first retry can be
    followed by one or more resumable batches without hand-editing JSON or
    accidentally discarding a credible baseline when a later Office launch
    happens to fail.
    """
    base = load(base_path)
    replacements: dict[str, dict] = {}
    for recovery_path in recovery_paths:
        for item in load(recovery_path).get("files", []):
            key = path_key(item.get("path"))
            if key:
                replacements[key] = item
    replaced = 0
    merged = []
    for item in base.get("files", []):
        candidate = replacements.get(path_key(item.get("path")))
        candidate_compared = candidate and candidate.get("baselineStatus") == "compared" and not candidate.get("error")
        base_compared = item.get("baselineStatus") == "compared" and not item.get("error")
        if candidate is not None and (candidate_compared or not base_compared):
            merged.append(candidate)
            replaced += 1
        else:
            merged.append(item)
    base["files"] = merged
    rebuild_aggregates(base)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(base, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"merged {replaced} records from {len(recovery_paths)} recovery reports into {output_path}")
    return 0


def rebuild_aggregates(report: dict) -> None:
    """Recalculate the report aggregates from its path-level records.

    Focused COM recovery deliberately writes a second report and later
    substitutes only the matching paths into the coverage report.  This
    helper keeps that operation auditable: no summary number survives merely
    because it happened to belong to the pre-recovery checkpoint.
    """
    files = report.get("files", [])
    compared = [x for x in files if x.get("baselineStatus") == "compared" and not x.get("error")]
    errors = [x for x in files if x.get("error")]
    office_images = sum(int(x.get("officeImages", 0) or 0) for x in compared)
    extracted_images = sum(int(x.get("extractedImages", 0) or 0) for x in compared)
    matched_tokens = sum(int(x.get("matchedTokens", 0) or 0) for x in compared)
    office_tokens = sum(int(x.get("officeTokens", 0) or 0) for x in compared)
    extracted_tokens = sum(int(x.get("extractedTokens", 0) or 0) for x in compared)
    ordered = [x for x in compared if x.get("orderedComparisonAvailable")]
    ordered_matched = sum(int(x.get("orderedMatchedTokens", 0) or 0) for x in ordered)
    ordered_office = sum(int(x.get("officeTokens", 0) or 0) for x in ordered)
    ordered_extracted = sum(int(x.get("extractedTokens", 0) or 0) for x in ordered)
    precision = matched_tokens / extracted_tokens if extracted_tokens else 0.0
    recall = matched_tokens / office_tokens if office_tokens else 0.0
    f1 = 2 * precision * recall / (precision + recall) if precision + recall else 0.0
    report["summary"] = {
        "total": len(files), "compared": len(compared), "errors": len(errors),
        "baselineUnavailable": sum(x.get("baselineStatus") == "baseline-unavailable" for x in errors),
        "officeImages": office_images, "extractedImages": extracted_images,
        "officeTokens": office_tokens, "extractedTokens": extracted_tokens,
        "matchedTokens": matched_tokens, "precision": precision, "recall": recall, "f1": f1,
        "orderedCompared": len(ordered), "orderedMatchedTokens": ordered_matched,
        "orderedOfficeTokens": ordered_office, "orderedExtractedTokens": ordered_extracted,
        "orderedPrecision": ordered_matched / ordered_extracted if ordered_extracted else 0.0,
        "orderedRecall": ordered_matched / ordered_office if ordered_office else 0.0,
    }
    diagnosis_counts = Counter(item.get("diagnosis") or classify_error(item) or "unclassified" for item in files)
    report["summary"]["diagnosisCounts"] = dict(sorted(diagnosis_counts.items()))
    ordered_precision = report["summary"]["orderedPrecision"]
    ordered_recall = report["summary"]["orderedRecall"]
    report["summary"]["orderedF1"] = 2 * ordered_precision * ordered_recall / (ordered_precision + ordered_recall) if ordered_precision + ordered_recall else 0.0
    by_ext: dict[str, dict] = {}
    for item in files:
        extension = (item.get("ext") or "").lower()
        if not extension:
            continue
        group = by_ext.setdefault(extension, {"total": 0, "compared": 0, "errors": 0, "baselineUnavailable": 0, "officeImages": 0, "extractedImages": 0, "officeTokens": 0, "extractedTokens": 0, "matchedTokens": 0, "orderedCompared": 0, "orderedMatchedTokens": 0, "orderedOfficeTokens": 0, "orderedExtractedTokens": 0, "diagnosisCounts": {}})
        group["total"] += 1
        if item.get("baselineStatus") == "compared" and not item.get("error"):
            group["compared"] += 1
            for key in ("officeImages", "extractedImages", "officeTokens", "extractedTokens", "matchedTokens"):
                group[key] += int(item.get(key, 0) or 0)
        if item.get("error"):
            group["errors"] += 1
            if item.get("baselineStatus") == "baseline-unavailable":
                group["baselineUnavailable"] += 1
        diagnosis = item.get("diagnosis") or classify_error(item) or "unclassified"
        group["diagnosisCounts"][diagnosis] = group["diagnosisCounts"].get(diagnosis, 0) + 1
        if item.get("baselineStatus") == "compared" and not item.get("error") and item.get("orderedComparisonAvailable"):
            group["orderedCompared"] += 1
            group["orderedMatchedTokens"] += int(item.get("orderedMatchedTokens", 0) or 0)
            group["orderedOfficeTokens"] += int(item.get("officeTokens", 0) or 0)
            group["orderedExtractedTokens"] += int(item.get("extractedTokens", 0) or 0)
    for group in by_ext.values():
        group["precision"] = group["matchedTokens"] / group["extractedTokens"] if group["extractedTokens"] else 0.0
        group["recall"] = group["matchedTokens"] / group["officeTokens"] if group["officeTokens"] else 0.0
        group["f1"] = 2 * group["precision"] * group["recall"] / (group["precision"] + group["recall"]) if group["precision"] + group["recall"] else 0.0
        group["orderedPrecision"] = group["orderedMatchedTokens"] / group["orderedExtractedTokens"] if group["orderedExtractedTokens"] else 0.0
        group["orderedRecall"] = group["orderedMatchedTokens"] / group["orderedOfficeTokens"] if group["orderedOfficeTokens"] else 0.0
        group["orderedF1"] = 2 * group["orderedPrecision"] * group["orderedRecall"] / (group["orderedPrecision"] + group["orderedRecall"]) if group["orderedPrecision"] + group["orderedRecall"] else 0.0
        group["diagnosisCounts"] = dict(sorted(group["diagnosisCounts"].items()))
    report["byExt"] = dict(sorted(by_ext.items()))
    visible = [x for x in compared if x.get("comparisonScope") == "office-visible"]
    excluded = [x.get("path") for x in compared if x.get("comparisonScope") != "office-visible" and x.get("path")]
    text_matches = sum(x.get("f1") == 1 for x in visible)
    image_matches = sum(bool(x.get("imageMatch")) for x in visible)
    fully_aligned = sum(x.get("f1") == 1 and bool(x.get("imageMatch")) for x in visible)
    denominator = len(visible)
    report["qualityGate"] = {
        "compared": len(compared),
        "baselineUnavailable": sum(x.get("baselineStatus") == "baseline-unavailable" for x in errors),
        "contentCompared": denominator, "contentTextMatches": text_matches,
        "contentImageMatches": image_matches, "contentFullyAligned": fully_aligned,
        "contentTextMatchRate": text_matches / denominator if denominator else 0.0,
        "contentImageMatchRate": image_matches / denominator if denominator else 0.0,
        "contentFullyAlignedRate": fully_aligned / denominator if denominator else 0.0,
        "excludedScopeMismatchFiles": excluded,
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("report", type=Path, nargs="?")
    parser.add_argument("--write-errors", type=Path)
    parser.add_argument("--category", choices=["office-baseline-unavailable", "office-policy-blocked", "office-session-unavailable", "office-password-protected", "office-baseline-issue"])
    parser.add_argument("--extensions", help="with --write-errors, comma-separated extension filter such as .xls,.xlsx")
    parser.add_argument("--write-diagnosis", type=Path, help="write paths with one diagnosis for focused reruns")
    parser.add_argument("--diagnosis", choices=["aligned", "text-mismatch", "image-mismatch", "text-and-image-mismatch", "baseline-scope-mismatch", "office-baseline-unavailable", "office-policy-blocked", "office-session-unavailable", "office-password-protected", "extractor-error"])
    parser.add_argument("--visible-only", action="store_true", help="with --write-diagnosis, retain only Office-visible comparisons")
    parser.add_argument("--suite", type=Path, metavar="REPORT_DIR", help="summarize the six path-level reports, applying explicit recovered-report overrides")
    parser.add_argument("--report-override", action="append", default=[], metavar="EXT=PATH", help="with --suite, replace one bulk report (for example .xlsx=reports/current.json); repeat per extension")
    parser.add_argument("--merge-xlsx-recovery", nargs=3, metavar=("BASE", "RECOVERY", "OUTPUT"), help="replace provisional XLSX unavailable records with one focused-retry report")
    parser.add_argument("--merge-recoveries", nargs="+", type=Path, metavar="PATH", help="BASE OUTPUT RECOVERY [RECOVERY ...]: merge multiple focused-recovery reports")
    args = parser.parse_args()
    if args.merge_xlsx_recovery:
        base, recovery, output = map(Path, args.merge_xlsx_recovery)
        return merge_xlsx_recovery(base, recovery, output)
    if args.merge_recoveries:
        if len(args.merge_recoveries) < 3:
            parser.error("--merge-recoveries requires BASE OUTPUT and at least one RECOVERY report")
        base, output, *recoveries = args.merge_recoveries
        return merge_recoveries(base, output, recoveries)
    if args.suite:
        overrides: dict[str, Path] = {}
        for value in args.report_override:
            extension, separator, path = value.partition("=")
            if not separator or not extension.startswith(".") or not path:
                parser.error("--report-override must be EXT=PATH, for example .xlsx=reports/current.json")
            overrides[extension.lower()] = Path(path)
        return suite_summary(args.suite, overrides)
    if args.report is None:
        parser.error("report is required unless --suite is used")
    report = load(args.report)
    if args.write_errors:
        extensions = None
        if args.extensions:
            extensions = {value.strip().lower() for value in args.extensions.split(",") if value.strip()}
            if not extensions or any(not value.startswith(".") for value in extensions):
                parser.error("--extensions must be a comma-separated list such as .xls,.xlsx")
        return write_errors(report, args.write_errors, args.category, extensions)
    if args.write_diagnosis:
        if not args.diagnosis:
            parser.error("--diagnosis is required with --write-diagnosis")
        return write_diagnosis(report, args.write_diagnosis, args.diagnosis, args.visible_only)
    return summary(report, args.report)


if __name__ == "__main__":
    raise SystemExit(main())
