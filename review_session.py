"""
review_session.py — save a review into one ZIP and re-open it later.

Why this exists: everything the review screen knows (which rows are
flagged, every height you adjusted, rows you deleted or added, which
screenshots you ticked as reviewed) lives only in the browser session.
A page refresh wipes it. And the report ZIP a person downloads used to
hold the finished CSV/PPTX/images plus a manifest that POINTS at the
original screenshots — but didn't contain them, so it couldn't be
re-opened on its own.

The bundle fixes both halves. Layout of the downloaded ZIP:

    images/ , <Type>_SEO_Report_*.csv/.pptx , review_manifest.json
        ^ exactly what the report ZIP always contained, same place
    _review_session/
        source/<Type>/<day>/<original screenshots>   (same layout the
                                                       manifest's paths
                                                       are relative to)
        template/<template>.pptx
        session_extras.json   (report type, reviewed ticks, CSV extras)

Nothing in here imports Streamlit, so it can be tested directly.
"""

import io
import json
import os
import shutil
import tempfile
import zipfile

SESSION_DIR = "_review_session"
SOURCE_SUBDIR = "source"
TEMPLATE_SUBDIR = "template"
EXTRAS_NAME = "session_extras.json"
MANIFEST_NAME = "review_manifest.json"

# Already-compressed formats gain nothing from deflate and just burn CPU
# (and time) re-compressing — store them as-is.
_ALREADY_COMPRESSED = (".jpg", ".jpeg", ".png", ".pptx", ".potx", ".docx", ".zip")


class MissingScreenshotsError(ValueError):
    """The review refers to screenshots that aren't in what was uploaded."""

    def __init__(self, missing):
        self.missing = list(missing)
        shown = ", ".join(os.path.basename(m) for m in self.missing[:5])
        more = f" (+{len(self.missing) - 5} more)" if len(self.missing) > 5 else ""
        super().__init__(
            f"{len(self.missing)} original screenshot file(s) are missing: {shown}{more}")


# ════════════════════════════════════════════════════════════════════════
# Saving
# ════════════════════════════════════════════════════════════════════════

def _jsonable_extra_csv(extra):
    if not extra:
        return {}
    return {kw: {str(wk): list(vals) for wk, vals in weeks.items()}
            for kw, weeks in extra.items()}


def _restore_extra_csv(data):
    if not data:
        return None
    return {kw: {int(wk): list(vals) for wk, vals in weeks.items()}
            for kw, weeks in data.items()}


def build_bundle_zip(core, report_root, manifest, report_type, template_path=None,
                     extra_csv_results=None, reviewed_indices=None,
                     include_sources=True) -> bytes:
    """Zip the report folder plus everything needed to re-open the review.

    The manifest on disk is rewritten from the CURRENT in-memory state
    first, so a bundle saved mid-review (before any build) carries the
    latest flags/heights, not whatever was last written to disk.
    """
    template_path = template_path or manifest.get("template_path")
    core.save_manifest(report_root, manifest["entries"], manifest["month_name"],
                       manifest["year"], manifest["month_num"], template_path)

    anchor = os.path.dirname(report_root)
    buf = io.BytesIO()
    seen = set()

    def add(zf, path, arcname):
        arcname = arcname.replace(os.sep, "/")
        if arcname in seen:
            return
        seen.add(arcname)
        comp = (zipfile.ZIP_STORED if path.lower().endswith(_ALREADY_COMPRESSED)
                else zipfile.ZIP_DEFLATED)
        zf.write(path, arcname, compress_type=comp)

    n_sources = 0
    with zipfile.ZipFile(buf, "w") as zf:
        # 1) the report itself, laid out exactly as it always was
        for root, dirs, files in os.walk(report_root):
            dirs[:] = [d for d in dirs if d != SESSION_DIR]
            for fname in files:
                full = os.path.join(root, fname)
                add(zf, full, os.path.relpath(full, report_root))

        # 2) the original screenshots, in the layout the manifest's
        #    relative paths expect (relative to the month folder)
        if include_sources:
            for entry in manifest["entries"]:
                raw = entry["raw_filepath"]
                if not os.path.exists(raw):
                    continue
                rel = os.path.relpath(raw, anchor)
                if rel.startswith(".."):
                    rel = os.path.join("_unmapped", os.path.basename(raw))
                add(zf, raw, f"{SESSION_DIR}/{SOURCE_SUBDIR}/{rel}")
                n_sources += 1

        # 3) the template, so a rebuild can still make the PPTX
        if template_path and os.path.exists(template_path):
            add(zf, template_path,
                f"{SESSION_DIR}/{TEMPLATE_SUBDIR}/{os.path.basename(template_path)}")

        extras = {
            "format_version": 1,
            "report_type": report_type,
            "includes_sources": bool(include_sources),
            "source_count": n_sources,
            "reviewed_indices": sorted(int(i) for i in (reviewed_indices or [])),
            "extra_csv_results": _jsonable_extra_csv(extra_csv_results),
        }
        zf.writestr(f"{SESSION_DIR}/{EXTRAS_NAME}", json.dumps(extras, indent=2))

    return buf.getvalue()


# ════════════════════════════════════════════════════════════════════════
# Restoring
# ════════════════════════════════════════════════════════════════════════

def _find_manifest(extract_dir):
    """Shallowest review_manifest.json — tolerates the ZIP having been
    unpacked and re-zipped with an extra wrapper folder around it."""
    best = None
    for root, dirs, files in os.walk(extract_dir):
        if MANIFEST_NAME in files and SESSION_DIR not in root.split(os.sep):
            depth = root[len(extract_dir):].count(os.sep)
            if best is None or depth < best[0]:
                best = (depth, os.path.join(root, MANIFEST_NAME))
    return best[1] if best else None


def _infer_report_type(bundle_root, manifest_json):
    for fname in os.listdir(bundle_root):
        low = fname.lower()
        if low.startswith("google_seo_report"):
            return "Google"
        if low.startswith("explorer_seo_report"):
            return "Explorer"
    for e in manifest_json.get("entries", []):
        first = e.get("raw_filepath", "").replace("\\", "/").split("/")[0].lower()
        if first == "google":
            return "Google"
        if first == "explorer":
            return "Explorer"
    return None


def _normalize_manifest_paths(manifest_path):
    """A manifest written on Windows stores backslash paths; on Linux
    (Streamlit Cloud) those are just odd filenames and nothing resolves.
    Rewrite them to forward slashes (which Windows accepts too) before
    loading, so a review started on one OS opens on the other."""
    with open(manifest_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    if data.get("template_path"):
        data["template_path"] = data["template_path"].replace("\\", "/")
    for e in data.get("entries", []):
        if e.get("raw_filepath"):
            e["raw_filepath"] = e["raw_filepath"].replace("\\", "/")
        for o in e.get("outputs", []):
            if o.get("path"):
                o["path"] = o["path"].replace("\\", "/")
    with open(manifest_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)
    return data


def _missing_raw(manifest):
    return [e["raw_filepath"] for e in manifest["entries"]
            if not os.path.exists(e["raw_filepath"])]


def restore_bundle(zip_bytes, cores, original_zip_bytes=None) -> dict:
    """Re-open a downloaded report/progress ZIP.

    cores: {"Google": <module>, "Explorer": <module>}
    original_zip_bytes: only for reports downloaded BEFORE bundles held
    the screenshots — the original month ZIP supplies them instead
    (its folder layout matches what those old manifests point at).
    """
    work_dir = tempfile.mkdtemp(prefix="seo_resume_")
    extract_dir = os.path.join(work_dir, "_extract")
    os.makedirs(extract_dir)
    try:
        with zipfile.ZipFile(io.BytesIO(bytes(zip_bytes))) as zf:
            zf.extractall(extract_dir)
    except zipfile.BadZipFile:
        raise ValueError("That isn't a valid ZIP file.")

    manifest_path = _find_manifest(extract_dir)
    if manifest_path is None:
        raise ValueError(
            f"No {MANIFEST_NAME} found inside — this doesn't look like a report ZIP "
            "downloaded from this app.")
    bundle_root = os.path.dirname(manifest_path)
    raw_manifest = _normalize_manifest_paths(manifest_path)

    session_dir = os.path.join(bundle_root, SESSION_DIR)
    extras = {}
    extras_path = os.path.join(session_dir, EXTRAS_NAME)
    if os.path.exists(extras_path):
        with open(extras_path, "r", encoding="utf-8") as f:
            extras = json.load(f)

    report_type = extras.get("report_type") or _infer_report_type(bundle_root, raw_manifest)
    if report_type not in cores:
        raise ValueError("Couldn't tell whether this is a Google or an Explorer report.")
    core = cores[report_type]

    # Rebuild the folder shape the manifest's relative paths expect:
    #   <month>/<Type> Report/   (report files + manifest)
    #   <month>/<Type>/<day>/    (original screenshots)
    month_dir = os.path.join(work_dir, "month")
    report_root = os.path.join(month_dir, f"{report_type} Report")
    os.makedirs(month_dir)
    shutil.move(bundle_root, report_root)
    session_dir = os.path.join(report_root, SESSION_DIR)

    src_dir = os.path.join(session_dir, SOURCE_SUBDIR)
    if os.path.isdir(src_dir):
        shutil.copytree(src_dir, month_dir, dirs_exist_ok=True)

    if original_zip_bytes is not None:
        with zipfile.ZipFile(io.BytesIO(bytes(original_zip_bytes))) as zf:
            zf.extractall(month_dir)

    template_path = None
    tpl_dir = os.path.join(session_dir, TEMPLATE_SUBDIR)
    if os.path.isdir(tpl_dir):
        names = [n for n in os.listdir(tpl_dir) if os.path.isfile(os.path.join(tpl_dir, n))]
        if names:
            keep_dir = os.path.join(work_dir, "template")
            os.makedirs(keep_dir)
            template_path = os.path.join(keep_dir, names[0])
            shutil.move(os.path.join(tpl_dir, names[0]), template_path)

    if os.path.isdir(session_dir):
        shutil.rmtree(session_dir, ignore_errors=True)
    shutil.rmtree(extract_dir, ignore_errors=True)

    manifest = core.load_manifest(report_root)
    if not manifest or not manifest.get("entries"):
        raise ValueError("The review inside that file has no screenshots in it.")

    missing = _missing_raw(manifest)
    if missing:
        raise MissingScreenshotsError(missing)

    manifest["template_path"] = template_path

    return {
        "work_dir": work_dir,
        "report_root": report_root,
        "manifest": manifest,
        "report_type": report_type,
        "core": core,
        "template_path": template_path,
        "extra_csv_results": _restore_extra_csv(extras.get("extra_csv_results")),
        "reviewed_indices": [i for i in extras.get("reviewed_indices", [])
                             if 0 <= i < len(manifest["entries"])],
    }


def load_existing_review(core, month_folder_path, report_type):
    """For the Drive app: the month folder it just downloaded already
    holds '<Type> Report/review_manifest.json' from an earlier build
    (the whole report folder gets uploaded back to Drive), next to the
    original screenshots. Load that instead of re-detecting from
    scratch. Returns None if there's no earlier review in the folder."""
    report_root = os.path.join(month_folder_path, f"{report_type} Report")
    manifest_path = os.path.join(report_root, MANIFEST_NAME)
    if not os.path.exists(manifest_path):
        return None
    _normalize_manifest_paths(manifest_path)
    manifest = core.load_manifest(report_root)
    if not manifest or not manifest.get("entries"):
        return None
    missing = _missing_raw(manifest)
    if missing:
        raise MissingScreenshotsError(missing)
    return {"report_root": report_root, "manifest": manifest}


def find_new_screenshots(core, month_folder_path, report_type, manifest):
    """Screenshots sitting in the month folder that the saved review does
    NOT contain. With screenshots being added week after week, a review
    re-opened from an earlier build would otherwise silently leave the
    newer ones out of the next report. Uses the scanner's own rules
    (extension, 'no results', marked files, keyword match, recognised
    day folder) so it reports exactly what a fresh scan would have picked
    up — no false alarms about files the scanner ignores anyway."""
    in_review = {os.path.normcase(os.path.abspath(e["raw_filepath"])) for e in manifest["entries"]}
    type_dir = None
    for d in core.list_dirs(month_folder_path):
        if d.lower() == report_type.lower():
            type_dir = os.path.join(month_folder_path, d)
            break
    if not type_dir:
        return []
    new = []
    for day in core.list_dirs(type_dir):
        if core.extract_day_number(day) is None:
            continue
        day_path = os.path.join(type_dir, day)
        for fname in core.list_files(day_path):
            stem, ext = os.path.splitext(fname)
            if ext.lower() not in core.IMAGE_EXTS:
                continue
            if core.is_no_results(stem) or core.is_marked(stem) or core.match_keyword(stem) is None:
                continue
            if os.path.normcase(os.path.abspath(os.path.join(day_path, fname))) not in in_review:
                new.append(os.path.join(day, fname))
    return sorted(new)
