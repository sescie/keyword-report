"""
app_zip_test.py — SEO Monthly Report Builder, ZIP-upload test version.

Same detect -> review -> build pipeline as app.py, but with NO Google
sign-in or Drive dependency at all — upload a ZIP, review in-browser,
download the finished report as a ZIP. Meant to be deployed as its own
separate Streamlit app first, specifically to confirm the core pipeline
(OCR, row detection, filename-based flagging, the review UI, PPTX
building) runs cleanly on Streamlit Cloud's servers before adding any
Google OAuth/Drive configuration on top.

Deploy this as a SEPARATE Streamlit app from the same GitHub repo (on
share.streamlit.io: "New app" -> same repo -> main file path:
app_zip_test.py) — it needs no Secrets configured at all to run.

Run locally with:
    streamlit run app_zip_test.py
"""

import io
import datetime
import os
import tempfile
import traceback
import zipfile

import streamlit as st
from PIL import Image

import monthly_seo_report_google as google_core
import monthly_seo_report_explorer as explorer_core
import review_session
import review_ui

CORES = {"Google": google_core, "Explorer": explorer_core}

# Streamlit Cloud's actual allocated memory is much smaller than what
# os.cpu_count() reports from the underlying host — the core scripts'
# own default (up to 8 parallel OCR/decode workers) assumes a
# reasonably resourced desktop machine. Running that many full
# screenshots decoded in memory AT ONCE is exactly the kind of thing
# that exhausts a constrained container's memory during a big batch.
# Capped lower here specifically for the web app; the desktop script's
# own default is untouched for people running it locally with more
# headroom to spare.
google_core.IMAGE_WORKERS = 2
explorer_core.IMAGE_WORKERS = 2

MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]

# Bump this string every time these files are handed over — shown at the
# bottom of the sidebar so it's possible to tell AT A GLANCE, just by
# looking at the running app, whether it's actually running the latest
# code or an older cached/undeployed version. No more guessing.
APP_BUILD = "2026-10-14-fast-review"

st.set_page_config(page_title="SEO Report Builder — ZIP test", layout="wide", page_icon="🧪")
st.sidebar.caption(f"Build: {APP_BUILD}")


# ════════════════════════════════════════════════════════════════════════
# Session state
# ════════════════════════════════════════════════════════════════════════

def _init_state():
    defaults = {
        "step": "upload", "work_dir": None, "report_root": None, "manifest": None,
        "core": None, "template_path": None, "extra_csv_results": None,
        "editing_entry_idx": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


_init_state()


# ════════════════════════════════════════════════════════════════════════
# Step 1: Upload
# ════════════════════════════════════════════════════════════════════════

def render_upload_step():
    st.title("🧪 SEO Report Builder — ZIP test")
    st.info("This is the no-Google-needed test version — upload a ZIP, review, download. "
             "Once this runs cleanly end-to-end, the real Drive-connected app is ready to configure.")

    col1, col2 = st.columns(2)
    with col1:
        report_type = st.radio("Report type", ["Google", "Explorer"], horizontal=True)
        # Defaults to the PREVIOUS calendar month, not a hardcoded
        # "July 2026" — reports are always for a month that's already
        # finished, so whatever "today" happens to be, last month is
        # the sensible starting guess rather than a stale leftover.
        _today = datetime.date.today()
        _default_month_date = (_today.replace(day=1) - datetime.timedelta(days=1))
        year = st.text_input("Year", value=str(_default_month_date.year))
        month_name = st.selectbox("Month", MONTH_NAMES, index=_default_month_date.month - 1)
    with col2:
        st.markdown(
            f"**ZIP should contain the report-type folder directly**, e.g.:\n\n"
            f"- `Google/17/Ashley Rocharam pg 5 no 2-4,8.jpg`\n"
            f"- `Explorer/3/Sudesh Rocharam no results.jpg` (skipped automatically)\n\n"
            "The numbers after **no** in the filename are exactly which rows get flagged."
        )
        zip_file = st.file_uploader("Upload month ZIP", type=["zip"])
        template_file = st.file_uploader("Upload a PPTX template (optional)", type=["pptx", "potx"])

    if st.button("🔎 Scan & Detect", type="primary", disabled=zip_file is None):
        with st.spinner("Extracting and detecting rows — this can take a while for a lot of screenshots…"):
            _run_detection(report_type, year, month_name, zip_file, template_file)

    st.divider()
    with st.expander("♻️ Re-open a review I downloaded earlier (also how to recover after a refresh)"):
        st.caption(
            "Upload the report ZIP you downloaded from this app — or a **progress file** saved "
            "mid-review — and you're back on the review screen with every flag, height change, "
            "added/deleted row and 'reviewed' tick exactly as you left them. Nothing is re-scanned.")
        resume_zip = st.file_uploader("Downloaded report or progress ZIP", type=["zip"], key="resume_zip")
        resume_orig = st.file_uploader(
            "Original month ZIP — only needed if the report was downloaded BEFORE this feature "
            "existed (older downloads don't contain the screenshots)",
            type=["zip"], key="resume_orig")
        if st.button("♻️ Re-open review", disabled=resume_zip is None):
            _run_resume(resume_zip, resume_orig)


def _run_resume(resume_zip, resume_orig):
    try:
        with st.spinner("Restoring your review…"):
            res = review_session.restore_bundle(
                resume_zip.getbuffer(), CORES,
                resume_orig.getbuffer() if resume_orig is not None else None)
    except review_session.MissingScreenshotsError as e:
        st.error(
            f"{e}. This file doesn't contain the original screenshots, so it can't be re-opened "
            "on its own. Either upload the original month ZIP in the second box above, or — if you "
            "still have the month ZIP — run Scan & Detect once and download again from this version "
            "(that download will be re-openable).")
        return
    except Exception as e:  # corrupt/foreign file — say so plainly instead of a stack trace
        st.error(f"Couldn't re-open that file: {e}")
        return

    st.session_state.update({
        "work_dir": res["work_dir"], "report_root": res["report_root"],
        "manifest": res["manifest"], "core": res["core"],
        "template_path": res["template_path"], "extra_csv_results": res["extra_csv_results"],
        "step": "review",
    })
    for key in ("build_done", "build_summary", "bundle_bytes", "bundle_sources", "progress_bytes"):
        st.session_state.pop(key, None)
    for idx in res["reviewed_indices"]:
        st.session_state[f"reviewed_{idx}"] = True
        res["manifest"]["entries"][idx]["reviewed"] = True
    st.rerun()


def run_detection_core(report_type, year, month_name, zip_bytes, template_bytes=None, template_filename=None):
    """Pure logic, no Streamlit dependency — kept separate so it's
    directly testable with real data rather than relying on Streamlit's
    own (limited) file-upload test simulation."""
    work_dir = tempfile.mkdtemp(prefix="seo_report_zip_")
    month_num = MONTH_NAMES.index(month_name) + 1

    # scan_and_detect expects <drive_root>/Sudesh <year>/<month>/<Google|Explorer>/<day>/
    # — the ZIP's contents (the report-type folder directly) get
    # extracted right into that exact shape.
    month_folder_path = os.path.join(work_dir, f"Sudesh {year}", month_name)
    os.makedirs(month_folder_path, exist_ok=True)
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        zf.extractall(month_folder_path)

    template_path = None
    if template_bytes is not None:
        template_path = os.path.join(work_dir, template_filename)
        with open(template_path, "wb") as f:
            f.write(template_bytes)

    core = google_core if report_type == "Google" else explorer_core
    core.prompt_choice = lambda q, o: o[0]
    if template_path:
        # .potx / .ppsx etc. are converted to something python-pptx can open;
        # anything that isn't a PowerPoint file fails HERE with a clear message
        template_path = core.normalize_template(template_path)

    report_root, manifest, extra_csv_results = core.scan_and_detect(
        work_dir, year, month_name, month_num, template_path=template_path)

    return work_dir, report_root, manifest, core, template_path, extra_csv_results


def _run_detection(report_type, year, month_name, zip_file, template_file):
    template_bytes = template_file.getbuffer() if template_file is not None else None
    template_filename = template_file.name if template_file is not None else None
    try:
        work_dir, report_root, manifest, core, template_path, extra_csv_results = run_detection_core(
            report_type, year, month_name, zip_file.getbuffer(), template_bytes, template_filename)
    except Exception as exc:
        # Shown in full: an uncaught exception on Streamlit Cloud is redacted to
        # "the original error message is redacted", which tells nobody anything.
        st.error(f"Couldn't process the upload: {exc}")
        with st.expander("Technical details"):
            st.code(traceback.format_exc())
        return

    if not manifest["entries"]:
        st.warning("No screenshots were found. Double-check the ZIP contains the "
                   f"'{report_type}' folder directly, with day-number subfolders inside it.")
        return

    st.session_state.update({
        "work_dir": work_dir, "report_root": report_root, "manifest": manifest,
        "core": core, "template_path": template_path, "extra_csv_results": extra_csv_results,
        "step": "review",
    })
    st.rerun()


# ════════════════════════════════════════════════════════════════════════
# Step 2: Review
# ════════════════════════════════════════════════════════════════════════

def _report_type_label(core):
    return "Google" if core is google_core else "Explorer"


def _reviewed_indices(n_entries):
    entries = st.session_state["manifest"]["entries"]
    return [i for i in range(n_entries) if entries[i].get("reviewed")]


def _state_signature(manifest):
    """Cheap fingerprint of everything worth saving — used to tell whether a
    previously prepared progress file still matches what's on screen."""
    import hashlib
    parts = []
    for i, e in enumerate(manifest["entries"]):
        parts.append((e["file"], bool(e.get("reviewed")),
                      tuple((r["row"], r["top"], r["bottom"], bool(r.get("flagged")))
                            for r in e["row_coords"])))
    return hashlib.md5(repr(parts).encode()).hexdigest()


@st.fragment(run_every="2s")
def render_progress_saver():
    """Review state lives only in the browser session, so a refresh wipes it.
    This produces a ZIP that can be re-opened from the upload page. Built on
    demand (not on every rerun) because it packs every original screenshot."""
    core = st.session_state["core"]
    manifest = st.session_state["manifest"]
    with st.expander("💾 Save my progress — protects your edits against a refresh or closed tab"):
        st.caption(
            "Click **Prepare progress file**, download it, and keep it. If the page refreshes, "
            "upload it under *Re-open a review* on the first screen and carry on where you "
            "stopped. Prepare it again after making more edits.")
        if st.button("📦 Prepare progress file", key="prep_progress"):
            with st.spinner("Packing your review (includes the original screenshots)…"):
                st.session_state["progress_bytes"] = review_session.build_bundle_zip(
                    core, st.session_state["report_root"], manifest, _report_type_label(core),
                    st.session_state.get("template_path"), st.session_state.get("extra_csv_results"),
                    _reviewed_indices(len(manifest["entries"])), include_sources=True)
                st.session_state["progress_sig"] = _state_signature(manifest)
        if "progress_bytes" in st.session_state:
            if st.session_state.get("progress_sig") == _state_signature(manifest):
                name = (f"{_report_type_label(core)}_SEO_Review_IN_PROGRESS_"
                        f"{manifest['month_name']}_{manifest['year']}.zip")
                size_mb = len(st.session_state["progress_bytes"]) / 1_000_000
                st.download_button(f"⬇️ Download progress file ({size_mb:.1f} MB)",
                                   data=st.session_state["progress_bytes"], file_name=name,
                                   mime="application/zip", key="dl_progress")
            else:
                st.warning("You've made changes since this file was prepared — click "
                           "**Prepare progress file** again so it includes them.")


def render_review_step():
    manifest = st.session_state["manifest"]
    st.title("Review")

    # Reserved here so it sits at the top, filled once the filters have been
    # applied (it needs to know how many screenshots are showing). It
    # refreshes itself, because edits only re-run their own screenshot.
    summary_slot = st.container()
    st.caption("Flags shown here were parsed straight from each filename's `no …` list — "
               "correct anything that needs it below before building.")
    render_progress_saver()
    st.divider()

    filter_col1, filter_col2, filter_col3, filter_col4 = st.columns(4)
    keywords = sorted({e["keyword"] for e in manifest["entries"]})
    weeks = sorted({e["week_num"] for e in manifest["entries"]})
    pages = sorted({e["page"] for e in manifest["entries"]})
    with filter_col1:
        kw_filter = st.selectbox("Keyword", ["All"] + keywords)
    with filter_col2:
        week_filter = st.selectbox("Week", ["All"] + [f"Week {w}" for w in weeks])
    with filter_col3:
        page_filter = st.selectbox("Page", ["All"] + [f"Page {p}" for p in pages])
    with filter_col4:
        show_filter = st.selectbox("Show", ["All screenshots", "Only flagged", "Only unflagged"])

    visible_indices = []
    for idx, entry in enumerate(manifest["entries"]):
        review_ui.ensure_row_fields(entry)
        if kw_filter != "All" and entry["keyword"] != kw_filter:
            continue
        if week_filter != "All" and f"Week {entry['week_num']}" != week_filter:
            continue
        if page_filter != "All" and f"Page {entry['page']}" != page_filter:
            continue
        has_flagged = len(entry["highlighted"]) > 0
        if show_filter == "Only flagged" and not has_flagged:
            continue
        if show_filter == "Only unflagged" and has_flagged:
            continue
        visible_indices.append(idx)

    with summary_slot:
        review_ui.render_summary(tuple(visible_indices))

    ec1, ec2, _ = st.columns([1, 1, 3])
    with ec1:
        if st.button("⬇️ Expand all", width="stretch"):
            for idx in visible_indices:
                st.session_state[f"expander_{idx}"] = True
            st.rerun()
    with ec2:
        if st.button("⬆️ Collapse all", width="stretch"):
            for idx in visible_indices:
                st.session_state[f"expander_{idx}"] = False
            st.rerun()

    # One fragment per screenshot: a click inside it re-runs only that block.
    for idx in visible_indices:
        review_ui.render_entry_block(idx)

    st.divider()
    if st.button("🚀 Build Report", type="primary"):
        # A fresh build must run EVERY time this is clicked, including
        # a second/third time after going back to fix something — the
        # old "only build once per session" check silently kept
        # re-showing the FIRST build's results even after real edits
        # (verified directly: deleting a flagged row updated the
        # manifest correctly, but the previously-cached build summary
        # still reported the old, pre-edit flagged count). Clearing
        # these here guarantees build_report_from_manifest actually
        # runs again against whatever the manifest looks like right now.
        st.session_state.pop("build_done", None)
        st.session_state.pop("build_summary", None)
        st.session_state.pop("progress_bytes", None)
        st.session_state["step"] = "build"
        st.rerun()


# ════════════════════════════════════════════════════════════════════════
# Step 3: Build + download
# ════════════════════════════════════════════════════════════════════════

def render_build_step():
    st.title("Build & Download")
    core = st.session_state["core"]
    manifest = st.session_state["manifest"]
    report_root = st.session_state["report_root"]

    if "build_done" not in st.session_state:
        try:
            with st.spinner("Building final images, CSV, and PPTX…"):
                total_images, total_highlighted, csv_path, pptx_path = core.build_report_from_manifest(
                    report_root, manifest, manifest.get("template_path"),
                    extra_csv_results=st.session_state.get("extra_csv_results"))
        except Exception as exc:
            st.error(f"The build failed: {type(exc).__name__}: {exc}")
            with st.expander("Technical details"):
                st.code(traceback.format_exc())
            if st.button("◀ Back to review"):
                st.session_state["step"] = "review"
                st.rerun()
            st.stop()
        st.session_state["build_done"] = True
        st.session_state["build_summary"] = (total_images, total_highlighted, csv_path, pptx_path)
        # A fresh build invalidates any previously packaged download.
        st.session_state.pop("bundle_bytes", None)
        st.session_state.pop("bundle_sources", None)

    total_images, total_highlighted, csv_path, pptx_path = st.session_state["build_summary"]
    st.success(f"✅ {total_images} image(s) built, {total_highlighted} row(s) flagged.")
    if manifest.get("_pptx_error"):
        st.warning(f"The CSV and images were built, but the PowerPoint couldn't be: "
                   f"{manifest['_pptx_error']}")

    include_sources = st.checkbox(
        "Include the original screenshots in the download (lets you re-open and re-review this "
        "report later, even after a refresh)", value=True, key="include_sources")

    # Packaged once per build (and again only if that checkbox changes) —
    # NOT on every rerun, since this now packs every original screenshot.
    if ("bundle_bytes" not in st.session_state
            or st.session_state.get("bundle_sources") != include_sources):
        with st.spinner("Packaging your download…"):
            st.session_state["bundle_bytes"] = review_session.build_bundle_zip(
                core, report_root, manifest, _report_type_label(core),
                manifest.get("template_path"), st.session_state.get("extra_csv_results"),
                _reviewed_indices(len(manifest["entries"])), include_sources=include_sources)
            st.session_state["bundle_sources"] = include_sources

    # Named after the report itself (type, month, year) rather than a
    # generic "report.zip" — matters once you're juggling downloads
    # from more than one month or more than one report type at a time.
    zip_filename = f"{_report_type_label(core)}_SEO_Report_{manifest['month_name']}_{manifest['year']}.zip"
    size_mb = len(st.session_state["bundle_bytes"]) / 1_000_000

    st.download_button(f"⬇️ Download full report (ZIP, {size_mb:.1f} MB)",
                        data=st.session_state["bundle_bytes"],
                        file_name=zip_filename, mime="application/zip", type="primary")
    if include_sources:
        st.caption("The CSV, PPTX and images are at the top level of the ZIP, same as always. "
                   "The `_review_session` folder inside holds your original screenshots and review "
                   "state — leave it in, and you can upload this exact file under *Re-open a review* "
                   "on the first screen to pick up where you left off.")
    else:
        st.caption("Report files only — this download can't be re-opened for review on its own "
                   "(you'd also need the original month ZIP).")

    col1, col2 = st.columns(2)
    with col1:
        if st.button("◀ Back to review (fix something else)"):
            # Going back doesn't need to clear build_done itself — the
            # "Build Report" button already clears it on the way back
            # IN, so whatever gets edited here is guaranteed to produce
            # a fresh build next time, not a stale re-show of this one.
            st.session_state["step"] = "review"
            st.rerun()
    with col2:
        if st.button("Start a new report"):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()


# ════════════════════════════════════════════════════════════════════════
# Router
# ════════════════════════════════════════════════════════════════════════

if st.session_state["step"] == "upload":
    render_upload_step()
elif st.session_state["step"] == "review":
    render_review_step()
elif st.session_state["step"] == "build":
    render_build_step()
