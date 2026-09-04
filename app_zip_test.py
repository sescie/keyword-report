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
import os
import tempfile
import zipfile

import streamlit as st
from PIL import Image

import monthly_seo_report_google as google_core
import monthly_seo_report_explorer as explorer_core

MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]

st.set_page_config(page_title="SEO Report Builder — ZIP test", layout="wide", page_icon="🧪")


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
        year = st.text_input("Year", value="2026")
        month_name = st.selectbox("Month", MONTH_NAMES, index=6)
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

    report_root, manifest, extra_csv_results = core.scan_and_detect(
        work_dir, year, month_name, month_num, template_path=template_path)

    return work_dir, report_root, manifest, core, template_path, extra_csv_results


def _run_detection(report_type, year, month_name, zip_file, template_file):
    template_bytes = template_file.getbuffer() if template_file is not None else None
    template_filename = template_file.name if template_file is not None else None
    work_dir, report_root, manifest, core, template_path, extra_csv_results = run_detection_core(
        report_type, year, month_name, zip_file.getbuffer(), template_bytes, template_filename)

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

def _renumber(entry):
    entry["row_coords"].sort(key=lambda r: r["top"])
    for i, rc in enumerate(entry["row_coords"], start=1):
        rc["row"] = i


def _commit(entry):
    entry["highlighted"] = sorted(rc["row"] for rc in entry["row_coords"] if rc["flagged"])
    entry["total_rows"] = len(entry["row_coords"])


def _ensure_flagged_field(entry):
    flagged_nums = set(entry.get("highlighted", []))
    for rc in entry["row_coords"]:
        rc.setdefault("flagged", rc["row"] in flagged_nums)


def _row_thumb(entry, rc, max_width=780):
    raw = Image.open(entry["raw_filepath"]).convert("RGB")
    crop = raw.crop((0, rc["top"], entry["x_right"], rc["bottom"]))
    if crop.width > max_width:
        scale = max_width / crop.width
        crop = crop.resize((max_width, max(1, int(crop.height * scale))))
    return crop


def render_review_step():
    manifest = st.session_state["manifest"]
    st.title("Review")

    total_rows = sum(len(e["row_coords"]) for e in manifest["entries"])
    total_flagged = sum(len(e.get("highlighted", [])) for e in manifest["entries"])
    c1, c2, c3 = st.columns(3)
    c1.metric("Screenshots", len(manifest["entries"]))
    c2.metric("Rows detected", total_rows)
    c3.metric("Rows flagged", total_flagged)
    st.caption("Flags shown here were parsed straight from each filename's `no …` list — "
               "correct anything that needs it below before building.")
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

    for idx, entry in enumerate(manifest["entries"]):
        _ensure_flagged_field(entry)
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

        flagged_n = sum(1 for r in entry["row_coords"] if r["flagged"])
        status_icon = "🔴" if flagged_n else "⚪"
        with st.expander(
            f"{status_icon} **{entry['keyword']}** · Page {entry['page']} · Week {entry['week_num']} · "
            f"{entry['date']} · {entry['file']} — {len(entry['row_coords'])} row(s), {flagged_n} flagged",
            expanded=False, key=f"expander_{idx}",
        ):
            for rc in sorted(entry["row_coords"], key=lambda r: r["top"]):
                c1, c2, c3 = st.columns([0.14, 0.1, 0.76])
                with c1:
                    new_val = st.checkbox(f"🚩 Flag #{rc['row']}", value=rc["flagged"],
                                           key=f"flag_{idx}_{rc['row']}_{id(rc)}")
                    if new_val != rc["flagged"]:
                        rc["flagged"] = new_val
                        _commit(entry)
                with c2:
                    if st.button("✕ Delete", key=f"del_{idx}_{rc['row']}_{id(rc)}"):
                        entry["row_coords"] = [r for r in entry["row_coords"] if r is not rc]
                        _renumber(entry)
                        _commit(entry)
                        st.rerun()
                with c3:
                    st.image(_row_thumb(entry, rc), width="stretch")

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
        with st.spinner("Building final images, CSV, and PPTX…"):
            total_images, total_highlighted, csv_path, pptx_path = core.build_report_from_manifest(
                report_root, manifest, manifest.get("template_path"),
                extra_csv_results=st.session_state.get("extra_csv_results"))
            st.session_state["build_done"] = True
            st.session_state["build_summary"] = (total_images, total_highlighted, csv_path, pptx_path)

    total_images, total_highlighted, csv_path, pptx_path = st.session_state["build_summary"]
    st.success(f"✅ {total_images} image(s) built, {total_highlighted} row(s) flagged.")

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(report_root):
            for fname in files:
                full = os.path.join(root, fname)
                zf.write(full, os.path.relpath(full, report_root))
    zip_buf.seek(0)

    st.download_button("⬇️ Download full report (ZIP)", data=zip_buf,
                        file_name="report.zip", mime="application/zip", type="primary")

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
