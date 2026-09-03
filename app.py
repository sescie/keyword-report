"""
app.py — SEO Monthly Report Builder, as a Streamlit web app.

Flow:
    1. Sign in with Google (session-only — nothing stored permanently).
    2. Paste the Drive link to the month's folder, pick report type,
       year, month, optional template.
    3. Scan & Detect — downloads that Drive folder locally, runs the
       exact same detection pipeline as the desktop tools.
    4. Review — flag/unflag rows, delete rows outright, or drag
       directly on a screenshot to redraw/add a row's boundary.
    5. Build Report — composites everything, then uploads the finished
       report straight back into that same Drive month folder.
"""

import os
import tempfile
from pathlib import Path

import streamlit as st
from PIL import Image

import drive_auth
import drive_utils
import monthly_seo_report_google as google_core
import monthly_seo_report_explorer as explorer_core

try:
    from streamlit_drawable_canvas import st_canvas
    HAS_CANVAS = True
except ImportError:
    HAS_CANVAS = False

MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]

st.set_page_config(page_title="SEO Monthly Report Builder", layout="wide", page_icon="📊")


# ════════════════════════════════════════════════════════════════════════
# Session state
# ════════════════════════════════════════════════════════════════════════

def _init_state():
    defaults = {
        "step": "setup", "work_dir": None, "report_root": None, "manifest": None,
        "core": None, "template_path": None, "extra_csv_results": None,
        "editing_entry_idx": None, "drive_month_folder_id": None, "report_folder_name": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


_init_state()


# ════════════════════════════════════════════════════════════════════════
# Step 1: Setup — Drive link, report type, year/month, template
# ════════════════════════════════════════════════════════════════════════

def render_setup_step(service):
    st.title("📊 SEO Monthly Report Builder")
    st.write("Point this at the month's folder in Drive — it reads the screenshots straight "
             "from there, and writes the finished report right back into the same place.")

    col1, col2 = st.columns([1, 1])
    with col1:
        drive_link = st.text_input(
            "Google Drive folder link for the month",
            placeholder="https://drive.google.com/drive/folders/….",
            help="Open the month's folder in Drive and copy the link from your browser's address bar.",
        )
        report_type = st.radio("Report type", ["Google", "Explorer"], horizontal=True)
        c1, c2 = st.columns(2)
        with c1:
            year = st.text_input("Year", value="2026")
        with c2:
            month_name = st.selectbox("Month", MONTH_NAMES, index=6)
    with col2:
        st.markdown(
            "**Expected folder shape inside that Drive folder:**\n\n"
            "- `Google/<day>/Ashley Rocharam pg 5 no 2-4,8.jpg`\n"
            "- `Explorer/<day>/Ashley Rocharam pg 5 no 2-4,8.jpg`\n\n"
            "The numbers after **no** are exactly which rows get flagged — "
            "you can still add, remove, or redraw rows in the review step."
        )
        template_link = st.text_input(
            "PPTX template Drive link (optional)",
            placeholder="https://drive.google.com/file/d/….",
        )

    if st.button("🔎 Scan & Detect", type="primary", disabled=not drive_link):
        with st.spinner("Downloading screenshots from Drive and detecting rows — "
                        "this can take a while for a lot of screenshots…"):
            _run_detection(service, report_type, year, month_name, drive_link, template_link)


def _run_detection(service, report_type, year, month_name, drive_link, template_link):
    work_dir = tempfile.mkdtemp(prefix="seo_report_")
    month_num = MONTH_NAMES.index(month_name) + 1
    month_folder_id = drive_utils.extract_folder_id(drive_link)

    # scan_and_detect expects <drive_root>/Sudesh <year>/<month>/<Google|Explorer>/<day>/
    # — mirror the Drive folder directly into that exact local shape so
    # the already-tested core scanning function runs completely
    # unchanged against the downloaded copy.
    month_folder_path = os.path.join(work_dir, f"Sudesh {year}", month_name)
    n_downloaded = drive_utils.download_folder_tree(service, month_folder_id, month_folder_path)
    st.toast(f"Downloaded {n_downloaded} file(s) from Drive.")

    template_path = None
    if template_link:
        template_id = drive_utils.extract_folder_id(template_link)
        template_path = os.path.join(work_dir, "template.pptx")
        drive_utils.download_file(service, template_id, template_path)

    core = google_core if report_type == "Google" else explorer_core
    core.prompt_choice = lambda q, o: o[0]

    report_root, manifest, extra_csv_results = core.scan_and_detect(
        work_dir, year, month_name, month_num, template_path=template_path)

    st.session_state.update({
        "work_dir": work_dir, "report_root": report_root, "manifest": manifest,
        "core": core, "template_path": template_path, "extra_csv_results": extra_csv_results,
        "drive_month_folder_id": month_folder_id,
        "report_folder_name": f"{report_type} Report",
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

    if st.session_state["editing_entry_idx"] is not None:
        render_crop_editor(manifest["entries"][st.session_state["editing_entry_idx"]])
        return

    filter_col1, filter_col2 = st.columns(2)
    keywords = sorted({e["keyword"] for e in manifest["entries"]})
    with filter_col1:
        kw_filter = st.selectbox("Filter by keyword", ["All"] + keywords)
    with filter_col2:
        show_filter = st.selectbox("Show", ["All screenshots", "Only flagged", "Only unflagged"])

    for idx, entry in enumerate(manifest["entries"]):
        _ensure_flagged_field(entry)
        if kw_filter != "All" and entry["keyword"] != kw_filter:
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
            expanded=False,
        ):
            bcol1, bcol2 = st.columns([1, 3])
            with bcol1:
                if st.button("✏️ Edit rows visually", key=f"editbtn_{idx}", width="stretch"):
                    st.session_state["editing_entry_idx"] = idx
                    st.rerun()

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
    left, right = st.columns([3, 1])
    with left:
        st.caption(f"Writing back to the same Drive folder as: **{st.session_state['report_folder_name']}**")
    with right:
        if st.button("🚀 Build & Upload to Drive", type="primary", width="stretch"):
            st.session_state["step"] = "build"
            st.rerun()


def render_crop_editor(entry):
    st.subheader(f"Visual editor — {entry['file']}")
    st.caption("Draw a rectangle over an existing row to redraw its boundary. "
               "Draw in empty space to add a row detection missed.")

    if not HAS_CANVAS:
        st.warning("Visual dragging isn't available right now — you can still flag/delete "
                   "rows on the main review screen.")
        if st.button("◀ Back to review"):
            st.session_state["editing_entry_idx"] = None
            st.rerun()
        return

    raw = Image.open(entry["raw_filepath"]).convert("RGB")
    max_width = 1000
    scale = min(max_width / raw.width, 1.0)
    disp_w, disp_h = int(raw.width * scale), int(raw.height * scale)
    display_img = raw.resize((disp_w, disp_h))

    initial_rects = {
        "version": "4.4.0",
        "objects": [
            {
                "type": "rect", "left": 0, "top": rc["top"] * scale,
                "width": disp_w, "height": (rc["bottom"] - rc["top"]) * scale,
                "fill": "rgba(200,45,35,0.12)", "stroke": "#c82d23", "strokeWidth": 2,
            }
            for rc in entry["row_coords"]
        ],
    }

    canvas_result = st_canvas(
        fill_color="rgba(255,204,0,0.25)", stroke_width=2, stroke_color="#ffcc00",
        background_image=display_img, height=disp_h, width=disp_w,
        drawing_mode="rect", initial_drawing=initial_rects, key=f"canvas_{entry['file']}",
    )

    col1, col2 = st.columns(2)
    with col1:
        if st.button("💾 Save these rows", type="primary"):
            if canvas_result.json_data is not None:
                new_rows = []
                for obj in canvas_result.json_data["objects"]:
                    top = int(obj["top"] / scale)
                    bottom = int((obj["top"] + obj["height"] * obj.get("scaleY", 1)) / scale)
                    if bottom > top:
                        new_rows.append({"top": top, "bottom": bottom})
                new_rows.sort(key=lambda r: r["top"])
                # A structural edit like this can add/remove/reorder rows
                # arbitrarily, so a row's NUMBER after renumbering has no
                # reliable relationship to what it meant before — flagged
                # state is reset here rather than risk silently
                # misattributing a flag to the wrong row.
                entry["row_coords"] = [
                    {"row": i + 1, "top": r["top"], "bottom": r["bottom"], "flagged": False}
                    for i, r in enumerate(new_rows)
                ]
                entry["total_rows"] = len(entry["row_coords"])
                entry["highlighted"] = []
                st.session_state["editing_entry_idx"] = None
                st.rerun()
    with col2:
        if st.button("◀ Cancel"):
            st.session_state["editing_entry_idx"] = None
            st.rerun()


# ════════════════════════════════════════════════════════════════════════
# Step 3: Build + upload back to Drive
# ════════════════════════════════════════════════════════════════════════

def render_build_step(service):
    st.title("Build & Upload")
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

    if "upload_done" not in st.session_state:
        with st.spinner("Uploading the finished report back to Drive…"):
            n_uploaded = drive_utils.upload_folder_tree(
                service, report_root, st.session_state["drive_month_folder_id"])
            st.session_state["upload_done"] = True
            st.session_state["n_uploaded"] = n_uploaded

    st.success(f"☁️ {st.session_state['n_uploaded']} file(s) uploaded back to the "
               f"**{st.session_state['report_folder_name']}** folder in Drive — same month folder you started from.")

    if st.button("Start a new report"):
        for key in list(st.session_state.keys()):
            if key != "drive_credentials":
                del st.session_state[key]
        st.rerun()


# ════════════════════════════════════════════════════════════════════════
# Router
# ════════════════════════════════════════════════════════════════════════

credentials = drive_auth.get_credentials()
if credentials is None:
    drive_auth.render_sign_in()
else:
    drive_auth.render_sign_out_button()
    drive_service = drive_utils.get_drive_service(credentials)

    if st.session_state["step"] == "setup":
        render_setup_step(drive_service)
    elif st.session_state["step"] == "review":
        render_review_step()
    elif st.session_state["step"] == "build":
        render_build_step(drive_service)
