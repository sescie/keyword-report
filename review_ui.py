"""
review_ui.py — the per-screenshot review panel, shared by both apps.

Why this is structured the way it is (all of it is about never making the
person wait after a click):

* Each screenshot is a Streamlit FRAGMENT. A click inside one only re-runs
  that one screenshot, not the filters, the metrics, the progress saver and
  every other screenshot's panel. (Streamlit runs the code inside
  collapsed expanders too, so before this, every click re-executed every
  checkbox, thumbnail and adjust-box of EVERY screenshot on the page.)

* Flag / delete / undo / bulk-flag / apply-height are widget CALLBACKS.
  Callbacks run before the rerun, so the state is already correct when the
  panel draws and no second `st.rerun()` is needed. (Delete used to call
  st.rerun() after handling the click — the whole page ran twice.)
  Several quick clicks are applied in order; none are dropped.

* Rows carry a stable id ("uid"). Row NUMBERS change when a row is deleted
  or resized (rows are renumbered top-to-bottom), so widget keys must not
  be built from them — otherwise a flag could end up attached to a
  different row after a delete.

* Thumbnails are cached by (file, top, bottom). Deleting a row renumbers
  the others but doesn't change any crop, so nothing is re-cropped.
  The decoded screenshot is held with cache_resource (shared, no copy);
  cache_data hands back a full COPY of the image on every call.
"""

import io

import streamlit as st
from PIL import Image


# ════════════════════════════════════════════════════════════════════════
# Data helpers
# ════════════════════════════════════════════════════════════════════════

def renumber(entry):
    """Reorder rows top-to-bottom and number them 1..N."""
    entry["row_coords"].sort(key=lambda r: r["top"])
    for i, rc in enumerate(entry["row_coords"], start=1):
        rc["row"] = i


def commit(entry):
    entry["highlighted"] = sorted(rc["row"] for rc in entry["row_coords"] if rc["flagged"])
    entry["total_rows"] = len(entry["row_coords"])


def ensure_row_fields(entry):
    """Every row needs a 'flagged' flag and a stable 'uid'. (A build replaces
    row_coords with fresh dicts that have neither, so this runs on every
    render and is a no-op once they exist.)"""
    flagged_nums = set(entry.get("highlighted", []))
    seq = int(entry.get("_uid_seq", 0))
    for rc in entry["row_coords"]:
        rc.setdefault("flagged", rc["row"] in flagged_nums)
        if "uid" not in rc:
            seq += 1
            rc["uid"] = seq
    seq = max([seq] + [int(rc["uid"]) for rc in entry["row_coords"]])
    entry["_uid_seq"] = seq


# ════════════════════════════════════════════════════════════════════════
# Cached image work
# ════════════════════════════════════════════════════════════════════════

@st.cache_resource(max_entries=8, show_spinner=False)
def load_raw_image(raw_filepath: str):
    """Decoded screenshot, shared across reruns WITHOUT copying (cache_data
    would pickle a full copy per call — hundreds of MB over a batch of
    thumbnails). Callers only ever crop it, which returns a new image.
    Capped at 8 so a batch of hundreds can't grow the cache unbounded."""
    return Image.open(raw_filepath).convert("RGB")


@st.cache_data(max_entries=2000, show_spinner=False)
def image_height(raw_filepath: str) -> int:
    with Image.open(raw_filepath) as im:      # reads the header only
        return int(im.height)


@st.cache_data(max_entries=600, show_spinner=False)
def thumb_png(raw_filepath: str, top: int, bottom: int, x_right: int, max_width: int = 780) -> bytes:
    raw = load_raw_image(raw_filepath)
    crop = raw.crop((0, int(top), int(x_right), int(bottom)))
    if crop.width > max_width:
        scale = max_width / crop.width
        crop = crop.resize((max_width, max(1, int(crop.height * scale))))
    buf = io.BytesIO()
    crop.save(buf, "PNG", compress_level=1)       # lossless: these are screenshots of text
    return buf.getvalue()


# ════════════════════════════════════════════════════════════════════════
# Callbacks (run BEFORE the rerun, so the panel always draws final state)
# ════════════════════════════════════════════════════════════════════════

def _entry(idx):
    return st.session_state["manifest"]["entries"][idx]


def _row(entry, uid):
    return next((r for r in entry["row_coords"] if r["uid"] == uid), None)


def _fk(idx, uid):
    return f"flag_{idx}_{uid}"


def _tk(idx, uid):
    return f"adjtop_{idx}_{uid}"


def _bk(idx, uid):
    return f"adjbot_{idx}_{uid}"


def _on_reviewed(idx):
    # Kept on the screenshot's own data, not just in the checkbox: Streamlit
    # throws away widget state for widgets that aren't drawn, so a tick would
    # vanish the moment a filter hid that screenshot.
    _entry(idx)["reviewed"] = bool(st.session_state.get(f"reviewed_{idx}", False))


def _on_flag(idx, uid):
    entry = _entry(idx)
    rc = _row(entry, uid)
    if rc is None:
        return
    rc["flagged"] = bool(st.session_state.get(_fk(idx, uid), rc["flagged"]))
    commit(entry)


def _on_flag_all(idx, value):
    entry = _entry(idx)
    for rc in entry["row_coords"]:
        rc["flagged"] = bool(value)
        # a checkbox that has rendered owns its value in session_state and
        # ignores anything else, so it has to be set here too
        st.session_state[_fk(idx, rc["uid"])] = bool(value)
    commit(entry)


def _on_delete(idx, uid):
    entry = _entry(idx)
    rc = _row(entry, uid)
    if rc is None:
        return
    st.session_state.setdefault(f"undo_{idx}", []).append(dict(rc))
    entry["row_coords"] = [r for r in entry["row_coords"] if r["uid"] != uid]
    renumber(entry)
    commit(entry)
    for k in (_fk(idx, uid), _tk(idx, uid), _bk(idx, uid), f"adjerr_{idx}_{uid}"):
        st.session_state.pop(k, None)
    if st.session_state.get(f"adjopen_{idx}") == uid:
        st.session_state[f"adjopen_{idx}"] = None


def _on_undo(idx):
    stack = st.session_state.get(f"undo_{idx}") or []
    if not stack:
        return
    entry = _entry(idx)
    rc = stack.pop()
    if _row(entry, rc["uid"]) is not None:       # can't happen, but never duplicate a row
        return
    entry["row_coords"].append(rc)
    renumber(entry)
    commit(entry)
    st.session_state[_fk(idx, rc["uid"])] = bool(rc["flagged"])
    st.session_state[_tk(idx, rc["uid"])] = int(rc["top"])
    st.session_state[_bk(idx, rc["uid"])] = int(rc["bottom"])


def _on_adjust_toggle(idx, uid):
    """Open the height controls for this row (closing any other row's), or
    close them if they're already open. Opening resets the numbers to the
    row's current values, so stale typed values never carry over."""
    k = f"adjopen_{idx}"
    if st.session_state.get(k) == uid:
        st.session_state[k] = None
        return
    entry = _entry(idx)
    rc = _row(entry, uid)
    if rc is None:
        return
    st.session_state[k] = uid
    st.session_state[_tk(idx, uid)] = int(rc["top"])
    st.session_state[_bk(idx, uid)] = int(rc["bottom"])
    st.session_state.pop(f"adjerr_{idx}_{uid}", None)


def _on_apply(idx, uid):
    entry = _entry(idx)
    rc = _row(entry, uid)
    if rc is None:
        return
    top = int(st.session_state.get(_tk(idx, uid), rc["top"]))
    bottom = int(st.session_state.get(_bk(idx, uid), rc["bottom"]))
    if bottom <= top:
        st.session_state[f"adjerr_{idx}_{uid}"] = "Bottom must be greater than Top."
        return
    st.session_state.pop(f"adjerr_{idx}_{uid}", None)
    rc["top"], rc["bottom"] = top, bottom
    renumber(entry)           # the row may now sit above/below its neighbours
    commit(entry)
    st.session_state[f"adjopen_{idx}"] = None


# ════════════════════════════════════════════════════════════════════════
# The panel
# ════════════════════════════════════════════════════════════════════════

def _render_panel(idx, entry):
    stack = st.session_state.get(f"undo_{idx}") or []
    a, b, c, d = st.columns([2.2, 1.3, 1.3, 1.4])
    with a:
        if f"reviewed_{idx}" not in st.session_state:
            st.session_state[f"reviewed_{idx}"] = bool(entry.get("reviewed"))
        st.checkbox("✅ Mark this screenshot as reviewed", key=f"reviewed_{idx}",
                    on_change=_on_reviewed, args=(idx,))
    with b:
        st.button("🚩 Flag all rows", key=f"flagall_{idx}", width="stretch",
                  on_click=_on_flag_all, args=(idx, True))
    with c:
        st.button("Clear all flags", key=f"unflagall_{idx}", width="stretch",
                  on_click=_on_flag_all, args=(idx, False))
    with d:
        st.button(f"↩ Undo delete ({len(stack)})", key=f"undo_btn_{idx}", width="stretch",
                  on_click=_on_undo, args=(idx,), disabled=not stack)
    st.divider()

    raw_path = entry["raw_filepath"]
    try:
        raw_h = image_height(raw_path)
    except Exception as exc:
        st.error(f"Couldn't open the screenshot file ({type(exc).__name__}: {exc}).")
        return

    for rc in sorted(entry["row_coords"], key=lambda r: r["top"]):
        uid = rc["uid"]
        fk, tk, bk = _fk(idx, uid), _tk(idx, uid), _bk(idx, uid)
        if fk not in st.session_state:
            st.session_state[fk] = bool(rc["flagged"])

        c1, c2, c3 = st.columns([0.14, 0.1, 0.76])
        with c1:
            st.checkbox(f"🚩 Flag #{rc['row']}", key=fk, on_change=_on_flag, args=(idx, uid))
        with c2:
            st.button("✕ Delete", key=f"del_{idx}_{uid}", on_click=_on_delete, args=(idx, uid))
            is_open = st.session_state.get(f"adjopen_{idx}") == uid
            st.button("▲ Close" if is_open else "↕️ Adjust", key=f"adj_{idx}_{uid}",
                      on_click=_on_adjust_toggle, args=(idx, uid))
        with c3:
            try:
                st.image(thumb_png(raw_path, rc["top"], rc["bottom"], entry["x_right"]), width="content")
            except Exception as exc:
                st.warning(f"Couldn't draw this row's preview ({type(exc).__name__}).")

        # Built only for the row being adjusted. Every row used to carry its
        # own (collapsed) expander with two number inputs and a button —
        # hundreds of widgets built on every render for something opened
        # rarely.
        if is_open:
            if tk not in st.session_state:
                st.session_state[tk] = int(rc["top"])
            if bk not in st.session_state:
                st.session_state[bk] = int(rc["bottom"])
            with st.container(border=True):
                st.caption(f"Adjust the height / position of row #{rc['row']}")
                ac1, ac2, ac3 = st.columns([1, 1, 1])
                with ac1:
                    adj_top = st.number_input("Top", min_value=0, max_value=raw_h, key=tk)
                with ac2:
                    adj_bottom = st.number_input("Bottom", min_value=0, max_value=raw_h, key=bk)
                with ac3:
                    st.write("")
                    st.write("")
                    st.button("Apply", key=f"apply_{idx}_{uid}", type="primary",
                              on_click=_on_apply, args=(idx, uid))
                err = st.session_state.get(f"adjerr_{idx}_{uid}")
                if err:
                    st.error(err)
                # The preview only shows while the numbers differ from the saved
                # row — otherwise it is a second copy of the thumbnail above it.
                pending = (int(adj_top) != int(rc["top"]) or int(adj_bottom) != int(rc["bottom"]))
                if pending and adj_bottom > adj_top:
                    st.image(thumb_png(raw_path, int(adj_top), int(adj_bottom), entry["x_right"]),
                             caption="Preview at these positions (click Apply to keep it)", width="stretch")


@st.fragment
def render_entry_block(idx):
    """One screenshot: its expander header plus the panel inside it. Being a
    fragment, a click in here re-runs only this block — so the header's row
    and flag counts stay live, without touching anything else on the page."""
    manifest = st.session_state.get("manifest")
    if not manifest or idx >= len(manifest["entries"]):
        return
    entry = manifest["entries"][idx]
    ensure_row_fields(entry)
    flagged_n = sum(1 for r in entry["row_coords"] if r["flagged"])
    status_icon = "🔴" if flagged_n else "⚪"
    reviewed_icon = "✅ " if entry.get("reviewed") else ""
    with st.expander(
        f"{reviewed_icon}{status_icon} **{entry['keyword']}** · Page {entry['page']} · "
        f"Week {entry['week_num']} · {entry['date']} · {entry['file']} — "
        f"{len(entry['row_coords'])} row(s), {flagged_n} flagged",
        expanded=False, key=f"expander_{idx}",
    ):
        _render_panel(idx, entry)


@st.fragment(run_every="2s")
def render_summary(visible_indices):
    """Totals at the top. Edits happen inside per-screenshot fragments, which
    don't re-run this, so it refreshes itself every couple of seconds
    instead of going stale until the next full-page rerun."""
    manifest = st.session_state.get("manifest")
    if not manifest:
        return
    total_rows = sum(len(e["row_coords"]) for e in manifest["entries"])
    total_flagged = sum(len(e.get("highlighted", [])) for e in manifest["entries"])
    c1, c2, c3 = st.columns(3)
    c1.metric("Screenshots", len(manifest["entries"]))
    c2.metric("Rows detected", total_rows)
    c3.metric("Rows flagged", total_flagged)
    reviewed_n = sum(1 for i in visible_indices if manifest["entries"][i].get("reviewed"))
    st.caption(f"✅ Marked reviewed: {reviewed_n} of {len(visible_indices)} shown")
