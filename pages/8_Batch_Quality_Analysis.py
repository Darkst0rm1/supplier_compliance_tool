"""Batch Quality Analysis — surface possible SAP batch-number and expiry-date
data-quality issues for human review.

Upload the SAP receiving export, run rule-based detection, review flagged issue
groups (with all related original records), optionally ask the AI assistant to
explain a single issue, record the human findings, and export the results. The
AI never suggests corrected values or edits SAP data.
"""
from __future__ import annotations

import io
import os

import pandas as pd
import streamlit as st

from src.batch_quality import ai_review as ai
from src.batch_quality.exporter import generate_excel
from src.batch_quality.issue_groups import analyze, build_related_records
from src.batch_quality.loader import (
    BatchQualityError,
    COL_BATCH,
    COL_BATCH_SLED,
    COL_MATERIAL,
    COL_MATERIAL_DESC,
    COL_PLANT,
    COL_PO,
    COL_PURCH_GROUP,
    COL_QTY,
    COL_RECEIVED,
    COL_SUPPLIER_NAME,
    COL_VENDOR_NAME,
    NORMALIZED_BATCH,
    load_batch_data,
)
from src.batch_quality import reviews as rv

st.title("Batch Quality Analysis")
st.caption(
    "Identify possible SAP batch-number and expiry-date data-quality issues. "
    "Rule-based detection runs first; the AI assistant only explains and "
    "organizes a selected issue — it never suggests corrected values or edits "
    "SAP data."
)

# Session-state keys are prefixed batch_quality_.
_AI_KEY = "batch_quality_ai_results"        # group_id -> AIReviewResult
_REVIEW_KEY = "batch_quality_reviews"       # group_id -> human review dict
st.session_state.setdefault(_AI_KEY, {})
st.session_state.setdefault(_REVIEW_KEY, {})


def _idx(options: list, value) -> int:
    """Index of ``value`` in ``options`` for pre-selecting a saved choice."""
    try:
        return options.index(value)
    except (ValueError, TypeError):
        return 0


@st.cache_data(show_spinner=False)
def _process(file_bytes: bytes):
    df = load_batch_data(io.BytesIO(file_bytes))
    result = analyze(df)
    related = build_related_records(result)
    return result, related


# ---------------------------------------------------------------------------
# Upload
# ---------------------------------------------------------------------------
uploaded = st.file_uploader("SAP receiving export (.xlsx)", type=["xlsx"], key="bq_upload")
if uploaded is None:
    st.info("Upload the SAP receiving export to begin.")
    st.stop()
if not st.button("Process File", type="primary"):
    st.stop()

with st.spinner("Loading export and running rule-based analysis..."):
    try:
        result, related = _process(uploaded.getvalue())
    except BatchQualityError as exc:
        st.error(str(exc))
        st.stop()
    except Exception as exc:  # noqa: BLE001
        st.error(f"Could not process the file: {exc}")
        st.stop()

df = result.df
flagged = result.flagged
s = result.summary


def _ai_review_df() -> pd.DataFrame:
    """AI results + saved human findings, one row per reviewed issue group."""
    rows = []
    ai_results = st.session_state[_AI_KEY]
    review_store = st.session_state[_REVIEW_KEY]
    ids = set(ai_results) | set(review_store)
    for gid in sorted(ids):
        grp = flagged[flagged["Issue Group ID"] == gid]
        base = grp.iloc[0].to_dict() if not grp.empty else {"Issue Group ID": gid}
        row = {
            "Issue Group ID": gid,
            "Risk Level": base.get("Risk Level", ""),
            "Issue Type": base.get("Issue Type", ""),
            "Material": base.get("Material", ""),
            "Material Description": base.get("Material Description", ""),
            "Supplier Name": base.get("Supplier Name", ""),
            "Original Batch Values": base.get("Original Batch Values", ""),
            "Expiry Dates": base.get("Expiry Dates", ""),
            "Purchase Orders": base.get("Purchase Orders", ""),
        }
        air = ai_results.get(gid)
        if air is not None:
            row.update({
                "AI Risk Level": air.risk_level,
                "AI Review Summary": air.review_summary,
                "AI Pattern Identified": air.pattern_identified,
                "AI Documents to Verify": "; ".join(air.documents_to_verify),
                "AI Possible Root Causes": "; ".join(air.possible_root_causes),
                "AI Recommended Review Steps": "; ".join(air.recommended_review_steps),
                "AI Review Note": air.review_note,
            })
        row.update(review_store.get(gid, {}))
        rows.append(row)
    return pd.DataFrame(rows)


tab_summary, tab_flagged, tab_history, tab_ai, tab_results = st.tabs([
    "Upload & Summary", "Flagged Issues", "Material Batch History",
    "AI-Assisted Review", "Review Results",
])

# ---------------------------------------------------------------------------
# 1. Summary
# ---------------------------------------------------------------------------
with tab_summary:
    st.subheader("Summary")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Receiving records", f"{s['total_records']:,}")
    c2.metric("Unique materials", f"{s['unique_materials']:,}")
    c3.metric("Unique batches", f"{s['unique_batches']:,}")
    c4.metric("Unique suppliers", f"{s['unique_suppliers']:,}")
    c5.metric("Unique POs", f"{s['unique_pos']:,}")

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Issue groups", f"{s['total_issue_groups']:,}")
    c2.metric("High risk", f"{s['high_risk']:,}")
    c3.metric("Medium risk", f"{s['medium_risk']:,}")
    c4.metric("Low risk", f"{s['low_risk']:,}")
    c5.metric("Materials: multi-batch & date", f"{s['multi_batch_materials']:,}")

    if not flagged.empty:
        st.markdown("**Issue groups by Issue Type / Risk Level**")
        col_a, col_b = st.columns(2)
        col_a.bar_chart(flagged["Issue Type"].value_counts())
        col_b.bar_chart(flagged["Risk Level"].value_counts())
        col_c, col_d = st.columns(2)
        if "Supplier Name" in flagged.columns:
            top_sup = flagged["Supplier Name"].replace("", "(none)").value_counts().head(15)
            col_c.markdown("By Supplier")
            col_c.bar_chart(top_sup)
        if "Plants" in flagged.columns:
            top_plant = flagged["Plants"].replace("", "(none)").value_counts().head(15)
            col_d.markdown("By Plant")
            col_d.bar_chart(top_plant)
    else:
        st.success("No rule-based issues were flagged in this export.")

# ---------------------------------------------------------------------------
# 2. Flagged Issues
# ---------------------------------------------------------------------------
FLAGGED_VIEW = [
    "Issue Group ID", "Risk Level", "Issue Type", "Material",
    "Material Description", "Supplier Name", "Normalized Batch",
    "Original Batch Values", "Expiry Dates", "Purchase Orders", "Plants",
    "Number of Records", "Reason Flagged",
]

with tab_flagged:
    st.subheader("Flagged Issues")
    if flagged.empty:
        st.info("No flagged issue groups.")
    else:
        fc1, fc2, fc3 = st.columns(3)
        risk_sel = fc1.multiselect("Risk Level", sorted(flagged["Risk Level"].unique()))
        type_sel = fc2.multiselect("Issue Type", sorted(flagged["Issue Type"].unique()))
        search = fc3.text_input("Search (material / supplier / batch / PO / description)")

        view = flagged.copy()
        if risk_sel:
            view = view[view["Risk Level"].isin(risk_sel)]
        if type_sel:
            view = view[view["Issue Type"].isin(type_sel)]
        if search:
            cols = ["Material", "Material Description", "Supplier Name",
                    "Original Batch Values", "Purchase Orders"]
            mask = pd.Series(False, index=view.index)
            for c in cols:
                if c in view.columns:
                    mask |= view[c].astype(str).str.contains(search, case=False, na=False)
            view = view[mask]

        st.caption(f"{len(view)} of {len(flagged)} issue groups")
        st.dataframe(view[FLAGGED_VIEW], use_container_width=True, hide_index=True)

        if not view.empty:
            gid = st.selectbox("Show related records for issue group", view["Issue Group ID"].tolist())
            idx = result.members.get(gid, [])
            st.markdown(f"**Original SAP records for {gid}** ({len(idx)})")
            st.dataframe(df.loc[idx].drop(columns=[NORMALIZED_BATCH], errors="ignore"),
                         use_container_width=True, hide_index=True)

# ---------------------------------------------------------------------------
# 3. Material Batch History
# ---------------------------------------------------------------------------
with tab_history:
    st.subheader("Material Batch History")
    materials = sorted(m for m in df[COL_MATERIAL].unique() if str(m).strip())
    if not materials:
        st.info("No materials found.")
    else:
        mat = st.selectbox("Material", materials, key="bq_history_material")
        g = df[df[COL_MATERIAL] == mat]
        has_sled = COL_BATCH_SLED in g.columns
        dates = g[COL_BATCH_SLED].dropna() if has_sled else pd.Series([], dtype="datetime64[ns]")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Distinct batches", g[COL_BATCH].replace("", pd.NA).nunique())
        m2.metric("Distinct expiry dates", dates.nunique())
        m3.metric("Suppliers", g[COL_SUPPLIER_NAME].replace("", pd.NA).nunique() if COL_SUPPLIER_NAME in g else 0)
        m4.metric("Purchase orders", g[COL_PO].replace("", pd.NA).nunique() if COL_PO in g else 0)
        if not dates.empty:
            e1, e2 = st.columns(2)
            e1.metric("Earliest expiry", dates.min().strftime("%Y-%m-%d"))
            e2.metric("Latest expiry", dates.max().strftime("%Y-%m-%d"))
        hist_cols = [c for c in [
            COL_SUPPLIER_NAME, COL_VENDOR_NAME, COL_PO, COL_PLANT, COL_BATCH,
            NORMALIZED_BATCH, COL_BATCH_SLED, COL_RECEIVED, COL_QTY,
        ] if c in g.columns]
        st.caption("Different production batches with different expiry dates may be normal — review, don't assume error.")
        st.dataframe(g[hist_cols], use_container_width=True, hide_index=True)

# ---------------------------------------------------------------------------
# 4. AI-Assisted Review
# ---------------------------------------------------------------------------
api_key = ai.get_api_key()
if not api_key:
    try:
        api_key = st.secrets.get("anthropic", {}).get("api_key")  # type: ignore[attr-defined]
    except Exception:  # noqa: BLE001
        api_key = None
ai_ready = bool(api_key)

with tab_ai:
    st.subheader("AI-Assisted Review")
    st.info("AI is assisting the review. It is not determining which batch or expiry date is correct.")
    if not ai_ready:
        st.warning(
            "AI review is disabled — set `ANTHROPIC_API_KEY` (and optionally "
            "`ANTHROPIC_MODEL`) in the environment or Streamlit secrets. All "
            "rule-based features remain available."
        )
    if flagged.empty:
        st.info("No issue groups to review.")
    else:
        gid = st.selectbox("Issue group", flagged["Issue Group ID"].tolist(), key="bq_ai_group")
        grp_row = flagged[flagged["Issue Group ID"] == gid].iloc[0].to_dict()
        idx = result.members.get(gid, [])
        related_records = df.loc[idx]

        st.markdown("**Issue group**")
        st.json({k: grp_row[k] for k in ["Issue Type", "Risk Level", "Material",
                                         "Normalized Batch", "Original Batch Values",
                                         "Expiry Dates", "Number of Records", "Reason Flagged"]})
        st.markdown("**Related SAP records**")
        st.dataframe(related_records.drop(columns=[NORMALIZED_BATCH], errors="ignore"),
                     use_container_width=True, hide_index=True)

        mat = related_records[COL_MATERIAL].iloc[0] if not related_records.empty else None
        nbs = [n for n in related_records[NORMALIZED_BATCH].unique() if str(n).strip()]
        material_records = df[df[COL_MATERIAL] == mat] if mat is not None else df.iloc[0:0]
        normbatch_records = df[df[NORMALIZED_BATCH].isin(nbs)] if nbs else df.iloc[0:0]

        if st.button("Review Selected Issue with AI", disabled=not ai_ready, key="bq_ai_run"):
            context = ai.build_ai_context(grp_row, related_records, material_records, normbatch_records)
            with st.spinner("Asking the AI assistant to review this issue..."):
                try:
                    st.session_state[_AI_KEY][gid] = ai.review_issue(context, api_key=api_key)
                except Exception as exc:  # noqa: BLE001
                    st.error(f"AI review failed: {exc}")

        air = st.session_state[_AI_KEY].get(gid)
        if air is not None:
            st.markdown("---")
            st.markdown(f"**AI Risk Level:** {air.risk_level}")
            st.markdown(f"**Review Summary**\n\n{air.review_summary}")
            st.markdown(f"**Pattern Identified**\n\n{air.pattern_identified}")
            st.markdown(f"**Reason for Review**\n\n{air.reason_for_review}")
            st.markdown(f"**Records Involved**\n\n{air.records_involved}")
            st.markdown(f"**Recurring Pattern**\n\n{air.recurring_pattern}")
            st.markdown("**Documents to Verify**")
            st.write(air.documents_to_verify)
            st.markdown("**Possible Root Causes**")
            st.write(air.possible_root_causes)
            st.markdown("**Recommended Review Steps**")
            st.write(air.recommended_review_steps)
            st.markdown("**Questions for Supplier or Receiver**")
            st.write(air.questions_for_supplier_or_receiver)
            st.markdown(f"**Review Note**\n\n{air.review_note}")

        # Human review fields
        st.markdown("---")
        st.markdown("**Record human findings**")
        existing = rv.get_review(st.session_state[_REVIEW_KEY], gid)
        with st.form(f"bq_review_form_{gid}"):
            confirmed = st.radio("Confirmed Issue", rv.CONFIRMED_OPTIONS,
                                 index=_idx(rv.CONFIRMED_OPTIONS, existing.get("Confirmed Issue")))
            root_cause = st.selectbox("Root Cause", rv.ROOT_CAUSE_OPTIONS,
                                      index=_idx(rv.ROOT_CAUSE_OPTIONS, existing.get("Root Cause")))
            responsible = st.selectbox("Responsible Area", rv.RESPONSIBLE_AREA_OPTIONS,
                                       index=_idx(rv.RESPONSIBLE_AREA_OPTIONS, existing.get("Responsible Area")))
            status = st.selectbox("Review Status", rv.REVIEW_STATUS_OPTIONS,
                                  index=_idx(rv.REVIEW_STATUS_OPTIONS, existing.get("Review Status")))
            follow = st.radio("Follow-Up Required", rv.FOLLOW_UP_OPTIONS,
                              index=_idx(rv.FOLLOW_UP_OPTIONS, existing.get("Follow-Up Required")))
            comment = st.text_area("Reviewer Comment", value=existing.get("Reviewer Comment", ""))
            if st.form_submit_button("Save Review", type="primary"):
                rv.save_review(st.session_state[_REVIEW_KEY], gid, {
                    "Confirmed Issue": confirmed,
                    "Root Cause": root_cause,
                    "Responsible Area": responsible,
                    "Review Status": status,
                    "Follow-Up Required": follow,
                    "Reviewer Comment": comment,
                })
                st.success(f"Saved review for {gid}.")

# ---------------------------------------------------------------------------
# 5. Review Results
# ---------------------------------------------------------------------------
with tab_results:
    st.subheader("Review Results")
    results_df = _ai_review_df()
    if results_df.empty:
        st.info("No AI reviews or human findings saved yet.")
    else:
        st.dataframe(results_df, use_container_width=True, hide_index=True)

# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------
st.markdown("---")
with st.spinner("Building workbook..."):
    xlsx = generate_excel(flagged, related, result.multi_batch, _ai_review_df())
st.download_button(
    "⬇️ Download Batch Quality Analysis",
    data=xlsx,
    file_name=f"Batch Quality Analysis {pd.Timestamp.today():%B %d %Y}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
