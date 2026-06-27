"""Streamlit page renderer for Batch Quality Analysis.

Kept out of the main app file: ``pages/8_Batch_Quality_Analysis.py`` just calls
``render()``. Upload the SAP receiving export, click *Process Batch Quality
File*, and the page runs rule-based detection, builds issue groups, and runs the
AI review agent automatically — there is no separate AI tab. Four tabs: Upload &
Summary, Flagged Issues, Material Batch History, Review Results.
"""
from __future__ import annotations

import io

import pandas as pd
import streamlit as st

from . import ai_agent as ai
from . import reviews as rv
from .exporter import generate_excel
from .issue_groups import (
    analyze,
    build_related_records,
    build_results_table,
)
from .loader import (
    BatchQualityError,
    COL_BATCH,
    COL_BATCH_SLED,
    COL_MATERIAL,
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

_REVIEW_KEY = "batch_quality_reviews"        # gid -> human review dict
_AI_CACHE_KEY = "batch_quality_ai_cache"     # fingerprint -> AIReviewResult
_PROCESSED_KEY = "batch_quality_processed_file"

# Flagged Issues tab view (subset of the merged results table).
FLAGGED_VIEW = [
    "Issue Group ID", "Rule Risk Level", "AI Review Priority", "Issue Type",
    "Material", "Material Description", "Supplier Name", "Original Batch Values",
    "Expiry Dates", "Purchase Orders", "Plants", "Received Dates",
    "Number of Records", "Reason Flagged", "AI Review Summary",
    "Pattern Identified", "Documents to Verify", "Possible Root Causes",
    "Recommended Review Steps", "Questions for Supplier or Receiver",
    "AI Review Note",
]


def _idx(options: list, value) -> int:
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


def _resolve_api_key() -> str | None:
    key = ai.get_api_key()
    if key:
        return key
    try:
        return st.secrets.get("anthropic", {}).get("api_key")  # type: ignore[attr-defined]
    except Exception:  # noqa: BLE001
        return None


def render() -> None:
    st.title("Batch Quality Analysis")
    st.caption(
        "Identify possible SAP batch-number and expiry-date data-quality issues "
        "for human review. Rule-based detection runs first; an AI review agent "
        "then runs automatically over the flagged issue groups. The tool never "
        "edits SAP, deletes records, or suggests a corrected batch or expiry."
    )

    st.session_state.setdefault(_REVIEW_KEY, {})
    st.session_state.setdefault(_AI_CACHE_KEY, {})

    uploaded = st.file_uploader("SAP receiving-history export (.xlsx)", type=["xlsx"], key="bq_upload")
    if uploaded is None:
        st.session_state.pop(_PROCESSED_KEY, None)
        st.info("Upload the SAP receiving-history export to begin.")
        return

    # Persist which file was processed so later reruns (Save Review, filters,
    # selectboxes) don't reset the page — a button is True only on its own click.
    if st.button("Process Batch Quality File", type="primary"):
        st.session_state[_PROCESSED_KEY] = uploaded.file_id
    if st.session_state.get(_PROCESSED_KEY) != uploaded.file_id:
        st.info("Click **Process Batch Quality File** to analyze the uploaded export.")
        return

    api_key = _resolve_api_key()
    ai_ready = bool(api_key)

    with st.status("Processing batch quality file…", expanded=True) as status:
        status.write("Loading file…")
        try:
            result, related = _process(uploaded.getvalue())
        except BatchQualityError as exc:
            status.update(label="Could not process the file", state="error")
            st.error(str(exc))
            return
        except Exception as exc:  # noqa: BLE001
            status.update(label="Could not process the file", state="error")
            st.error(f"Could not process the file: {exc}")
            return
        status.write("Running batch rules…")
        status.write("Building issue groups…")

        ai_map: dict = {}
        if not ai_ready:
            status.write("AI review agent unavailable — set `ANTHROPIC_API_KEY`. Running rule-based checks only.")
        elif result.flagged.empty:
            status.write("No issue groups to review with AI.")
        else:
            status.write("Running AI review agent on issue groups…")
            bar = st.progress(0.0)
            max_issues = ai.get_max_issues()

            def _cb(done: int, total: int) -> None:
                bar.progress(min(done / total, 1.0) if total else 1.0)

            try:
                ai_map = ai.run_agent(
                    result, related,
                    cache=st.session_state[_AI_CACHE_KEY],
                    api_key=api_key, model=ai.get_model(),
                    max_issues=max_issues, progress_cb=_cb,
                )
            except Exception as exc:  # noqa: BLE001
                ai_map = {}
                st.warning(f"AI review agent could not run: {exc}. Rule-based results are unaffected.")
            bar.empty()
        status.write("Preparing results…")
        status.update(label="Batch quality analysis complete", state="complete")

    df = result.df
    flagged = result.flagged
    s = result.summary
    results_full = build_results_table(
        flagged, ai_map, st.session_state[_REVIEW_KEY],
        unavailable_label=("Not Available" if not ai_ready else "Not Reviewed"),
    )

    tab_summary, tab_flagged, tab_history, tab_results = st.tabs([
        "Upload & Summary", "Flagged Issues", "Material Batch History", "Review Results",
    ])

    # ---- Tab 1: Summary ---------------------------------------------------
    with tab_summary:
        st.subheader("Summary")
        c = st.columns(5)
        c[0].metric("Receiving records", f"{s['total_records']:,}")
        c[1].metric("Unique materials", f"{s['unique_materials']:,}")
        c[2].metric("Unique batches", f"{s['unique_batches']:,}")
        c[3].metric("Unique suppliers", f"{s['unique_suppliers']:,}")
        c[4].metric("Unique POs", f"{s['unique_pos']:,}")
        c = st.columns(5)
        c[0].metric("Issue groups", f"{s['total_issue_groups']:,}")
        c[1].metric("High risk", f"{s['high_risk']:,}")
        c[2].metric("Medium risk", f"{s['medium_risk']:,}")
        c[3].metric("Low risk", f"{s['low_risk']:,}")
        c[4].metric("Multi-batch & date materials", f"{s['multi_batch_materials']:,}")

        if not flagged.empty:
            st.markdown("**Issue groups by Issue Type / Rule Risk Level**")
            a, b = st.columns(2)
            a.bar_chart(flagged["Issue Type"].value_counts())
            b.bar_chart(flagged["Rule Risk Level"].value_counts())
            cc, dd = st.columns(2)
            cc.markdown("By Supplier")
            cc.bar_chart(flagged["Supplier Name"].replace("", "(none)").value_counts().head(15))
            dd.markdown("By Plant")
            dd.bar_chart(flagged["Plants"].replace("", "(none)").value_counts().head(15))
        else:
            st.success("No rule-based issues were flagged in this export.")

    # ---- Tab 2: Flagged Issues -------------------------------------------
    with tab_flagged:
        st.subheader("Flagged Issues")
        if results_full.empty:
            st.info("No flagged issue groups.")
        else:
            f1, f2, f3, f4 = st.columns(4)
            risk_sel = f1.multiselect("Rule Risk Level", sorted(results_full["Rule Risk Level"].unique()))
            prio_sel = f2.multiselect("AI Review Priority", sorted(results_full["AI Review Priority"].unique()))
            type_sel = f3.multiselect("Issue Type", sorted(results_full["Issue Type"].unique()))
            sup_sel = f4.multiselect("Supplier Name", sorted(x for x in results_full["Supplier Name"].unique() if str(x).strip()))
            f5, f6, f7, f8 = st.columns(4)
            ven_sel = f5.multiselect("Vendor Name", sorted(x for x in results_full["Vendor Name"].unique() if str(x).strip()))
            plant_sel = f6.multiselect("Plant", sorted(x for x in results_full["Plants"].unique() if str(x).strip()))
            pg_sel = f7.multiselect("Purchasing Group", sorted(x for x in results_full["Purchasing Group"].unique() if str(x).strip()))
            mat_sel = f8.multiselect("Material", sorted(x for x in results_full["Material"].unique() if str(x).strip()))
            search = st.text_input("Search (material / description / supplier / batch / PO)")

            view = results_full.copy()
            if risk_sel:
                view = view[view["Rule Risk Level"].isin(risk_sel)]
            if prio_sel:
                view = view[view["AI Review Priority"].isin(prio_sel)]
            if type_sel:
                view = view[view["Issue Type"].isin(type_sel)]
            if sup_sel:
                view = view[view["Supplier Name"].isin(sup_sel)]
            if ven_sel:
                view = view[view["Vendor Name"].isin(ven_sel)]
            if plant_sel:
                view = view[view["Plants"].isin(plant_sel)]
            if pg_sel:
                view = view[view["Purchasing Group"].isin(pg_sel)]
            if mat_sel:
                view = view[view["Material"].isin(mat_sel)]
            if search:
                cols = ["Material", "Material Description", "Supplier Name",
                        "Original Batch Values", "Purchase Orders"]
                mask = pd.Series(False, index=view.index)
                for col in cols:
                    mask |= view[col].astype(str).str.contains(search, case=False, na=False)
                view = view[mask]

            st.caption(f"{len(view)} of {len(results_full)} issue groups")
            st.dataframe(view[FLAGGED_VIEW], use_container_width=True, hide_index=True)

            if not view.empty:
                gid = st.selectbox("Show original SAP records for issue group", view["Issue Group ID"].tolist())
                idx = result.members.get(gid, [])
                st.markdown(f"**Original SAP records for {gid}** ({len(idx)})")
                st.dataframe(df.loc[idx].drop(columns=[NORMALIZED_BATCH], errors="ignore"),
                             use_container_width=True, hide_index=True)

    # ---- Tab 3: Material Batch History -----------------------------------
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
            m = st.columns(4)
            m[0].metric("Distinct batches", g[COL_BATCH].replace("", pd.NA).nunique())
            m[1].metric("Distinct expiry dates", dates.nunique())
            m[2].metric("Suppliers", g[COL_SUPPLIER_NAME].replace("", pd.NA).nunique() if COL_SUPPLIER_NAME in g else 0)
            m[3].metric("Purchase orders", g[COL_PO].replace("", pd.NA).nunique() if COL_PO in g else 0)
            if not dates.empty:
                e = st.columns(2)
                e[0].metric("Earliest expiry", dates.min().strftime("%Y-%m-%d"))
                e[1].metric("Latest expiry", dates.max().strftime("%Y-%m-%d"))
            hist_cols = [c for c in [
                COL_SUPPLIER_NAME, COL_VENDOR_NAME, COL_PO, COL_PLANT, COL_BATCH,
                NORMALIZED_BATCH, COL_BATCH_SLED, COL_RECEIVED, COL_QTY,
            ] if c in g.columns]
            st.caption("Different production batches with different expiry dates may be normal — review, don't assume error.")
            st.dataframe(g[hist_cols], use_container_width=True, hide_index=True)

            st.markdown("---")
            st.markdown("**Materials With Multiple Batches and Expiry Dates**")
            if result.multi_batch.empty:
                st.info("No materials carry more than one batch and more than one expiry date.")
            else:
                st.dataframe(result.multi_batch, use_container_width=True, hide_index=True)

    # ---- Tab 4: Review Results -------------------------------------------
    with tab_results:
        st.subheader("Review Results")
        if flagged.empty:
            st.info("No issue groups to review.")
        else:
            gid = st.selectbox("Issue group", flagged["Issue Group ID"].tolist(), key="bq_review_group")
            grp = results_full[results_full["Issue Group ID"] == gid].iloc[0].to_dict()
            left, right = st.columns(2)
            with left:
                st.markdown("**AI review (assist only)**")
                st.markdown(f"- **AI Priority:** {grp.get('AI Review Priority', '')}")
                st.markdown(f"- **Summary:** {grp.get('AI Review Summary', '')}")
                st.markdown(f"- **Pattern:** {grp.get('Pattern Identified', '')}")
                st.markdown(f"- **Documents to verify:** {grp.get('Documents to Verify', '')}")
                st.markdown(f"- **Possible root causes:** {grp.get('Possible Root Causes', '')}")
                st.markdown(f"- **Recommended steps:** {grp.get('Recommended Review Steps', '')}")
                st.markdown(f"- **Questions:** {grp.get('Questions for Supplier or Receiver', '')}")
                st.markdown(f"- **Note:** {grp.get('AI Review Note', '')}")
            with right:
                st.markdown("**Human findings (you decide)**")
                existing = rv.get_review(st.session_state[_REVIEW_KEY], gid)
                with st.form(f"bq_review_form_{gid}"):
                    confirmed = st.radio("Confirmed Issue", rv.CONFIRMED_OPTIONS,
                                         index=_idx(rv.CONFIRMED_OPTIONS, existing.get("Confirmed Issue")))
                    root_cause = st.selectbox("Root Cause", rv.ROOT_CAUSE_OPTIONS,
                                              index=_idx(rv.ROOT_CAUSE_OPTIONS, existing.get("Root Cause")))
                    responsible = st.selectbox("Responsible Area", rv.RESPONSIBLE_AREA_OPTIONS,
                                               index=_idx(rv.RESPONSIBLE_AREA_OPTIONS, existing.get("Responsible Area")))
                    rstatus = st.selectbox("Review Status", rv.REVIEW_STATUS_OPTIONS,
                                           index=_idx(rv.REVIEW_STATUS_OPTIONS, existing.get("Review Status")))
                    follow = st.radio("Follow-Up Required", rv.FOLLOW_UP_OPTIONS,
                                      index=_idx(rv.FOLLOW_UP_OPTIONS, existing.get("Follow-Up Required")))
                    comment = st.text_area("Reviewer Comment", value=existing.get("Reviewer Comment", ""))
                    if st.form_submit_button("Save Review", type="primary"):
                        rv.save_review(st.session_state[_REVIEW_KEY], gid, {
                            "Confirmed Issue": confirmed,
                            "Root Cause": root_cause,
                            "Responsible Area": responsible,
                            "Review Status": rstatus,
                            "Follow-Up Required": follow,
                            "Reviewer Comment": comment,
                        })
                        st.success(f"Saved review for {gid}.")

            st.markdown("---")
            st.markdown("**All issue groups (AI review + human findings)**")
            st.dataframe(results_full, use_container_width=True, hide_index=True)

    # ---- Export ----------------------------------------------------------
    st.markdown("---")
    # Rebuild with the latest saved reviews so the workbook reflects them.
    export_flagged = build_results_table(
        flagged, ai_map, st.session_state[_REVIEW_KEY],
        unavailable_label=("Not Available" if not ai_ready else "Not Reviewed"),
    )
    with st.spinner("Building workbook…"):
        xlsx = generate_excel(export_flagged, related, result.multi_batch)
    st.download_button(
        "⬇️ Download Batch Quality Analysis",
        data=xlsx,
        file_name=f"Batch Quality Analysis {pd.Timestamp.today():%B %d %Y}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
