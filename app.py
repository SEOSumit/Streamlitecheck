from __future__ import annotations

import hashlib

import streamlit as st

from link_checker import output_filename, process_workbook, workbook_sheet_names
from sitemap_checker import (
    collect_rendered_links_browser,
    create_html_sitemap_audit_from_links,
    load_queries_from_workbook,
    sitemap_output_filename,
    try_fetch_raw_source,
)


st.set_page_config(page_title="SEO Toolkit", page_icon="🔎", layout="wide")
st.sidebar.title("SEO Toolkit")
selected_tool = st.sidebar.radio(
    "Select a tool",
    ["Internal Link Checker", "HTML Sitemap Audit"],
)


def internal_link_checker() -> None:
    st.title("Internal Link Implementation Checker")
    st.caption(
        "Verify whether each suggested anchor is linked to the suggested URL inside the specified paragraph."
    )

    with st.expander("Required sheet format"):
        st.markdown(
            "- **Column C:** source page URL on a page header row; linking paragraph on the rows below\n"
            "- **Column D:** suggested anchor text\n"
            "- **Column E:** suggested hyperlink\n\n"
            "A row is marked **Implemented** only when the exact anchor is linked to the suggested URL "
            "within the specified paragraph."
        )

    uploaded = st.file_uploader("Upload Excel file", type=["xlsx"], key="internal_file")
    if uploaded:
        workbook_bytes = uploaded.getvalue()
        upload_signature = hashlib.sha256(workbook_bytes).hexdigest()
        if st.session_state.get("internal_active_upload") != upload_signature:
            for key in ("internal_output", "internal_results", "internal_filename"):
                st.session_state.pop(key, None)
            st.session_state["internal_active_upload"] = upload_signature

        try:
            sheets = workbook_sheet_names(workbook_bytes)
        except Exception as exc:
            st.error(f"The workbook could not be opened: {exc}")
            return

        selected_sheet = st.selectbox("Sheet to check", sheets)
        workers = st.slider("Pages checked simultaneously", 1, 10, 5)

        if st.button("Check internal links", type="primary", use_container_width=True):
            progress_bar = st.progress(0, text="Preparing checks…")

            def update_progress(done: int, total: int) -> None:
                progress_bar.progress(done / total, text=f"Checked {done} of {total} source pages")

            try:
                output_bytes, results = process_workbook(
                    workbook_bytes, selected_sheet, max_workers=workers, progress=update_progress
                )
            except Exception as exc:
                progress_bar.empty()
                st.error(str(exc))
            else:
                progress_bar.progress(1.0, text="Check complete")
                st.session_state["internal_output"] = output_bytes
                st.session_state["internal_results"] = results
                st.session_state["internal_filename"] = output_filename(uploaded.name)

    if st.session_state.get("internal_results"):
        results = st.session_state["internal_results"]
        implemented = sum(row["Implementation Status"] == "Implemented" for row in results)
        col1, col2, col3 = st.columns(3)
        col1.metric("Suggestions checked", len(results))
        col2.metric("Implemented", implemented)
        col3.metric("Not Implemented", len(results) - implemented)
        st.download_button(
            "Download checked Excel file",
            st.session_state["internal_output"],
            st.session_state["internal_filename"],
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
        st.subheader("Detailed results")
        visible = [
            "Excel Row", "Source Page URL", "Suggested Anchor Text", "Suggested Hyperlink",
            "Implementation Status", "Actual Hyperlink Found", "Check Details", "Page HTTP Status",
        ]
        st.dataframe(
            [{key: row.get(key, "") for key in visible} for row in results],
            use_container_width=True,
            hide_index=True,
        )


def html_sitemap_gap_checker() -> None:
    st.title("HTML Sitemap Audit")
    st.caption(
        "Collect links from the rendered HTML sitemap DOM, validate the collection, then create the exact Excel audit."
    )

    url_status_file = st.file_uploader(
        "1. Upload URL and status-code list",
        type=["xlsx", "csv", "zip"],
        key="sitemap_complete",
        help="Row 1 contains headings. Column A: URL, Column B: status code, Column C: final redirect URL.",
    )
    sitemap_url = st.text_input(
        "2. HTML sitemap page URL",
        placeholder="https://www.example.com/site-map/",
    )

    scope_mode = st.radio(
        "3. Audit scope",
        ["Complete HTML sitemap audit", "Only URLs containing a specific folder or text"],
    )
    scope_pattern = ""
    if scope_mode != "Complete HTML sitemap audit":
        scope_pattern = st.text_input("Folder or URL text", placeholder="/servers-storage/")

    queries: list[str] = []
    gemini_api_key = None
    if url_status_file:
        queries = load_queries_from_workbook(url_status_file.getvalue(), url_status_file.name)
        try:
            gemini_api_key = st.secrets["GEMINI_API_KEY"]
        except Exception:
            gemini_api_key = None
        if queries:
            if gemini_api_key:
                st.caption(
                    f"Found {len(queries)} queries in the 'Query' sheet — AI anchor suggestions will run automatically."
                )
            else:
                gemini_api_key = st.text_input(
                    "Gemini API key (a 'Query' sheet was found — needed for AI anchor suggestions)",
                    type="password",
                    key="gemini_api_key_input",
                )

    signature = None
    if url_status_file and sitemap_url:
        signature = hashlib.sha256(
            url_status_file.getvalue() + sitemap_url.strip().encode("utf-8")
        ).hexdigest()
        if st.session_state.get("sitemap_input_signature") != signature:
            for key in (
                "sitemap_collected_links", "sitemap_raw_source", "sitemap_output",
                "sitemap_existing", "sitemap_missing", "sitemap_stats", "sitemap_filename",
            ):
                st.session_state.pop(key, None)
            st.session_state["sitemap_input_signature"] = signature

    if url_status_file and sitemap_url and st.button(
        "Collect and preview sitemap links", type="primary", use_container_width=True
    ):
        with st.spinner("Collecting links from the rendered sitemap…"):
            try:
                collected_links = collect_rendered_links_browser(sitemap_url)
                raw_source = try_fetch_raw_source(sitemap_url)
            except Exception as exc:
                st.error(str(exc))
            else:
                st.session_state["sitemap_collected_links"] = collected_links
                st.session_state["sitemap_raw_source"] = raw_source
                st.session_state.pop("confirm_sitemap_collection", None)
                for key in ("sitemap_output", "sitemap_existing", "sitemap_missing", "sitemap_stats", "sitemap_filename"):
                    st.session_state.pop(key, None)

    if st.session_state.get("sitemap_collected_links"):
        links = st.session_state["sitemap_collected_links"]
        unique_links = len({link["normalized"] for link in links})
        if scope_mode == "Complete HTML sitemap audit":
            scoped_links = links
        else:
            scoped_links = [
                link for link in links if scope_pattern.casefold().strip() in link["url"].casefold()
            ]
        c1, c2, c3 = st.columns(3)
        c1.metric("Rendered link occurrences", len(links))
        c2.metric("Unique rendered URLs", unique_links)
        c3.metric("Links in selected scope", len(scoped_links))
        if len(links) < 10:
            st.error(
                "Only a few links were collected. Do not generate the audit—switch to Chrome paste mode."
            )
        if st.session_state.get("sitemap_raw_source") is None:
            st.warning(
                "Raw source was blocked, so In HTML Source and Source Code Type will remain blank. "
                "Rendered sitemap matching is still available."
            )
        st.markdown("**Collection sample**")
        st.table(
            [
                {"URL": link["url"], "Anchor Text": link.get("anchor", "")}
                for link in scoped_links[:10]
            ]
        )
        confirmed = st.checkbox(
            "I confirm the collected link count and sample look correct",
            key="confirm_sitemap_collection",
        )
        if st.button(
            "Create Excel Audit (.xlsx)",
            type="primary",
            use_container_width=True,
            disabled=not confirmed or not url_status_file or len(links) < 10,
        ):
            use_ai = bool(queries) and bool(gemini_api_key)
            ai_progress_bar = st.empty()

            def _update_ai_progress(done: int, total: int) -> None:
                if total:
                    ai_progress_bar.progress(
                        done / total, text=f"Generating AI anchor suggestions… {done} of {total}"
                    )

            with st.spinner("Matching URLs and building the Excel template…"):
                try:
                    output, existing, missing, stats = create_html_sitemap_audit_from_links(
                        url_status_file.getvalue(),
                        url_status_file.name,
                        links,
                        sitemap_url,
                        scope_mode,
                        scope_pattern,
                        st.session_state.get("sitemap_raw_source"),
                        queries=queries,
                        api_key=gemini_api_key,
                        progress=_update_ai_progress if use_ai else None,
                    )
                except Exception as exc:
                    st.error(str(exc))
                else:
                    ai_progress_bar.empty()
                    st.session_state["sitemap_output"] = output
                    st.session_state["sitemap_existing"] = existing
                    st.session_state["sitemap_missing"] = missing
                    st.session_state["sitemap_stats"] = stats
                    st.session_state["sitemap_filename"] = sitemap_output_filename(url_status_file.name)

    if st.session_state.get("sitemap_stats"):
        stats = st.session_state["sitemap_stats"]
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Uploaded URLs", stats["uploaded_urls"])
        c2.metric("Sitemap links", stats["sitemap_links"])
        c3.metric("Missing URLs", stats["missing_urls"])
        c4.metric("Sitemap-only links", stats["sitemap_only"])
        st.download_button(
            "Download Excel Audit (.xlsx)",
            data=st.session_state["sitemap_output"],
            file_name=st.session_state["sitemap_filename"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
        st.success("Excel workbook ready with Summary, Existing URLs, and URLs Not in HTML Sitemap sheets.")


if selected_tool == "Internal Link Checker":
    internal_link_checker()
else:
    html_sitemap_gap_checker()
