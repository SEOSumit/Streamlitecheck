from __future__ import annotations

import hashlib

import streamlit as st

from link_checker import output_filename, process_workbook, workbook_sheet_names
from sitemap_checker import create_html_sitemap_audit, fetch_sitemap_html, sitemap_output_filename


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
        "Create one complete HTML sitemap audit workbook. The tool fetches only the sitemap page once "
        "and never visits every listed URL."
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
    source_mode = "Fetch sitemap page once"
    html_upload = None
    pasted_html = ""
    with st.expander("Fallback only if automatic sitemap fetch is blocked"):
        source_mode = st.radio(
            "Sitemap source method",
            ["Fetch sitemap page once", "Upload saved HTML file", "Paste HTML source"],
            horizontal=True,
        )
        if source_mode == "Upload saved HTML file":
            html_upload = st.file_uploader("Upload HTML file", type=["html", "htm", "txt"], key="html_source_file")
        elif source_mode == "Paste HTML source":
            pasted_html = st.text_area("Paste the complete HTML source", height=220)

    scope_mode = st.radio(
        "3. Audit scope",
        ["Complete HTML sitemap audit", "Only URLs containing a specific folder or text"],
    )
    scope_pattern = ""
    if scope_mode != "Complete HTML sitemap audit":
        scope_pattern = st.text_input("Folder or URL text", placeholder="/servers-storage/")

    if url_status_file and sitemap_url and st.button(
        "Create HTML sitemap audit", type="primary", use_container_width=True
    ):
        with st.spinner("Reading sitemap source and preparing the audit…"):
            try:
                if source_mode == "Fetch sitemap page once":
                    sitemap_source, final_sitemap_url, _ = fetch_sitemap_html(sitemap_url)
                elif source_mode == "Upload saved HTML file":
                    if not html_upload:
                        raise ValueError("Upload the saved HTML source file.")
                    sitemap_source = html_upload.getvalue().decode("utf-8", errors="replace")
                    final_sitemap_url = sitemap_url
                else:
                    if not pasted_html.strip():
                        raise ValueError("Paste the HTML source before creating the audit.")
                    sitemap_source = pasted_html
                    final_sitemap_url = sitemap_url

                output, existing, missing, stats = create_html_sitemap_audit(
                    url_status_file.getvalue(),
                    url_status_file.name,
                    sitemap_source,
                    final_sitemap_url,
                    scope_mode,
                    scope_pattern,
                )
            except Exception as exc:
                st.error(str(exc))
            else:
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
            "Download HTML sitemap audit",
            st.session_state["sitemap_output"],
            st.session_state["sitemap_filename"],
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
        st.subheader("URLs not in HTML sitemap")
        st.dataframe(st.session_state["sitemap_missing"], use_container_width=True, hide_index=True)


if selected_tool == "Internal Link Checker":
    internal_link_checker()
else:
    html_sitemap_gap_checker()
