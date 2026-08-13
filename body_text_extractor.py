import io
import hashlib
import time
import subprocess
import streamlit as st
import openpyxl
import concurrent.futures
import queue
from playwright.sync_api import sync_playwright, Error as PlaywrightError

def ensure_playwright_installed():
    """
    Attempts to run playwright browsers. If it fails due to missing executable,
    it runs `playwright install chromium`.
    """
    try:
        with sync_playwright() as p:
            # Just try to launch the browser quickly to see if it's installed
            browser = p.chromium.launch(headless=True)
            browser.close()
    except PlaywrightError as e:
        if "Executable doesn't exist" in str(e) or "playwright install" in str(e):
            st.info("First run detected: Installing Chromium for Playwright. This may take a minute...")
            subprocess.run(["playwright", "install", "chromium"], check=True)
            st.success("Chromium installed successfully!")
        else:
            raise e

def _process_chunk(chunk, progress_queue):
    """
    Processes a chunk of URLs in a single Playwright browser context.
    Yields progress back to the main thread via progress_queue.
    """
    results = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()

        for item in chunk:
            url = item["url"]
            row_idx = item["row"]
            
            page = None
            try:
                page = context.new_page()
                page.goto(url, wait_until='domcontentloaded', timeout=15000)
                page.wait_for_timeout(2000) # 2s fixed wait
                
                # Extract innerText
                inner_text = page.evaluate("document.body.innerText")
                
                if inner_text:
                    cleaned_text = " ".join(inner_text.split())
                    results.append({"row": row_idx, "url": url, "text": cleaned_text, "error": None})
                else:
                    results.append({"row": row_idx, "url": url, "text": None, "error": "ERROR: No body text found"})
                    
            except Exception as e:
                error_msg = f"ERROR: {str(e).splitlines()[0]}"
                results.append({"row": row_idx, "url": url, "text": None, "error": error_msg})
            finally:
                if page:
                    try:
                        page.close()
                    except Exception:
                        pass
            
            # Delay between requests
            time.sleep(0.5)
            progress_queue.put(1)
            
        browser.close()
    return results

def process_body_text_extractor(workbook_bytes, progress_callback, max_workers):
    """
    Processes the uploaded workbook, extracts body text from URLs,
    and returns (output_excel_bytes, markdown_text, stats).
    """
    wb = openpyxl.load_workbook(io.BytesIO(workbook_bytes))
    
    if "Pages" not in wb.sheetnames:
        raise ValueError("The uploaded workbook must contain a sheet named 'Pages'.")
    
    sheet = wb["Pages"]
    
    # Check headers
    if sheet.cell(row=1, column=1).value != "Address":
        raise ValueError("Column A header (row 1) must be 'Address'.")
    if sheet.cell(row=1, column=2).value != "Status Code":
        raise ValueError("Column B header (row 1) must be 'Status Code'.")
    if sheet.cell(row=1, column=3).value != "Body Content":
        # Create it if it doesn't exist to be safe, though instructions say it will be empty
        sheet.cell(row=1, column=3, value="Body Content")

    urls_to_process = []
    # Collect all URLs
    for row_idx in range(2, sheet.max_row + 1):
        url = sheet.cell(row=row_idx, column=1).value
        status_code = sheet.cell(row=row_idx, column=2).value
        if url and isinstance(url, str) and url.strip():
            if status_code == 200 or str(status_code).strip() == "200":
                urls_to_process.append({"row": row_idx, "url": url.strip()})
    
    total_urls = len(urls_to_process)
    if total_urls == 0:
        raise ValueError("No URLs found in Column A matching status code 200.")

    ensure_playwright_installed()

    # Split into chunks of 50 to prevent browser memory leaks on very large lists
    chunk_size = 50
    chunks = [urls_to_process[i:i + chunk_size] for i in range(0, total_urls, chunk_size)]
    
    all_results = []
    progress_queue = queue.Queue()
    done_count = 0

    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = [executor.submit(_process_chunk, chunk, progress_queue) for chunk in chunks]
        
        while done_count < total_urls:
            try:
                # Wait for a progress tick from any thread
                progress_queue.get(timeout=1.0)
                done_count += 1
                progress_callback(done_count, total_urls)
            except queue.Empty:
                # If queue is empty for a while, check if all futures are done (e.g. they crashed)
                if all(f.done() for f in futures):
                    break
        
        for f in futures:
            try:
                all_results.extend(f.result())
            except Exception as e:
                # Fallback if a whole chunk fails completely
                pass

    markdown_lines = []
    success_count = 0
    error_count = 0
    
    # Write results back to the workbook
    for res in all_results:
        row_idx = res["row"]
        url = res["url"]
        
        if res["text"]:
            sheet.cell(row=row_idx, column=3, value=res["text"])
            markdown_lines.append(f"## {url}\n\n{res['text']}\n\n---")
            success_count += 1
        else:
            sheet.cell(row=row_idx, column=3, value=res["error"])
            error_count += 1

    # Save to bytes
    output_io = io.BytesIO()
    wb.save(output_io)
    output_bytes = output_io.getvalue()
    
    markdown_output = "\n".join(markdown_lines)
    
    stats = {
        "total": total_urls,
        "success": success_count,
        "error": error_count
    }
    
    return output_bytes, markdown_output, stats

def bulk_body_text_extractor_tool() -> None:
    st.title("Bulk Body Text Extractor")
    st.caption("Upload an Excel file with a 'Pages' sheet to extract the rendered body text from a list of URLs.")
    
    with st.expander("Required sheet format"):
        st.markdown(
            "- **Sheet name:** `Pages`\n"
            "- **Column A (Row 1):** `Address` (with URLs on rows 2+)\n"
            "- **Column B (Row 1):** `Status Code` (only URLs with 200 are processed)\n"
            "- **Column C (Row 1):** `Body Content` (to be filled by this tool)"
        )
        
    workers = st.slider("Max Concurrent Browsers", 1, 5, 3, help="Higher is faster but uses more memory. Reduce to 1 or 2 if the app crashes on large files.")
    
    uploaded = st.file_uploader("Upload Excel file", type=["xlsx"], key="body_text_file")
    
    if uploaded:
        st.markdown(
            """
            <style>
            [data-testid="stFileUploadDropzone"] {
                display: none;
            }
            </style>
            """,
            unsafe_allow_html=True,
        )
        workbook_bytes = uploaded.getvalue()
        upload_signature = hashlib.sha256(workbook_bytes).hexdigest()
        
        # Reset state if new file is uploaded
        if st.session_state.get("body_text_active_upload") != upload_signature:
            for key in ("body_text_output_xlsx", "body_text_output_md", "body_text_stats"):
                st.session_state.pop(key, None)
            st.session_state["body_text_active_upload"] = upload_signature
            
        if st.button("Extract Body Text", type="primary", use_container_width=True):
            progress_bar = st.progress(0, text="Initializing Playwright...")
            
            def update_progress(done: int, total: int) -> None:
                progress_bar.progress(done / total, text=f"Processing {done}/{total} URLs...")

            with st.spinner("Extracting body text. This may take a while depending on the number of URLs..."):
                try:
                    output_xlsx, output_md, stats = process_body_text_extractor(workbook_bytes, update_progress, max_workers=workers)
                except Exception as exc:
                    progress_bar.empty()
                    st.error(str(exc))
                else:
                    progress_bar.progress(1.0, text="Extraction complete")
                    st.session_state["body_text_output_xlsx"] = output_xlsx
                    st.session_state["body_text_output_md"] = output_md
                    st.session_state["body_text_stats"] = stats

    # Display results and download buttons
    if st.session_state.get("body_text_stats"):
        stats = st.session_state["body_text_stats"]
        col1, col2, col3 = st.columns(3)
        col1.metric("Total URLs Processed", stats["total"])
        col2.metric("Successfully Extracted", stats["success"])
        col3.metric("Errors", stats["error"])
        
        st.success("Extraction finished! Download your files below.")
        
        dl_col1, dl_col2 = st.columns(2)
        
        with dl_col1:
            st.download_button(
                "Download Updated Excel (.xlsx)",
                st.session_state["body_text_output_xlsx"],
                "Extracted_Body_Text.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )
            
        with dl_col2:
            st.download_button(
                "Download Markdown (.md)",
                st.session_state["body_text_output_md"],
                "Extracted_Body_Text.md",
                "text/markdown",
                use_container_width=True,
            )
