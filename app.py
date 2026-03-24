"""
URL Redirect Mapping Tool
=========================
Replaces the Excel VBA macro for mapping old henkel-adhesives.com URLs
to new next.henkel-adhesives.com URLs, with async HTTP status checks.

Run with:  streamlit run app.py
"""

import asyncio
import csv
import io
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor

import aiohttp
import openpyxl
import pandas as pd
import streamlit as st

# =========================================================
# Constants
# =========================================================
CATEGORIES_CSV = Path(__file__).parent / "categories.csv"
NEW_BASE_DOMAIN = "https://next.henkel-adhesives.com"
OLD_DOMAIN_MARKER = ".com/"
HTTP_CONCURRENCY = 30
HTTP_TIMEOUT = aiohttp.ClientTimeout(total=12, connect=5)

# =========================================================
# Data Loading
# =========================================================
@st.cache_data
def load_cat_dict() -> dict[str, str]:
    """Load Product Categories CSV into {norm_key: henkel_code} dict."""
    cat_dict: dict[str, str] = {}
    with open(CATEGORIES_CSV, encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            nk = row["norm_key"].strip()
            hc = row["henkel_code"].strip()
            if nk and hc and nk not in cat_dict:
                cat_dict[nk] = hc
    return cat_dict


@st.cache_data
def load_cat_df() -> pd.DataFrame:
    """Load Product Categories as DataFrame (for output Excel)."""
    return pd.read_csv(CATEGORIES_CSV, encoding="utf-8")


def read_sitemap_urls(file_obj) -> list[str]:
    """Read URLs from column A of an uploaded Excel file."""
    wb = openpyxl.load_workbook(file_obj, read_only=True, data_only=True)
    ws = wb.active
    urls: list[str] = []
    for row in ws.iter_rows(min_row=1, max_col=1, values_only=True):
        val = row[0]
        if val is not None:
            s = str(val).strip()
            if s.startswith("http"):
                urls.append(s)
    wb.close()
    return urls


# =========================================================
# String Helpers
# =========================================================
def norm_text(s: str) -> str:
    """Normalize text for category matching (mirrors VBA NormText)."""
    t = s.lower().strip()
    t = t.replace("-", "").replace(" ", "").replace("\u00a0", "")
    return t


def strip_query_hash(url: str) -> str:
    p = url.find("?")
    if p >= 0:
        url = url[:p]
    p = url.find("#")
    if p >= 0:
        url = url[:p]
    return url


def remove_trailing_slash(url: str) -> str:
    return url.rstrip("/")


def remove_html_ext(s: str) -> str:
    if s.lower().endswith(".html"):
        return s[:-5]
    return s


def contains_text(haystack: str, needle: str) -> bool:
    return needle in haystack.lower()


# =========================================================
# URL Mapping Logic
# =========================================================
def build_product_url(
    parts: list[str], base: str, fallback: str, cat_dict: dict[str, str]
) -> str:
    """Try to match URL slugs against Product Categories. Last slug first,
    then walk backwards through path segments (skip country/lang/section)."""
    last_idx = len(parts) - 1
    if last_idx < 0:
        return fallback

    # 1) Try last slug (filename without .html)
    candidate = remove_html_ext(parts[last_idx])
    norm = norm_text(candidate)
    if norm and norm in cat_dict:
        return f"{base}/products.html/producttype_{cat_dict[norm]}.html"

    # 2) Walk backwards from second-to-last to index 3
    if last_idx >= 3:
        for i in range(last_idx - 1, 2, -1):
            candidate = parts[i]
            norm = norm_text(candidate)
            if norm and norm in cat_dict:
                return f"{base}/products.html/producttype_{cat_dict[norm]}.html"

    return fallback


def build_new_url(old_url: str, cat_dict: dict[str, str]) -> str:
    """Map an old henkel-adhesives.com URL to a new next.henkel-adhesives.com URL."""
    url = strip_query_hash(old_url)
    url = remove_trailing_slash(url)

    # Extract path after .com/
    p = url.lower().find(OLD_DOMAIN_MARKER)
    if p < 0:
        return ""
    after_com = url[p + len(OLD_DOMAIN_MARKER) :]

    parts = after_com.split("/")
    if len(parts) < 2:
        return ""

    country = parts[0].strip()
    lang = remove_html_ext(parts[1].strip())  # Bug fix: strip .html from lang
    section = parts[2].strip() if len(parts) > 2 else ""

    if not country or not lang:
        return ""

    base = f"{NEW_BASE_DOMAIN}/{country}/{lang}"
    prod_generic = f"{base}/products.html/producttype_industrial-root-producttype.html"

    if not section:
        return f"{base}.html"

    if contains_text(section, "applications"):
        return f"{base}/applications.html"
    elif contains_text(section, "industries"):
        return f"{base}/industries.html"
    elif contains_text(section, "insights"):
        return f"{base}/knowledge.html"
    elif contains_text(section, "search"):
        return prod_generic
    elif contains_text(section, "services"):
        return f"{base}/support.html"
    elif contains_text(section, "spotlights"):
        return f"{base}/knowledge.html"
    elif contains_text(section, "about"):
        return f"{base}.html"
    elif contains_text(section, "product"):
        return build_product_url(parts, base, prod_generic, cat_dict)
    else:
        return f"{base}.html"


# =========================================================
# HTTP Status Labels (mirrors VBA exactly)
# =========================================================
def status_label_ow(code: int, error_msg: str = "") -> str:
    """OW check label (Module2 - follows redirects)."""
    if code <= 0:
        return f"ERROR ({error_msg})" if error_msg else "ERROR"
    if 200 <= code <= 299:
        return f"OK ({code})"
    if code in (301, 302):
        return f"Redirect OK ({code})"
    if 300 <= code <= 399:
        return f"Redirect ({code})"
    if code == 404:
        return "404 Not Found"
    if 400 <= code <= 499:
        return f"ERROR ({code})"
    if 500 <= code <= 599:
        return f"Server Error ({code})"
    return f"HTTP {code}"


def status_label_dep(code: int) -> str:
    """DEP check label (Module1 - no redirect follow)."""
    if 200 <= code <= 299:
        return "OK"
    if 300 <= code <= 399:
        return "Redirect"
    if code == 401:
        return "Unauthorized (401)"
    if code == 403:
        return "Forbidden (403)"
    if code == 404:
        return "Not Found (404)"
    if code == 408:
        return "Request Timeout (408)"
    if code == 429:
        return "Too Many Requests (429)"
    if 500 <= code <= 599:
        return f"Server Error ({code})"
    if code == -1:
        return "Request failed"
    return f"HTTP {code}"


# =========================================================
# Async HTTP Checker
# =========================================================
async def _check_one(
    session: aiohttp.ClientSession,
    url: str,
    sem: asyncio.Semaphore,
    follow_redirects: bool,
) -> tuple[str, int, str]:
    """Check a single URL. Returns (url, status_code, error_msg)."""
    async with sem:
        # Try HEAD first
        try:
            async with session.head(
                url, allow_redirects=follow_redirects, timeout=HTTP_TIMEOUT
            ) as resp:
                return (url, resp.status, "")
        except Exception:
            pass

        # Fallback to GET
        try:
            async with session.get(
                url, allow_redirects=follow_redirects, timeout=HTTP_TIMEOUT
            ) as resp:
                return (url, resp.status, "")
        except Exception as e:
            return (url, -1, str(e)[:80])


async def _check_all(
    urls: list[str],
    follow_redirects: bool,
) -> dict[str, tuple[int, str]]:
    """Check all URLs async with concurrency limit. Returns {url: (code, err)}."""
    sem = asyncio.Semaphore(HTTP_CONCURRENCY)
    connector = aiohttp.TCPConnector(limit=HTTP_CONCURRENCY, ssl=False)
    async with aiohttp.ClientSession(
        connector=connector,
        headers={"User-Agent": "Mozilla/5.0 (URL-Mapping-Tool/1.0)"},
    ) as session:
        tasks = [_check_one(session, u, sem, follow_redirects) for u in urls]
        results: dict[str, tuple[int, str]] = {}
        for coro in asyncio.as_completed(tasks):
            url, code, err = await coro
            results[url] = (code, err)
        return results


def run_checks(
    urls: list[str], follow_redirects: bool
) -> dict[str, tuple[int, str]]:
    """Sync wrapper for async HTTP checks. Safe to call from Streamlit."""
    unique_urls = list(dict.fromkeys(urls))

    with ThreadPoolExecutor(max_workers=1) as pool:
        future = pool.submit(
            asyncio.run,
            _check_all(unique_urls, follow_redirects),
        )
        return future.result()


# =========================================================
# Excel Output Builder
# =========================================================
def build_output_excel(
    old_urls: list[str],
    new_urls: list[str],
    ow_statuses: dict[str, tuple[int, str]],
    dep_statuses: dict[str, tuple[int, str]],
    cat_df: pd.DataFrame,
) -> bytes:
    """Build output Excel with 5 sheets matching the VBA template format."""
    wb = openpyxl.Workbook()

    # Sheet 1: Instructions
    ws_instr = wb.active
    ws_instr.title = "Instructions"
    ws_instr["B3"] = "INSTRUCTIONS"
    ws_instr["C3"] = "RULES"
    ws_instr["B4"] = "Generated by URL Redirect Mapping Tool (Python)"
    ws_instr["C4"] = "applications -> {base}/applications.html"
    ws_instr["C5"] = "industries -> {base}/industries.html"
    ws_instr["C6"] = "insights -> {base}/knowledge.html"
    ws_instr["C7"] = "search -> {base}/products.html/producttype_..."
    ws_instr["C8"] = "services -> {base}/support.html"
    ws_instr["C9"] = "spotlights -> {base}/knowledge.html"
    ws_instr["C10"] = "about -> {base}.html"
    ws_instr["C11"] = "product/products -> Product Category lookup"
    ws_instr["C12"] = "Anything else -> {base}.html"

    # Sheet 2: OW URL (Paste Here)
    ws_ow = wb.create_sheet("OW URL (Paste Here)")
    ws_ow["A1"] = "OW URL (paste the  OW URL's in this column) "
    ws_ow["B1"] = "DEP URL`S"
    for i, (old, new) in enumerate(zip(old_urls, new_urls), start=2):
        ws_ow.cell(row=i, column=1, value=old)
        ws_ow.cell(row=i, column=2, value=new)

    # Sheet 3: Error Check OW URL
    ws_chk_ow = wb.create_sheet("Error Check OW URL")
    ws_chk_ow["A1"] = "OW URL "
    ws_chk_ow["B1"] = "ERROR (404)"
    for i, url in enumerate(old_urls, start=2):
        ws_chk_ow.cell(row=i, column=1, value=url)
        code, err = ow_statuses.get(url, (-1, "Not checked"))
        ws_chk_ow.cell(row=i, column=2, value=status_label_ow(code, err))

    # Sheet 4: error check DEP urls
    ws_chk_dep = wb.create_sheet("error check DEP urls")
    ws_chk_dep["A1"] = "URL"
    ws_chk_dep["B1"] = "Status"
    for i, url in enumerate(new_urls, start=2):
        ws_chk_dep.cell(row=i, column=1, value=url)
        code, err = dep_statuses.get(url, (-1, "Not checked"))
        ws_chk_dep.cell(row=i, column=2, value=status_label_dep(code))

    # Sheet 5: Product categories
    ws_cat = wb.create_sheet("Product categories")
    ws_cat["A1"] = "Product Category/Sub Category"
    ws_cat["B1"] = "Level"
    ws_cat["C1"] = "Henkel Code"
    ws_cat["D1"] = "NormKey"
    for i, row in enumerate(cat_df.itertuples(index=False), start=2):
        ws_cat.cell(row=i, column=1, value=row.category)
        ws_cat.cell(row=i, column=2, value=row.level)
        ws_cat.cell(row=i, column=3, value=row.henkel_code)
        ws_cat.cell(row=i, column=4, value=row.norm_key)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# =========================================================
# Streamlit UI
# =========================================================
def main():
    st.set_page_config(page_title="URL Redirect Mapping Tool", layout="wide")
    st.title("URL Redirect Mapping Tool")
    st.caption("Henkel Adhesives: OW -> DEP URL Migration")

    # Load categories
    try:
        cat_dict = load_cat_dict()
        cat_df = load_cat_df()
    except FileNotFoundError:
        st.error(f"Categories file not found: {CATEGORIES_CSV}")
        st.stop()

    # Sidebar: debug + reset
    st.sidebar.header("Status")
    st.sidebar.metric("Product Categories", len(cat_dict))
    st.sidebar.caption(f"Session keys: {list(st.session_state.keys())}")
    if st.sidebar.button("Reset (new file)"):
        st.session_state.clear()
        st.rerun()

    # --- Step 1: Upload ---
    st.header("1. Upload Sitemap")
    uploaded = st.file_uploader(
        "Drop a sitemap Excel file here (column A = URLs)",
        type=["xlsx", "xls"],
    )

    if uploaded and "old_urls" not in st.session_state:
        st.session_state["old_urls"] = read_sitemap_urls(uploaded)

    if "old_urls" not in st.session_state:
        st.info("Upload a sitemap Excel file to get started.")
        return

    old_urls = st.session_state["old_urls"]
    st.success(f"{len(old_urls)} URLs loaded")

    # --- Step 2: Generate Mappings ---
    st.header("2. Generate New URLs")
    if st.button("Generate URL Mappings", type="primary"):
        st.session_state["new_urls"] = [build_new_url(u, cat_dict) for u in old_urls]

    if "new_urls" not in st.session_state:
        return

    new_urls = st.session_state["new_urls"]
    empty_count = sum(1 for u in new_urls if not u)
    col1, col2 = st.columns(2)
    col1.metric("URLs mapped", len(new_urls))
    col2.metric("Failed to map", empty_count)

    with st.expander("Preview (first 50 rows)", expanded=False):
        preview_df = pd.DataFrame(
            {"Old URL": old_urls[:50], "New URL": new_urls[:50]}
        )
        st.dataframe(preview_df, use_container_width=True)

    # --- Step 3: HTTP Checks ---
    st.header("3. HTTP Status Checks")

    col_a, col_b = st.columns(2)

    with col_a:
        st.subheader("OW URLs (follow redirects)")
        if st.button("Check OW URLs"):
            with st.spinner("Checking OW URLs..."):
                st.session_state["ow_statuses"] = run_checks(
                    old_urls, follow_redirects=True
                )

        if "ow_statuses" in st.session_state:
            ow = st.session_state["ow_statuses"]
            ok_count = sum(1 for c, _ in ow.values() if 200 <= c <= 299)
            err_count = sum(1 for c, _ in ow.values() if c == 404)
            st.metric("OK", ok_count)
            st.metric("404", err_count)

    with col_b:
        st.subheader("DEP URLs (no redirects)")
        if st.button("Check DEP URLs"):
            with st.spinner("Checking DEP URLs..."):
                st.session_state["dep_statuses"] = run_checks(
                    new_urls, follow_redirects=False
                )

        if "dep_statuses" in st.session_state:
            dep = st.session_state["dep_statuses"]
            ok_count = sum(1 for c, _ in dep.values() if 200 <= c <= 299)
            redir_count = sum(1 for c, _ in dep.values() if 300 <= c <= 399)
            st.metric("OK", ok_count)
            st.metric("Redirect", redir_count)

    # --- Step 4: Export ---
    if "ow_statuses" in st.session_state and "dep_statuses" in st.session_state:
        st.header("4. Download Result")
        if "excel_bytes" not in st.session_state:
            st.session_state["excel_bytes"] = build_output_excel(
                old_urls,
                new_urls,
                st.session_state["ow_statuses"],
                st.session_state["dep_statuses"],
                cat_df,
            )
        st.download_button(
            "Download Output Excel",
            data=st.session_state["excel_bytes"],
            file_name="redirects_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )


if __name__ == "__main__":
    main()
