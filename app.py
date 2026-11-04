"""
Ambenix Lead Finder
Search businesses by profession + location (city / area / state / country),
collect details from Google Places API, find emails & social links from their
websites, and export everything to Excel.

Run:  streamlit run app.py
"""
import io
import re
import time
from concurrent.futures import ThreadPoolExecutor
from urllib.parse import urljoin

import pandas as pd
import requests
import streamlit as st

PLACES_URL = "https://places.googleapis.com/v1/places:searchText"
FIELD_MASK = ",".join([
    "places.displayName", "places.formattedAddress", "places.websiteUri",
    "places.nationalPhoneNumber", "places.internationalPhoneNumber",
    "places.googleMapsUri", "places.rating", "places.userRatingCount",
    "places.businessStatus", "nextPageToken",
])

EMAIL_RE = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
BAD_EMAIL_END = (".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg", ".css", ".js")
SOCIAL_PATTERNS = {
    "Facebook": r"https?://(?:www\.)?facebook\.com/[^\s\"'<>)]+",
    "Instagram": r"https?://(?:www\.)?instagram\.com/[^\s\"'<>)]+",
    "LinkedIn": r"https?://(?:[a-z]{2,3}\.)?linkedin\.com/(?:company|in)/[^\s\"'<>)]+",
    "Twitter/X": r"https?://(?:www\.)?(?:twitter|x)\.com/[^\s\"'<>)]+",
    "YouTube": r"https?://(?:www\.)?youtube\.com/[^\s\"'<>)]+",
    "Pinterest": r"https?://(?:www\.)?pinterest\.[a-z.]+/[^\s\"'<>)]+",
}
HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; LeadFinder/1.0)"}


def search_places(api_key, query, max_results):
    """Google Places API (New) Text Search with pagination (max 60 per query)."""
    results, token = [], None
    while len(results) < max_results:
        body = {"textQuery": query, "pageSize": 20}
        if token:
            body["pageToken"] = token
        r = requests.post(
            PLACES_URL,
            json=body,
            headers={"X-Goog-Api-Key": api_key, "X-Goog-FieldMask": FIELD_MASK},
            timeout=30,
        )
        if r.status_code != 200:
            raise RuntimeError(f"Google API error {r.status_code}: {r.text[:300]}")
        data = r.json()
        results.extend(data.get("places", []))
        token = data.get("nextPageToken")
        if not token:
            break
        time.sleep(1.5)
    return results[:max_results]


def fetch(url):
    try:
        r = requests.get(url, headers=HEADERS, timeout=10)
        if r.ok and "text" in r.headers.get("Content-Type", ""):
            return r.text
    except requests.RequestException:
        pass
    return ""


def enrich_website(site):
    """Crawl homepage + contact/about pages for emails and social links."""
    out = {"Emails": "", **{k: "" for k in SOCIAL_PATTERNS}}
    if not site:
        return out
    pages = [site] + [urljoin(site, p) for p in ("/contact", "/contact-us", "/about", "/about-us")]
    html = ""
    for p in pages:
        html += fetch(p)
    emails = {
        e.lower() for e in EMAIL_RE.findall(html)
        if not e.lower().endswith(BAD_EMAIL_END)
    }
    out["Emails"] = ", ".join(sorted(emails)[:5])
    for name, pat in SOCIAL_PATTERNS.items():
        m = re.findall(pat, html)
        if m:
            out[name] = m[0].rstrip("/\\.,;")
    return out


def whatsapp_link(phone):
    digits = re.sub(r"\D", "", phone or "")
    return f"https://wa.me/{digits}" if digits else ""


def build_dataframe(places, enrich):
    rows = []
    for p in places:
        rows.append({
            "Business Name": p.get("displayName", {}).get("text", ""),
            "Owner Name": "",  # not available from Google; fill manually / via LinkedIn
            "Address": p.get("formattedAddress", ""),
            "Phone": p.get("internationalPhoneNumber") or p.get("nationalPhoneNumber", ""),
            "WhatsApp Link": whatsapp_link(p.get("internationalPhoneNumber", "")),
            "Website": p.get("websiteUri", ""),
            "Has Website": "Yes" if p.get("websiteUri") else "No",
            "Rating": p.get("rating", ""),
            "Reviews": p.get("userRatingCount", ""),
            "Status": p.get("businessStatus", ""),
            "Google Maps URL": p.get("googleMapsUri", ""),
        })
    df = pd.DataFrame(rows)
    if df.empty:
        return df
    if enrich:
        with ThreadPoolExecutor(max_workers=8) as ex:
            extra = list(ex.map(enrich_website, df["Website"]))
        df = pd.concat([df, pd.DataFrame(extra)], axis=1)
    return df


def to_excel(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Leads")
        ws = w.sheets["Leads"]
        for col in ws.columns:
            width = max(len(str(c.value)) if c.value else 0 for c in col[:50])
            ws.column_dimensions[col[0].column_letter].width = min(max(width + 2, 12), 50)
        ws.freeze_panes = "A2"
    return buf.getvalue()


# ----------------------------- UI -----------------------------
st.set_page_config(page_title="Ambenix Lead Finder", layout="wide")
st.title("Ambenix Lead Finder")
st.caption("Find businesses on Google Maps, collect contact details, export to Excel.")

with st.sidebar:
    api_key = st.text_input("Google Places API key", type="password")
    st.markdown("Get a key: Google Cloud Console → enable **Places API (New)**.")

c1, c2, c3 = st.columns([2, 1, 2])
profession = c1.text_input("Profession / business type", placeholder="e.g. dentist, restaurant, gym")
loc_type = c2.selectbox("Location type", ["City", "Area", "State", "Country"])
location = c3.text_input(f"{loc_type} name", placeholder="e.g. Ahmedabad")

c4, c5, c6 = st.columns(3)
max_results = c4.slider("Max results", 20, 60, 60, step=20)
only_no_site = c5.checkbox("Only businesses WITHOUT a website", value=False)
enrich = c6.checkbox("Find emails & social links from websites", value=True)

if st.button("Find leads", type="primary"):
    if not api_key or not profession or not location:
        st.error("Enter API key, profession and location.")
        st.stop()
    query = f"{profession} in {location}"
    try:
        with st.spinner("Searching Google Maps..."):
            places = search_places(api_key, query, max_results)
        with st.spinner("Collecting contact details..."):
            df = build_dataframe(places, enrich)
    except Exception as e:
        st.error(str(e))
        st.stop()
    st.session_state["df"] = df

df = st.session_state.get("df")
if df is not None:
    if only_no_site and not df.empty:
        df = df[df["Has Website"] == "No"]
    st.success(f"{len(df)} leads found")
    st.dataframe(df, use_container_width=True)
    if not df.empty:
        st.download_button(
            "Download Excel",
            to_excel(df),
            file_name="leads.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
