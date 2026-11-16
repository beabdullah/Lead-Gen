"""Lead Finder - free lead generation tool.
Data: OpenStreetMap (Nominatim + Overpass). Emails: scraped from business websites.
Optional AI: OpenRouter (free models) to understand any business keyword.
"""
import io
import json
import os
import re
from concurrent.futures import ThreadPoolExecutor
from urllib.parse import urljoin

import requests
from dotenv import load_dotenv
from flask import Flask, jsonify, request, send_file
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

load_dotenv()
ON_RENDER = bool(os.getenv("RENDER"))  # Render sets this automatically
app = Flask(__name__, static_folder="static", static_url_path="")

UA = {"User-Agent": "LeadFinder/1.0 (free personal lead tool)"}
OVERPASS = [
    "https://overpass-api.de/api/interpreter",
    "https://overpass.kumi.systems/api/interpreter",
    "https://overpass.private.coffee/api/interpreter",
]

# Built-in keyword -> OSM tags (used when AI is off or fails)
TAGS = {
    "dentist": [("amenity", "dentist"), ("healthcare", "dentist")],
    "doctor": [("amenity", "doctors"), ("healthcare", "doctor")],
    "clinic": [("amenity", "clinic"), ("healthcare", "clinic")],
    "hospital": [("amenity", "hospital"), ("healthcare", "hospital")],
    "pharmacy": [("amenity", "pharmacy")],
    "restaurant": [("amenity", "restaurant")],
    "cafe": [("amenity", "cafe")],
    "fast food": [("amenity", "fast_food")],
    "bakery": [("shop", "bakery")],
    "hotel": [("tourism", "hotel")],
    "guest house": [("tourism", "guest_house")],
    "gym": [("leisure", "fitness_centre")],
    "salon": [("shop", "hairdresser"), ("shop", "beauty")],
    "barber": [("shop", "hairdresser")],
    "spa": [("leisure", "spa"), ("shop", "massage")],
    "supermarket": [("shop", "supermarket")],
    "school": [("amenity", "school")],
    "college": [("amenity", "college")],
    "university": [("amenity", "university")],
    "bank": [("amenity", "bank")],
    "lawyer": [("office", "lawyer")],
    "accountant": [("office", "accountant")],
    "real estate": [("office", "estate_agent")],
    "insurance": [("office", "insurance")],
    "travel agency": [("shop", "travel_agency"), ("office", "travel_agent")],
    "car repair": [("shop", "car_repair")],
    "mechanic": [("shop", "car_repair")],
    "car dealer": [("shop", "car")],
    "mobile shop": [("shop", "mobile_phone")],
    "electronics": [("shop", "electronics")],
    "clothes": [("shop", "clothes")],
    "jewellery": [("shop", "jewelry")],
    "furniture": [("shop", "furniture")],
    "bookshop": [("shop", "books")],
    "photographer": [("craft", "photographer"), ("shop", "photo")],
    "it company": [("office", "it"), ("office", "company")],
    "veterinary": [("amenity", "veterinary")],
    "optician": [("shop", "optician")],
    "laundry": [("shop", "laundry"), ("shop", "dry_cleaning")],
    "florist": [("shop", "florist")],
    "petrol pump": [("amenity", "fuel")],
}
ALIASES = {"dentists": "dentist", "doctors": "doctor", "hotels": "hotel", "restaurants": "restaurant",
           "cafes": "cafe", "gyms": "gym", "salons": "salon", "schools": "school", "banks": "bank",
           "lawyers": "lawyer", "pharmacies": "pharmacy", "hospitals": "hospital", "clinics": "clinic",
           "software house": "it company", "software company": "it company", "advocate": "lawyer",
           "beauty parlour": "salon", "beauty salon": "salon", "vet": "veterinary", "bakeries": "bakery"}

JUNK_EMAIL = re.compile(r"\.(png|jpe?g|gif|svg|webp|css|js)$|sentry|wixpress|example\.|domain\.com|your@|@2x", re.I)
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+")
TEL_RE = re.compile(r'href=["\']tel:([^"\']+)', re.I)
MAILTO_RE = re.compile(r'href=["\']mailto:([^"\'?]+)', re.I)


# ---------- keyword -> tags ----------
def builtin_tags(keyword):
    k = keyword.lower().strip()
    k = ALIASES.get(k, k)
    if k in TAGS:
        return TAGS[k]
    for name, tags in TAGS.items():
        if name in k or k in name:
            return tags
    return []


def ai_tags(keyword):
    """Ask OpenRouter for OSM tags. Returns [] on any problem."""
    key = os.getenv("OPENROUTER_API_KEY", "").strip()
    if not key:
        return []
    model = os.getenv("OPENROUTER_MODEL", "meta-llama/llama-3.3-70b-instruct:free")
    prompt = (
        f'Give the best OpenStreetMap tags to find "{keyword}" businesses. '
        'Reply ONLY with JSON: {"tags": [["key","value"], ...]} (max 4 pairs). '
        "Example for dentists: {\"tags\": [[\"amenity\",\"dentist\"]]}"
    )
    try:
        r = requests.post(
            "https://openrouter.ai/api/v1/chat/completions",
            headers={"Authorization": f"Bearer {key}"},
            json={"model": model, "messages": [{"role": "user", "content": prompt}], "temperature": 0},
            timeout=25,
        )
        text = r.json()["choices"][0]["message"]["content"]
        data = json.loads(re.search(r"\{.*\}", text, re.S).group(0))
        out = []
        for k, v in data.get("tags", [])[:4]:
            if re.fullmatch(r"[a-z_:]+", str(k)) and re.fullmatch(r"[a-z0-9_]+", str(v)):
                out.append((k, v))
        return out
    except Exception:
        return []


# ---------- OpenStreetMap ----------
def geocode(location):
    r = requests.get(
        "https://nominatim.openstreetmap.org/search",
        params={"q": location, "format": "json", "limit": 1},
        headers=UA, timeout=20,
    )
    r.raise_for_status()
    data = r.json()
    if not data:
        return None
    s, n, w, e = data[0]["boundingbox"]
    return f"{s},{w},{n},{e}", data[0].get("display_name", location)


def overpass(query):
    last = None
    for url in OVERPASS:
        try:
            r = requests.post(url, data={"data": query}, headers=UA, timeout=35)
            if r.status_code == 200:
                return r.json().get("elements", [])
            last = f"HTTP {r.status_code}"
        except Exception as ex:
            last = str(ex)
    raise RuntimeError(f"Map server is busy ({last}). Try again in a minute.")


SYN = {"dentist": "dent|orthodont", "doctor": "doctor|physician|medical", "clinic": "clinic|medical",
       "restaurant": "restaurant|grill|bbq|kitchen", "salon": "salon|parlou?r|beauty", "gym": "gym|fitness",
       "lawyer": "law|advocate|legal", "real estate": "real estate|estate|property"}


def name_regex(keyword):
    k = ALIASES.get(keyword.lower().strip(), keyword.lower().strip())
    if k in SYN:
        return SYN[k]
    return re.sub(r"[^a-z0-9 ]", "", k.rstrip("s")) or k


def build_query(bbox, tags, keyword, limit):
    if tags:
        body = "".join(f'nwr["{k}"="{v}"]["name"]({bbox});' for k, v in tags)
    else:
        body = f'nwr["name"~"{name_regex(keyword)}",i]({bbox});'
    return f"[out:json][timeout:30];({body});out center tags {min(limit * 3, 600)};"


def to_lead(el):
    t = el.get("tags", {})
    name = t.get("name") or t.get("name:en")
    if not name:
        return None
    parts = [t.get("addr:housenumber"), t.get("addr:street"), t.get("addr:suburb"),
             t.get("addr:city"), t.get("addr:postcode")]
    address = ", ".join(p for p in parts if p) or t.get("addr:full", "")
    c = el.get("center") or {"lat": el.get("lat"), "lon": el.get("lon")}
    website = t.get("website") or t.get("contact:website") or t.get("url") or ""
    if website and not website.startswith("http"):
        website = "http://" + website
    return {
        "name": name,
        "phone": t.get("phone") or t.get("contact:phone") or t.get("mobile") or t.get("contact:mobile") or "",
        "email": t.get("email") or t.get("contact:email") or "",
        "website": website,
        "address": address,
        "map": f"https://www.openstreetmap.org/?mlat={c['lat']}&mlon={c['lon']}#map=18/{c['lat']}/{c['lon']}" if c.get("lat") else "",
    }


# ---------- Google Places (optional, much more complete) ----------
def google_search(keyword, location, limit):
    key = os.getenv("GOOGLE_API_KEY", "").strip()
    if not key:
        return [], None
    queries = [f"{keyword} in {location}", f"best {keyword} in {location}",
               f"top rated {keyword} in {location}", f"{keyword} near {location} city center"]
    headers = {
        "X-Goog-Api-Key": key, "Content-Type": "application/json",
        "X-Goog-FieldMask": "places.id,places.displayName,places.formattedAddress,places.nationalPhoneNumber,"
                            "places.internationalPhoneNumber,places.websiteUri,places.googleMapsUri,nextPageToken",
    }
    out, seen = [], set()
    try:
        for q in queries:
            token = None
            for _ in range(3):  # 3 pages x 20 = up to 60 per query
                body = {"textQuery": q, "pageSize": 20}
                if token:
                    body["pageToken"] = token
                r = requests.post("https://places.googleapis.com/v1/places:searchText",
                                  headers=headers, json=body, timeout=25)
                if r.status_code != 200:
                    msg = r.json().get("error", {}).get("message", f"HTTP {r.status_code}")
                    return out, f"Google Places error: {msg}"
                data = r.json()
                for p in data.get("places", []):
                    name = p.get("displayName", {}).get("text", "")
                    if not name or p.get("id") in seen or name.lower() in seen:
                        continue
                    seen.add(name.lower())
                    out.append({
                        "name": name,
                        "phone": p.get("nationalPhoneNumber") or p.get("internationalPhoneNumber") or "",
                        "email": "",
                        "website": p.get("websiteUri", ""),
                        "address": p.get("formattedAddress", ""),
                        "map": p.get("googleMapsUri", ""),
                    })
                token = data.get("nextPageToken")
                if not token:
                    break
            if len(out) >= limit:
                break
    except Exception as ex:
        return out, f"Google Places error: {ex}"
    return out[:limit], None


# ---------- Google Maps scraper (no API key, uses a headless browser) ----------
def gmaps_scrape(keyword, location, limit):
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        return [], "Playwright is not installed. Run: pip install playwright && playwright install chromium"
    from urllib.parse import quote
    out = []
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_context(locale="en-US", viewport={"width": 1280, "height": 900}).new_page()
            page.goto(f"https://www.google.com/maps/search/{quote(keyword + ' in ' + location)}?hl=en", timeout=45000)
            try:
                page.locator('button:has-text("Accept all")').first.click(timeout=2500)
            except Exception:
                pass
            page.wait_for_selector('div[role="feed"]', timeout=25000)
            feed = page.locator('div[role="feed"]')

            links = []
            for _ in range(40):  # scroll the results list until we have enough
                hrefs = page.eval_on_selector_all('div[role="feed"] a[href*="/maps/place/"]',
                                                  "els => els.map(e => e.href)")
                links = list(dict.fromkeys(hrefs))
                if len(links) >= limit or page.locator("text=You've reached the end of the list").count():
                    break
                feed.evaluate("el => el.scrollTo(0, el.scrollHeight)")
                page.wait_for_timeout(1200)

            for href in links[:limit]:
                try:
                    page.goto(href, timeout=30000)
                    page.wait_for_selector("h1", timeout=8000)
                    page.wait_for_timeout(700)
                    name = page.locator("h1").first.inner_text().strip()

                    def attr(sel, a):
                        loc = page.locator(sel)
                        return (loc.first.get_attribute(a) or "") if loc.count() else ""

                    address = attr('button[data-item-id="address"]', "aria-label").replace("Address:", "").strip()
                    phone = attr('button[data-item-id^="phone"]', "aria-label").replace("Phone:", "").strip()
                    website = attr('a[data-item-id="authority"]', "href")
                    if name and name.lower() != "results":
                        out.append({"name": name, "phone": phone, "email": "", "website": website,
                                    "address": address, "map": href})
                except Exception:
                    continue
            browser.close()
    except Exception as ex:
        return out, f"Google Maps scraper failed: {str(ex)[:160]}"
    if not out:
        return [], "Google Maps scraper found nothing (Google may have changed its page or blocked the request)."
    return out, None


def merge(primary, extra):
    """Add leads from `extra` that are not already in `primary` (by name); fill blanks."""
    idx = {re.sub(r"[^a-z0-9]", "", l["name"].lower()): l for l in primary}
    for l in extra:
        k = re.sub(r"[^a-z0-9]", "", l["name"].lower())
        if k in idx:
            for f in ("phone", "email", "website", "address", "map"):
                if not idx[k][f] and l[f]:
                    idx[k][f] = l[f]
        else:
            primary.append(l)
            idx[k] = l
    return primary


# ---------- email / phone enrichment (free: scrape business website) ----------
def enrich(lead):
    if not lead["website"] or (lead["email"] and lead["phone"]):
        return lead
    emails, phones = set(), set()
    base = lead["website"]
    for path in ("", "contact", "contact-us"):
        try:
            r = requests.get(urljoin(base if base.endswith("/") else base + "/", path),
                             headers=UA, timeout=6)
            if r.status_code != 200 or "text/html" not in r.headers.get("content-type", ""):
                continue
            html = r.text[:400000]
            emails.update(MAILTO_RE.findall(html))
            emails.update(EMAIL_RE.findall(html))
            phones.update(p.strip() for p in TEL_RE.findall(html))
            if emails:
                break
        except Exception:
            continue
    emails = [e for e in emails if not JUNK_EMAIL.search(e)]
    if emails and not lead["email"]:
        lead["email"] = sorted(emails, key=len)[0]
    if phones and not lead["phone"]:
        lead["phone"] = sorted(phones, key=len)[0]
    return lead


# ---------- routes ----------
@app.get("/")
def index():
    return app.send_static_file("index.html")


@app.get("/api/status")
def status():
    return jsonify(ai=bool(os.getenv("OPENROUTER_API_KEY", "").strip()), google=bool(os.getenv("GOOGLE_API_KEY", "").strip()), gmaps=not ON_RENDER)


@app.post("/api/search")
def search():
    d = request.get_json(force=True)
    keyword = (d.get("keyword") or "").strip()
    location = (d.get("location") or "").strip()
    limit = max(5, min(int(d.get("limit") or 50), 200))
    if not keyword or not location:
        return jsonify(error="Enter a business type and a location."), 400
    try:
        geo = geocode(location)
        if not geo:
            return jsonify(error=f'Could not find "{location}". Try a city name.'), 404
        bbox, place = geo

        tags = builtin_tags(keyword)
        used_ai = False
        if not tags and d.get("use_ai"):
            tags = ai_tags(keyword)
            used_ai = bool(tags)

        elements, osm_error = [], None
        try:
            elements = overpass(build_query(bbox, tags, keyword, limit)) if tags else []
            if len(elements) < limit:  # also match business names (e.g. "Smile Dental Care")
                elements += overpass(build_query(bbox, [], keyword, limit))
        except RuntimeError as ex:
            osm_error = str(ex)  # not fatal: Google sources can still work

        seen, leads = set(), []
        for el in elements:
            lead = to_lead(el)
            if not lead:
                continue
            key = (lead["name"].lower(), lead["address"].lower())
            if key in seen:
                continue
            seen.add(key)
            leads.append(lead)

        sources = ["OpenStreetMap"] if not osm_error else []
        warnings = [osm_error] if osm_error else []

        g_leads, g_warn = google_search(keyword, location, limit)
        src = "Google Places"
        if g_warn:
            warnings.append(g_warn)
        if not g_leads and d.get("gmaps"):
            if ON_RENDER:
                warnings.append("The Google Maps scraper is disabled on this server. "
                                "Run Lead Finder on your own computer to use it.")
            else:
                g_leads, w2 = gmaps_scrape(keyword, location, limit)
                src = "Google Maps"
                if w2:
                    warnings.append(w2)
        if g_leads:
            leads = merge(g_leads, leads)
            sources.insert(0, src)
        if not leads and warnings:
            return jsonify(error=" ".join(warnings)), 503
        warning = None if g_leads else (" ".join(warnings) or None)

        # richest leads first, then cut to limit
        score = lambda l: sum(bool(l[k]) for k in ("phone", "email", "website", "address"))
        leads.sort(key=score, reverse=True)
        leads = leads[:limit]

        if d.get("find_emails", True):
            with ThreadPoolExecutor(max_workers=12) as ex:
                leads = list(ex.map(enrich, leads))

        return jsonify(leads=leads, place=place, ai=used_ai, sources=sources, warning=warning)
    except RuntimeError as ex:
        return jsonify(error=str(ex)), 503
    except Exception as ex:
        return jsonify(error=f"Something went wrong: {ex}"), 500


@app.post("/api/export")
def export():
    d = request.get_json(force=True)
    leads = d.get("leads", [])
    wb = Workbook()
    ws = wb.active
    ws.title = "Leads"
    heads = ["Name", "Phone", "Email", "Website", "Address", "Map"]
    ws.append(heads)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="000000")
    for l in leads:
        ws.append([l.get("name", ""), l.get("phone", ""), l.get("email", ""),
                   l.get("website", ""), l.get("address", ""), l.get("map", "")])
    for col, w in zip("ABCDEF", (32, 20, 30, 34, 50, 30)):
        ws.column_dimensions[col].width = w
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(vertical="top", wrap_text=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    fname = re.sub(r"[^a-z0-9]+", "_", f"leads_{d.get('keyword','')}_{d.get('location','')}".lower()).strip("_")
    return send_file(buf, as_attachment=True, download_name=f"{fname}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if __name__ == "__main__":
    port = int(os.getenv("PORT", "5000"))
    app.run(host="0.0.0.0" if ON_RENDER else "127.0.0.1", port=port, debug=False)