"""
Jumia Kenya – daily prohibited / restricted item monitor.

Flow: load rules from the Google Sheet -> search jumia.co.ke for each term ->
match listings -> classify (DELIST / REVIEW) -> compare with yesterday ->
write Excel report -> email it.

Env vars: SMTP_USER, SMTP_PASS, REPORT_TO (comma separated)
Optional: SHEET_XLSX (local xlsx path instead of Google), MAX_PAGES, DELAY
"""
import io, json, os, re, smtplib, time, unicodedata, datetime as dt
from email.message import EmailMessage
from pathlib import Path
from urllib.parse import quote_plus

import pandas as pd
import requests
from bs4 import BeautifulSoup

SHEET_ID = "1-oD2sxIh8f0LectY_ooKYDpOFunUqK0EBWV83ByGNWk"
SHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx"
BASE = "https://www.jumia.co.ke"
MAX_PAGES = int(os.getenv("MAX_PAGES", "2"))
DELAY = float(os.getenv("DELAY", "1.5"))
STATE = Path("state.json")
OUT = Path("reports"); OUT.mkdir(exist_ok=True)
HDRS = {"User-Agent": "Mozilla/5.0 (compatible; JumiaKE-ComplianceBot/1.0)",
        "Accept-Language": "en-KE,en;q=0.9"}

# Generic words that appear in normal product names. For these we only flag
# when the listing's BRAND field equals the term (not just the title).
NOISE = {"pan", "git", "kin", "kings", "origin", "floral", "charcoal", "mega", "alpha",
         "dura", "tiger", "action", "ace", "axe", "eno", "mst", "vega", "diane", "wisdom",
         "optimal", "renewed", "evolve", "contour", "pastel", "senna", "moza", "gripe",
         "the game", "cotton wool", "fay", "brett", "dura", "prisma", "coach", "zara",
         "hydrogen", "touch", "tears", "denk", "mika", "vim", "omo", "gyno"}

# Keyword rules for categories in the "Prohibited Products" tab that are not brand names.
PROHIBITED_KEYWORDS = [
    "viagra", "cialis", "kamagra", "sildenafil", "ritalin", "apetamin", "fibroid tea",
    "slimming tea", "slimming pills", "slimming capsules", "weight loss pills",
    "enhancement pills", "vape", "e-liquid", "e liquid", "shisha", "hookah", "cigarette",
    "rolling papers", "taser", "stun gun", "ammunition", "drone", "hydroquinone",
    "skin lightening", "skin bleaching", "whitening cream", "mercury soap",
    "feminine wash", "cry babies", "baby formula", "fire extinguisher",
    "pregnancy test", "hiv test", "test kit", "dry ice", "betnovate", "diprosone",
]


def norm(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s.lower()).strip()


def clean(v):
    return None if pd.isna(v) or not str(v).strip() else str(v).strip()


# ---------------------------------------------------------------- rules
def load_rules():
    src = os.getenv("SHEET_XLSX")
    if src:
        sheets = pd.read_excel(src, sheet_name=None, header=None)
    else:
        r = requests.get(SHEET_URL, headers=HDRS, timeout=60)
        r.raise_for_status()  # fails if the sheet is not shared "anyone with link: viewer"
        sheets = pd.read_excel(io.BytesIO(r.content), sheet_name=None, header=None)

    rules = {}  # norm term -> dict(term, kind, vendors, source)

    def add(term, kind, vendors=None, source=""):
        t = clean(term)
        if not t or len(t) < 2:
            return
        k = norm(t)
        if k in rules and rules[k]["kind"] == "brand":
            if vendors and not rules[k]["vendors"]:
                rules[k]["vendors"] = vendors
            return
        rules[k] = {"term": t, "kind": kind, "vendors": vendors, "source": source}

    for name, df in sheets.items():
        n = name.lower()
        if n.startswith("restricted brands"):
            for _, row in df.iloc[1:].iterrows():
                add(row[0], "brand", None, f"Restricted brand – {clean(row[1])}")
        elif n.startswith("pharmaceutical"):
            for _, row in df.iloc[1:].iterrows():
                add(row[0], "pharma", None, "Pharma OTC")
                add(row[1], "pharma", None, "Pharma Rx")
                add(row[3], "brand", clean(row[5]), "Restricted brand (allowed vendors list)")
        elif n.startswith("prohibited"):
            for v in df.values.ravel():
                v = clean(v)
                if v and "•" in v:
                    for line in v.splitlines():
                        line = line.replace("•", "").strip(" ,\t")
                        if len(line) > 3:
                            add(line, "prohibited", None, "Prohibited product (named)")
    for kw in PROHIBITED_KEYWORDS:
        add(kw, "prohibited", None, "Prohibited category keyword")
    return rules


def vendor_policy(vendors):
    """Returns list of allowed seller names, [] if nobody allowed, None if unknown/central."""
    if not vendors:
        return []
    v = vendors.lower()
    if v.startswith(("none", "banned")):
        return []
    if "central restricted" in v:
        return None
    return [x.strip() for x in re.split(r",|/", vendors) if x.strip()]


# ---------------------------------------------------------------- scraping
session = requests.Session(); session.headers.update(HDRS)


def get(url, tries=3):
    for i in range(tries):
        try:
            r = session.get(url, timeout=30)
            if r.status_code == 200:
                return r.text
            if r.status_code in (403, 429):
                time.sleep(10 * (i + 1))
        except requests.RequestException:
            time.sleep(3)
    return None


def search(term):
    items = []
    for p in range(1, MAX_PAGES + 1):
        html = get(f"{BASE}/catalog/?q={quote_plus(term)}&sort=newest&page={p}")
        time.sleep(DELAY)
        if not html:
            break
        soup = BeautifulSoup(html, "html.parser")
        cards = soup.select("article.prd")
        if not cards:
            break
        for c in cards:
            a = c.select_one("a.core") or c.find("a", href=True)
            if not a or not a.get("href"):
                continue
            name = a.get("data-gtm-name") or (c.select_one(".name") or a).get_text(" ", strip=True)
            brand = a.get("data-gtm-brand") or a.get("data-brand") or ""
            prc = c.select_one(".prc")
            price = a.get("data-gtm-price") or (prc.get_text(strip=True) if prc else "")
            items.append({"name": name, "brand": brand, "price": price,
                          "url": BASE + a["href"] if a["href"].startswith("/") else a["href"]})
    return items


_seller_cache = {}


def seller_of(url):
    if url in _seller_cache:
        return _seller_cache[url]
    html = get(url); time.sleep(DELAY)
    seller = None
    if html:
        soup = BeautifulSoup(html, "html.parser")
        # Heuristic: link/text right after the "Seller Information" heading. VERIFY on live page.
        h = soup.find(string=re.compile(r"Seller Information", re.I))
        if h:
            box = h.find_parent().find_next(["a", "p", "div"])
            seller = box.get_text(" ", strip=True) if box else None
        if not seller:
            m = re.search(r'"seller"\s*:\s*\{[^}]*"name"\s*:\s*"([^"]+)"', html)
            seller = m.group(1) if m else None
    _seller_cache[url] = seller
    return seller


def matches(rule_key, rule, item):
    pat = re.compile(rf"(?<![a-z0-9]){re.escape(rule_key)}(?![a-z0-9])")
    nm, br = norm(item["name"]), norm(item["brand"])
    if rule_key in NOISE or len(rule_key) <= 3:
        return br == rule_key
    return bool(pat.search(nm) or pat.search(br))


# ---------------------------------------------------------------- main
def run():
    today = dt.date.today().isoformat()
    state = json.loads(STATE.read_text()) if STATE.exists() else {"open": {}}
    rules = load_rules()
    print(f"{len(rules)} rule terms loaded")

    found = {}  # url -> record
    for i, (key, rule) in enumerate(rules.items(), 1):
        if i % 100 == 0:
            print(f"  {i}/{len(rules)}")
        for it in search(rule["term"]):
            if it["url"] in found or not matches(key, rule, it):
                continue
            status, seller = "DELIST", ""
            if rule["kind"] == "brand":
                allowed = vendor_policy(rule["vendors"])
                if allowed is None:
                    status = "REVIEW"
                elif allowed:
                    seller = seller_of(it["url"]) or ""
                    if not seller:
                        status = "REVIEW"
                    elif any(norm(a) in norm(seller) or norm(seller) in norm(a) for a in allowed):
                        continue  # authorised vendor – fine
                else:
                    seller = seller_of(it["url"]) or ""
                    if "jumia" in norm(seller):
                        status = "REVIEW"
            found[it["url"]] = {"Status": status, "Matched term": rule["term"], "Rule": rule["source"],
                                "Product": it["name"], "Brand": it["brand"], "Price": it["price"],
                                "Seller": seller, "URL": it["url"]}

    prev = state["open"]
    rows = []
    for url, rec in found.items():
        first = prev.get(url, {}).get("first_seen", today)
        rec["First seen"] = first
        rec["Days live"] = (dt.date.fromisoformat(today) - dt.date.fromisoformat(first)).days
        rec["New today"] = "YES" if url not in prev else ""
        rows.append(rec)
    resolved = [{"Product": v["Product"], "URL": u, "First seen": v["first_seen"], "Gone on": today}
                for u, v in prev.items() if u not in found]

    state["open"] = {u: {"first_seen": r["First seen"], "Product": r["Product"]} for u, r in found.items()}
    STATE.write_text(json.dumps(state, indent=1))

    df = pd.DataFrame(rows, columns=["Status", "New today", "Days live", "First seen", "Matched term", "Rule",
                                     "Product", "Brand", "Price", "Seller", "URL"])
    if not df.empty:
        df = df.sort_values(["Status", "New today", "Days live"], ascending=[True, False, False])
    path = OUT / f"jumia_prohibited_{today}.xlsx"
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        df[df.Status == "DELIST"].to_excel(xw, "Delist", index=False)
        df[df.Status == "REVIEW"].to_excel(xw, "Review", index=False)
        pd.DataFrame(resolved).to_excel(xw, "Resolved since yesterday", index=False)
        for ws in xw.book.worksheets:
            ws.freeze_panes = "A2"
            for col, w in zip("ABCDEFGHIJK", (10, 10, 10, 12, 22, 30, 50, 18, 10, 22, 60)):
                ws.column_dimensions[col].width = w

    n_del, n_rev = (df.Status == "DELIST").sum(), (df.Status == "REVIEW").sum()
    n_new = (df["New today"] == "YES").sum() if not df.empty else 0
    body = (f"Jumia KE prohibited-items check – {today}\n\n"
            f"To delist: {n_del}\nTo review: {n_rev}\nNew today: {n_new}\n"
            f"Resolved since yesterday: {len(resolved)}\n\nFull list attached.")
    send(f"[Jumia KE] Prohibited items {today}: {n_del} to delist, {n_new} new", body, path)
    print(body)


def send(subject, body, attachment):
    if not os.getenv("SMTP_USER"):
        print("SMTP not configured – report saved only"); return
    m = EmailMessage()
    m["Subject"], m["From"], m["To"] = subject, os.environ["SMTP_USER"], os.environ["REPORT_TO"]
    m.set_content(body)
    m.add_attachment(attachment.read_bytes(), maintype="application",
                     subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=attachment.name)
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
        s.login(os.environ["SMTP_USER"], os.environ["SMTP_PASS"])
        s.send_message(m)


if __name__ == "__main__":
    run()
