"""
Crisdel Group, Inc. — Virginia Public Opportunities Scraper
============================================================
Sources:
  • MWAA Current & Upcoming contracting opportunities
  • VDOT Northern Virginia District construction projects
  • Fairfax County solicitation portal
  • Loudoun County bid portal
  • Arlington County procurement (Vendor Registry portal)

Outputs a daily PDF report and emails it to all configured recipients.
"""
import os
import re
import json
import smtplib
import schedule
import time
import requests
import msal
import base64
from dotenv import load_dotenv
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from bs4 import BeautifulSoup
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
)
from reportlab.lib.enums import TA_CENTER

# ── CONFIG ───────────────────────────────────────────────────────────────────
# Load environment variables from .env file
load_dotenv()

# Azure/Microsoft Graph Configuration
CLIENT_ID = os.getenv("crisdel_client_id")
TENANT_ID = os.getenv("crisdel_tenant_id")
CLIENT_SECRET = os.getenv("crisdel_client_secret")

# Email Configuration
EMAIL_SENDER   = "analytics@crisdel.com"  # Mailbox to send from
EMAIL_PASSWORD = os.environ.get("CRISDEL_EMAIL_PASSWORD", "")  # Deprecated (no longer used)

# Add/remove recipients here — at least one entry required in EMAIL_TO
EMAIL_TO  = ["frankc@crisdel.com", "barryh@crisdel.com", "mpollio@crisdel.com", "groti@crisdel.com", "michaelc@crisdel.com", "franksr@crisdel.com"]           # Primary recipients
EMAIL_CC  = ["rmacak@crisdel.com"]                            # CC recipients, e.g. ["boss@crisdel.com", "team@crisdel.com"]

PDF_FILE     = "Crisdel Virginia Public Opportunities {date}.pdf"   # {date} filled at runtime
HISTORY_FILE = "seen_opportunities.json"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    )
}

# ── SOURCES ──────────────────────────────────────────────────────────────────
MWAA_SOURCES = {
    "MWAA Current":  "https://www.mwaa.com/business/current-contracting-opportunities",
    "MWAA Upcoming": "https://www.mwaa.com/business/upcoming-contracting-opportunities",
}

VDOT_SOURCES = {
    "VDOT Northern Virginia": "https://www.vdot.virginia.gov/projects/northern-virginia-district/",
}

# All three county portals are scrapable with requests.
COUNTY_SOURCES = {
    "Fairfax County": "https://www.fairfaxcounty.gov/solicitation/",
    "Loudoun County": "https://www.loudoun.gov/bids.aspx",
}

ARLINGTON_PORTAL = "https://vrapp.vendorregistry.com/Bids/View/BidsList?BuyerId=a596c7c4-0123-4202-bf15-3583300ee088"

# ── CONSTANTS ────────────────────────────────────────────────────────────────
# MWAA solicitation-number prefixes that signal a real procurement entry
MWAA_PREFIXES = ("IFB-", "RFP-", "RFQ-", "ITB-", "IFQ-", "RFQI-")

MWAA_LOCATION_KEYWORDS = [
    "dca", "iad", "dtr", "dulles", "reagan", "national airport",
    "fairfax", "loudoun", "loudon", "arlington", "virginia",
]

# Keywords that flag a solicitation as construction-related
CONSTRUCTION_KEYWORDS = [
    "construction", "renovation", "rehabilitation", "repair", "infrastructure",
    "road", "bridge", "pavement", "sidewalk", "building", "facility", "trail",
    "drainage", "utility", "grading", "excavation", "concrete", "asphalt",
    "electrical", "mechanical", "plumbing", "roofing", "structural",
    "contractor", "invitation for bid", "site work", "earthwork", "demolition",
    "waterline", "sewer", "retaining wall", "extension", "widening",
    "intersection", "signal", "corridor", "streetscape",
]

NON_CONSTRUCTION_KEYWORDS = [
    "snow removal", "snow plow", "brining", "de-icing", "deicing",
    "janitorial", "cleaning service", "custodial", "housekeeping",
    "staffing analysis", "staffing service", "temporary personnel",
    "lawn care", "mowing", "turf management", "grounds maintenance",
    "it hardware", "it service", "software", "cybersecurity",
    "wireless network", "network implementation", "network maintenance",
    "food service", "catering", "vending", "concession",
    "security guard", "guard service", "security patrol", "security service",
    "office supply", "office supplies", "printing service",
    "legal service", "legal counsel",
    "shuttle service", "bus service", "transit service",
    "pest control", "exterminator",
    "upholstery", "furniture repair", "furniture replacement",
    "relay testing", "testing service",
    "uniforms", "clothing", "apparel",
    "marketing service", "advertising service", "public relations",
    "photography", "videography",
    "accounting service", "auditing service", "financial advisory",
    "medical service", "health service", "nursing",
    "disparity study", "recreational program",
    "towing service", "courier service",
    "laundry service", "dry cleaning",
    "landscaping service", "roofing", "building rehabilitation", "CEI Services", "inspection", "consulting",
]

# VDOT filters (unchanged from original)
VDOT_EXCLUDE_KEYWORDS = [
    "public hearing", "appendix", "newsletter", "survey", "frequently asked",
    "faq", "transcript", "comment", "presentation", "study report", "final report",
    "notice of", "skip to", "policies", "accept", "proceed", "translation",
    "nondiscrimination", "civil rights", "featured site", "myvdot",
    "project pipeline", "how projects", "request assistance", "open file",
    "bus rapid", "embark", "stars study", "flashing yellow",
    "categorical exclusion", "typical section", "smart scale", "interactive map",
    "simulation video", "community meeting", "civic association",
    "homeowners association", "de minimis", "section 4", "willingness to hold",
    "virtual public information", "public information meeting", "public comment",
    "recently issued", "weekly forecast", "quarterly forecast", "week of",
    "pdf presentation", "arcgis", "vimeo", "publicinput", "conta.cc",
    "view more", "read more", "learn more", "click here", "more information",
]

VDOT_EXCLUDE_URL_FRAGMENTS = [
    "/media/", "arcgis.com", "vimeo.com", "publicinput.com", "conta.cc",
    "fairfaxcounty.gov", "arlingtonva.us", "novatransit.org", "smartportal",
    "vaprojectpipeline", "my.vdot.virginia.gov", "improve81.vdot",
    "495next.vdot", "64expresslanes", "mailto:", "#main-content",
    "/policies/", "/roads-funded/", "/project-planning/", "/site-assets/",
]

VDOT_DISTRICTS = ["northern-virginia-district"]

VDOT_EXCLUDE_COUNTIES = [
    "prince william", "stafford county", "fauquier county",
    "manassas park", "city of manassas",
]

VDOT_TARGET_LOCATIONS = [
    "fairfax", "arlington", "loudoun", "loudon",
    "herndon", "reston", "mclean", "tysons", "centreville", "chantilly",
    "springfield", "burke", "lorton", "vienna", "oakton", "annandale",
    "falls church", "merrifield", "great falls",
    "leesburg", "ashburn", "sterling", "purcellville",
    "dulles", "south riding",
]

VDOT_CONSTRUCTION_SIGNALS = [
    "improvement", "widening", "construction", "bridge", "interchange",
    "intersection", "corridor", "extension", "realignment", "lane", "road",
    "route", "highway", "street", "boulevard", "avenue", "connector",
    "overpass", "underpass", "ramp", "path", "trail", "safety", "signal",
    "drainage", "utility", "pavement", "resurfac", "sidewalk", "pedestrian",
    "multimodal", "transit", "i-64", "i-95", "i-81", "i-495", "i-66",
    "i-77", "i-85", "rte", "rt.", "u.s.", "us-", "sr-",
]

VDOT_NAV_PHRASES = [
    "view project", "read more", "learn more", "click here", "more information",
    "project page", "view more", "skip to", "accept & proceed", "proceed to open",
    "request assistance", "all policies", "how projects are funded", "featured site",
    "myvdot", "project pipeline", "view more information", "public input",
    "bristol district", "culpeper district", "fredericksburg district",
    "hampton roads district", "lynchburg district", "northern virginia district",
    "richmond district", "salem district", "staunton district",
]

VIRGINIA_KEYWORDS = [
    "virginia", " va ", " va.", " va,",
    "fairfax", "loudoun", "loudon", "arlington", "alexandria",
    "richmond", "norfolk", "hampton roads", "virginia beach",
    "northern virginia", "tysons", "reston", "herndon",
    "sterling", "ashburn", "leesburg", "manassas", "woodbridge",
    "springfield", "falls church", "mclean", "centreville",
    "chantilly", "dulles", "reagan national",
]

# ── HELPERS ──────────────────────────────────────────────────────────────────
def truncate_text(text: str, max_chars: int = 460) -> str:
    """Truncate at a word boundary and append '...' if needed."""
    text = text.strip()
    if len(text) <= max_chars:
        return text
    cut = text[:max_chars]
    last_space = cut.rfind(" ")
    if last_space > max_chars * 0.75:
        cut = cut[:last_space]
    return cut + "..."


def is_construction_related(title: str, description: str = "") -> bool:
    combined = (title + " " + description).lower()
    if any(kw in combined for kw in NON_CONSTRUCTION_KEYWORDS):
        return False
    return any(kw in combined for kw in CONSTRUCTION_KEYWORDS)


def extract_sol_num(title: str) -> str:
    """Pull the base solicitation number out of a title string."""
    m = re.match(r"^((?:IFB|RFP|RFQ|ITB|IFQ|RFQI)-\d{2}-\d+[a-zA-Z]?)",
                 title, re.IGNORECASE)
    return m.group(1).upper() if m else title


def _xml_escape(text: str) -> str:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


# ── HISTORY ──────────────────────────────────────────────────────────────────
def load_history() -> set:
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE) as f:
            return set(json.load(f))
    return set()


def save_history(seen: set) -> None:
    with open(HISTORY_FILE, "w") as f:
        json.dump(list(seen), f)


# ── MWAA SCRAPER ─────────────────────────────────────────────────────────────
def scrape_mwaa(url: str, label: str) -> list:
    """
    Parse an MWAA contracting opportunities page.

    Improvements over v4:
    - Parses the DOM structure rather than raw text so descriptions are
      captured in full (up to the truncation limit).
    - Extracts the due date and appends it to the description for quick scanning.
    - Prefers a direct MWAA solicitation page URL when available; otherwise
      uses the CLM system link or falls back to the section landing page.
    """
    print(f"  🔍 Scraping {label}: {url}")
    try:
        resp = requests.get(url, headers=HEADERS, timeout=20)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
    except Exception as e:
        print(f"  ❌ Failed: {e}")
        return []

    results = []
    seen_titles: set = set()

    for strong in soup.find_all(["strong", "b"]):
        raw = strong.get_text(strip=True)
        # Strip trailing status badges like ***UPDATED*** or ***NEW***
        clean_title = re.sub(
            r"\s*\*{1,3}\s*(UPDATED|NEW)\s*\*{1,3}\s*$", "", raw, flags=re.IGNORECASE
        ).strip()

        if not any(clean_title.upper().startswith(p) for p in MWAA_PREFIXES):
            continue
        if clean_title in seen_titles:
            continue
        seen_titles.add(clean_title)

        # Find the nearest block-level ancestor that contains the full entry
        container = None
        for tag in ("p", "div", "li", "td", "section"):
            container = strong.find_parent(tag)
            if container:
                break
        if not container:
            continue

        full_text = re.sub(r"\s+", " ", container.get_text(separator=" ", strip=True))

        # Extract description (text between "Description:" and the next label)
        desc_m = re.search(
            r"Description[:\s]+(.+?)(?=Solicitation Issue Date|Due Date|Anticipated|"
            r"Amendments Issued|SLBE Requirement|Additional Information|$)",
            full_text, re.IGNORECASE | re.DOTALL,
        )
        description = desc_m.group(1).strip() if desc_m else ""

        # Extract due / anticipated submittal date and prepend it
        due_m = re.search(
            r"(?:Due Date for Submissions|Anticipated Submittal Due Date)[:\s]+([^|]+?)(?:\s+Amendments|\s+SLBE|$)",
            full_text, re.IGNORECASE,
        )
        if due_m:
            due_str = due_m.group(1).strip()
            description = f"Due: {due_str}  |  {description}".strip(" |")

        # Find specific /contracting-opportunity/ page on mwaa.com
        link = url
        search_area = [container]
        sib = container.find_next_sibling()
        for _ in range(5):
            if sib is None or sib.name in ("h2", "h3"):
                break
            search_area.append(sib)
            sib = sib.find_next_sibling()

        for el in search_area:
            for a in el.find_all("a", href=True):
                href = a["href"].strip()
                if not href or href.startswith("#"):
                    continue
                full_href = href if href.startswith("http") else "https://www.mwaa.com" + href
                if "mwaa.com/contracting-opportunity" in full_href:
                    link = full_href
                    break
            if "contracting-opportunity" in link:
                break

        # Page-wide fallback: search entire page for a detail URL with this sol number
        if "contracting-opportunity" not in link:
            sol_slug = extract_sol_num(clean_title).lower()
            for a in soup.find_all("a", href=True):
                href = a["href"].strip()
                full_href = href if href.startswith("http") else "https://www.mwaa.com" + href
                if "contracting-opportunity" in full_href and sol_slug in full_href.lower():
                    link = full_href
                    break

        if not is_construction_related(clean_title, description):
            continue

        combined_lower = (clean_title + " " + description).lower()
        if not any(loc in combined_lower for loc in MWAA_LOCATION_KEYWORDS):
            continue

        results.append({
            "title":       clean_title,
            "description": truncate_text(description or "See link for full solicitation details."),
            "url":         link,
        })

    print(f"  ✅ Found {len(results)} solicitations in {label}")
    return results


# ── VDOT SCRAPERS ────────────────────────────────────────────────────────────
def _valid_vdot_url(url: str) -> bool:
    u = url.lower()
    if "vdot.virginia.gov" not in u or "/projects/" not in u:
        return False
    if not any(d in u for d in VDOT_DISTRICTS):
        return False
    for d in VDOT_DISTRICTS:
        if d in u:
            slug = u.split(d)[-1].strip("/").split("/")[0]
            if len(slug) < 3:
                return False
            break
    return not any(frag in u for frag in VDOT_EXCLUDE_URL_FRAGMENTS)


def _valid_vdot_title(text: str) -> bool:
    if not text or len(text) < 8:
        return False
    tl = text.lower()
    if any(kw in tl for kw in VDOT_EXCLUDE_KEYWORDS):
        return False
    if tl in [p.lower() for p in VDOT_NAV_PHRASES]:
        return False
    return any(sig in tl for sig in VDOT_CONSTRUCTION_SIGNALS)


def _fetch_vdot_description(url: str) -> str:
    try:
        resp = requests.get(url, headers=HEADERS, timeout=15)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
        for sel in [
            "div.field--type-text-with-summary",
            "div.field--name-body",
            "article .content",
            "div.views-field-body",
        ]:
            el = soup.select_one(sel)
            if el:
                t = el.get_text(separator=" ", strip=True)
                if len(t) > 40:
                    return truncate_text(t)
        for tag in soup.select("main p, article p"):
            t = tag.get_text(strip=True)
            if len(t) > 60:
                return truncate_text(t)
    except Exception as e:
        print(f"    ⚠️  Could not fetch description: {e}")
    return ""


def scrape_vdot_district(url: str, label: str) -> list:
    print(f"  🔍 Scraping {label}: {url}")

    results = []
    seen_urls: set = set()
    base_url = url.rstrip("/")

    for page_num in range(1, 12):
        page_url = base_url + "/" if page_num == 1 else f"{base_url}/{page_num}/index.php"

        try:
            resp = requests.get(page_url, headers=HEADERS, timeout=20)
            resp.raise_for_status()
            soup = BeautifulSoup(resp.text, "html.parser")
        except Exception as e:
            if page_num == 1:
                print(f"  ❌ Failed: {e}")
            break

        page_found = 0
        for a in soup.find_all("a", href=True):
            href = a["href"].strip()
            text = a.get_text(separator=" ", strip=True)
            full_url = href if href.startswith("http") else "https://www.vdot.virginia.gov" + href

            if not _valid_vdot_url(full_url) or not _valid_vdot_title(text):
                continue
            if full_url in seen_urls:
                continue
            seen_urls.add(full_url)

            description = ""
            parent = a.find_parent(["li", "div", "p", "article", "td", "section"])
            if parent:
                dt = parent.get_text(separator=" ", strip=True)
                dt = re.sub(re.escape(text), "", dt).strip()
                dt = re.sub(r"\s+", " ", dt).strip()
                if len(dt) > 30:
                    description = dt

            if not description or len(description) < 40:
                print(f"    🌐 Fetching project page: {text[:60]}...")
                description = _fetch_vdot_description(full_url)

            combined_loc = (text + " " + description).lower()
            if any(exc in combined_loc for exc in VDOT_EXCLUDE_COUNTIES):
                continue
            if not any(loc in combined_loc for loc in VDOT_TARGET_LOCATIONS):
                continue

            results.append({
                "title":       text,
                "description": truncate_text(description or "Visit the project page for full details."),
                "url":         full_url,
            })
            page_found += 1

        if page_found == 0 and page_num > 1:
            break
        if page_num > 1:
            print(f"    📄 Page {page_num}: found {page_found} projects")

    print(f"  ✅ Found {len(results)} projects in {label}")
    return results


# ── FAIRFAX COUNTY SCRAPER ───────────────────────────────────────────────────
def scrape_fairfax(url: str, label: str) -> list:
    """
    Scrape Fairfax County's solicitation listing page and filter for
    construction-related items.

    Note: Full solicitation documents require registration on the Bonfire portal
    at https://fairfaxcounty.bonfirehub.com — the link in each entry directs
    there automatically.
    """
    print(f"  🔍 Scraping {label}: {url}")
    try:
        resp = requests.get(url, headers=HEADERS, timeout=20)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
    except Exception as e:
        print(f"  ❌ Failed: {e}")
        return []

    results = []
    seen: set = set()

    # Strategy 1 — look for <h3>/<h4> headings that introduce each solicitation
    for heading in soup.find_all(["h3", "h4"]):
        title = heading.get_text(strip=True)
        if not title or len(title) < 5 or title in seen:
            continue

        # Collect the body text that follows this heading
        body_parts = []
        node = heading.find_next_sibling()
        for _ in range(8):
            if node is None or node.name in ("h3", "h4", "hr"):
                break
            body_parts.append(node.get_text(separator=" ", strip=True))
            node = node.find_next_sibling()
        body = " ".join(body_parts)

        if not is_construction_related(title, body):
            continue

        bid_m   = re.search(r"BID NUMBER[:\s]+(\S+)",                body, re.IGNORECASE)
        close_m = re.search(r"CLOSING DATE[/\s\w]*[:\s]+([^\.]+\w)", body, re.IGNORECASE)
        bid_num   = bid_m.group(1)   if bid_m   else ""
        close_str = close_m.group(1).strip() if close_m else ""

        desc = re.sub(r"BID NUMBER[:\s]+\S+", "", body, flags=re.IGNORECASE)
        desc = re.sub(r"CLOSING DATE[\w\s/]*[:\s]+[^\.]+", "", desc, flags=re.IGNORECASE)
        desc = re.sub(r"\s+", " ", desc).strip()
        if close_str:
            desc = f"Closes: {close_str}  |  {desc}".strip(" |")
        if bid_num:
            desc = f"Bid #: {bid_num}  |  {desc}".strip(" |")

        # Find the best link for this solicitation
        sol_link = "https://fairfaxcounty.bonfirehub.com/portal/?tab=openOpportunities"
        search_nodes = [heading]
        sib_node = heading.find_next_sibling()
        for _ in range(8):
            if sib_node is None or sib_node.name in ("h3", "h4", "hr"):
                break
            search_nodes.append(sib_node)
            sib_node = sib_node.find_next_sibling()

        for el in search_nodes:
            for a in el.find_all("a", href=True):
                href = a["href"].strip()
                if not href or href.startswith("#"):
                    continue
                full_href = href if href.startswith("http") else "https://www.fairfaxcounty.gov/solicitation/" + href
                if "bonfirehub.com" in full_href and "/opportunities/" in full_href:
                    sol_link = full_href
                    break
                if "DownloadPDF" in href or "AttachmentID" in href:
                    sol_link = full_href
                    break
                if "bonfirehub.com" in full_href and sol_link.endswith("openOpportunities"):
                    sol_link = full_href
            if sol_link != "https://fairfaxcounty.bonfirehub.com/portal/?tab=openOpportunities":
                break

        seen.add(title)
        results.append({
            "title":       title,
            "description": truncate_text(desc or "See Fairfax County procurement portal for details."),
            "url":         sol_link,
        })

    # Strategy 2 — text blocks separated by horizontal rules (fallback)
    if not results:
        full_text = soup.get_text(separator="\n")
        blocks = re.split(r"\n-{3,}\n", full_text)
        for block in blocks:
            lines = [ln.strip() for ln in block.split("\n") if ln.strip()]
            if not lines:
                continue
            title = lines[0]
            if len(title) < 5 or len(title) > 200 or title in seen:
                continue
            body = " ".join(lines[1:])
            if not is_construction_related(title, body):
                continue

            bid_m   = re.search(r"BID NUMBER[:\s]+(\S+)",                body, re.IGNORECASE)
            close_m = re.search(r"CLOSING DATE[/\s\w]*[:\s]+([^\n]+\w)", body, re.IGNORECASE)
            bid_num   = bid_m.group(1)          if bid_m   else ""
            close_str = close_m.group(1).strip() if close_m else ""
            desc = body[:460]
            if close_str:
                desc = f"Closes: {close_str}  |  {desc}"
            if bid_num:
                desc = f"Bid #: {bid_num}  |  {desc}"

            seen.add(title)
            results.append({
                "title":       title,
                "description": truncate_text(desc or "See Fairfax County procurement portal."),
                "url":         "https://fairfaxcounty.bonfirehub.com/portal/?tab=openOpportunities",
            })  # Fallback path has no HTML to search for links

    print(f"  ✅ Found {len(results)} construction solicitations in {label}")
    return results


# ── LOUDOUN COUNTY SCRAPER ───────────────────────────────────────────────────
def scrape_loudoun(url: str, label: str) -> list:
    """
    Scrape Loudoun County's public bid listing. Each bid has a unique URL
    (bids.aspx?bidID=NNN); construction-related bids are kept.
    """
    print(f"  🔍 Scraping {label}: {url}")
    try:
        resp = requests.get(url, headers=HEADERS, timeout=20)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
    except Exception as e:
        print(f"  ❌ Failed: {e}")
        return []

    results = []
    seen: set = set()
    seen_bid_ids: set = set()

    for a in soup.find_all("a", href=True):
        href = a["href"]
        if "bids.aspx?bidID=" not in href:
            continue

        bid_id_m = re.search(r'bidID=(\d+)', href)
        bid_id = bid_id_m.group(1) if bid_id_m else None
        if bid_id and bid_id in seen_bid_ids:
            continue

        title = a.get_text(strip=True)
        if not title or title in seen:
            continue

        title_normalized = re.sub(r'\s+', ' ', title.replace('\xa0', ' '))
        if "read on" in title_normalized.lower():
            continue

        if bid_id:
            seen_bid_ids.add(bid_id)

        if href.startswith("http"):
            full_url = href
        elif href.startswith("/"):
            full_url = "https://www.loudoun.gov" + href
        else:
            full_url = "https://www.loudoun.gov/" + href

        # Pull surrounding context for description and close date
        parent = a.find_parent(["li", "div", "p", "td"])
        desc = ""
        if parent:
            ctx = parent.get_text(separator=" ", strip=True)
            ctx = re.sub(re.escape(title), "", ctx).strip()
            ctx = re.sub(r"\s+", " ", ctx).strip()

            close_m = re.search(r"Closes[:\s]+([^\[]+?)(?:\s+Open|\s*$)", ctx, re.IGNORECASE)
            if close_m:
                desc = f"Closes: {close_m.group(1).strip()}  |  {ctx}"
            else:
                desc = ctx

        if not is_construction_related(title, desc):
            continue

        seen.add(title)
        results.append({
            "title":       title,
            "description": truncate_text(desc or "See Loudoun County bid portal for details."),
            "url":         full_url,
        })

    print(f"  ✅ Found {len(results)} construction bids in {label}")
    return results


# ── ARLINGTON COUNTY SCRAPER (Vendor Registry) ────────────────────────────────
def scrape_arlington(url: str, label: str) -> list:
    """
    Scrape Arlington County's Vendor Registry procurement portal for
    construction-related solicitations.
    """
    print(f"  🔍 Scraping {label}: {url}")
    try:
        resp = requests.get(url, headers=HEADERS, timeout=20)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
    except Exception as e:
        print(f"  ❌ Failed: {e}")
        return []

    results = []
    seen: set = set()

    for row in soup.select("tr"):
        cells = row.find_all("td")
        if len(cells) < 3:
            continue

        title = ""
        link = url
        for cell in cells:
            a = cell.find("a", href=True)
            cell_text = cell.get_text(strip=True)
            if a and len(cell_text) > len(title):
                title = cell_text
                href = a["href"]
                link = href if href.startswith("http") else f"https://vrapp.vendorregistry.com{href}"
            elif not title and len(cell_text) > 10:
                title = cell_text

        if not title or len(title) < 5 or title in seen:
            continue

        detail_text = row.get_text(separator=" ", strip=True)

        if not is_construction_related(title, detail_text):
            continue

        deadline_m = re.search(r"(\d{1,2}/\d{1,2}/\d{2,4})", detail_text)
        desc = detail_text
        if deadline_m:
            desc = f"Deadline: {deadline_m.group(1)}  |  {desc}"

        seen.add(title)
        results.append({
            "title":       title,
            "description": truncate_text(desc or "See Arlington County Vendor Registry for details."),
            "url":         link,
        })

    print(f"  ✅ Found {len(results)} construction solicitations in {label}")
    return results

# ── MAIN SCRAPE ──────────────────────────────────────────────────────────────
def run_scrape():
    print("\n🚀 Starting scrape...\n")
    history = load_history()

    all_data: dict    = {}
    total_count: int  = 0
    new_count: int    = 0

    # Track MWAA solicitation numbers across Current + Upcoming to avoid duplicates
    mwaa_seen_nums: set = set()

    def tag_and_store(label: str, entries: list, dedup_set: set = None):
        nonlocal total_count, new_count
        tagged = []
        for e in entries:
            if dedup_set is not None:
                sol_num = extract_sol_num(e["title"])
                if sol_num in dedup_set:
                    continue
                dedup_set.add(sol_num)
            key    = e["title"]
            is_new = key not in history
            tagged.append({**e, "is_new": is_new})
        all_data[label] = tagged
        total_count    += len(tagged)
        new_count      += sum(1 for e in tagged if e["is_new"])

    # ── MWAA (deduplicate Upcoming vs Current) ──
    mwaa_current_entries = scrape_mwaa(MWAA_SOURCES["MWAA Current"], "MWAA Current")
    tag_and_store("MWAA Current", mwaa_current_entries, mwaa_seen_nums)

    mwaa_upcoming_entries = scrape_mwaa(MWAA_SOURCES["MWAA Upcoming"], "MWAA Upcoming")
    tag_and_store("MWAA Upcoming", mwaa_upcoming_entries, mwaa_seen_nums)

    # ── VDOT ──
    for label, url in VDOT_SOURCES.items():
        entries = scrape_vdot_district(url, label)
        tag_and_store(label, entries)

    # ── Counties ──
    for label, url in COUNTY_SOURCES.items():
        if "fairfax" in url:
            entries = scrape_fairfax(url, label)
        elif "loudoun" in url:
            entries = scrape_loudoun(url, label)
        else:
            entries = []
        tag_and_store(label, entries)

    # ── Arlington County (Playwright) ──
    arlington_entries = scrape_arlington(ARLINGTON_PORTAL, "Arlington County")
    tag_and_store("Arlington County", arlington_entries)

    # ── Persist history ──
    all_keys = {e["title"] for entries in all_data.values() for e in entries}
    save_history(history | all_keys)

    print(f"\n📊 Total: {total_count} | New: {new_count}")
    return all_data, total_count, new_count


# ── PDF GENERATION ────────────────────────────────────────────────────────────
def build_pdf(all_data: dict, total_count: int, new_count: int) -> str:
    today      = datetime.now().strftime("%m-%d-%Y")
    pdf_path   = PDF_FILE.format(date=today)
    now_str    = datetime.now().strftime("%B %d, %Y")
    date_str   = datetime.now().strftime("%m/%d/%Y")

    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=letter,
        leftMargin=0.75 * inch,
        rightMargin=0.75 * inch,
        topMargin=0.75 * inch,
        bottomMargin=0.75 * inch,
    )

    # ── Styles ──
    s_title = ParagraphStyle(
        "ReportTitle",
        fontSize=24, fontName="Helvetica-Bold",
        textColor=colors.HexColor("#1A2E4A"),
        leading=30, spaceAfter=7, alignment=TA_CENTER,
    )
    s_subtitle = ParagraphStyle(
        "Subtitle",
        fontSize=10, fontName="Helvetica",
        textColor=colors.HexColor("#555555"),
        leading=14, spaceAfter=6, alignment=TA_CENTER,
    )
    s_section = ParagraphStyle(
        "SectionHeader",
        fontSize=13, fontName="Helvetica-Bold",
        textColor=colors.white,
        spaceAfter=0, spaceBefore=0, leftIndent=8,
    )
    s_proj_title = ParagraphStyle(
        "ProjectTitle",
        fontSize=10, fontName="Helvetica-Bold",
        textColor=colors.HexColor("#1A2E4A"),
        spaceAfter=3,
    )
    s_desc = ParagraphStyle(
        "Description",
        fontSize=9, fontName="Helvetica",
        textColor=colors.HexColor("#333333"),
        spaceAfter=4, leading=14,
    )
    s_link_label = ParagraphStyle(
        "LinkLabel",
        fontSize=8, fontName="Helvetica-Bold",
        textColor=colors.HexColor("#555555"),
        spaceAfter=1,
    )
    s_link = ParagraphStyle(
        "Link",
        fontSize=8, fontName="Helvetica-Oblique",
        textColor=colors.HexColor("#0066CC"),
        spaceAfter=6, wordWrap="LTR",
    )
    # ── Section header colours ──
    section_colors = {
        "MWAA Current":          "#1A2E4A",
        "MWAA Upcoming":         "#2E6DA4",
        "VDOT Northern Virginia":"#2E7D32",
        "Fairfax County":        "#6A1B9A",
        "Loudoun County":        "#B71C1C",
        "Arlington County":      "#00695C",   # teal
    }

    story = []

    # ── Header ──
    story.append(Spacer(1, 0.1 * inch))
    story.append(Paragraph("Crisdel Group, Inc.", s_title))
    story.append(Paragraph("Virginia Public Construction Opportunities", s_subtitle))
    story.append(Paragraph(f"Generated: {now_str}", s_subtitle))
    story.append(Spacer(1, 0.1 * inch))

    # ── Summary banner ──
    def sum_style(color="#1A2E4A"):
        return ParagraphStyle(
            f"SC_{color}", fontSize=11, fontName="Helvetica-Bold",
            textColor=colors.HexColor(color), alignment=TA_CENTER,
        )

    active_sections = sum(1 for v in all_data.values() if v)
    summary_data = [[
        Paragraph(f"<b>Total Opportunities</b><br/>{total_count}", sum_style()),
        Paragraph(f"<b>New Today</b><br/>{new_count}",             sum_style("#C62828")),
        Paragraph(f"<b>Active Sources</b><br/>{active_sections}",  sum_style()),
        Paragraph(f"<b>Report Date</b><br/>{date_str}",            sum_style()),
    ]]
    summary_tbl = Table(summary_data, colWidths=[1.6 * inch] * 4)
    summary_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F0F4F8")),
        ("BACKGROUND", (1, 0), (1,  0),  colors.HexColor("#FFF0F0")),
        ("BOX",        (0, 0), (-1, -1), 1,   colors.HexColor("#CCCCCC")),
        ("INNERGRID",  (0, 0), (-1, -1), 0.5, colors.HexColor("#DDDDDD")),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ("VALIGN",     (0, 0), (-1, -1), "MIDDLE"),
    ]))
    story.append(summary_tbl)
    story.append(Spacer(1, 0.2 * inch))
    story.append(HRFlowable(width="100%", thickness=1.5,
                            color=colors.HexColor("#1A2E4A"), spaceAfter=14))

    # ── Sections ──
    for section, entries in all_data.items():
        if not entries:
            continue

        sec_color = section_colors.get(section, "#333333")

        # Section header bar
        hdr_data = [[Paragraph(
            f"  {section}  —  {len(entries)} opportunit{'y' if len(entries)==1 else 'ies'}",
            s_section,
        )]]
        hdr_tbl  = Table(hdr_data, colWidths=[7.0 * inch])
        hdr_tbl.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor(sec_color)),
            ("TOPPADDING", (0, 0), (-1, -1), 8),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ("VALIGN",     (0, 0), (-1, -1), "MIDDLE"),
        ]))
        story.append(hdr_tbl)
        story.append(Spacer(1, 0.06 * inch))

        # Entry rows
        for i, entry in enumerate(entries):
            row_bg = "#FFFFFF" if i % 2 == 0 else "#F7F9FC"
            badge  = ' <font color="#C62828"><b>[NEW]</b></font>' if entry.get("is_new") else ""

            title_para = Paragraph(f"{_xml_escape(entry['title'])}{badge}", s_proj_title)

            desc_text = _xml_escape(entry.get("description", "").strip())
            desc_para = Paragraph(desc_text, s_desc) if desc_text else Paragraph("No description available.", s_desc)

            link_url   = entry.get("url", "")
            safe_url   = _xml_escape(link_url)
            link_label = Paragraph("View Opportunity:", s_link_label)
            link_para  = Paragraph(
                f"<a href='{safe_url}' color='#0066CC'>{safe_url}</a>" if link_url else "Link unavailable",
                s_link,
            )

            cell_content = [title_para, desc_para, link_label, link_para]

            row_data = [[cell_content]]
            row_tbl  = Table(row_data, colWidths=[6.8 * inch])
            row_tbl.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor(row_bg)),
                ("BOX",        (0, 0), (-1, -1), 0.25, colors.HexColor("#DDDDDD")),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
                ("VALIGN",     (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING",  (0, 0), (-1, -1), 10),
                ("RIGHTPADDING", (0, 0), (-1, -1), 10),
            ]))
            story.append(row_tbl)
            story.append(Spacer(1, 0.04 * inch))

        story.append(Spacer(1, 0.15 * inch))

    # ── Footer ──
    story.append(HRFlowable(width="100%", thickness=1,
                            color=colors.HexColor("#CCCCCC"), spaceBefore=6))
    story.append(Paragraph(
        "Auto-generated by the Crisdel Opportunity Tracker · "
        "Sources: MWAA, VDOT Northern Virginia, Fairfax County, Loudoun County and Arlington County",
        ParagraphStyle("Footer", fontSize=7, fontName="Helvetica",
                       textColor=colors.HexColor("#999999"), alignment=TA_CENTER),
    ))

    doc.build(story)
    print(f"  ✅ PDF saved: {pdf_path}")
    return pdf_path

# ── EMAIL ───────────────────────────────────────────────────────────────────────────
def send_email(pdf_path: str, all_data: dict, total_count: int, new_count: int) -> None:
    # Authenticate with Azure/Microsoft Graph
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    scope = ["https://graph.microsoft.com/.default"]
   
    app = msal.ConfidentialClientApplication(
        client_id=CLIENT_ID,
        client_credential=CLIENT_SECRET,
        authority=authority
    )
   
    token_result = app.acquire_token_for_client(scopes=scope)
   
    if "access_token" not in token_result:
        print(f"  ❌ Token Error: {token_result.get('error')} - {token_result.get('error_description')}")
        return
   
    access_token = token_result["access_token"]
   
    # Build email subject
    subject = (
        f"VA Public Bid Opportunities — "
        f"{datetime.now().strftime('%B %d, %Y')}"
        f"({total_count} total, {new_count} new)"
    )
   
    # Build a section-by-section summary for the email body
    section_lines = []
    for section, entries in all_data.items():
        if entries:
            new_in_section = sum(1 for e in entries if e.get("is_new"))
            line = f"  • {section}: {len(entries)} item(s)"
            if new_in_section:
                line += f"  ({new_in_section} NEW)"
            section_lines.append(line)

    # Build a detailed listing of just the NEW opportunities (title, description, link)
    new_items_blocks = []
    for section, entries in all_data.items():
        new_entries = [e for e in entries if e.get("is_new")]
        if not new_entries:
            continue
        new_items_blocks.append(f"\n{section}\n{'-' * len(section)}")
        for e in new_entries:
            title = e.get("title", "").strip()
            desc  = e.get("description", "").strip() or "No description available."
            link  = e.get("url", "").strip() or "Link unavailable"
            new_items_blocks.append(
                f"\n  {title}\n"
                f"    {desc}\n"
                f"    Link: {link}\n"
            )

    if new_items_blocks:
        new_items_section = (
            "NEW OPPORTUNITIES TODAY\n"
            "========================\n"
            + "".join(new_items_blocks)
        )
    else:
        new_items_section = "NEW OPPORTUNITIES TODAY\n========================\n  (No new opportunities today.)\n"

    body = f"""Hello,

Please find attached the daily Virginia Public Construction Opportunities Report for Crisdel Group, Inc.

SUMMARY
  Total Opportunities: {total_count}
  New Today: {new_count}
  Report Date: {datetime.now().strftime('%B %d, %Y')}

BREAKDOWN BY SOURCE
{chr(10).join(section_lines)}

{new_items_section}
Sources: MWAA, VDOT Northern Virginia, Fairfax County, Loudoun County, and Arlington County.

This is an automated daily report. Please do not reply to this email.

— Crisdel Opportunity Tracker
"""
   
    # Read PDF for attachment
    with open(pdf_path, "rb") as f:
        pdf_content = f.read()
   
    # Encode PDF as base64 for attachment
    pdf_base64 = base64.b64encode(pdf_content).decode('utf-8')
   
    # Prepare Graph API request
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
   
    all_recipients = EMAIL_TO + EMAIL_CC
    to_recipients = []
    cc_recipients = []
   
    for email in EMAIL_TO:
        to_recipients.append({"emailAddress": {"address": email}})
    for email in EMAIL_CC:
        cc_recipients.append({"emailAddress": {"address": email}})
   
    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "Text", "content": body},
            "toRecipients": to_recipients,
            "ccRecipients": cc_recipients,
            "attachments": [
                {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": os.path.basename(pdf_path),
                    "contentBytes": pdf_base64
                }
            ]
        },
        "saveToSentItems": "true"
    }
   
    try:
        response = requests.post(
            f"https://graph.microsoft.com/v1.0/users/{EMAIL_SENDER}/sendMail",
            headers=headers,
            json=payload
        )
       
        if response.status_code in [200, 202]:
            print(f"  ✅ Email sent to: {', '.join(all_recipients)}")
        else:
            print(f"  ❌ Email failed: Status {response.status_code} - {response.text}")
    except Exception as e:
        print(f"  ❌ Email failed: {e}")


# ── JOB ─────────────────────────────────────────────────────────────────────────────
def job():
    all_data, total_count, new_count = run_scrape()
    pdf_path = build_pdf(all_data, total_count, new_count)
    send_email(pdf_path, all_data, total_count, new_count)
    print("\n✅ Job complete.\n")


# ── ENTRY POINT ─────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    job()  
