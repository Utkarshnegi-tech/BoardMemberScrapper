import os
import re
import time
import pandas as pd
from urllib.parse import urlparse, unquote
from seleniumbase import Driver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

#NLP model
try:
    import spacy
    try:
        _nlp = spacy.load("en_core_web_sm")
    except OSError:
        import subprocess, sys
        subprocess.run(
            [sys.executable, "-m", "pip", "install", "spacy", "--system"],
            check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
        )
        subprocess.run(
            [sys.executable, "-m", "spacy", "download", "en_core_web_sm"],
            check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
        )
        _nlp = spacy.load("en_core_web_sm")
    _SPACY_AVAILABLE = True
except Exception:
    _nlp = None
    _SPACY_AVAILABLE = False

#Config
import sys

_BASE_DIR  = os.path.dirname(os.path.abspath(sys.argv[0]))
file_path  = os.path.join(_BASE_DIR, "malmad-title-list-9-january-2023.xlsx")
column_name    = "title"
how_many       = 1
preferred_site = "tandfonline.com"

df = pd.read_excel(file_path)
df.columns = df.columns.str.strip().str.lower()

driver = Driver(uc=True, headed=True)
driver.set_page_load_timeout(20)
wait = WebDriverWait(driver, 10)

all_data         = []
current_title    = ""
current_acronym  = ""
_member_id_counter = 0

CSV_COLUMNS = [
    "EdBoardMemberID", "journal_title", "journal_acronym",
    "section", "role",
    "first_name", "last_name", "contact_title",
    "affiliation", "country", "notes"
]

# Compiled patterns
_CONTACT_SKIP = re.compile(
    r"^\s*(phone|tel|fax|mobile|email|e-mail|contact|address|website|url|linkedin|twitter|orcid)\s*[:\-]",
    re.IGNORECASE
)
_HONORIFICS = re.compile(
    r"^(Dr\.?|Prof\.?|Professor|Mr\.?|Mrs\.?|Ms\.?|Sir|Dame)\s+",
    re.IGNORECASE
)

# Known countries set
_KNOWN_COUNTRIES = {
    "Afghanistan","Albania","Algeria","Argentina","Armenia","Australia","Austria",
    "Azerbaijan","Bangladesh","Belarus","Belgium","Bolivia","Bosnia","Brazil",
    "Bulgaria","Cambodia","Cameroon","Canada","Chile","China","Colombia","Croatia",
    "Cuba","Cyprus","Czech Republic","Denmark","Ecuador","Egypt","Estonia",
    "Ethiopia","Finland","France","Georgia","Germany","Ghana","Greece","Guatemala",
    "Hungary","India","Indonesia","Iran","Iraq","Ireland","Israel","Italy",
    "Jamaica","Japan","Jordan","Kazakhstan","Kenya","Latvia","Lebanon","Lithuania",
    "Luxembourg","Malaysia","Mexico","Moldova","Morocco","Netherlands","New Zealand",
    "Nigeria","Norway","Pakistan","Palestine","Peru","Philippines","Poland",
    "Portugal","Romania","Russia","Saudi Arabia","Serbia","Singapore","Slovakia",
    "Slovenia","South Africa","South Korea","Spain","Sri Lanka","Sweden",
    "Switzerland","Syria","Taiwan","Tanzania","Thailand","Tunisia","Turkey",
    "UAE","Uganda","UK","Ukraine","United Arab Emirates","United Kingdom",
    "United States","United States of America","USA","Uruguay","Uzbekistan",
    "Venezuela","Vietnam","Zimbabwe",
}
_KNOWN_COUNTRIES_LOWER = {c.lower() for c in _KNOWN_COUNTRIES}

def _is_country(text):
    """Return True only if text is a recognised country name."""
    return text.strip().lower() in _KNOWN_COUNTRIES_LOWER


def wait_for_cloudflare():
    try:
        WebDriverWait(driver, 20).until(lambda d: "just a moment" not in d.page_source.lower())
    except:
        pass


def split_name(name):
    """Return (first_name, last_name, contact_title)."""
    m = _HONORIFICS.match(name)
    contact_title = m.group(0).strip().rstrip(".").strip() if m else None
    clean = _HONORIFICS.sub("", name).strip()
    parts = clean.split()
    if len(parts) > 1:
        return parts[0], " ".join(parts[1:]), contact_title
    return clean, None, contact_title


def extract_country(aff):
    if not aff:
        return None, aff
    if "," in aff:
        parts = [p.strip() for p in aff.split(",")]
        last  = parts[-1]
       
        if _is_country(last):
            return last, ",".join(parts[:-1]).strip()
    return None, aff


def record(section, role, name, aff):
    global _member_id_counter
    _member_id_counter += 1
    country, aff = extract_country(aff)
    fn, ln, title = split_name(name)
    all_data.append({
        "EdBoardMemberID": _member_id_counter,
        "journal_title":   current_title,
        "journal_acronym": current_acronym,
        "section":         section,
        "role":            role,
        "first_name":      fn,
        "last_name":       ln,
        "contact_title":   title,
        "affiliation":     aff,
        "country":         country,
        "notes":           None
    })
    print(f"+ {fn} {ln} | {role}")

#Main Extraction Logic
def _looks_like_role_heading(line):
    words = line.split()

    if len(words) > 7:
        return False
    if any(sep in line for sep in ["–", " - "]):
        return False
    if "," in line:
        return False
    role_keywords = [
        "editor", "board", "chair", "director", "advisor", "adviser",
        "reviewer", "committee", "managing", "founding", "honorary",
        "emeritus", "consultant",
    ]
    low = line.lower()
    if not any(kw in low for kw in role_keywords):
        return False
    if 2 <= len(words) <= 4 and all(w[0].isupper() for w in words if w):
        return False
    return True


def _is_person_name(text):
    """
    Return True if text looks like a person name.
    Uses spaCy NER as a hint but falls back to heuristics so hyphenated
    names, names with middle initials, and non-English names aren't dropped.
    """
    text = text.strip()
    if not text:
        return False

    # Reject if starts with a lowercase word (e.g. "in The Autobiography Society")
    if text[0].islower():
        return False

    # Reject known role/org words
    _REJECT_WORDS = {
        "editor", "editors", "board", "committee", "chair", "director",
        "advisor", "adviser", "reviewer", "managing", "associate", "senior",
        "founding", "honorary", "emeritus", "consultant", "assistant", "intern",
        "interns", "editorial", "digital", "content", "review", "book",
        # org/society words
        "society", "association", "institute", "foundation", "center", "centre",
        "network", "group", "council", "academy", "press", "journal", "about",
        "membership", "information", "subscribe", "submission", "contact",
        "copyright", "aims", "scope", "news", "announcement",
    }
    words_lower = {w.lower().rstrip("s") for w in text.split()}
    if words_lower & _REJECT_WORDS:
        return False

    # Reject if it contains "The" followed by a multi-word proper noun (org pattern)
    if re.match(r"^The\s+[A-Z]", text) and len(text.split()) > 3:
        return False

    words = text.split()
    if 2 <= len(words) <= 5 and not re.search(r"\d", text):
        cap_words = [w for w in words if w[0].isupper()]
        if len(cap_words) >= len(words) - 1:   # allow at most one lowercase particle
            return True

    # spaCy as a secondary check
    if _SPACY_AVAILABLE:
        doc = _nlp(text)
        if any(ent.label_ == "PERSON" for ent in doc.ents):
            return True

    return False


def _parse_member_line(line, section, role):
    if len(line) < 4:
        return
    if re.match(r"^\d", line):
        return
    if "@" in line:
        return
    if line.lower().startswith("http"):
        return
    if _CONTACT_SKIP.match(line):
        return

    # Reject obvious non-person lines
    _NAV_SKIP = re.compile(
        r"^(about|membership|subscribe|submission|contact|copyright|aims|scope"
        r"|news|announcement|the\s+\w+\s+(society|association|institute|foundation"
        r"|network|group|council|academy|press|journal))",
        re.IGNORECASE
    )
    if _NAV_SKIP.match(line):
        return

    role_prefixes = [
        "editorial board", "editors", "associate editors", "managing editors",
        "consulting editors", "editorial assistant", "editorial interns",
        "book review editor", "digital content editor", "editor-in-chief",
        "senior editors", "advisory board", "founding editors"
    ]
    line_lower = line.lower()
    for prefix in role_prefixes:
        if line_lower.startswith(prefix):
            line = line[len(prefix):].strip()
            break

    name        = None
    affiliation = None

    for sep in [" – ", "–", " - "]:
        if sep in line:
            parts       = line.split(sep, 1)
            name        = parts[0].strip()
            affiliation = parts[1].strip() if len(parts) > 1 else None
            break

    if name is None:
        parts     = line.split(",", 1)
        candidate = parts[0].strip()
        if 2 <= len(candidate.split()) <= 5 and not re.search(r"\d", candidate):
            name = candidate
            affiliation = parts[1].strip() if len(parts) > 1 else None
        else:
            return
        
    if not name or len(name.split()) < 2:
        return
    if re.search(r"\d", name):
        return

    # NLP guard — only record if spaCy agrees this looks like a person name
    clean_name = _HONORIFICS.sub("", name).strip()
    if not _is_person_name(clean_name):
        return

    record(section, role, name, affiliation)


def _split_into_entries(text):
    """
    Split a block of text into individual person entries.
    Handles cases where entries are separated by newlines OR where multiple
    entries are concatenated without newlines (e.g. "...Storrs G. Thomas Couser – Hofstra...").
    Strategy: split on the em-dash separator, then reconstruct name–affiliation pairs.
    """
    lines = [l.strip() for l in text.split("\n") if l.strip()]

    result = []
    for line in lines:
        parts_by_emdash = line.split(" – ")
        if len(parts_by_emdash) <= 2:
            result.append(line)
        else:
            i = 0
            while i < len(parts_by_emdash) - 1:
                name_part  = parts_by_emdash[i].strip()
                affil_part = parts_by_emdash[i + 1].strip()

                split_match = re.search(
                    r'([a-z\-])([A-Z][a-z]*\.?\s+[A-Z])',
                    affil_part
                )
                if split_match:
                    split_pos = split_match.start(2)
                    real_affil = affil_part[:split_pos].strip()
                    parts_by_emdash[i + 1] = affil_part[split_pos:].strip()
                    result.append(f"{name_part} – {real_affil}")
                else:
                    result.append(f"{name_part} – {affil_part}")
                    i += 1
                i += 1

    return result


def extract_editorial_board():
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight)")
    time.sleep(2)

    count_before = len(all_data)

    try:
        headers = driver.find_elements(
            By.XPATH,
            "//*[self::h2 or self::h3][contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', "
            "'abcdefghijklmnopqrstuvwxyz'), 'editorial')]"
        )

        if not headers:
            return False

        for header in headers:
            section = header.text.strip() or "Editorial Board"
            current_role = section

            siblings = header.find_elements(By.XPATH, "following-sibling::*")

            for el in siblings:
                tag = el.tag_name.lower()

                # Treat all h2/h3 as role/sub-section headings — don't stop
                # when another editorial header is found, since each header
                # already handles its own section in the outer loop.
                if tag in ["h2", "h3"]:
                    role_candidate = el.text.strip()
                    if role_candidate:
                        current_role = role_candidate
                    continue

                if tag in ["h4", "h5", "strong", "b"]:
                    role_candidate = el.text.strip()
                    if role_candidate:
                        current_role = role_candidate
                    continue

                text = el.text.strip()
                if not text:
                    continue

                # Use the new splitting function to handle concatenated entries
                entries = _split_into_entries(text)
                for line in entries:
                    if _looks_like_role_heading(line):
                        current_role = line
                        continue
                    _parse_member_line(line, section, current_role)

    except Exception as e:
        print("Heading-sibling strategy error:", e)

    if len(all_data) > count_before:
        return True

    _extract_generic_text()
    return len(all_data) > count_before


def _extract_generic_text():
    try:
        body      = driver.find_element(By.TAG_NAME, "body")
        full_text = body.text

        in_editorial = False
        current_role = "Editorial Board"
        section      = "Editorial Board"
        seen_names   = set()

        for raw_line in full_text.split("\n"):
            raw_line = raw_line.strip()
            if not raw_line:
                continue

            low = raw_line.lower()

            if "editorial board" in low or "editors" in low:
                in_editorial = True
                section      = raw_line
                current_role = raw_line
                continue

            if not in_editorial:
                continue

            if any(kw in low for kw in ["aims and scope", "submit", "subscribe", "copyright"]):
                in_editorial = False
                continue

            # Split concatenated entries on this line before processing
            sub_entries = _split_into_entries(raw_line)
            for line in sub_entries:
                if _looks_like_role_heading(line):
                    current_role = line
                    continue

                # deduplicate: skip if we've already recorded this line
                if line in seen_names:
                    continue
                seen_names.add(line)

                _parse_member_line(line, section, current_role)

    except Exception as e:
        print("Generic text strategy error:", e)


def _wait_if_captcha(drv):
    """If a CAPTCHA/bot-check is detected, pause until the user solves it."""
    captcha_signals = [
        "i am not a robot", "are you a robot",
        "verify you are human", "just a moment",
        "checking your browser", "enable javascript and cookies",
    ]
    # Also detect Google's /sorry page (reCAPTCHA checkbox)
    try:
        current_url = drv.current_url.lower()
    except Exception:
        current_url = ""

    url_is_captcha = "google.com/sorry" in current_url or "recaptcha" in current_url

    try:
        body_text = drv.find_element(By.TAG_NAME, "body").text.lower()
    except Exception:
        body_text = ""

    text_is_captcha = any(sig in body_text for sig in captcha_signals)

    if not url_is_captcha and not text_is_captcha:
        return

    print("\n⚠️  CAPTCHA / bot-check detected! Please solve it in the browser window.")
    attempts = 0
    while attempts < 20:
        input("    Press Enter once you've solved it to resume...")
        try:
            new_url  = drv.current_url.lower()
            body_now = drv.find_element(By.TAG_NAME, "body").text.lower()
        except Exception:
            break
        still_captcha = (
            "google.com/sorry" in new_url
            or "recaptcha" in new_url
            or any(sig in body_now for sig in captcha_signals)
        )
        if not still_captcha:
            print("   ✅ CAPTCHA cleared, resuming...")
            break
        attempts += 1
        print("   ⚠️  Still detecting a bot check — please finish solving it.")


def _safe_get(drv, url, timeout=20):
    """Navigate to url, return True on success, False on timeout/error."""
    old_timeout = None
    try:
        old_timeout = drv.timeouts.page_load
    except Exception:
        pass
    try:
        drv.set_page_load_timeout(timeout)
        drv.get(url)
        _wait_if_captcha(drv)
        return True
    except Exception as e:
        print(f"  Page load failed ({url}): {type(e).__name__}")
        try:
            drv.execute_script("window.stop();")
        except Exception:
            pass
        # still check for captcha on partial load
        _wait_if_captcha(drv)
        return False
    finally:
        try:
            if old_timeout is not None:
                drv.set_page_load_timeout(old_timeout)
        except Exception:
            pass


def _cleanup_extra_tabs(drv, keep_handle=None):
    """Close all tabs except keep_handle (defaults to first handle)."""
    try:
        handles = drv.window_handles
        if not handles:
            return
        target = keep_handle or handles[0]
        for h in handles:
            if h != target:
                try:
                    drv.switch_to.window(h)
                    drv.close()
                except Exception:
                    pass
        drv.switch_to.window(target)
    except Exception:
        pass


#navigation function
def open_about_page():
    try:
        current_url = driver.current_url

        # --- TandF-specific: build the editorial board URL directly from the journal code ---
        # Patterns: /journals/XXXX20  or  /toc/XXXX20  or  /loi/XXXX20
        tandf_match = re.search(
            r"tandfonline\.com/(?:journals|toc|loi)/([a-z0-9]+)",
            current_url, re.IGNORECASE
        )
        if tandf_match:
            journal_code = tandf_match.group(1)
            # Try the direct editorial board URL first
            eb_url = f"https://www.tandfonline.com/action/journalInformation?show=editorialBoard&journalCode={journal_code}"
            print(f"  Navigating to editorial board: {eb_url}")
            if _safe_get(driver, eb_url):
                time.sleep(3)
                wait_for_cloudflare()
                if extract_editorial_board():
                    return
            # Fallback: about-this-journal page
            about_url = f"https://www.tandfonline.com/journals/{journal_code}/about-this-journal#editorial-board"
            print(f"  Trying about page: {about_url}")
            if _safe_get(driver, about_url):
                time.sleep(3)
                wait_for_cloudflare()
                if extract_editorial_board():
                    return

        # --- Generic fallback: scan page links ---
        links = driver.find_elements(By.XPATH, "//a")

        def _rank(link):
            href = (link.get_attribute("href") or "").lower()
            text = link.text.lower()
            if "editorialboard" in href or "editorial-board" in href or "editorial_board" in href:
                return 0
            if "about-this-journal" in href or "about_this_journal" in href:
                return 1
            if "editorial" in text and "board" in text:
                return 2
            if "about" in text and "journal" in text:
                return 3
            if "about" in text:
                return 4
            return 99

        candidates = sorted(
            [l for l in links if _rank(l) < 99],
            key=_rank
        )

        if not candidates:
            print("  About/editorial board link not found, trying extract on current page.")
            extract_editorial_board()
            return

        best = candidates[0]
        href = best.get_attribute("href") or ""
        if href:
            if not _safe_get(driver, href):
                print("  About page timed out, trying extract on current page.")
        else:
            try:
                driver.execute_script("arguments[0].click();", best)
            except Exception:
                pass

        time.sleep(3)
        wait_for_cloudflare()
        extract_editorial_board()

    except Exception as e:
        print("About page error:", e)


def search_and_extract(query):
    global current_acronym
    driver.delete_all_cookies()

    if not _safe_get(driver, "https://www.google.com"):
        print("  Could not load Google, skipping.")
        return

    try:
        btn = WebDriverWait(driver, 8).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Accept')]"))
        )
        try:
            btn.click()
        except Exception:
            driver.execute_script("arguments[0].click();", btn)
        time.sleep(1)
    except Exception:
        pass

    try:
        search_box = wait.until(EC.element_to_be_clickable((By.NAME, "q")))
    except Exception:
        print("  Google search box not found, skipping.")
        return

    try:
        search_box.click()
        time.sleep(0.3)
        search_box.clear()
        search_box.send_keys(f'"{query}" site:{preferred_site}')
        search_box.send_keys(Keys.RETURN)
    except Exception:
        try:
            driver.execute_script(
                "arguments[0].value = arguments[1];", search_box,
                f'"{query}" site:{preferred_site}'
            )
            search_box.send_keys(Keys.RETURN)
        except Exception as e:
            print(f"  Could not interact with search box: {e}, skipping.")
            return

    try:
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "h3")))
    except Exception:
        print("  No search results loaded, skipping.")
        return
    time.sleep(2)

    results = driver.find_elements(By.CSS_SELECTOR, "h3")
    if not results:
        print("  No search results found")
        return

    link = results[0].find_element(By.XPATH, "..").get_attribute("href")
    if not link:
        print("  Could not get result URL")
        return

    main_handle = driver.window_handles[0]
    try:
        driver.execute_script("window.open(arguments[0]);", link)
        driver.switch_to.window(driver.window_handles[-1])
    except Exception as e:
        print(f"  Could not open result tab: {e}")
        _cleanup_extra_tabs(driver, main_handle)
        return

    # wait for page to settle (don't hard-fail on timeout)
    try:
        WebDriverWait(driver, 15).until(lambda d: d.execute_script("return document.readyState") == "complete")
    except Exception:
        try:
            driver.execute_script("window.stop();")
        except Exception:
            pass

    print("  Opened:", driver.current_url)

    _acr_match = re.search(r"/(?:toc|about|loi)/([A-Z]{2,8})\d*", driver.current_url, re.IGNORECASE)
    if _acr_match:
        current_acronym = _acr_match.group(1).upper()
    else:
        _path = urlparse(driver.current_url).path.rstrip("/")
        _seg  = _path.split("/")[-1]
        current_acronym = re.sub(r"\d.*", "", _seg).upper() or ""

    time.sleep(3)
    wait_for_cloudflare()

    try:
        open_about_page()
    except Exception as e:
        print(f"  open_about_page error: {e}")

    time.sleep(2)
    _cleanup_extra_tabs(driver, main_handle)


#Main processing loop

for i in range(min(how_many, len(df))):
    current_title   = str(df.iloc[i].get(column_name, "")).strip()
    current_acronym = ""   # will be set once we land on the TandF page

    if not current_title or current_title.lower() == "nan":
        continue

    print(f"\n[{i+1}] Processing: {current_title}")

    try:
        search_and_extract(current_title)
        time.sleep(3)
    except Exception as e:
        print("Row error:", e)

output_csv  = os.path.join(_BASE_DIR, "editorial_board_output.csv")
output_path = os.path.abspath(output_csv)

#saving the data
df_out = pd.DataFrame(all_data, columns=CSV_COLUMNS).fillna("null")
df_out.to_csv(output_csv, index=False, encoding="utf-8-sig")

print("\n==============================")
print("Total records:", len(all_data))
print("Saved CSV at:", output_path)
print("==============================")

try:
    driver.quit()
except Exception:
    pass

#----------------------------------------------------------------------------------------------------------------------------------------


# Google person search + page scraping
PERSON_CSV_COLUMNS = [
    "EdBoardMemberID", "journal_title", "journal_acronym",
    "first_name", "last_name", "contact_title", "orcid",
    "institution", "city", "state", "country", "email"
]


_RE_EMAIL = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")
_RE_ORCID = re.compile(r"\b(\d{4}-\d{4}-\d{4}-\d{3}[\dX])\b")


def _is_clean_place(text):
    """A valid city/state/country token: short, no digits, no long sentences."""
    t = text.strip()
    if not t or len(t) > 60:
        return False
    if re.search(r"\d", t):
        return False
    if len(t.split()) > 6:          # more than 6 words → probably not a place name
        return False
    return True
_RE_ZIP_US2   = re.compile(r"\b(\d{5})(?:-\d{4})?\s+([A-Z]{2})\b")           
_RE_ZIP_CA    = re.compile(r"\b([A-Z]\d[A-Z])\s*(\d[A-Z]\d)\b")              
_RE_ZIP_UK    = re.compile(r"\b([A-Z]{1,2}\d{1,2}[A-Z]?\s*\d[A-Z]{2})\b")   
_RE_ZIP_GEN   = re.compile(r"\b\d{4,6}\b")                                   
_US_STATES = {
    "AL","AK","AZ","AR","CA","CO","CT","DE","FL","GA","HI","ID","IL","IN","IA",
    "KS","KY","LA","ME","MD","MA","MI","MN","MS","MO","MT","NE","NV","NH","NJ",
    "NM","NY","NC","ND","OH","OK","OR","PA","RI","SC","SD","TN","TX","UT","VT",
    "VA","WA","WV","WI","WY","DC"
}


def _parse_location_from_text(text):
    """
    Try to extract city, state, country from a short address-like text snippet.
    Rejects long blobs of page text. Returns dict: city, state, country (all may be None).
    """
    loc = {"city": None, "state": None, "country": None}

    # reject page blobs — only process short address-like lines
    if len(text) > 200 or len(text.split()) > 30:
        return loc

    # --- US ZIP: "City, ST 12345" or "City, ST 12345-6789" ---
    m = re.search(r"([A-Za-z\s\-]+),\s*([A-Z]{2})\s+\d{5}(?:-\d{4})?", text)
    if m:
        city  = m.group(1).strip()
        state = m.group(2).strip()
        if _is_clean_place(city) and state in _US_STATES:
            loc["city"]    = city
            loc["state"]   = state
            loc["country"] = "USA"
            return loc

    # --- Canadian postal: "City, Province A1A 1A1" ---
    m = re.search(r"([A-Za-z\s\-]+),\s*([A-Za-z\s]+)\s+[A-Z]\d[A-Z]\s*\d[A-Z]\d", text)
    if m:
        city  = m.group(1).strip()
        prov  = m.group(2).strip()
        if _is_clean_place(city):
            loc["city"]    = city
            loc["state"]   = prov
            loc["country"] = "Canada"
            return loc

    # --- UK postcode ---
    m = re.search(r"([A-Za-z\s\-]+),?\s+([A-Z]{1,2}\d{1,2}[A-Z]?\s*\d[A-Z]{2})\b", text)
    if m:
        city = m.group(1).strip()
        if _is_clean_place(city):
            loc["city"]    = city
            loc["country"] = "UK"
            return loc

    # --- Generic: "City, State, Country" — only if country is recognised ---
    parts = [p.strip() for p in text.split(",")]
    parts = [p for p in parts if p and not _RE_ZIP_GEN.fullmatch(p) and _is_clean_place(p)]
    if len(parts) >= 3 and _is_country(parts[-1]):
        loc["city"]    = parts[0]
        loc["state"]   = parts[1]
        loc["country"] = parts[-1]
        return loc
    if len(parts) == 2 and _is_country(parts[-1]):
        loc["city"]    = parts[0]
        loc["country"] = parts[-1]
        return loc
    # single token that is a known country
    if len(parts) == 1 and _is_country(parts[0]):
        loc["country"] = parts[0]
        return loc

    # --- NLP GPE fallback — only accept if GPE is a known country or short place ---
    if _SPACY_AVAILABLE:
        doc  = _nlp(text)
        gpes = [ent.text for ent in doc.ents if ent.label_ == "GPE" and _is_clean_place(ent.text)]
        countries = [g for g in gpes if _is_country(g)]
        cities    = [g for g in gpes if not _is_country(g)]
        if countries:
            loc["country"] = countries[-1]
        if cities:
            loc["city"] = cities[0]

    return loc


_UNI_TLD_RE = re.compile(
    r"\.edu(/|$)"
    r"|\.ac\.[a-z]{2}(/|$)"
    r"|\.edu\.[a-z]{2}(/|$)"
    , re.IGNORECASE
)

_UNI_DOMAIN_RE = re.compile(
    r"(university|univer|univers|univ[^e])"
    r"|(institute|instit)"
    r"|(college)"
    r"|(hospital)"
    r"|(academia)"
    r"|(polytechnic|polytech)"
    r"|(faculty)"
    , re.IGNORECASE
)

# Matches domains like ualberta.ca, utoronto.ca, ubc.ca, nyu.edu, mit.edu etc.
_UNI_ABBREV_RE = re.compile(
    r"^(u[a-z]{2,}|[a-z]{2,}u)\.(ca|edu|ac\.[a-z]{2}|edu\.[a-z]{2})$",
    re.IGNORECASE
)


def _is_university_url(url):
  
  
    try:
  
  
        parsed = urlparse(url)
        domain = parsed.netloc.lower()
        full   = url.lower()
    except Exception:
        domain = url.lower()
        full   = url.lower()

    # strip subdomains to get root domain for abbreviation check
    root_domain = ".".join(domain.replace("www.", "").split(".")[-3:])

    if _UNI_TLD_RE.search(domain):
        return True

   
   
    if _UNI_DOMAIN_RE.search(domain):
        return True



    parts = domain.replace("www.", "").split(".")
    for part in parts:
        if _UNI_DOMAIN_RE.search(part):
            return True

    # match university abbreviation domains: ualberta.ca, utoronto.ca, ubc.ca etc.
    if _UNI_ABBREV_RE.match(root_domain):
        return True

    # match subdomains of known university abbreviation domains
    # e.g. apps.ualberta.ca → ualberta.ca
    if len(parts) >= 3:
        parent = ".".join(parts[-3:])
        if _UNI_ABBREV_RE.match(parent) or _UNI_ABBREV_RE.match(".".join(parts[-2:])):
            return True

    academic_path_kw = ["faculty", "staff", "research", "department", "dept", "profile", "people", "directory"]
    if any(kw in full for kw in academic_path_kw):
  
  
        non_academic = ["linkedin", "researchgate", "twitter", "facebook",
                        "instagram", "wikipedia", "youtube", "google"]
        if not any(na in domain for na in non_academic):
            return True

    return False

_ACADEMIC_PROFILE_RE = re.compile(
    r"orcid\.org"
    r"|researchgate\.net"
    r"|scholar\.google"
    r"|academia\.edu"
    r"|semanticscholar\.org"
    r"|pubmed\.ncbi"
    r"|linkedin\.com/in"
    , re.IGNORECASE
)


def _is_valid_profile_url(url):
    """Return True if url is a university/academic domain OR a known academic profile site."""
    if _is_university_url(url):
        return True
    if _ACADEMIC_PROFILE_RE.search(url):
        return True
    return False


def _ensure_driver2_window():
    """Switch driver2 to the last open window handle, recovering from closed-window errors."""
    try:
        handles = driver2.window_handles
        if not handles:
            return False
        driver2.switch_to.window(handles[-1])
        return True
    except Exception:
        return False


def _restart_driver2():
    """Quit and restart driver2 after a session crash."""
    global driver2
    print("  Restarting driver2 after session crash...")
    try:
        driver2.quit()
    except Exception:
        pass
    try:
        driver2 = Driver(uc=True, headed=True)
        driver2.set_page_load_timeout(30)
        print("  driver2 restarted successfully.")
    except Exception as e:
        print(f"  Failed to restart driver2: {e}")



def _scrape_person_details(row_data, first_name, last_name):
    """Extract profile details from the currently loaded page."""
    result = {col: "null" for col in PERSON_CSV_COLUMNS}
    result["EdBoardMemberID"]  = row_data.get("EdBoardMemberID", "null")
    result["journal_title"]    = row_data.get("journal_title",   "null")
    result["journal_acronym"]  = row_data.get("journal_acronym", "null")
    result["first_name"]       = first_name
    result["last_name"]        = last_name
    result["contact_title"]    = row_data.get("contact_title",   "null")

    _wait_if_captcha(driver2)



    try:
        page_text = driver2.find_element(By.TAG_NAME, "body").text
    except Exception:
        return result



    full_name = f"{first_name} {last_name}"
    current_url = driver2.current_url.lower()

    # --- Email ---
    emails = _RE_EMAIL.findall(page_text)
    # filter out noreply/webmaster/generic addresses
    skip_prefixes = ("noreply", "no-reply", "webmaster", "info@", "contact@", "support@", "admin@")
    for em in emails:
        if not any(em.lower().startswith(p) for p in skip_prefixes):
            result["email"] = em
            break

    # --- ORCID ---
    orcids = _RE_ORCID.findall(page_text)
    if orcids:
        result["orcid"] = orcids[0]

    # --- Institution: try structured selectors first ---
    inst_selectors = [
        # schema.org
        "[itemprop='affiliation']", "[itemprop='organization']",
        # common class names
        ".affiliation", ".institution", ".org", ".department",
        "[class*='affil']", "[class*='instit']", "[class*='department']",
        # ORCID page
        ".affiliation-name", ".org-name",
        # ResearchGate
        ".nova-legacy-e-text--theme-bare",
        # generic
        "h2.title", ".profile-institution",
    ]
    for sel in inst_selectors:
        try:
            el  = driver2.find_element(By.CSS_SELECTOR, sel)
            txt = el.text.strip()
            if txt and len(txt) < 200:
                result["institution"] = txt
                break
        except Exception:
            pass

    # --- Location: structured selectors ---
    for sel in ["[itemprop='addressLocality']", "[class='city']", "[class='location-city']", ".city"]:
        try:
            el  = driver2.find_element(By.CSS_SELECTOR, sel)
            txt = el.text.strip()
            if txt and _is_clean_place(txt):
                result["city"] = txt
                break
        except Exception:
            pass

    for sel in ["[itemprop='addressRegion']", "[class='state']", "[class='region']", ".state"]:
        try:
            el  = driver2.find_element(By.CSS_SELECTOR, sel)
            txt = el.text.strip()
            if txt and _is_clean_place(txt):
                result["state"] = txt
                break
        except Exception:
            pass

    for sel in ["[itemprop='addressCountry']", "[class='country']", ".country"]:
        try:
            el  = driver2.find_element(By.CSS_SELECTOR, sel)
            txt = el.text.strip()
            if txt and _is_country(txt):
                result["country"] = txt
                break
        except Exception:
            pass

    # --- Text-based fallback: scan lines near the person's name ---
    needs_inst    = result["institution"] == "null"
    needs_location = any(result[k] == "null" for k in ["city", "state", "country"])

    if needs_inst or needs_location:
        lines = [l.strip() for l in page_text.split("\n") if l.strip()]

        # find lines near the person's name
        name_indices = [i for i, l in enumerate(lines) if full_name.lower() in l.lower()]
        # if name not found, scan the whole page (e.g. ORCID/RG pages list name in title)
        if not name_indices:
            name_indices = [len(lines) // 2]

        for idx in name_indices:
            window = lines[max(0, idx - 5): idx + 15]

            if needs_inst and result["institution"] == "null":
                if _SPACY_AVAILABLE:
                    doc = _nlp(" ".join(window))
                    for ent in doc.ents:
                        if ent.label_ == "ORG":
                            result["institution"] = ent.text
                            break
                if result["institution"] == "null":
                    for wline in window:
                        if any(kw in wline.lower() for kw in
                               ["university", "institute", "college", "hospital",
                                "school", "center", "centre", "department", "faculty"]):
                            if len(wline) < 150:
                                result["institution"] = wline
                                break

            if needs_location:
                for wline in window:
                    if result["city"] != "null" and result["state"] != "null" and result["country"] != "null":
                        break
                    if len(wline) > 200 or len(wline.split()) > 30:
                        continue
                    if not ("," in wline or re.search(r"\d{4,}", wline)):
                        continue
                    loc = _parse_location_from_text(wline)
                    if result["city"]    == "null" and loc["city"]:
                        result["city"]    = loc["city"]
                    if result["state"]   == "null" and loc["state"]:
                        result["state"]   = loc["state"]
                    if result["country"] == "null" and loc["country"]:
                        result["country"] = loc["country"]

    print(f"    email={result['email']} | orcid={result['orcid']} | institution={result['institution']}")
    return result


def _extract_real_url(href):
    
    if not href:
        return None
    
    match = re.search(r"[?&]q=(https?://[^&]+)", href)
    if match:
    
        return unquote(match.group(1))
    return href


def google_search_person(row_data, first_name, last_name, affiliation):
    """Search Google for the person, open the first university/academic result and scrape."""
    # Build a targeted query: name + institution (if known) + "professor" hint
    query_parts = [f'"{first_name} {last_name}"']
    if affiliation and affiliation.lower() not in ("null", ""):
        query_parts.append(affiliation)
    else:
        query_parts.append("professor OR faculty OR researcher")
    query = " ".join(query_parts)

    print(f"\n  Googling: {query}")

    if not _ensure_driver2_window():
        print("  driver2 has no open windows — skipping.")
        return None

    if not _safe_get(driver2, "https://www.google.com", timeout=20):
        print("  Could not load Google for person search, skipping.")
        return None
    _wait_if_captcha(driver2)

    try:
        btn = WebDriverWait(driver2, 8).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Accept')]"))
        )
        try:
            btn.click()
        except Exception:
            driver2.execute_script("arguments[0].click();", btn)
        time.sleep(1)
    except Exception:
        pass

    # wait for search box to be fully interactable (not just present)
    try:
        search_box = WebDriverWait(driver2, 10).until(
            EC.element_to_be_clickable((By.NAME, "q"))
        )
    except Exception:
        print("    Google search box not found/interactable, skipping.")
        return None

    try:
        search_box.click()
        time.sleep(0.3)
        search_box.clear()
        search_box.send_keys(query)
        search_box.send_keys(Keys.RETURN)
    except Exception:
        # JS fallback if element still not interactable
        try:
            driver2.execute_script(
                "arguments[0].value = arguments[1];", search_box, query
            )
            search_box.send_keys(Keys.RETURN)
        except Exception as e:
            print(f"    Could not interact with search box: {e}, skipping.")
            return None

    try:
        WebDriverWait(driver2, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "h3"))
        )
        time.sleep(1)

        _wait_if_captcha(driver2)
    except Exception:
        print("    Search results did not load, skipping.")
        return None

    result_elements = driver2.find_elements(By.CSS_SELECTOR, "h3")
    if not result_elements:
        print("    No results found")
        return None

    # collect all result links, unwrapping Google redirects to get real URLs
    links = []
    for el in result_elements:
        try:
            raw_href = el.find_element(By.XPATH, "..").get_attribute("href")
            real_url = _extract_real_url(raw_href)
            if real_url:
                links.append(real_url)
        except Exception:
            pass

    if not links:
        print("    No result links found, skipping.")
        return None

    # Only try the FIRST result — if it's not academic, skip entirely
    link = links[0]

    if not _is_valid_profile_url(link):
        print(f"    First result is not an academic URL, skipping: {link}")
        return None

    if not _safe_get(driver2, link, timeout=45):
        # page timed out but may have partially loaded — try scraping anyway
        print(f"    Page timed out, attempting scrape on partial load: {link}")
    else:
        print(f"    Loaded: {link}")

    _cleanup_extra_tabs(driver2)
    time.sleep(2)
    _wait_if_captcha(driver2)

    # verify the final URL after redirects is still academic
    if not _is_valid_profile_url(driver2.current_url):
        print(f"    Redirected to non-academic domain, skipping: {driver2.current_url}")
        return None

    # verify the person's name appears on the page
    # use flexible matching: first+last, or just last name (handles middle initials)
    try:
        page_preview = driver2.find_element(By.TAG_NAME, "body").text[:5000].lower()
    except Exception:
        page_preview = ""

    # strip middle initials for flexible match: "Eva C. Karpinski" → check "eva karpinski"
    first_clean = first_name.split()[0].lower()  # just the first word of first_name
    last_clean  = last_name.lower()

    name_found = (
        f"{first_clean} {last_clean}" in page_preview
        or last_clean in page_preview
    )

    if not name_found:
        print(f"    Name '{first_name} {last_name}' not found on page, skipping: {driver2.current_url}")
        return None

    print(f"    Opened: {driver2.current_url}")
    return _scrape_person_details(row_data, first_name, last_name)


print("\nStarting Google person search + scraping for all scraped members...")

driver2 = Driver(uc=True, headed=True)
driver2.set_page_load_timeout(30) 

df_members      = pd.read_csv(output_csv)
all_person_data = []

for idx, row in df_members.iterrows():
    fn  = str(row.get("first_name",   "")).strip()
    ln  = str(row.get("last_name",    "")).strip()
    aff = str(row.get("affiliation",  "")).strip()

    if not fn or fn == "null" or not ln or ln == "null":
        continue

    # Build a baseline record with nulls so the person always appears in output
    baseline = {col: "null" for col in PERSON_CSV_COLUMNS}
    baseline["EdBoardMemberID"] = str(row.get("EdBoardMemberID", "null"))
    baseline["journal_title"]   = str(row.get("journal_title",   "null"))
    baseline["journal_acronym"] = str(row.get("journal_acronym", "null"))
    baseline["first_name"]      = fn
    baseline["last_name"]       = ln
    baseline["contact_title"]   = str(row.get("contact_title",   "null"))

    try:
        details = google_search_person(row.to_dict(), fn, ln, aff)
        all_person_data.append(details if details else baseline)
        time.sleep(2)
    except Exception as e:
        all_person_data.append(baseline)
        msg = str(e)
        if "invalid session id" in msg or "session deleted" in msg or "chrome not reachable" in msg.lower():
            print(f"  Browser session crashed for {fn} {ln} — restarting driver2")
            _restart_driver2()
        elif "Read timed out" in msg or "HTTPConnectionPool" in msg:
            print(f"  Browser timed out for {fn} {ln} — continuing")
        elif "no such window" in msg or "target window already closed" in msg:
            print(f"  Window closed for {fn} {ln} — attempting recovery")
            _ensure_driver2_window()
        else:
            print(f"  Error for {fn} {ln}: {e}")

try:
    driver2.quit()
except Exception:
    pass

person_csv      = os.path.join(_BASE_DIR, "validation_output.csv")
person_csv_path = os.path.abspath(person_csv)


_RENAME = {
    "EdBoardMemberID": "EdBoardMemberID",
    "journal_title":   "Journal Title",
    "journal_acronym": "Journal Acronym",
    "first_name":      "First Name",
    "last_name":       "Surname",
    "contact_title":   "ContactTitle",
    "orcid":           "ORCiD",
    "institution":     "Institution",
    "city":            "City",
    "state":           "State",
    "country":         "Country",
    "email":           "Email Address",
}

df_person = (
    pd.DataFrame(all_person_data, columns=PERSON_CSV_COLUMNS)
    .fillna("null")
    .rename(columns=_RENAME)
)
df_person.to_csv(person_csv, index=False, encoding="utf-8-sig")

print("\n==============================")
print("Google person search complete.")
print(f"Total records : {len(all_person_data)}")
print(f"Saved CSV at  : {person_csv_path}")
print("Columns: EdBoardMemberID | Journal Acronym | First Name | Surname | ContactTitle | ORCiD | Institution | City | State | Country | Email Address")


print("==============================")
