import pandas as pd
import time
import os
import re
from seleniumbase import Driver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

file_path = r"malmad-title-list-9-january-2023.xlsx"
column_name = "title"
how_many = 1
preferred_site = "tandfonline.com"

df = pd.read_excel(file_path)
df.columns = df.columns.str.strip().str.lower()

driver = Driver(uc=True, headed=True)
driver.maximize_window()
wait = WebDriverWait(driver, 25)

all_data = []
current_title = ""

# CSV columns
CSV_COLUMNS = ["journal_title", "section", "role", "first_name", "last_name", "affiliation", "country", "notes"]


def wait_for_cloudflare():
    try:
        WebDriverWait(driver, 30).until(
            lambda d: "just a moment" not in d.page_source.lower()
        )
        time.sleep(2)
    except:
        pass


def parse_country_from_affiliation(affiliation):
    """Try to extract country from the end of an affiliation string."""
    if not affiliation:
        return None, affiliation

    # Common country names to detect at end of affiliation
    known_countries = [
        "USA", "United States", "UK", "United Kingdom", "Canada", "Australia",
        "Germany", "France", "Netherlands", "China", "Japan", "India", "Italy",
        "Spain", "Sweden", "Norway", "Denmark", "Finland", "Belgium", "Switzerland",
        "Austria", "Poland", "Portugal", "Ireland", "Scotland", "New Zealand",
        "South Africa", "Brazil", "Colombia", "Taiwan", "Hong Kong", "Singapore",
        "South Korea", "Korea", "Israel", "Turkey", "Greece", "Hungary", "Romania",
        "Czech Republic", "Slovakia", "Croatia", "Serbia", "Bulgaria", "Ukraine",
        "Russia", "Mexico", "Argentina", "Chile", "Peru", "Indonesia", "Malaysia",
        "Thailand", "Vietnam", "Philippines", "Pakistan", "Bangladesh", "Egypt",
        "Nigeria", "Kenya", "Ghana", "Ethiopia", "Tanzania", "Uganda", "Zimbabwe",
        "SAR", "England", "Wales", "Northern Ireland",
    ]

    for country in known_countries:
        # match country at end of string (after comma or space)
        pattern = r",?\s*" + re.escape(country) + r"\s*$"
        if re.search(pattern, affiliation, re.IGNORECASE):
            clean_affiliation = re.sub(pattern, "", affiliation, flags=re.IGNORECASE).strip().rstrip(",").strip()
            return country, clean_affiliation

    return None, affiliation


def record(section, role, person_name, affiliation_raw, notes=None):
    """Append one cleaned record to all_data."""
    country, affiliation = parse_country_from_affiliation(affiliation_raw)

    # split into first and last name
    name_parts = person_name.strip().split() if person_name else []
    first_name = name_parts[0] if name_parts else None
    last_name  = " ".join(name_parts[1:]) if len(name_parts) > 1 else None

    all_data.append({
        "journal_title": current_title,
        "section":       section     or None,
        "role":          role        or None,
        "first_name":    first_name  or None,
        "last_name":     last_name   or None,
        "affiliation":   affiliation or None,
        "country":       country     or None,
        "notes":         notes       or None,
    })
    print(f"  + {first_name} {last_name} | {role} | {affiliation} | {country}")


def extract_editorial_board():
    """Extract editorial board members from the current page."""
    global all_data, current_title

    try:
        # scroll to load lazy content
        for _ in range(6):
            driver.execute_script("window.scrollBy(0, 1200);")
            time.sleep(0.8)

        # wait for any editorial-related text
        WebDriverWait(driver, 20).until(
            lambda d: "editorial" in d.page_source.lower()
        )
        time.sleep(2)

        #  1: structured dl/dt/dd lists (common on T&F) 
        extracted = _extract_dl_structure()
        if extracted:
            return

        #2: heading + sibling paragraphs 
        extracted = _extract_heading_siblings()
        if extracted:
            return

        # 3: generic text scan 
        _extract_generic_text()

    except Exception as e:
        print("Extraction error:", e)


def _extract_dl_structure():
    """
    Handle pages where roles are in <dt> and members in <dd> or <p> siblings.
    Returns True if anything was captured.
    """
    count_before = len(all_data)

    try:
        # find all dt elements that look like role headings
        dts = driver.find_elements(By.XPATH, "//dt | //h4 | //h5")
        for dt in dts:
            role_text = dt.text.strip()
            if not role_text or len(role_text.split()) > 8:
                continue

            # collect following dd / p siblings until next dt/h4
            siblings = dt.find_elements(By.XPATH, "following-sibling::dd | following-sibling::p")
            section = "Editorial Board"

            for sib in siblings:
                sib_tag = sib.tag_name.lower()
                if sib_tag in ["dt", "h4", "h5"]:
                    break
                text = sib.text.strip()
                if not text:
                    continue
                lines = [l.strip() for l in text.split("\n") if l.strip()]
                for line in lines:
                    _parse_member_line(line, section, role_text)

    except Exception as e:
        print("DL strategy error:", e)

    return len(all_data) > count_before


def _extract_heading_siblings():
    """
    Find h2/h3 that contains 'Editorial', then walk following siblings.
    Returns True if anything was captured.
    """
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

                # new top-level section → stop
                if tag in ["h2", "h3"]:
                    break

                # sub-role heading
                if tag in ["h4", "h5", "strong", "b"]:
                    role_candidate = el.text.strip()
                    if role_candidate:
                        current_role = role_candidate
                    continue

                text = el.text.strip()
                if not text:
                    continue

                lines = [l.strip() for l in text.split("\n") if l.strip()]
                for line in lines:
                    # detect inline role headings (short, no dash separator)
                    if _looks_like_role_heading(line):
                        current_role = line
                        continue
                    _parse_member_line(line, section, current_role)

    except Exception as e:
        print("Heading-sibling strategy error:", e)

    return len(all_data) > count_before


def _extract_generic_text():
    """
    Last-resort: grab all visible text blocks and parse name–affiliation lines.
    """
    try:
        body = driver.find_element(By.TAG_NAME, "body")
        full_text = body.text

        in_editorial = False
        current_role = "Editorial Board"
        section = "Editorial Board"

        for line in full_text.split("\n"):
            line = line.strip()
            if not line:
                continue

            low = line.lower()

            # enter editorial section
            if "editorial board" in low or "editors" in low:
                in_editorial = True
                section = line
                current_role = line
                continue

            if not in_editorial:
                continue

            # exit on unrelated sections
            if any(kw in low for kw in ["aims and scope", "submit", "subscribe", "contact", "copyright"]):
                in_editorial = False
                continue

            if _looks_like_role_heading(line):
                current_role = line
                continue

            _parse_member_line(line, section, current_role)

    except Exception as e:
        print("Generic text strategy error:", e)


def _looks_like_role_heading(line):
    """Return True if the line looks like a role label, not a person entry."""
    words = line.split()
    if len(words) > 7:
        return False
    if any(sep in line for sep in ["–", " - ", ", "]):
        return False
    role_keywords = [
        "editor", "board", "chair", "director", "advisor", "adviser",
        "reviewer", "committee", "associate", "managing", "senior",
        "founding", "honorary", "emeritus", "consultant",
    ]
    low = line.lower()
    return any(kw in low for kw in role_keywords)


def _parse_member_line(line, section, role):
    """
    Parse a single text line into a person record.
    Handles formats:
      - "First Last – Affiliation, Country"
      - "First Last - Affiliation"
      - "First Last, Affiliation, Country"
    """
    # skip obvious non-person lines
    if len(line) < 4:
        return
    if re.match(r"^\d", line):          # starts with digit
        return
    if "@" in line:                      # email address
        return
    if line.lower().startswith("http"):  # URL
        return

    name = None
    affiliation = None

    # separator: em-dash or spaced hyphen
    for sep in [" – ", "–", " - "]:
        if sep in line:
            parts = line.split(sep, 1)
            name = parts[0].strip()
            affiliation = parts[1].strip() if len(parts) > 1 else None
            break

    # fallback: comma split — only if first token looks like a name
    if name is None:
        parts = line.split(",", 1)
        candidate = parts[0].strip()
        # a name usually has 2-4 words, no digits
        if 2 <= len(candidate.split()) <= 4 and not re.search(r"\d", candidate):
            name = candidate
            affiliation = parts[1].strip() if len(parts) > 1 else None
        else:
            return  # can't parse

    # basic name sanity check
    if not name or len(name.split()) < 2:
        return
    if re.search(r"\d", name):
        return

    record(section, role, name, affiliation)

# ── Navigation helpers ──────────────────────────────────────────────────────

def open_about_page():
    """Find and navigate to the 'About this journal' page, then extract."""
    try:
        links = driver.find_elements(By.XPATH, "//a")

        # prefer explicit about-this-journal href
        for link in links:
            href = link.get_attribute("href") or ""
            if "about-this-journal" in href or "about_this_journal" in href:
                driver.execute_script("window.location.href = arguments[0];", href)
                time.sleep(5)
                wait_for_cloudflare()
                extract_editorial_board()
                return

        # fallback: any link with "about" in text
        for link in links:
            text = link.text.lower()
            if "about" in text and "journal" in text:
                driver.execute_script("arguments[0].click();", link)
                time.sleep(5)
                wait_for_cloudflare()
                extract_editorial_board()
                return

        # last fallback: any "about" link
        for link in links:
            if "about" in (link.text.lower()):
                driver.execute_script("arguments[0].click();", link)
                time.sleep(5)
                wait_for_cloudflare()
                extract_editorial_board()
                return

        print("  About page link not found")

    except Exception as e:
        print("About page error:", e)


def search_and_extract(query):
    """Google the journal title, open the first T&F result, navigate to About, extract."""
    driver.delete_all_cookies()
    driver.get("https://www.google.com")

    # accept cookies if prompted
    try:
        btn = WebDriverWait(driver, 8).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//button[contains(., 'Accept')]")
            )
        )
        btn.click()
    except:
        pass

    search_box = wait.until(EC.presence_of_element_located((By.NAME, "q")))
    search_box.clear()
    search_box.send_keys(f'"{query}" site:{preferred_site}')
    search_box.send_keys(Keys.RETURN)

    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "h3")))
    time.sleep(2)

    results = driver.find_elements(By.CSS_SELECTOR, "h3")
    if not results:
        print("  No search results found")
        return

    link = results[0].find_element(By.XPATH, "..").get_attribute("href")
    if not link:
        print("  Could not get result URL")
        return

    # open in new tab
    driver.execute_script("window.open(arguments[0]);", link)
    driver.switch_to.window(driver.window_handles[-1])
    print("  Opened:", driver.current_url)

    time.sleep(5)
    wait_for_cloudflare()

    open_about_page()

    time.sleep(3)
    driver.close()
    driver.switch_to.window(driver.window_handles[0])

# ── Main loop ───────────────────────────────────────────────────────────────

for i in range(min(how_many, len(df))):
    current_title = str(df.iloc[i].get(column_name, "")).strip()

    if not current_title or current_title.lower() == "nan":
        continue

    print(f"\n[{i+1}] Processing: {current_title}")

    try:
        search_and_extract(current_title)
        time.sleep(3)
    except Exception as e:
        print("Row error:", e)

# ── Save CSV ────────────────────────────────────────────────────────────────

output_csv = "editorial_board_output.csv"
output_path = os.path.abspath(output_csv)

df_out = pd.DataFrame(all_data, columns=CSV_COLUMNS)
df_out = df_out.fillna("null")
df_out.to_csv(output_csv, index=False, encoding="utf-8-sig")

print("\n==============================")
print("Total records:", len(all_data))
print("Saved CSV at:", output_path)
print("==============================")

# ── ORCID lookup ─────────────────────────────────────────────────────────────

def fill_angular_input(driver, css_selector, value):
    """Fill an Angular Material input by triggering native input/change events."""
    try:
        el = WebDriverWait(driver, 8).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, css_selector))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", el)
        time.sleep(0.2)
        driver.execute_script("arguments[0].click();", el)
        time.sleep(0.2)
        driver.execute_script("""
            var el = arguments[0];
            var setter = Object.getOwnPropertyDescriptor(
                window.HTMLInputElement.prototype, 'value').set;
            setter.call(el, '');
            el.dispatchEvent(new Event('input', { bubbles: true }));
        """, el)
        el.send_keys(value)
        driver.execute_script(
            "arguments[0].dispatchEvent(new Event('change', { bubbles: true }));", el
        )
        time.sleep(0.3)
        return True
    except Exception as e:
        print(f"  fill_angular_input failed '{css_selector}': {e}")
        return False


def search_orcid(first_name, last_name, affiliation):
    """
    Fill the ORCID advanced search form and scrape all result cards.
    Returns a list of dicts with columns matching the ORCID results page.
    """
    results_data = []

    try:
        full_name = f"{first_name} {last_name}".strip()

        driver.get("https://orcid.org/orcid-search/search")
        wait_for_cloudflare()

        # dismiss cookie banner
        try:
            cookie_btn = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH,
                    "//button[contains(.,'Accept') or contains(.,'accept')]"))
            )
            driver.execute_script("arguments[0].click();", cookie_btn)
            time.sleep(1)
        except:
            pass

        # wait for Angular inputs to render
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[formcontrolname]"))
        )
        time.sleep(1)

        # open advanced panel if collapsed
        adv_inputs = driver.find_elements(By.CSS_SELECTOR, "input[formcontrolname='firstName']")
        if not adv_inputs or not adv_inputs[0].is_displayed():
            try:
                toggle = WebDriverWait(driver, 8).until(
                    EC.element_to_be_clickable((By.XPATH,
                        "//*[contains(translate(text(),'abcdefghijklmnopqrstuvwxyz',"
                        "'ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'ADVANCED SEARCH')]"))
                )
                driver.execute_script("arguments[0].click();", toggle)
                time.sleep(2)
            except Exception as e:
                print(f"  Toggle not found: {e}")

        # fill fields
        fill_angular_input(driver, "input[formcontrolname='firstName']", first_name)
        fill_angular_input(driver, "input[formcontrolname='lastName']",  last_name)
        if affiliation and affiliation != "null":
            fill_angular_input(driver, "input[formcontrolname='institution']", affiliation)

        # submit
        try:
            search_btn = WebDriverWait(driver, 8).until(
                EC.element_to_be_clickable((By.XPATH,
                    "//button[normalize-space(.)='SEARCH' or normalize-space(.)='Search']"))
            )
            driver.execute_script("arguments[0].click();", search_btn)
            print(f"  Searching: {full_name}")
            time.sleep(5)
        except Exception as e:
            print(f"  Search button not found: {e}")
            return results_data

        # ── scrape results table ──────────────────────────────────────────────
        # ORCID renders results as a table with columns:
        # ORCID ID | First Name | Last Name | Other Names | Affiliations
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//table//tr | //app-search-result"))
            )
        except:
            print(f"  No results table found for: {full_name}")

        rows = driver.find_elements(By.XPATH, "//table//tbody//tr")

        # fallback: if no <table>, try row-based divs
        if not rows:
            rows = driver.find_elements(By.XPATH,
                "//app-search-result | //div[contains(@class,'search-result')]")

        print(f"  Found {len(rows)} result row(s)")

        for row in rows:
            cells = row.find_elements(By.XPATH, ".//td")

            if len(cells) >= 5:
                # proper table row: td[0]=ORCID ID, td[1]=First, td[2]=Last,
                #                   td[3]=Other Names, td[4]=Affiliations
                orcid_id     = cells[0].text.strip() or "null"
                result_first = cells[1].text.strip() or "null"
                result_last  = cells[2].text.strip() or "null"
                other_names  = cells[3].text.strip() or "null"
                affiliations = cells[4].text.strip() or "null"

                # also try to get ORCID iD from the link href if text is empty
                if orcid_id == "null":
                    try:
                        href = cells[0].find_element(By.TAG_NAME, "a").get_attribute("href") or ""
                        m = re.search(r"(\d{4}-\d{4}-\d{4}-\d{3}[\dX])", href)
                        if m:
                            orcid_id = m.group(1)
                    except:
                        pass

            else:
                # fallback for card-based layout
                orcid_id     = "null"
                result_first = "null"
                result_last  = "null"
                other_names  = "null"
                affiliations = "null"

                try:
                    links = row.find_elements(By.XPATH, ".//a[@href]")
                    for lnk in links:
                        href = lnk.get_attribute("href") or ""
                        m = re.search(r"(\d{4}-\d{4}-\d{4}-\d{3}[\dX])", href)
                        if m:
                            orcid_id = m.group(1)
                            break
                except:
                    pass

                lines = [l.strip() for l in row.text.split("\n") if l.strip()]
                if lines:
                    parts = lines[0].split(None, 1)
                    result_first = parts[0]
                    result_last  = parts[1] if len(parts) > 1 else "null"
                if len(lines) > 1:
                    affiliations = " | ".join(lines[1:])

            results_data.append({
                "searched_first_name": first_name,
                "searched_last_name":  last_name,
                "orcid_id":            orcid_id,
                "first_name":          result_first,
                "last_name":           result_last,
                "other_names":         other_names,
                "affiliations":        affiliations,
            })
            print(f"    → {result_first} {result_last} | {orcid_id}")

        if not results_data:
            print(f"  No results for: {full_name}")

    except Exception as e:
        print(f"  ORCID search error for {first_name} {last_name}: {e}")

    return results_data


print("\nStarting ORCID lookup for all records...")

df_orcid        = pd.read_csv(output_csv)
all_orcid_results = []

for idx, row in df_orcid.iterrows():
    fn = str(row.get("first_name",  "")).strip()
    ln = str(row.get("last_name",   "")).strip()
    af = str(row.get("affiliation", "null")).strip()

    if not fn or fn == "null" or not ln or ln == "null":
        continue

    results = search_orcid(fn, ln, af)
    all_orcid_results.extend(results)
    time.sleep(2)

# save ORCID results to a separate CSV
orcid_results_csv  = "orcid_results.csv"
orcid_results_path = os.path.abspath(orcid_results_csv)

df_orcid_out = pd.DataFrame(all_orcid_results, columns=[
    "searched_first_name",
    "searched_last_name",
    "orcid_id",
    "first_name",
    "last_name",
    "other_names",
    "affiliations",
])
df_orcid_out = df_orcid_out.fillna("null")
df_orcid_out.to_csv(orcid_results_csv, index=False, encoding="utf-8-sig")

print("\n==================")
print("ORCID lookup complete.")
print(f"ORCID results CSV : {orcid_results_path}")
print(f"Total records     : {len(all_orcid_results)}")
print("==============================")

driver.quit()

