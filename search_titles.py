import os
import re
import time
import pandas as pd
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
wait = WebDriverWait(driver, 20)

all_data = []
current_title = ""

CSV_COLUMNS = ["journal_title", "section", "role", "first_name", "last_name", "affiliation", "country", "notes"]


def wait_for_cloudflare():
    try:
        WebDriverWait(driver, 20).until(lambda d: "just a moment" not in d.page_source.lower())
    except:
        pass


def split_name(name):
    parts = name.split()
    return (parts[0], " ".join(parts[1:])) if len(parts) > 1 else (name, None)


def extract_country(aff):
    if not aff:
        return None, aff
    if "," in aff:
        parts = aff.split(",")
        return parts[-1].strip(), ",".join(parts[:-1]).strip()
    return None, aff


def record(section, role, name, aff):
    country, aff = extract_country(aff)
    fn, ln = split_name(name)
    all_data.append({
        "journal_title": current_title,
        "section": section,
        "role": role,
        "first_name": fn,
        "last_name": ln,
        "affiliation": aff,
        "country": country,
        "notes": None
    })
    print(f"+ {fn} {ln} | {role}")

#Main Extraction Logic
def _looks_like_role_heading(line):
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
    if len(line) < 4:
        return
    if re.match(r"^\d", line):
        return
    if "@" in line:
        return
    if line.lower().startswith("http"):
        return

    # skip lines that are contact/metadata labels, not person names
    _CONTACT_SKIP = re.compile(
        r"^\s*(phone|tel|fax|mobile|email|e-mail|contact|address|website|url|linkedin|twitter|orcid)\s*[:\-]",
        re.IGNORECASE
    )
    if _CONTACT_SKIP.match(line):
        return

    name = None
    affiliation = None

    for sep in [" – ", "–", " - "]:
        if sep in line:
            parts = line.split(sep, 1)
            name = parts[0].strip()
            affiliation = parts[1].strip() if len(parts) > 1 else None
            break

    if name is None:
        parts = line.split(",", 1)
        candidate = parts[0].strip()
        if 2 <= len(candidate.split()) <= 4 and not re.search(r"\d", candidate):
            name = candidate
            affiliation = parts[1].strip() if len(parts) > 1 else None
        else:
            return

    if not name or len(name.split()) < 2:
        return
    if re.search(r"\d", name):
        return
        

    record(section, role, name, affiliation)


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

                if tag in ["h2", "h3"]:
                    break

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

            if "editorial board" in low or "editors" in low:
                in_editorial = True
                section = line
                current_role = line
                continue

            if not in_editorial:
                continue

            if any(kw in low for kw in ["aims and scope", "submit", "subscribe", "copyright"]):
                in_editorial = False
                continue

            if _looks_like_role_heading(line):
                current_role = line
                continue

            _parse_member_line(line, section, current_role)

    except Exception as e:
        print("Generic text strategy error:", e)

#navigation function
def open_about_page():
    try:
        links = driver.find_elements(By.XPATH, "//a")

        for link in links:
            href = link.get_attribute("href") or ""
            if "about-this-journal" in href or "about_this_journal" in href:
                driver.execute_script("window.location.href = arguments[0];", href)
                time.sleep(5)
                wait_for_cloudflare()
                extract_editorial_board()
                return

        for link in links:
            text = link.text.lower()
            if "about" in text and "journal" in text:
                driver.execute_script("arguments[0].click();", link)
                time.sleep(5)
                wait_for_cloudflare()
                extract_editorial_board()
                return

        for link in links:
            if "about" in link.text.lower():
                driver.execute_script("arguments[0].click();", link)
                time.sleep(5)
                wait_for_cloudflare()
                extract_editorial_board()
                return

        print("  About page link not found")

    except Exception as e:
        print("About page error:", e)


def search_and_extract(query):
    driver.delete_all_cookies()
    driver.get("https://www.google.com")

    try:
        btn = WebDriverWait(driver, 8).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Accept')]"))
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

    driver.execute_script("window.open(arguments[0]);", link)
    driver.switch_to.window(driver.window_handles[-1])
    print("  Opened:", driver.current_url)

    time.sleep(5)
    wait_for_cloudflare()

    open_about_page()

    time.sleep(3)
    driver.close()
    driver.switch_to.window(driver.window_handles[0])


#Main processing loop

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

output_csv = "editorial_board_output.csv"
output_path = os.path.abspath(output_csv)

#saving the data
df_out = pd.DataFrame(all_data, columns=CSV_COLUMNS).fillna("null")
df_out.to_csv(output_csv, index=False, encoding="utf-8-sig")

print("\n==============================")
print("Total records:", len(all_data))
print("Saved CSV at:", output_path)
print("==============================")

driver.quit()

# ============================================================
# Google person search + page scraping
# ============================================================

PERSON_CSV_COLUMNS = [
    "searched_name", "source_url",
    "name", "email", "orcid", "institution", "city", "state", "country"
]

# regex patterns for extracting fields from raw page text
_RE_EMAIL = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")
_RE_ORCID = re.compile(r"\b(\d{4}-\d{4}-\d{4}-\d{3}[\dX])\b")

# Academic TLD patterns and domain name indicators
_UNI_TLD_RE = re.compile(
    r"\.edu(/|$)"                     # .edu TLD  (mit.edu, harvard.edu)
    r"|\.ac\.[a-z]{2}(/|$)"           # .ac.uk / .ac.jp / .ac.in etc.
    r"|\.edu\.[a-z]{2}(/|$)"          # .edu.au / .edu.br etc.
    , re.IGNORECASE
)

_UNI_DOMAIN_RE = re.compile(
    r"(university|univer|univers|univ[^e])"   # "university" or abbreviations like "univie"
    r"|(institute|instit)"
    r"|(college)"
    r"|(hospital)"
    r"|(academia)"
    r"|(polytechnic|polytech)"
    r"|(faculty)"
    , re.IGNORECASE
)


def _is_university_url(url):
    """Return True if the URL looks like a university/academic site."""
    try:
        from urllib.parse import urlparse
        parsed = urlparse(url)
        domain = parsed.netloc.lower()
        full   = url.lower()
    except Exception:
        domain = url.lower()
        full   = url.lower()

    # check TLD first (.edu, .ac.uk, .edu.au etc.)
    if _UNI_TLD_RE.search(domain):
        return True

    # check domain name itself for university indicators
    if _UNI_DOMAIN_RE.search(domain):
        return True

    # check subdomains — e.g. profiles.laps.yorku.ca → parts: ['profiles','laps','yorku','ca']
    # if any subdomain part contains a university keyword, accept it
    parts = domain.replace("www.", "").split(".")
    for part in parts:
        if _UNI_DOMAIN_RE.search(part):
            return True

    # check full URL path for academic signals (e.g. /faculty/, /staff/, /research/)
    academic_path_kw = ["faculty", "staff", "research", "department", "dept", "profile", "people"]
    if any(kw in full for kw in academic_path_kw):
        # only count this if the domain itself looks institutional (not linkedin, researchgate etc.)
        non_academic = ["linkedin", "researchgate", "twitter", "facebook",
                        "instagram", "wikipedia", "youtube", "google"]
        if not any(na in domain for na in non_academic):
            return True

    return False


def _wait_if_captcha(drv):
    """If a CAPTCHA/bot-check is detected, pause until the user presses Enter."""
    captcha_signals = [
        "captcha", "recaptcha", "i am not a robot", "are you a robot",
        "verify you are human", "just a moment", "checking your browser",
        "enable javascript and cookies", "cf-challenge",
    ]
    page = drv.page_source.lower()
    if any(sig in page for sig in captcha_signals):
        print("\n⚠️  CAPTCHA detected! Please solve it in the browser window.")
        input("    Press Enter once you've solved the CAPTCHA to resume...")
        print("    ✅ Resuming...")


def _scrape_person_details(first_name, last_name):
    """Extract profile details from the currently loaded page."""
    result = {col: "null" for col in PERSON_CSV_COLUMNS}
    result["searched_name"] = f"{first_name} {last_name}"
    result["source_url"]    = driver2.current_url

    _wait_if_captcha(driver2)

    # only scrape if we're actually on a university/academic domain
    if not _is_university_url(driver2.current_url):
        print(f"    Skipping scrape — not a university domain: {driver2.current_url}")
        return None

    try:
        page_text = driver2.find_element(By.TAG_NAME, "body").text
    except Exception:
        return result

    # ---- name: look for the searched name on the page ----
    full_name = f"{first_name} {last_name}"
    if full_name.lower() in page_text.lower():
        result["name"] = full_name

    # ---- email ----
    emails = _RE_EMAIL.findall(page_text)
    if emails:
        result["email"] = emails[0]

    # ---- orcid ----
    orcids = _RE_ORCID.findall(page_text)
    if orcids:
        result["orcid"] = orcids[0]

    # ---- institution / city / state / country ----
    for sel in ["[itemprop='affiliation']", "[itemprop='organization']",
                ".affiliation", ".institution", "[class*='affil']", "[class*='instit']"]:
        try:
            el = driver2.find_element(By.CSS_SELECTOR, sel)
            txt = el.text.strip()
            if txt:
                result["institution"] = txt
                break
        except Exception:
            pass

    for sel in ["[itemprop='addressLocality']", "[class*='city']"]:
        try:
            el = driver2.find_element(By.CSS_SELECTOR, sel)
            txt = el.text.strip()
            if txt:
                result["city"] = txt
                break
        except Exception:
            pass

    for sel in ["[itemprop='addressRegion']", "[class*='state']", "[class*='region']"]:
        try:
            el = driver2.find_element(By.CSS_SELECTOR, sel)
            txt = el.text.strip()
            if txt:
                result["state"] = txt
                break
        except Exception:
            pass

    for sel in ["[itemprop='addressCountry']", "[class*='country']"]:
        try:
            el = driver2.find_element(By.CSS_SELECTOR, sel)
            txt = el.text.strip()
            if txt:
                result["country"] = txt
                break
        except Exception:
            pass

    if result["institution"] == "null" or result["country"] == "null":
        lines = [l.strip() for l in page_text.split("\n") if l.strip()]
        for i, line in enumerate(lines):
            if full_name.lower() in line.lower():
                window = lines[max(0, i-3): i+6]
                for wline in window:
                    if result["institution"] == "null" and any(
                        kw in wline.lower() for kw in
                        ["university", "institute", "college", "hospital",
                         "school", "center", "centre", "department", "faculty"]
                    ):
                        result["institution"] = wline

                    if result["country"] == "null" and re.match(
                        r"^[A-Za-z\s,.\-]{3,60}$", wline
                    ) and "," in wline:
                        parts = [p.strip() for p in wline.split(",")]
                        if len(parts) >= 2:
                            result["country"] = parts[-1]
                            if len(parts) >= 3:
                                result["city"]  = parts[0]
                                result["state"] = parts[1]
                            elif len(parts) == 2:
                                result["city"]  = parts[0]
                break

    print(f"    name={result['name']} | email={result['email']} | orcid={result['orcid']} | institution={result['institution']}")
    return result


def _extract_real_url(href):
    """Unwrap Google redirect URLs to get the actual destination."""
    if not href:
        return None
    # Google wraps links as: https://www.google.com/url?q=ACTUAL_URL&...
    match = re.search(r"[?&]q=(https?://[^&]+)", href)
    if match:
        from urllib.parse import unquote
        return unquote(match.group(1))
    return href


def google_search_person(first_name, last_name, affiliation):
    """Search Google for the person, open the first university/academic result and scrape."""
    query_parts = [first_name, last_name]
    if affiliation and affiliation.lower() not in ("null", ""):
        query_parts.append(affiliation)
    query = " ".join(query_parts)

    print(f"\n  Googling: {query}")

    driver2.get("https://www.google.com")
    _wait_if_captcha(driver2)

    try:
        btn = WebDriverWait(driver2, 8).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Accept')]"))
        )
        btn.click()
    except:
        pass

    search_box = WebDriverWait(driver2, 10).until(
        EC.presence_of_element_located((By.NAME, "q"))
    )
    search_box.clear()
    search_box.send_keys(query)
    search_box.send_keys(Keys.RETURN)

    try:
        WebDriverWait(driver2, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "h3"))
        )
        time.sleep(1)

        _wait_if_captcha(driver2)

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

        # try each link — only visit and scrape university/academic URLs
        for link in links:
            if not _is_university_url(link):
                print(f"    Skipping non-university URL: {link}")
                continue

            driver2.get(link)
            print(f"    Opened: {driver2.current_url}")
            time.sleep(3)
            _wait_if_captcha(driver2)

            # also verify the final URL after any redirects is still academic
            if not _is_university_url(driver2.current_url):
                print(f"    Redirected to non-university domain, skipping: {driver2.current_url}")
                driver2.back()
                time.sleep(2)
                continue

            return _scrape_person_details(first_name, last_name)

        print("    No university/academic URL found in results, skipping.")
        return None

    except Exception as e:
        print(f"    Google search error: {e}")
        return None


print("\nStarting Google person search + scraping for all scraped members...")

driver2 = Driver(uc=True, headed=True)

df_members      = pd.read_csv(output_csv)
all_person_data = []

for idx, row in df_members.iterrows():
    fn  = str(row.get("first_name",   "")).strip()
    ln  = str(row.get("last_name",    "")).strip()
    aff = str(row.get("affiliation",  "")).strip()

    if not fn or fn == "null" or not ln or ln == "null":
        continue

    try:
        details = google_search_person(fn, ln, aff)
        if details:
            all_person_data.append(details)
        time.sleep(2)
    except Exception as e:
        print(f"  Error for {fn} {ln}: {e}")

driver2.quit()

person_csv      = "person_details_output.csv"
person_csv_path = os.path.abspath(person_csv)

df_person = pd.DataFrame(all_person_data, columns=PERSON_CSV_COLUMNS).fillna("null")
df_person.to_csv(person_csv, index=False, encoding="utf-8-sig")

print("\n==============================")
print("Google person search complete.")
print(f"Total records : {len(all_person_data)}")
print(f"Saved CSV at  : {person_csv_path}")
print("==============================")


# ============================================================
# ORCID part — commented out
# ============================================================

# def fill_angular_input(driver, css_selector, value):
#     try:
#         el = WebDriverWait(driver, 8).until(
#             EC.presence_of_element_located((By.CSS_SELECTOR, css_selector))
#         )
#         driver.execute_script("arguments[0].scrollIntoView(true);", el)
#         time.sleep(0.2)
#         driver.execute_script("arguments[0].click();", el)
#         time.sleep(0.2)
#         driver.execute_script("""
#             var el = arguments[0];
#             var setter = Object.getOwnPropertyDescriptor(
#                 window.HTMLInputElement.prototype, 'value').set;
#             setter.call(el, '');
#             el.dispatchEvent(new Event('input', { bubbles: true }));
#         """, el)
#         el.send_keys(value)
#         driver.execute_script(
#             "arguments[0].dispatchEvent(new Event('change', { bubbles: true }));", el
#         )
#         time.sleep(0.3)
#         return True
#     except Exception as e:
#         print(f"  fill_angular_input failed '{css_selector}': {e}")
#         return False


# def search_orcid(first_name, last_name, affiliation):
#     results_data = []
#
#     try:
#         full_name = f"{first_name} {last_name}".strip()
#
#         driver.get("https://orcid.org/orcid-search/search")
#         wait_for_cloudflare()
#
#         try:
#             cookie_btn = WebDriverWait(driver, 5).until(
#                 EC.element_to_be_clickable((By.XPATH,
#                     "//button[contains(.,'Accept') or contains(.,'accept')]"))
#             )
#             driver.execute_script("arguments[0].click();", cookie_btn)
#             time.sleep(1)
#         except:
#             pass
#
#         WebDriverWait(driver, 15).until(
#             EC.presence_of_element_located((By.CSS_SELECTOR, "input[formcontrolname]"))
#         )
#         time.sleep(1)
#
#         adv_inputs = driver.find_elements(By.CSS_SELECTOR, "input[formcontrolname='firstName']")
#         if not adv_inputs or not adv_inputs[0].is_displayed():
#             try:
#                 toggle = WebDriverWait(driver, 8).until(
#                     EC.element_to_be_clickable((By.XPATH,
#                         "//*[contains(translate(text(),'abcdefghijklmnopqrstuvwxyz',"
#                         "'ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'ADVANCED SEARCH')]"))
#                 )
#                 driver.execute_script("arguments[0].click();", toggle)
#                 time.sleep(2)
#             except Exception as e:
#                 print(f"  Toggle not found: {e}")
#
#         fill_angular_input(driver, "input[formcontrolname='firstName']", first_name)
#         fill_angular_input(driver, "input[formcontrolname='lastName']", last_name)
#         if affiliation and affiliation != "null":
#             fill_angular_input(driver, "input[formcontrolname='institution']", affiliation)
#
#         try:
#             search_btn = WebDriverWait(driver, 8).until(
#                 EC.element_to_be_clickable((By.XPATH,
#                     "//button[normalize-space(.)='SEARCH' or normalize-space(.)='Search']"))
#             )
#             driver.execute_script("arguments[0].click();", search_btn)
#             print(f"  Searching: {full_name}")
#             time.sleep(5)
#         except Exception as e:
#             print(f"  Search button not found: {e}")
#             return results_data
#
#         try:
#             WebDriverWait(driver, 10).until(
#                 EC.presence_of_element_located((By.XPATH, "//table//tr | //app-search-result"))
#             )
#         except:
#             print(f"  No results table found for: {full_name}")
#
#         rows = driver.find_elements(By.XPATH, "//table//tbody//tr")
#
#         if not rows:
#             rows = driver.find_elements(By.XPATH,
#                 "//app-search-result | //div[contains(@class,'search-result')]")
#
#         print(f"  Found {len(rows)} result row(s)")
#
#         for row in rows:
#             cells = row.find_elements(By.XPATH, ".//td")
#
#             if len(cells) >= 5:
#                 orcid_id     = cells[0].text.strip() or "null"
#                 result_first = cells[1].text.strip() or "null"
#                 result_last  = cells[2].text.strip() or "null"
#                 other_names  = cells[3].text.strip() or "null"
#                 affiliations = cells[4].text.strip() or "null"
#
#                 if orcid_id == "null":
#                     try:
#                         href = cells[0].find_element(By.TAG_NAME, "a").get_attribute("href") or ""
#                         m = re.search(r"(\d{4}-\d{4}-\d{4}-\d{3}[\dX])", href)
#                         if m:
#                             orcid_id = m.group(1)
#                     except:
#                         pass
#             else:
#                 orcid_id     = "null"
#                 result_first = "null"
#                 result_last  = "null"
#                 other_names  = "null"
#                 affiliations = "null"
#
#                 try:
#                     links = row.find_elements(By.XPATH, ".//a[@href]")
#                     for lnk in links:
#                         href = lnk.get_attribute("href") or ""
#                         m = re.search(r"(\d{4}-\d{4}-\d{4}-\d{3}[\dX])", href)
#                         if m:
#                             orcid_id = m.group(1)
#                             break
#                 except:
#                     pass
#
#                 lines = [l.strip() for l in row.text.split("\n") if l.strip()]
#                 if lines:
#                     parts = lines[0].split(None, 1)
#                     result_first = parts[0]
#                     result_last  = parts[1] if len(parts) > 1 else "null"
#                 if len(lines) > 1:
#                     affiliations = " | ".join(lines[1:])
#
#             results_data.append({
#                 "searched_first_name": first_name,
#                 "searched_last_name":  last_name,
#                 "orcid_id":            orcid_id,
#                 "first_name":          result_first,
#                 "last_name":           result_last,
#                 "other_names":         other_names,
#                 "affiliations":        affiliations,
#             })
#             print(f"    → {result_first} {result_last} | {orcid_id}")
#
#         if not results_data:
#             print(f"  No results for: {full_name}")
#
#     except Exception as e:
#         print(f"  ORCID search error for {first_name} {last_name}: {e}")
#
#     return results_data


# print("\nStarting ORCID lookup for all records...")
#
# df_orcid = pd.read_csv(output_csv)
# all_orcid_results = []
#
# for idx, row in df_orcid.iterrows():
#     fn = str(row.get("first_name",  "")).strip()
#     ln = str(row.get("last_name",   "")).strip()
#     af = str(row.get("affiliation", "null")).strip()
#
#     if not fn or fn == "null" or not ln or ln == "null":
#         continue
#
#     results = search_orcid(fn, ln, af)
#     all_orcid_results.extend(results)
#     time.sleep(2)
#
# orcid_results_csv  = "orcid_results.csv"
# orcid_results_path = os.path.abspath(orcid_results_csv)
#
# df_orcid_out = pd.DataFrame(all_orcid_results, columns=[
#     "searched_first_name",
#     "searched_last_name",
#     "orcid_id",
#     "first_name",
#     "last_name",
#     "other_names",
#     "affiliations",
# ]).fillna("null")
# df_orcid_out.to_csv(orcid_results_csv, index=False, encoding="utf-8-sig")
#
# print("\n==================")
# print("ORCID lookup complete.")
# print(f"ORCID results CSV : {orcid_results_path}")
# print(f"Total records     : {len(all_orcid_results)}")
# print("==============================")
#
# driver.quit()
