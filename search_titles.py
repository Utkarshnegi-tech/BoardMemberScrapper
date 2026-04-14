import pandas as pd
import time
import random
from seleniumbase import Driver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

file_path = r"malmad-title-list-9-january-2023.xlsx"
column_name = "title"
how_many = 2
preferred_site = "tandfonline.com"

df = pd.read_excel(file_path)
df.columns = df.columns.str.strip().str.lower()

driver = Driver(uc=True, headed=True)
driver.maximize_window()
wait = WebDriverWait(driver, 25)

def wait_for_cloudflare():
    try:
        WebDriverWait(driver, 30).until(
            lambda d: "just a moment" not in d.page_source.lower()
            and "performing security verification" not in d.page_source.lower()
        )
        time.sleep(random.uniform(2, 4))
    except:
        try:
            driver.uc_gui_click_captcha()
            time.sleep(8)
        except:
            pass

def extract_editorial_board():
    try:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(3)

        section = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//h2[contains(.,'Editorial board')]")
            )
        )

        driver.execute_script("arguments[0].scrollIntoView({block:'start'});", section)
        time.sleep(2)

        elements = driver.find_elements(
            By.XPATH,
            "//h2[contains(.,'Editorial board')]/following::*"
        )

        names = []
        current_role = ""

        for el in elements:
            tag = el.tag_name.lower()
            text = el.text.strip()

            if not text:
                continue

            if tag in ["h2"] and "editorial" not in text.lower():
                break

            if tag in ["h3", "strong", "b"]:
                current_role = text
                continue

            if "–" in text:
                names.append(f"{text} | {current_role}")

        print("\nEditorial Board:")
        for n in names:
            print("-", n)

        return names

    except Exception as e:
        print("Extraction error:", e)
        return []

def open_about_page():
    try:
        for by, value in [
            (By.LINK_TEXT, "About this journal"),
            (By.PARTIAL_LINK_TEXT, "About this journal"),
            (By.XPATH, "//a[contains(@href,'about-this-journal')]")
        ]:
            try:
                link = wait.until(EC.element_to_be_clickable((by, value)))
                link.click()
                time.sleep(5)
                wait_for_cloudflare()
                extract_editorial_board()
                return
            except:
                continue

        print("About page not found")

    except Exception as e:
        print("About page error:", e)

def search_and_extract(query):
    driver.get("https://www.google.com")
    time.sleep(random.uniform(2, 4))

    try:
        btn = wait.until(
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

    results = driver.find_elements(By.CSS_SELECTOR, "h3")

    if not results:
        print("No results")
        return

    try:
        results[0].click()
        time.sleep(5)
        wait_for_cloudflare()
        open_about_page()
    except Exception as e:
        print("Open result error:", e)

for i in range(min(how_many, len(df))):
    title = str(df.iloc[i].get(column_name, "")).strip()

    if not title or title.lower() == "nan":
        continue

    print("\n" + "=" * 80)
    print(f"Processing: {title}")

    try:
        search_and_extract(title)
        time.sleep(random.uniform(5, 8))
    except Exception as e:
        print("Row error:", e)

input("\nPress Enter to close...")
driver.quit()