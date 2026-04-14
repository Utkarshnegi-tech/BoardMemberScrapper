import pandas as pd
import time
import random
import os
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

all_data = []
current_title = ""

def wait_for_cloudflare():
    try:
        WebDriverWait(driver, 30).until(
            lambda d: "just a moment" not in d.page_source.lower()
        )
        time.sleep(2)
    except:
        pass

def extract_editorial_board():
    global all_data, current_title
    try:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(3)

        wait.until(EC.presence_of_element_located(
            (By.XPATH, "//h2[contains(.,'Editorial board')]")
        ))

        elements = driver.find_elements(
            By.XPATH,
            "//h2[contains(.,'Editorial board')]/following::*[self::p or self::li or self::div or self::span]"
        )

        lines = []

        for el in elements:
            txt = el.text
            if txt:
                txt = txt.encode('utf-8', 'ignore').decode('utf-8')
                txt = txt.replace("â€“", "–")
                lines.extend(txt.split("\n"))

        lines = [l.strip() for l in lines if l.strip()]

        current_role = ""

        for line in lines:

            if line.lower().startswith("editorial board"):
                continue

            # ROLE
            if "–" not in line:
                current_role = line
                continue

            # NAME + AFFILIATION
            if "–" in line:
                parts = line.split("–")
                name_aff = parts[0].strip()
                affiliation = parts[1].strip() if len(parts) > 1 else ""

                # Handle multiple names in one line (comma separated)
                if "," in name_aff and current_role.lower().startswith("founding"):
                    split_names = [n.strip() for n in name_aff.split(",")]

                    for nm in split_names:
                        all_data.append({
                            "journal_title": current_title,
                            "name_affiliation": nm,
                            "role": current_role,
                            "other_details": affiliation
                        })
                else:
                    all_data.append({
                        "journal_title": current_title,
                        "name_affiliation": name_aff,
                        "role": current_role,
                        "other_details": affiliation
                    })

    except Exception as e:
        print("Extraction error:", e)

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
    except Exception as e:
        print("About page error:", e)

def search_and_extract(query):
    driver.get("https://www.google.com")
    time.sleep(2)

    try:
        btn = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Accept')]"))
        )
        btn.click()
    except:
        pass

    search_box = wait.until(EC.presence_of_element_located((By.NAME, "q")))
    search_box.send_keys(f'"{query}" site:{preferred_site}')
    search_box.send_keys(Keys.RETURN)

    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "h3")))
    results = driver.find_elements(By.CSS_SELECTOR, "h3")

    if results:
        results[0].click()
        time.sleep(5)
        wait_for_cloudflare()
        open_about_page()

for i in range(min(how_many, len(df))):
    current_title = str(df.iloc[i].get(column_name, "")).strip()

    if not current_title or current_title.lower() == "nan":
        continue

    print("\nProcessing:", current_title)

    try:
        search_and_extract(current_title)
        time.sleep(5)
    except Exception as e:
        print("Row error:", e)

# ✅ SAVE OUTPUT
output_csv = "editorial_board_output.csv"
output_excel = "editorial_board_output.xlsx"

df_out = pd.DataFrame(all_data)

df_out.to_csv(output_csv, index=False)
df_out.to_excel(output_excel, index=False)

print("\nSaved CSV:", os.path.abspath(output_csv))
print("Saved Excel:", os.path.abspath(output_excel))

input("Press Enter to close...")
driver.quit()