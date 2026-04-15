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
how_many = 5
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
        for _ in range(5):
            driver.execute_script("window.scrollBy(0, 1000);")
            time.sleep(1)

        wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//h2[contains(.,'Editorial')]")
            )
        )

        elements = driver.find_elements(
            By.XPATH,
            "//h2[contains(.,'Editorial')]/following-sibling::*"
        )

        print("Total elements found:", len(elements))

        current_role = ""

        for el in elements:
            tag = el.tag_name.lower()
            txt = el.text.strip()

            if not txt:
                continue

            txt = txt.encode('utf-8', 'ignore').decode('utf-8')
            txt = txt.replace("â€“", "–")

            if tag == "h2":
                break

            lines = [l.strip() for l in txt.split("\n") if l.strip()]

            for line in lines:
                if len(line) > 120:
                    continue

                if line.count(",") > 3:
                    continue

                if "–" not in line:
                    if len(line.split()) <= 6:
                        current_role = line
                    continue

                parts = line.split("–")

                name = parts[0].strip()
                affiliation = parts[1].strip() if len(parts) > 1 else ""

                if not (2 <= len(name.split()) <= 6):
                    continue

                if any(x in name.lower() for x in [
                    "journal", "society", "publication", "study"
                ]):
                    continue

                print(f"Captured: {name} | {affiliation} | {current_role}")

                all_data.append({
                    "journal_title": current_title,
                    "name_affiliation": name,
                    "role": current_role,
                    "other_details": affiliation
                })

    except Exception as e:
        print("Extraction error:", e)


def open_about_page():
    try:
        links = driver.find_elements(By.XPATH, "//a")

        for link in links:
            try:
                href = link.get_attribute("href") or ""
                text = link.text.lower()

                if "about" in text or "about-this-journal" in href:
                    driver.execute_script("arguments[0].click();", link)
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
    driver.delete_all_cookies()
    driver.get("https://www.google.com")

    try:
        btn = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//button[contains(., 'Accept')]")
            )
        )
        btn.click()
    except:
        pass

    search_box = wait.until(
        EC.presence_of_element_located((By.NAME, "q"))
    )
    search_box.send_keys(f'"{query}" site:{preferred_site}')
    search_box.send_keys(Keys.RETURN)

    wait.until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "h3"))
    )
    results = driver.find_elements(By.CSS_SELECTOR, "h3")

    if results:
        link = results[0].find_element(By.XPATH, "..").get_attribute("href")

        driver.execute_script("window.open(arguments[0]);", link)
        driver.switch_to.window(driver.window_handles[-1])

        print("Opened URL:", driver.current_url)

        time.sleep(5)
        wait_for_cloudflare()
        open_about_page()

        driver.close()
        driver.switch_to.window(driver.window_handles[0])


for i in range(min(how_many, len(df))):
    current_title = str(df.iloc[i].get(column_name, "")).strip()

    if not current_title or current_title.lower() == "nan":
        continue

    print("\nProcessing:", current_title)

    try:
        search_and_extract(current_title)
        time.sleep(3)
    except Exception as e:
        print("Row error:", e)


output_csv = "editorial_board_output.csv"
output_path = os.path.abspath(output_csv)

df_out = pd.DataFrame(all_data)
df_out.to_csv(output_csv, index=False)

print("\n==============================")
print("Total records:", len(all_data))
print("Saved CSV at:", output_path)
print("==============================")

driver.quit()