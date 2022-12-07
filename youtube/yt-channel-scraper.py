from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException

import os
import sys
import time
from datetime import datetime
from dateutil.relativedelta import relativedelta
import pandas as pd
from tqdm import tqdm
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows


# Accept cookies popup when it appears, timeout=20s.
def accept_cookies(driver):
    cookies_visible = EC.visibility_of_element_located((By.CSS_SELECTOR, "[aria-label='Accept all']"))
    cookie_consent = WebDriverWait(driver, 20).until(cookies_visible)
    cookie_consent.click()

# Function to check if element exists.
def element_exists(driver, xpath):
    try:
        driver.find_element(By.XPATH, xpath)
    except NoSuchElementException:
        return False
    return True

# Scrolls to bottom of passed element.
def scroll_to_bottom(driver):
    while True:
        html = driver.find_element(By.TAG_NAME, 'html')
        html.send_keys(Keys.END)
        # Stop scrolling if youtube item renderer doesn't exist.
        if not element_exists(driver, '//*[@id="contents"]/ytd-continuation-item-renderer'):
            break

# Get details from individual video element.
def get_video_details(video):
    
    title = video.find_element(By.XPATH, './/*[@id="video-title"]').text
    metadata = video.find_elements(By.XPATH, './/*[@id="metadata-line"]/span')

    if len(metadata) == 2:
        views = metadata[0].text[:-6]
        posted = metadata[1].text
    # Youtube Originals (no view count)
    else:
        views = None
        posted = metadata[0].text

    # Convert views string to integer.
    if views:
        view_suffixes = {"K": 1000, "M": 1000000, "B": 1000000000}
        for k, v in view_suffixes.items():
            if views[-1] == k:
                views_int = int(float(views[:-1]) * v)
                break
        else:
            views_int = int(float(views))
    else:
        views_int = views

    # Convert posted string to date.
    date_suffixes = {}
    date_suffix = posted.split(" ")[1]
    date_value = posted.split(" ")[0]
    if date_suffix[-1] != "s":
        date_suffix+="s"
    date_suffixes[date_suffix] = int(date_value)
    date = datetime.now() - relativedelta(**date_suffixes)

    video_details = {
        "title": title,
        "views": views_int,
        "date": date
    }
    return video_details

# Loop through all videos.
def get_videos_details(driver):
    videos = driver.find_elements(By.XPATH, '//*[@id="content"]')[1:-1]
    videos_details = []
    for video in tqdm(videos):  
        video_details = get_video_details(video)
        videos_details.append(video_details)
    return videos_details

# Saves results to spreadsheet from dataframe.
def save_to_xlsx(videos_details, filename):
    df = pd.DataFrame(videos_details)
    wb = Workbook()
    ws = wb.active
    for r in dataframe_to_rows(df, index=True, header=True):
        ws.append(r)
    os.makedirs(f"{sys.path[0]}/output", exist_ok=True)
    wb.save(f"{sys.path[0]}/output/{filename}.xlsx")

def main():
    print("Enter channel username:")
    channel_username = input()
    start = time.time()

    service = Service("C:\Program Files (x86)\chromedriver.exe")
    options = Options()
    driver = webdriver.Chrome(service=service, options=options)
    driver.get(f"https://www.youtube.com/@{channel_username}/videos")

    accept_cookies(driver)
    scroll_to_bottom(driver)
    videos_details = get_videos_details(driver)
    save_to_xlsx(videos_details, channel_username)

    end = time.time()
    print(f"Channel @{channel_username} scraped in {round(end-start)}s.")
    time.sleep(5)
    driver.quit()

if __name__ == "__main__":
    main()