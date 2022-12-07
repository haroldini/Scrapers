from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from bs4 import BeautifulSoup
from tqdm import tqdm
from datetime import datetime
import requests
import pandas as pd
import time
import sys
import os


# Get new random page.
def shuffle_article():
    new_url = "https://en.wikipedia.org/wiki/Special:Random"
    new_res = requests.get(new_url)
    new_doc = BeautifulSoup(new_res.text, "html.parser")
    return new_doc

# Read relevant information from page.
def read_article(doc):
    # Article name.
    title = doc.find("h1", id="firstHeading").find().text

    # Date last modified.
    last_mod = doc.find("li", id="footer-info-lastmod").text[30:-7]
    last_mod = datetime.strptime(last_mod, "%d %B %Y, at %H:%M")

    # Categories.
    cats = doc.find("div", id="mw-normal-catlinks").find("ul").find_all("li")
    cats = str([cat.text for cat in cats])[1:-1]
    
    # Number of references.
    refs = doc.find("ol", class_="references")
    if refs:
        num_refs = len(refs.find_all("li"))
    else:
        num_refs = 0

    article_properties = {
        "title": title,
        "num_refs": num_refs,
        "last_mod": last_mod,
        "categories": cats,
    }
    return article_properties

# Get properties from n random articles.
def get_articles(n_articles):
    articles = []
    for i in tqdm(range(n_articles)):
        doc = shuffle_article()
        articles.append(read_article(doc))
    return articles

# Save results to an excel file.
def save_to_xlsx(articles):
    df = pd.DataFrame(articles)
    wb = Workbook()
    ws = wb.active
    for r in dataframe_to_rows(df, index=True, header=True):
        ws.append(r)
    os.makedirs(f'{sys.path[0]}/output', exist_ok=True)
    wb.save(f"{sys.path[0]}/output/output.xlsx")


def main():
    # Start scraper.
    print("Enter number of random articles to scrape.")
    n_articles = int(input())
    start = time.time()

    # Run scraper.
    articles = get_articles(n_articles)
    save_to_xlsx(articles)

    # End scraper.
    end = time.time()
    print(f"{n_articles} random Wikipedia articles scraped in {round(end-start)}s.")

if __name__ == "__main__":
    main()