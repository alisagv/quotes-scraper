import time

import requests
from bs4 import BeautifulSoup
from openpyxl.workbook import Workbook

wb = Workbook()
ws = wb.active
ws.append(["text", "author", "tags"])
response = requests.get("https://quotes.toscrape.com/")
page = 1
while response.status_code == 200:
    soup = BeautifulSoup(response.text, "html.parser")
    quotes = soup.find_all("div", {"class" : "quote"})

    rows = []
    for i in range(0, len(quotes)):
        tags = [tag.text for tag in quotes[i].find("div", {"class": "tags"}).find_all("a")]
        row = {
            "text": quotes[i].find("span", {"class": "text"}).text,
            "author": quotes[i].find("small", {"class": "author"}).text,
            "tags" : tags,
        }
        rows.append(row)

    for row in rows:
        tags = ""
        for tag in row["tags"]:
            tags += f" - {tag}"
        ws.append([row["text"], row["author"], tags])


    print(f"page {page} done")
    page += 1

    nav_link = soup.find("nav")
    next_link = nav_link.find("li", {"class": "next"})
    if next_link is not None:
        time.sleep(1)
        response = requests.get("https://quotes.toscrape.com/" + next_link.find("a")["href"])
    else:
        break

wb.save("quotes.xlsx")
print("file saved to  quotes.xlsx")


