# build command: pyinstaller --onefile taxlien.py
# expected input file  = input.xlsx
# expected output file = taxlien.xlsx

from bs4 import BeautifulSoup
import requests
import pandas as pd
from tabulate import tabulate
import time
import xlsxwriter

# to prevent being blocked as bot
# set agent as if request came from Firefox browser
headers = requests.utils.default_headers()
headers.update(
    {
        "User-Agent": "Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:52.0) Gecko/20100101 Firefox/52.0",
    }
)

# read input.xlsx and get all parcels
startPointer = 0
in_df = pd.read_excel("input.xlsx", sheet_name="Sheet1", dtype=str)
parcels = in_df["Parcel Number"]

writer = pd.ExcelWriter("taxlien.xlsx", engine="xlsxwriter")

for parcel in parcels:
    print(f"processing parcel# {parcel}...")

    html_text = requests.get(
        f"https://www.calcasieuassessor.org/Details?parcelNumber={str(parcel)}/0",
        headers=headers,
    ).text

    # replace all html line break tags "br" with new line "\n"
    soup = BeautifulSoup(html_text, "lxml")
    for br in soup.find_all("br"):
        br.replace_with("\n")

    parcelDetail = soup.find("div", id="parcelDetails")

    spans = parcelDetail.find_all("span")
    parcelNumber = spans[1].text
    primaryOwner = spans[3].text
    pAddress = spans[5].text
    pType = spans[9].text

    data = [[parcelNumber, primaryOwner, pAddress, pType]]
    df1 = pd.DataFrame(data, columns=["Parcel#", "PrimaryOwner", "Address", "Type"])
    print(tabulate(df1, headers="keys", tablefmt="psql"))
    df1.to_excel(writer, sheet_name="TaxLien", startrow=startPointer, startcol=0)

    # output Parcel Items table
    table = soup.find_all("table")[0]
    df = pd.read_html(str(table))
    print(tabulate(df[0], headers="keys", tablefmt="psql"))
    df[0].to_excel(writer, sheet_name="TaxLien", startrow=startPointer, startcol=5)
    table_0_rows = df[0].shape[0]

    # output Deeds table
    table = soup.find_all("table")[1]
    df = pd.read_html(str(table))
    print(tabulate(df[0], headers="keys", tablefmt="psql"))
    df[0].to_excel(writer, sheet_name="TaxLien", startrow=startPointer, startcol=11)
    table_1_rows = df[0].shape[0]

    # output Ownership History table
    table = soup.find_all("table")[2]
    df = pd.read_html(str(table))
    print(tabulate(df[0], headers="keys", tablefmt="psql"))
    df[0].to_excel(writer, sheet_name="TaxLien", startrow=startPointer, startcol=19)
    table_2_rows = df[0].shape[0]

    startPointer += max(table_0_rows, table_1_rows, table_2_rows) + 3
    print("waiting 60s for next parcel...\n")
    time.sleep(60)

writer.close()
print("All done, check taxlien.xlsx.")
