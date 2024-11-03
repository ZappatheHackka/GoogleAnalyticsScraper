import datetime
import xlsxwriter
from typing import Dict
from google.oauth2 import service_account
from googleapiclient.discovery import build, Resource
from googleapiclient.errors import HttpError

# Constants for query data
CLICKS = []
IMPRESSIONS = []
CTR = []
KEYS = []

# Constants for previous month's query data
MAclicks = []
MAimps = []
MActr = []
MAkeys = []

# Constants for page data
PAGE_CLICKS = []
PAGE_IMPS = []
PAGE_CTR = []
PAGES = []

# Constants for previous month's page data
PAGE_MA_PAGES = []
PAGE_MA_CLICKS = []
PAGE_MA_IMPS = []
PAGE_MA_CTR = []

# define general constants
API_SERVICE_NAME = "webmasters"
API_VERSION = "v3"
SCOPE = ["https://www.googleapis.com/auth/webmasters"]

# getting the range of dates I want to query
date = datetime.datetime.now()
monthago = date - datetime.timedelta(days=30)
twomonthago = date - datetime.timedelta(days=60)
monthago = monthago.strftime("%Y-%m-%d")
date = date.strftime("%Y-%m-%d")
twomonthago = twomonthago.strftime('%Y-%m-%d')


def auth_using_key_file(key_filepath: str) -> Resource:
    """Authenticate using a service account key file saved locally"""

    credentials = service_account.Credentials.from_service_account_file(
        key_filepath, scopes=SCOPE
    )
    service = build(API_SERVICE_NAME, API_VERSION, credentials=credentials)

    return service


def list_sites(service: Resource):
    """List sites from Search Console to verify permissions."""
    try:
        response = service.sites().list().execute()
        if 'siteEntry' in response and response['siteEntry']:
            return response['siteEntry']
        else:
            print("No sites found for this service account.")
            return []
    except HttpError as err:
        print(f"API Error: {err}")
        raise


def query(client: Resource, payload: Dict[str, str]) -> Dict[str, any]:
    response = client.searchanalytics().query(siteUrl=domain, body=payload).execute()
    return response


# filepath location of your service account key json file
KEY_FILE = "YOUR SECRET"

# authenticate session
service = auth_using_key_file(key_filepath=KEY_FILE)

# verify your service account has permissions to your domain - had to print
results = service.sites().list().execute()
domain = results['siteEntry'][0]['siteUrl']

# search for which queries result in which page click-through.
payload = {
        "startDate": monthago,
        "endDate": date,
        "dimensions": ['query'],
        "rowLimit": 25_000,
        "startRow": 0 * 25_000,
}

response = query(client=service, payload=payload)


# Organizing the json data
response = response['rows']
# gather only the rows that have clicks > 1
# print(response[0])

# Compiling data for the QUERY sheet of the xl spreadsheet
# response is list of dicts holding click, impression, kw, and ctr data for every query
for result in response:
    if result['clicks'] >= 4:   # Ensuring the data captured is from keywords resulting in at least 4 clicks
        keyword = result['keys']
        KEYS.append(keyword)

        clicks = result['clicks']
        CLICKS.append(clicks)

        impress = result['impressions']
        IMPRESSIONS.append(impress)

        ctr = result['ctr']
        CTR.append(ctr)


# Converting every item of KEYS to a string
KEYS = [str(key) for key in KEYS]
KEYS = [key[1:-1] for key in KEYS]
# Rounding to 2 decimal places for CTR
CTR = [round(n, 2) for n in CTR]

# Calculating totals & averages
total_clicks = sum(CLICKS)
total_imps = sum(IMPRESSIONS)
average_ctr = sum(CTR) / len(CTR)
average_ctr = round(average_ctr, 2)

# Gathering previous month data for comparison

payload = {
        "startDate": twomonthago,
        "endDate": monthago,
        "dimensions": ['query'],
        "rowLimit": 25_000,
        "startRow": 0 * 25_000,
}


response = query(client=service, payload=payload)
response = response['rows']


for result in response:
    if result['clicks'] >= 4:   # Ensuring the data captured is from keywords resulting in at least 4 clicks
        keyword = result['keys']
        MAkeys.append(keyword)

        clicks = result['clicks']
        MAclicks.append(clicks)

        impress = result['impressions']
        MAimps.append(impress)

        ctr = result['ctr']
        MActr.append(ctr)

# Converting every item of KEYS to a string
MAkeys = [str(key) for key in MAkeys]
MAkeys = [key[1:-1] for key in MAkeys]
# Rounding to 2 decimal places for CTR
MActr = [round(n, 2) for n in MActr]

# Calculating totals & averages
ma_total_clicks = sum(MAclicks)
ma_total_imps = sum(MAimps)
ma_average_ctr = sum(MActr) / len(MActr)
ma_average_ctr = round(ma_average_ctr, 2)

# Getting % changes from month to month -> ((new - old) / old) * 100 and rounding + adding "%" sign
chng_in_clicks = ((total_clicks - ma_total_clicks) / ma_total_clicks) * 100
print(f"total clicks: {total_clicks}, {ma_total_clicks}")

chng_in_imps = ((total_imps - ma_total_imps) / ma_total_imps) * 100
chng_in_ctr = ((average_ctr - ma_average_ctr) / ma_average_ctr) * 100

chng_in_clicks = str(round(chng_in_clicks, 2)) + "%"
chng_in_imps = str(round(chng_in_imps, 2)) + '%'
chng_in_ctr = str(round(chng_in_ctr, 2)) + '%'

# --------------BEGINNING XLSX WRITING FOR QUERIES-----------------

# Creating xlsx file with scraped data
workbook = xlsxwriter.Workbook(f'{monthago} to {date} Web Traffic Data.xlsx')
worksheet1 = workbook.add_worksheet('Queries & Keywords')

# Styling my header cells
header_format = workbook.add_format({
    'bold': True,
    'align': 'center',
    'valign': 'center',
    'bg_color': '#Ff6600',
    'font_color': '#13124B'
})

# Adding left alignment options for text data
left_align_format = workbook.add_format({'align': 'left'})

# Defining headers
headers = ['Keyword', 'Clicks', 'Impressions', 'CTR', 'Total Clicks',
           'Total Impressions', 'Average CTR', 'Last Month\'s Total Impressions',
           'Last Month\'s Average CTR', 'Last Month\'s Total Clicks', '% Change in Imps from Last Month',
           '% Change in CTR from Last Month', '% Change in Clicks from Last Month']


# Looping through headers and writing header to correlating index value in xlsx
for col, header in enumerate(headers):
    if 4 <= col <= 6:
        row = len(KEYS) + 3
        # Write to columns 0-2 instead of adjusting the column number
        worksheet1.write(row, col % 4, header, header_format)
    elif 7 <= col < 10:
        row = len(KEYS) + 7
        # Write to columns 0-2 instead of adjusting the column number
        worksheet1.write(row, col % 3, header, header_format)
    elif col >= 10:
        row = len(KEYS) + 11
        # Write to columns 0-2 instead of adjusting the column number
        worksheet1.write(row, col % 3, header, header_format)
    else:
        worksheet1.write(0, col, header, header_format)

# Writing keywords to document
for index, keyword in enumerate(KEYS):
    row = index + 1
    worksheet1.write(row, 0, keyword)

# Adjusting width of columns
worksheet1.set_column(0, len(headers), width=38)

# Adding number of clicks to column 2
for index, num in enumerate(CLICKS):
    row = index + 1
    worksheet1.write(row, 1, num, left_align_format)

# Adding impressions to column 3
for index, num in enumerate(IMPRESSIONS):
    row = index + 1
    worksheet1.write(row, 2, num, left_align_format)

# Adding CTR to column 4
for index, num in enumerate(CTR):
    row = index + 1
    worksheet1.write(row, 3, num, left_align_format)

# Adding total clicks to Headers in row 9
worksheet1.write(10, 0, total_clicks, left_align_format)

worksheet1.write(10, 1, total_imps, left_align_format)

worksheet1.write(10, 2, average_ctr, left_align_format)

# Adding last month's data
worksheet1.write(14, 0, ma_total_clicks, left_align_format)

worksheet1.write(14, 1, ma_total_imps, left_align_format)

worksheet1.write(14, 2, ma_average_ctr, left_align_format)

# Adding month-to-month calcs
worksheet1.write(18, 0, chng_in_clicks, left_align_format)

worksheet1.write(18, 1, chng_in_imps, left_align_format)

worksheet1.write(18, 2, chng_in_ctr, left_align_format)

# --------------WRITING PAGES CLICKED TAB------------------

# Querying GSC for Page data
payload = {
        "startDate": monthago,
        "endDate": date,
        "dimensions": ['page'],
        "rowLimit": 25_000,
        "startRow": 0 * 25_000,
}

response = query(client=service, payload=payload)
response = response['rows']

# Looping through gsc response and compiling data into constants
for entry in response:
    if entry['clicks'] >= 10:    # Ensuring only entries with >5 clicks are noted
        clicks = entry['clicks']
        PAGE_CLICKS.append(clicks)

        imps = entry['impressions']
        PAGE_IMPS.append(imps)

        page = entry['keys']
        PAGES.append(page)

        ctr = entry['ctr']
        PAGE_CTR.append(ctr)

# Cleaning data
PAGE_CTR = [round(ctr, 2) for ctr in PAGE_CTR]
PAGES = [str(page) for page in PAGES]
PAGES = [page[1:-1] for page in PAGES]

# print(PAGE_CTR)
# print(PAGES)
# print(PAGE_IMPS)
# print(PAGE_CLICKS)

# Combining idhoops.com entries (http + https)
if "'http://idhoops.com/'" and "'https://idhoops.com/'" in PAGES:
    rm = PAGES.index("'http://idhoops.com/'")
    keep = PAGES.index("'https://idhoops.com/'")

    new_ctr = PAGE_CTR[rm] + PAGE_CTR[keep]  # Updating values to reflect http / https combining
    PAGE_CTR[keep] = new_ctr
    PAGE_CTR.remove(PAGE_CTR[rm])

    new_imps = PAGE_IMPS[rm] + PAGE_IMPS[keep]
    PAGE_IMPS[keep] = new_imps
    PAGE_IMPS.remove(PAGE_IMPS[rm])

    new_clicks = PAGE_CLICKS[rm] + PAGE_CLICKS[keep]
    PAGE_CLICKS[keep] = new_clicks
    PAGE_CLICKS.remove(PAGE_CLICKS[rm])

    PAGES.remove(PAGES[rm])

# print(PAGE_CTR)
# print(PAGES)
# print(PAGE_IMPS)
# print(PAGE_CLICKS)

# Adding new tab to xl spreadsheet
worksheet2 = workbook.add_worksheet('Page Data')

# Defining headers, writing, & formatting them
headers = ['Top Pages', 'Clicks', 'Impressions', 'CTR', 'Last Month\'s Top Pages', 'Last Month\'s Clicks',
           'Last Month\'s Impressions', 'Last Month\'s CTR']

for col, header in enumerate(headers):
    row = len(PAGES) + 4
    if col >= 4:
        worksheet2.write(row, col - 4, header, header_format)
    else:
        worksheet2.write(0, col, header, header_format)

# Ensuring columns are wide enough to fit all URL text
worksheet2.set_column(0, len(headers), width=44)

# Writing top pages into the 'Top Pages' column
for col, page, in enumerate(PAGES):
    worksheet2.write(col + 1, 0, page, left_align_format)

# Writing clicks
for col, click in enumerate(PAGE_CLICKS):
    worksheet2.write(col + 1, 1, click, left_align_format)

# Writing Imps
for col, imp in enumerate(PAGE_IMPS):
    worksheet2.write(col + 1, 2, imp, left_align_format)

# Writing Page CTR
for col, ctr, in enumerate(PAGE_CTR):
    worksheet2.write(col + 1, 3, ctr, left_align_format)

# Getting & compiling last month's data

payload = {
        "startDate": twomonthago,
        "endDate": monthago,
        "dimensions": ['page'],
        "rowLimit": 25_000,
        "startRow": 0 * 25_000,
}

response = query(client=service, payload=payload)
response = response['rows']

for entry in response:
    if entry['clicks'] >= 10:    # Ensuring only entries with >5 clicks are noted
        clicks = entry['clicks']
        PAGE_MA_CLICKS.append(clicks)

        imps = entry['impressions']
        PAGE_MA_IMPS.append(imps)

        page = entry['keys']
        PAGE_MA_PAGES.append(page)

        ctr = entry['ctr']
        PAGE_MA_CTR.append(ctr)

# Cleaning the Data
PAGE_MA_CTR = [round(ctr, 2) for ctr in PAGE_MA_CTR]
PAGE_MA_PAGES = [str(page) for page in PAGE_MA_PAGES]
PAGE_MA_PAGES = [page[1:-1] for page in PAGE_MA_PAGES]

# Combining idhoops.com entries into one
if "'http://idhoops.com/'" and "'https://idhoops.com/'" in PAGE_MA_PAGES:
    rm = PAGE_MA_PAGES.index("'http://idhoops.com/'")
    keep = PAGE_MA_PAGES.index("'https://idhoops.com/'")

    new_ctr = PAGE_MA_CTR[rm] + PAGE_MA_CTR[keep]  # Updating values to reflect http / https combining
    PAGE_MA_CTR[keep] = new_ctr
    PAGE_MA_CTR.remove(PAGE_MA_CTR[rm])

    new_imps = PAGE_MA_IMPS[rm] + PAGE_MA_IMPS[keep]
    PAGE_MA_IMPS[keep] = new_imps
    PAGE_MA_IMPS.remove(PAGE_MA_IMPS[rm])

    new_clicks = PAGE_MA_CLICKS[rm] + PAGE_MA_CLICKS[keep]
    PAGE_MA_CLICKS[keep] = new_clicks
    PAGE_MA_CLICKS.remove(PAGE_MA_CLICKS[rm])

    PAGE_MA_PAGES.remove(PAGE_MA_PAGES[rm])

# Writing top pages into the 'Top Pages' column
for col, page, in enumerate(PAGE_MA_PAGES):
    row = col + 11
    worksheet2.write(row, 0, page, left_align_format)

# Writing clicks
for col, click in enumerate(PAGE_MA_CLICKS):
    row = col + 11
    worksheet2.write(row, 1, click, left_align_format)

# Writing Imps
for col, imp in enumerate(PAGE_MA_IMPS):
    row = col + 11
    worksheet2.write(row, 2, imp, left_align_format)

# Writing Page CTR
for col, ctr, in enumerate(PAGE_MA_CTR):
    row = col + 11
    worksheet2.write(row, 3, ctr, left_align_format)

workbook.close()
