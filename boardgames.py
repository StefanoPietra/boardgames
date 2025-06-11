#! python3
""" boardgames.py - Update board game prices from boardgameprices.co.uk website,
saving them in a new sheet of the Excel file"""

import logging, openpyxl, datetime, requests, bs4
from openpyxl.styles import PatternFill, Font

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s", datefmt="%H:%M:%S")
# logging.disable(logging.INFO)

### PARAMETERS
website = "https://boardgameprices.co.uk/item/show/"
excelFile = "Data\Boardgames.xlsx"
priceChange = 5     # amount over which the price will be considered changed
redFill = PatternFill(fill_type="solid", start_color="F2DCDB", end_color="F2DCDB")
greenFill = PatternFill(fill_type="solid", start_color="EBF1DE", end_color="EBF1DE")
noFill = PatternFill(fill_type="none")

logging.info("Starting program")

### Create a new sheet
logging.info("Opening Excel file: " + excelFile)
workbook = openpyxl.load_workbook(excelFile)
latestSheet = workbook[workbook.sheetnames[-1]]

logging.info("Building the new sheet name")
dt = datetime.datetime.now()
newSheetName = dt.strftime("%Y-%m")     # year and month, 2024-07
if newSheetName == latestSheet.title :
    newSheetName = newSheetName+"(1)"       # if the month already exists -> 2024-07(1)

logging.info("Duplicating the latest sheet")
newSheet = workbook.copy_worksheet(latestSheet)

logging.info("Renaming the new sheet as: " + newSheetName)
newSheet.title = newSheetName

### Loop over row1 to find column Nr of "Other best"
for cellObj in list(newSheet.rows)[0]:
    if cellObj.value == "Other best":
        otherBestColumn = cellObj.column
        logging.debug('Found "Other best" column: ' + str(otherBestColumn))
        break

### Loop over the games
for row in range(1, newSheet.max_row):
    if newSheet.cell(row=row, column=1).value == None:  # skips empty rows
        continue
    game = newSheet.cell(row=row, column=1).value

    ### Visit website and scrape the game's price and availability
    logging.info("Building URL for " + game)
    url = website + newSheet.cell(row=row, column=2).value

    logging.info("Getting web page")
    logging.debug(url)
    try:
        res = requests.get(url)
        res.raise_for_status()

        logging.info("Parsing page to find cheapest price") # (with shipping)
        gamePage = bs4.BeautifulSoup(res.text, "html.parser")
        otherBestPriceElem = gamePage.select("#vendorlist > div:nth-child(8) > div > div.total.grand-total")
        otherBestPrice = round(float(otherBestPriceElem[0].text[1:]))    # transform price into a number
        logging.info("Best price for " + game + " is " + str(otherBestPrice))
        newSheet.cell(row=row, column=otherBestColumn).value = otherBestPrice
        
    except:
        logging.error("Impossible to retrieve info for " + game)
        continue
    
    logging.info("Checking availability")   # the green light
    availabilityElem = gamePage.select("#vendorlist > div:nth-child(7) > div.infocontainer.multicell > div.vendorstock > span")
    availability = availabilityElem[0].text
    logging.debug(availability)
    if availability == "No":
        logging.info(game + " is NOT available")
        newSheet.cell(row=row, column=otherBestColumn).font = Font(color="FF0000")
    elif availability != "Yes":
        logging.error("Impossible to retrieve availability info for " + game)

    logging.info("Comparing price with the previous sheet")
    priceDifference = otherBestPrice - latestSheet.cell(row=row, column=otherBestColumn).value
    if priceDifference > priceChange:
        newSheet.cell(row=row, column=otherBestColumn).fill = redFill
        logging.info("New price is HIGHER: filling cell red")
    elif priceDifference < -priceChange:
        newSheet.cell(row=row, column=otherBestColumn).fill = greenFill
        logging.info("New price is LOWER: filling cell green")
    else:
        newSheet.cell(row=row, column=otherBestColumn).fill = noFill

logging.info("Saving the Excel file")
workbook.active.views.sheetView[0].tabSelected = False  # disable the current active sheet
workbook.active = newSheet  # set new active sheet
workbook.save(excelFile)
