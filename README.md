**boardgames.py**
=

A little tracker for boardgame prices in Python.

## Purpose
Keeping track of the prices of boardgames on my wishlist by scraping two websites and updating
an Excel file that contains information about each game.

## Features
- Detailed logging of all steps
- New sheet created according to the current year and month
- Scraping of games' prices (including postage) and availability
- Cells formatted based on availability and price variation

## Usage
Reads list of games from the Excel file, which contains specific urls in hidden columns,
and overwrites file adding an additional sheet that highlights changes from the previous version.

## Preview
![Boardgames](https://github.com/user-attachments/assets/de09899f-a999-49ea-8c6a-2dcf9641d34c)
