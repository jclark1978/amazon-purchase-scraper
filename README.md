# Amazon Purchase Scraper

A Chrome extension that scrapes your Amazon order history and exports it to an XLSX spreadsheet — with computed return eligibility status, days until return deadline, and ASIN extraction.

## What It Exports

| Column | Description |
|---|---|
| Status | Eligible, Urgent - Return Soon, Window Closed, Returned, Not Yet Arrived, Unknown |
| Days Until Deadline | Countdown to return window close date |
| Order Date | Date the order was placed |
| Order Total | Order total in dollars |
| Order Number | Amazon order ID |
| Items | Pipe-delimited list of product titles |
| ASINs | Pipe-delimited list of Amazon ASINs |
| Order Link | Clickable link to the order detail page |
| Return Eligible | Yes / No |
| Return Date | Return window open/close date |
| Notes | Empty column for your own notes |

## Installation

This extension is not published to the Chrome Web Store — you load it manually as an unpacked extension.

1. **Download the code**
   - Click the green **Code** button on this page → **Download ZIP**
   - Unzip the folder somewhere permanent (e.g. `Documents/amazon-scraper`)
   - Or clone the repo: `git clone https://github.com/jclark1978/amazon-purchase-scraper.git`

2. **Open Chrome Extensions**
   - In Chrome, go to `chrome://extensions`
   - Or: Chrome menu → More Tools → Extensions

3. **Enable Developer Mode**
   - Toggle **Developer mode** on (top-right corner of the Extensions page)

4. **Load the extension**
   - Click **Load unpacked**
   - Select the folder you unzipped/cloned (the one containing `manifest.json`)

5. **Pin it (optional)**
   - Click the puzzle piece icon in the Chrome toolbar
   - Click the pin icon next to **Amazon Purchase Scraper**

## Usage

1. Go to your Amazon order history: [amazon.com/gp/css/order-history](https://www.amazon.com/gp/css/order-history)
2. Click the extension icon in your toolbar
3. Choose a scrape mode:
   - **Scrape Current Page** — scrapes only the orders visible on the current page
   - **Scrape All Pages (Auto Next)** — automatically pages through all of your order history and downloads a single combined file when done
4. The XLSX file will download automatically

## Updating

If you downloaded the ZIP, re-download and replace the folder, then click the refresh icon on `chrome://extensions`.

If you cloned with git, run `git pull` in the folder, then click the refresh icon on `chrome://extensions`.
