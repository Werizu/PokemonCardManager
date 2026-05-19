# Pokemon Card Manager

A macOS desktop app for managing your Pokemon card collection. Track inventory, grading, sales, and photos — all stored locally in an Excel workbook.

## Features

- **Add cards** by pasting a Cardmarket URL (auto-fills name, set, card number) or manually
- **Collection view** with filtering (All / Collection / For Sale / Sold / Grading)
- **Grading tracker** — submit cards for grading and record results
- **Partial selling** — sell part of your stock, track sales with profit calculation
- **eBay integration** — create draft listings or publish directly from your inventory
- **Photo upload** from your phone via local WiFi (built-in web server)
- **Cardmarket links** — double-click a card to open its Cardmarket page
- **Statistics** tab with investment totals, grading overview, and sales summary
- **Auto-update** — check for new versions directly from the app
- **Excel-based storage** — your data lives in a standard `.xlsx` file you can open anytime

## Requirements

- macOS
- Python 3.10+
- Dependencies: `openpyxl`, `Pillow`

## Installation

```bash
git clone https://github.com/Werizu/PokemonCardManager.git
cd PokemonCardManager
chmod +x install.sh
./install.sh
```

This will:
1. Create `~/Pokemon-Sammlung/` with your data folders
2. Set up a Python virtual environment and install dependencies
3. Create an empty Excel workbook with the correct structure
4. Place a **Poke Inv** app on your Desktop (with icon)

## Usage

Double-click **Poke Inv** on your Desktop, or run:

```bash
~/Pokemon-Sammlung/.venv/bin/python3 ~/Pokemon-Sammlung/pokemon_card_manager.py
```

### Adding Cards

1. Go to the **Add** tab
2. Paste a Cardmarket URL — name, set, and card number fill automatically
3. Select language, condition, source, and enter price/quantity
4. Click **Save**

### Photos

When the app is running, it starts a local web server for phone uploads:
1. Look at the bottom of the window for the upload URL (e.g. `http://192.168.1.x:8080`)
2. Open that URL on your phone (same WiFi network)
3. Select the card ID and upload a photo

### Selling Cards

1. Select a card in the **Collection** tab
2. Click **Mark Sold**
3. Enter sale price, platform, fees, and quantity to sell
4. The inventory quantity decreases; a sales record is created

## eBay Integration

The app can create listings on eBay directly from your inventory — either as API drafts or published live.

### One-Time Setup

#### 1. Create an eBay Developer Account

1. Go to [developer.ebay.com](https://developer.ebay.com) and sign in with your eBay account
2. Go to **My Account > Application Access Keys**
3. Click **Create a keyset** and choose a name (e.g. "Pokemon Card Manager")
4. If your keyset shows "disabled", click **Apply for exemption** under Marketplace Account Deletion and select "I do not persist eBay data"
5. Note your **App ID (Client ID)** and **Cert ID (Client Secret)**

#### 2. Configure OAuth Redirect URL

1. On the Application Keys page, click **User Tokens** next to your keyset
2. Select **OAuth (new security)**
3. Click **Get a Token from eBay via Your Application**
4. Click **Add eBay Redirect URL** and enter:
   - **Auth Accepted URL**: `https://yourdomain.com/ebay/callback` (any HTTPS URL you own)
   - **Auth Declined URL**: `https://yourdomain.com/ebay/declined`
   - **Privacy Policy URL**: `https://yourdomain.com/privacy`
   - **Display Title**: e.g. "Pokemon Card Manager"
5. Save — eBay will generate a **RuName** (e.g. `Your_Name-AppName-PRD-xxxxxx`)

> **Note:** eBay requires a real HTTPS domain — `localhost` is not accepted. The redirect URL does not need to actually host anything. After authorizing, you simply copy the URL from your browser and paste it into the app.

#### 3. Connect in the App

1. Open the app and click **eBay Settings**
2. Enter your **Client ID**, **Client Secret**, and **RuName**
3. Select your **Marketplace** (e.g. `EBAY_DE` for Germany)
4. Click **Save & Connect**
5. Your browser opens the eBay login page — log in and grant access
6. After authorization, eBay redirects to your redirect URL. The page may show an error — **that's normal**
7. **Copy the full URL** from your browser's address bar (it contains `?code=...`)
8. Paste it into the dialog in the app
9. Done — the app exchanges the code for an access token automatically

### Creating Listings

1. Select a card in the **Collection** tab
2. Click **List on eBay**
3. Set your price and quantity
4. Choose:
   - **Save as Draft** — stores the offer in the eBay API for later publishing
   - **Publish Now** — creates a live listing on eBay immediately

The app automatically handles:
- **Condition mapping** — ungraded cards use eBay's card condition descriptors (Near Mint, Excellent, etc.); graded cards include the grading service (PSA, BGS, CGC, ...) and grade
- **Photo upload** — if the card has a photo, it is uploaded to eBay automatically
- **Duplicate detection** — if an offer already exists for the card, it is updated instead of creating a duplicate

## Auto-Update

Click **Check for Updates** in the Collection tab to check for a newer version on GitHub. If an update is available, the app downloads the latest code and restarts automatically.

Only the application code is updated — your data is never touched:

| Preserved (your data) | Updated |
|------------------------|---------|
| `Pokemon-Inventar.xlsx` | `pokemon_card_manager.py` |
| `Fotos/`, `Karten/` | |
| `.ebay_config.json` | |
| `.app_config.json`, `.card_urls.json` | |

## Data Storage

All data is stored locally in `~/Pokemon-Sammlung/`:

| Path | Content |
|------|---------|
| `Pokemon-Inventar.xlsx` | Inventory, sales, and dashboard |
| `Fotos/` | Card photos (named by card ID) |
| `Karten/` | Card data as JSON files |
| `.card_urls.json` | Cardmarket URLs per card |
| `.app_config.json` | Window size and preferences |
| `.ebay_config.json` | eBay credentials and tokens (do not share) |

## License

MIT
