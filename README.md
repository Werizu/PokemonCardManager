# Pokemon Card Manager

A macOS desktop app for managing your Pokemon card collection. Track inventory, grading, sales, and photos — all stored locally in an Excel workbook.

## Features

- **Gallery with real card images**: every card shows its official artwork (via the free [TCGdex](https://tcgdex.dev) API, in the card's language where available), so cards with the same name are easy to tell apart. Sealed products show the set logo. Wrong match? Right-click a card, **Kartenbild ändern**, and pick the right one.
- **Market prices**: Cardmarket trend and 30-day average per card, refreshed daily; gallery, detail panel and statistics show value and gain/loss against the purchase price (raw cards only, slabs are excluded)
- **Detail panel** for the selected card with all info, own photos and context actions (sell, eBay, grading, edit, delete)
- **Gallery or list view**, search across name/set/number/ID, filters (Bestand, Zum Verkauf, Beim Grading, Vorgemerkt, Gegradet, Verkauft, Alle), sorting by date, name, set, price, market value or gain
- **Add cards** by pasting a Cardmarket URL (also picked up automatically from the clipboard): name, set and number are filled in and the image preview appears with the current Cardmarket trend; all fields stay editable, so manual entry works too
- **Grading board**: queued, submitted and graded cards side by side; submit all queued cards to one service at once
- **Sales view** with all sales, profit per sale, totals and **undo** for mistaken entries
- **Statistics**: stock, total invested, market value, unrealized and realized gain, purchases per month, value by set, most valuable cards
- **Partial selling** with live profit preview; platforms Cardmarket, eBay, Whatnot, Trade, Other
- **eBay integration**: create draft listings or publish directly (fixed price with Best Offer, or auction)
- **Photo upload** from your phone via local WiFi (built-in web page with search and thumbnails)
- **Keyboard shortcuts**: Cmd+1 to Cmd+6 switch sections, Cmd+N add, Cmd+F search, Cmd+R reload, arrow keys move through the gallery, Space shows the image large
- **Auto-update** from GitHub, **Excel-based storage**: your data lives in a standard `.xlsx` file you can open anytime

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

1. Copy a Cardmarket URL and open **Hinzufügen** (Cmd+N); the link is taken from the clipboard automatically
2. Name, set and number fill in, the card image and Cardmarket trend appear on the right
3. Pick language, condition, source, enter price and quantity, press **Enter**
4. Language, condition and source are remembered for the next card

### Photos

When the app is running, it starts a local web server for phone uploads:
1. A QR code and the upload URL are shown at the bottom of the sidebar (e.g. `http://192.168.1.x:8080`); click either to copy the URL
2. Scan the QR code with your phone camera (same WiFi network)
3. Find the card and tap „Foto“ to take or pick a photo

### Selling Cards

1. Select a card in **Sammlung**
2. Click **Verkauf eintragen**
3. Enter quantity, total sale price, platform, fees and your shipping cost; the profit is shown live
4. The inventory quantity decreases and a sales record is created. Mistake? **Verkäufe**, select the sale, **Verkauf rückgängig**

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

1. Select a card in **Sammlung**
2. Click **Auf eBay listen** (the price is pre-filled with the Cardmarket trend)
3. Set your price and quantity
4. Choose:
   - **Save as Draft** — stores the offer in the eBay API for later publishing
   - **Publish Now** — creates a live listing on eBay immediately

The app automatically handles:
- **Condition mapping** — ungraded cards use eBay's card condition descriptors (Near Mint, Excellent, etc.); graded cards include the grading service (PSA, BGS, CGC, ...) and grade
- **Photo upload** — if the card has a photo, it is uploaded to eBay automatically
- **Duplicate detection** — if an offer already exists for the card, it is updated instead of creating a duplicate

## Auto-Update

The app checks for updates on startup; you can also click **Nach Updates suchen** in **Einstellungen** to check for a newer version on GitHub. If an update is available, the app downloads the latest code and restarts automatically.

Only the application code is updated — your data is never touched:

| Preserved (your data) | Updated |
|------------------------|---------|
| `Pokemon-Inventar.xlsx` | `pokemon_card_manager.py` |
| `Fotos/`, `Karten/` | |
| `.ebay_config.json` | |
| `.app_config.json`, `.card_urls.json`, `.card_meta.json` | |
| `.bilder/`, `.cache/` | |

## Data Storage

All data is stored locally in `~/Pokemon-Sammlung/`:

| Path | Content |
|------|---------|
| `Pokemon-Inventar.xlsx` | Inventory, sales, and dashboard |
| `Fotos/` | Card photos (named by card ID) |
| `Karten/` | Card data as JSON files |
| `.card_urls.json` | Cardmarket URLs per card |
| `.app_config.json` | Window size and preferences |
| `.card_meta.json` | Matched TCGdex card per entry and cached market prices |
| `.bilder/` | Downloaded official card images |
| `.cache/` | Set lists and thumbnails (safe to delete) |
| `.ebay_config.json` | eBay credentials and tokens (do not share) |

## License

MIT
