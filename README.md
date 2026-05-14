# Poke Inv

A macOS desktop app for managing your Pokemon card collection. Track inventory, grading, sales, and photos — all stored locally in an Excel workbook.

## Features

- **Add cards** by pasting a Cardmarket URL (auto-fills name, set, card number) or manually
- **Collection view** with filtering (All / Collection / For Sale / Sold / Grading)
- **Grading tracker** — submit cards for grading and record results
- **Partial selling** — sell part of your stock, track sales with profit calculation
- **eBay integration** — create draft listings directly from your inventory
- **Photo upload** from your phone via local WiFi (built-in web server)
- **Cardmarket links** — double-click a card to open its Cardmarket page
- **Statistics** tab with investment totals, grading overview, and sales summary
- **Excel-based storage** — your data lives in a standard `.xlsx` file you can open anytime

## Requirements

- macOS
- Python 3.10+ (pre-installed on modern macOS, or install via [python.org](https://www.python.org/downloads/))
- Xcode Command Line Tools (`xcode-select --install`) for the native launcher

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

### eBay Listings

1. Go to **eBay Settings** and connect your eBay developer account (one-time setup)
2. Select a card and click **List on eBay**
3. Confirm the price — a draft listing is created in your eBay Seller Hub

## Data Storage

All data is stored locally in `~/Pokemon-Sammlung/`:

| Path | Content |
|------|---------|
| `Pokemon-Inventar.xlsx` | Inventory, sales, and dashboard |
| `Fotos/` | Card photos (named by card ID) |
| `Karten/` | Additional card files |
| `.card_urls.json` | Cardmarket URLs per card |
| `.app_config.json` | Window size and preferences |

## License

MIT
