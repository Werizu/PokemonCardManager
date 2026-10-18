#!/usr/bin/env python3
import base64
import hashlib
import html
import io
import json
import os
import queue
import re
import shutil
import socket
import subprocess
import sys
import threading
import time as _time
import tkinter as tk
import tkinter.font as tkfont
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
from http.server import HTTPServer, BaseHTTPRequestHandler
from tkinter import ttk, messagebox, simpledialog, filedialog
from datetime import date, datetime
from urllib.parse import parse_qs

from openpyxl import load_workbook
from PIL import Image, ImageTk, ImageOps, ImageDraw, ImageEnhance

VERSION = "2.5.2"
GITHUB_RAW_URL = "https://raw.githubusercontent.com/Werizu/PokemonCardManager/main/pokemon_card_manager.py"
GITHUB_API_URL = "https://api.github.com/repos/Werizu/PokemonCardManager/contents/pokemon_card_manager.py?ref=main"

IS_MAC = sys.platform == "darwin"

# POKEINV_HOME erlaubt eine Test-Sammlung neben der echten.
BASE_DIR = os.environ.get("POKEINV_HOME") or os.path.join(os.path.expanduser("~"), "Pokemon-Sammlung")
EXCEL_PATH = os.path.join(BASE_DIR, "Pokemon-Inventar.xlsx")
KARTEN_DIR = os.path.join(BASE_DIR, "Karten")
CONFIG_PATH = os.path.join(BASE_DIR, ".app_config.json")
CACHE_DIR = os.path.join(BASE_DIR, ".cache")
IMG_DIR = os.path.join(BASE_DIR, ".bilder")
META_PATH = os.path.join(BASE_DIR, ".card_meta.json")

SETS = [
    "151", "30th Celebration", "Ancient Origins", "Aquapolis", "Arceus",
    "Ascended Heroes",
    "Astral Radiance", "Base Set", "Base Set 2", "Battle Styles",
    "Black & White", "Black Bolt", "Boundaries Crossed", "BREAKpoint",
    "BREAKthrough", "Brilliant Stars", "Burning Shadows", "Call of Legends",
    "Celebrations", "Celestial Storm", "Champion's Path", "Chaos Rising",
    "Chilling Reign",
    "Cosmic Eclipse", "Crimson Invasion", "Crown Zenith", "Dark Explorers",
    "Darkness Ablaze", "Delta Reign", "Destined Rivals", "Detective Pikachu",
    "Diamond & Pearl", "Double Crisis", "Dragon Majesty", "Dragon Vault",
    "Dragons Exalted", "Emerging Powers", "Evolutions", "Evolving Skies",
    "EX Crystal Guardians", "EX Delta Species", "EX Deoxys", "EX Dragon",
    "EX Dragon Frontiers", "EX Emerald", "EX FireRed & LeafGreen",
    "EX Hidden Legends", "EX Holon Phantoms", "EX Legend Maker",
    "EX Power Keepers", "EX Ruby & Sapphire", "EX Sandstorm",
    "EX Team Magma vs Team Aqua", "EX Team Rocket Returns",
    "EX Unseen Forces", "Expedition Base Set", "Fates Collide", "Flashfire",
    "Forbidden Light", "Fossil", "Furious Fists", "Fusion Strike",
    "Generations", "Great Encounters", "Guardians Rising", "Gym Challenge",
    "Gym Heroes", "HeartGold & SoulSilver", "Hidden Fates", "HS Triumphant",
    "HS Undaunted", "HS Unleashed", "Journey Together", "Jungle",
    "Kalos Starter Set", "Legendary Collection", "Legendary Treasures",
    "Legends Awakened",
    "Lost Origin", "Lost Thunder", "Majestic Dawn", "Mega Evolution",
    "Mysterious Treasures", "Neo Destiny", "Neo Discovery", "Neo Genesis",
    "Neo Revelation", "Next Destinies", "Noble Victories", "Obsidian Flames",
    "Other", "Paldea Evolved", "Paldean Fates", "Paradox Rift",
    "Perfect Order", "Phantasmal Flames", "Phantom Forces", "Pitch Black",
    "Plasma Blast", "Plasma Freeze", "Plasma Storm", "Platinum", "Pokemon GO",
    "Primal Clash",
    "Prismatic Evolutions", "Promo", "Rebel Clash", "Rising Rivals",
    "Roaring Skies", "Scarlet & Violet", "Secret Wonders", "Shining Fates",
    "Shining Legends", "Shrouded Fable", "Silver Tempest", "Skyridge",
    "Southern Islands", "Steam Siege", "Stellar Crown", "Stormfront",
    "Sun & Moon",
    "Supreme Victors", "Surging Sparks", "Sword & Shield", "Team Rocket",
    "Team Up", "Temporal Forces", "Twilight Masquerade", "Ultra Prism",
    "Unbroken Bonds", "Unified Minds", "Vivid Voltage", "White Flare", "XY",
]

LANGUAGES = ["EN", "DE", "JP", "KR", "FR", "IT", "ES", "CN", "PT"]
CONDITIONS = ["Mint", "Near Mint", "Excellent", "Good", "Light Played", "Played", "Poor"]
SOURCES = ["Cardmarket", "eBay", "Convention", "Trade", "Booster", "Other"]
PLATFORMS = ["Cardmarket", "eBay", "Whatnot", "Trade", "Other"]
GRADING_SERVICES = ["PSA", "CGC", "Beckett", "ACE", "TAG", "Other"]

# Col mapping: 1=ID, 2=Name, 3=Set, 4=Code, 5=Lang, 6=Condition,
# 7=Grading Status, 8=Grading Service, 9=Grade, 10=Cert#,
# 11=Purchase Price, 12=Purchase Date, 13=Source,
# 14=Quantity, 15=Status, 16=Grading Vormerkung, 17=Vormerkung Notiz,
# 18=Vorheriger Service, 19=Vorheriger Grade, 20=Vorheriger Cert#
# Sales: 1=ID, 2=Name, 3=Datum, 4=Plattform, 5=Preis, 6=Gebuehren,
# 7=Versand, 8=Netto (Formel), 9=Einkauf, 10=Gewinn (Formel),
# 11=Kaeufer, 12=Tracking, 13=Anzahl

GRADE_ALIASES = {
    "P10": "Pristine 10", "P 10": "Pristine 10",
    "B10": "Black Label 10", "B 10": "Black Label 10",
}

# eBay API
EBAY_CONFIG_PATH = os.path.join(BASE_DIR, ".ebay_config.json")
EBAY_SCOPES = ("https://api.ebay.com/oauth/api_scope/sell.inventory"
               " https://api.ebay.com/oauth/api_scope/sell.account")

CONDITION_TO_EBAY_CARD = {
    "Mint": "400010", "Near Mint": "400010",
    "Excellent": "400015", "Good": "400016",
    "Light Played": "400016", "Played": "400017", "Poor": "400017",
}

GRADING_SERVICE_TO_EBAY = {
    "PSA": "275010", "BCCG": "275011", "BVG": "275012", "BGS": "275013",
    "BECKETT": "275013",
    "CGC": "275015", "SGC": "275016", "KSA": "275017", "GMA": "275018",
    "HGA": "275019", "ISA": "2750110", "PCA": "2750111", "GSG": "2750112",
    "PGS": "2750113", "MNT": "2750114", "TAG": "2750115", "Rare": "2750116",
    "RCG": "2750117", "PCG": "2750118", "ACE": "2750119", "CGA": "2750120",
    "TCG": "2750121", "ARK": "2750122", "AGS": "2750124", "DSG": "2750125",
}

GRADE_TO_EBAY = {
    "10": "275020", "9.5": "275021", "9": "275022", "8.5": "275023",
    "8": "275024", "7.5": "275025", "7": "275026", "6.5": "275027",
    "6": "275028", "5.5": "275029", "5": "2750210", "4.5": "2750211",
    "4": "2750212", "3.5": "2750213", "3": "2750214", "2.5": "2750215",
    "2": "2750216", "1.5": "2750217", "1": "2750218",
}

SEALED_WORDS = {"booster", "bundle", "display", "box", "elite", "trainer", "case",
                "pack", "tin", "collection", "premium", "blister", "build", "battle"}


def parse_grade(grade_str):
    upper = grade_str.upper().strip()
    if upper in GRADE_ALIASES:
        return GRADE_ALIASES[upper]
    return float(grade_str.replace(",", "."))


def fmt_grade(grade):
    if grade in (None, ""):
        return ""
    if isinstance(grade, float) and grade.is_integer():
        return str(int(grade))
    return str(grade)


def fmt_eur(value, sign=False):
    if value is None:
        return "–"
    s = f"{abs(value):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    if sign:
        prefix = "+" if value > 0.004 else ("−" if value < -0.004 else "±")
        return f"{prefix}{s} €"
    return f"−{s} €" if value < -0.004 else f"{s} €"


def fmt_date(d):
    if not d:
        return "–"
    if isinstance(d, datetime):
        d = d.date()
    if isinstance(d, date):
        return d.strftime("%d.%m.%Y")
    return str(d)


def parse_eur(text, default=None):
    text = (text or "").strip().replace("€", "").replace(" ", "")
    if not text:
        if default is None:
            raise ValueError("empty")
        return default
    if "," in text:
        text = text.replace(".", "").replace(",", ".")
    return float(text)


def is_sealed_name(name, number=""):
    if number:
        return False
    words = set(re.split(r"[\s\-]+", (name or "").lower()))
    return bool(words & SEALED_WORDS)


def _norm_words(text):
    text = text.lower().replace("'", "").replace("’", "")
    return [w for w in text.split() if w != "&"]


def _match_set_name(slug):
    set_name = slug.replace("-", " ").replace(" and ", " & ")
    slug_words = _norm_words(set_name)
    slug_set = set(slug_words)

    best_set = None
    best_rank = None
    for s in SETS:
        s_words = _norm_words(s)
        if s_words == slug_words:
            return s
        # Nur Sets akzeptieren, deren Name KOMPLETT im Slug steckt. Sonst
        # gewinnt bei "Pitch-Black" schon ein einzelnes Wort und macht
        # faelschlich "Black & White" daraus.
        if not s_words or not all(w in slug_set for w in s_words):
            continue
        # Laengster/spezifischster Treffer gewinnt: "Base-Set-2" -> "Base Set 2",
        # nicht "Base Set"; "XY-Evolutions" -> "Evolutions", nicht "XY".
        rank = (len(s_words), sum(len(w) for w in s_words))
        if best_rank is None or rank > best_rank:
            best_rank = rank
            best_set = s
    # Unbekanntes Set (z.B. brandneue Erweiterung): lieber den Slug-Namen
    # uebernehmen als es einem falschen Set zuzuordnen.
    return best_set or set_name


def parse_cardmarket_url(url):
    url = url.strip()

    # Singles: /Products/Singles/{SetSlug}/{CardSlug}
    m_singles = re.search(
        r"cardmarket\.com/\w+/Pokemon/Products/Singles/([^/]+)/([^/?#]+)", url
    )
    if m_singles:
        set_slug = m_singles.group(1)
        card_slug = m_singles.group(2)
        parts = card_slug.split("-")
        card_code = ""
        known_variants = {"V1", "V2", "V3", "V4", "RH", "FA", "SAR", "SIR", "IR",
                          "AR", "HR", "SR", "UR", "RR", "GG", "TG", "GR"}
        if parts and re.match(r'^[A-Z]{2,5}\w*\d+$', parts[-1]):
            card_code = parts.pop()
        elif parts and parts[-1].isdigit():
            card_code = parts.pop()
        if parts and parts[-1] in known_variants:
            parts.pop()
        elif len(parts) >= 2 and f"{parts[-2]}-{parts[-1]}" == "Alt-Art":
            parts.pop()
            parts.pop()
        return {
            "name": " ".join(parts),
            "set": _match_set_name(set_slug),
            "number": card_code,
        }

    # Sealed products: /Products/{Type}/{ProductSlug}  (Boosters, Booster-Boxes, etc.)
    m_sealed = re.search(
        r"cardmarket\.com/\w+/Pokemon/Products/([^/]+)/([^/?#]+)", url
    )
    if m_sealed:
        product_type = m_sealed.group(1)
        if product_type == "Singles":
            return None
        product_slug = m_sealed.group(2)
        parts = product_slug.split("-")

        product_labels = {"Booster", "Bundle", "Display", "Box", "Elite", "Trainer",
                          "Case", "Pack", "Tin", "Mini", "Collection", "Premium",
                          "Pin", "Set", "10"}
        name_parts = []
        set_parts = []
        hit_label = False
        for p in parts:
            if p in product_labels:
                hit_label = True
            if hit_label:
                name_parts.append(p)
            else:
                set_parts.append(p)

        if not set_parts:
            set_parts = parts
            name_parts = parts

        set_slug = "-".join(set_parts)
        product_name = " ".join(name_parts) if name_parts else " ".join(parts)

        # Cardmarket-Slugs nutzen "Booster Box" (EN), die dt. Seite nennt
        # dasselbe Produkt "Display" - Sammlerjargon, kein eigener Produkttyp.
        if product_type == "Booster-Boxes" and product_name == "Booster Box":
            product_name = "Display"

        return {
            "name": product_name,
            "set": _match_set_name(set_slug),
            "number": "",
        }

    return None


# ============================================================
# EXCEL OPERATIONS
# ============================================================
# Frueher waren alle Schleifen fest auf Zeile 2-201 begrenzt: ab Karte 201
# hat save_card dann Zeile 2 ueberschrieben. Jetzt wird bis max_row gelesen.

def _to_int(v):
    try:
        return int(v)
    except (TypeError, ValueError):
        return None


def _data_rows(ws):
    return range(2, ws.max_row + 1)


def _find_row(ws, card_id):
    for r in _data_rows(ws):
        if _to_int(ws.cell(row=r, column=1).value) == card_id:
            return r
    return None


def _first_empty_row(ws):
    for r in _data_rows(ws):
        if ws.cell(row=r, column=1).value is None:
            return r
    return ws.max_row + 1


def _num(v, default=0.0):
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def _row_to_card(ws, r):
    c = lambda col: ws.cell(row=r, column=col).value
    qty = c(14)
    return {
        "row": r,
        "id": int(c(1)),
        "name": str(c(2) or ""),
        "set": str(c(3) or ""),
        "number": str(c(4) or ""),
        "language": c(5) or "",
        "condition": c(6) or "Near Mint",
        "grading_status": c(7) or "Not Graded",
        "grading_service": c(8) or "",
        "grade": c(9) if c(9) is not None else "",
        "cert": str(c(10) or ""),
        "price": _num(c(11)),
        "purchase_date": c(12),
        "source": c(13) or "",
        "quantity": int(_num(qty, 1)) if qty is not None else 1,
        "status": c(15) or "Collection",
        "vormerkung": (c(16) or "") == "Ja",
        "vormerkung_notiz": c(17) or "",
        "prev_service": c(18) or "",
        "prev_grade": c(19) if c(19) is not None else "",
        "prev_cert": str(c(20) or ""),
    }


def get_next_id():
    wb = load_workbook(EXCEL_PATH)
    ws = wb["Inventar"]
    max_id = 0
    for row in _data_rows(ws):
        v = _to_int(ws.cell(row=row, column=1).value)
        if v is not None:
            max_id = max(max_id, v)
    wb.close()
    return max_id + 1


def save_card(name, set_name, card_number, language, condition, price, source,
              quantity=1, grading_status="Not Graded", grading_service="",
              grade=None, cert_number="", url=""):
    card_id = get_next_id()

    safe_set = set_name.replace(" ", "-").replace("&", "and").replace("'", "")
    safe_name = re.sub(r'[\\/:*?"<>|]', "", name.replace(" ", "-"))
    folder_name = f"{card_id:03d}-{safe_name}"
    folder_path = os.path.join(KARTEN_DIR, safe_set, folder_name)
    os.makedirs(folder_path, exist_ok=True)

    wb = load_workbook(EXCEL_PATH)
    ws = wb["Inventar"]
    new_row = _first_empty_row(ws)

    ws.cell(row=new_row, column=1, value=card_id)
    ws.cell(row=new_row, column=2, value=name)
    ws.cell(row=new_row, column=3, value=set_name)
    ws.cell(row=new_row, column=4, value=card_number)
    ws.cell(row=new_row, column=5, value=language)
    ws.cell(row=new_row, column=6, value=condition)
    ws.cell(row=new_row, column=7, value=grading_status)
    if grading_service:
        ws.cell(row=new_row, column=8, value=grading_service)
    if grade is not None:
        ws.cell(row=new_row, column=9, value=grade)
    if cert_number:
        ws.cell(row=new_row, column=10, value=cert_number)
    ws.cell(row=new_row, column=11, value=price)
    ws.cell(row=new_row, column=11).number_format = '#,##0.00 "EUR"'
    ws.cell(row=new_row, column=12, value=date.today())
    ws.cell(row=new_row, column=12).number_format = "YYYY-MM-DD"
    ws.cell(row=new_row, column=13, value=source)
    ws.cell(row=new_row, column=14, value=quantity)
    ws.cell(row=new_row, column=15, value="Collection")

    wb.save(EXCEL_PATH)
    wb.close()
    if url:
        _save_url(card_id, url)
    return card_id


def load_all_cards():
    wb = load_workbook(EXCEL_PATH)
    ws = wb["Inventar"]
    cards = []
    for row in _data_rows(ws):
        if _to_int(ws.cell(row=row, column=1).value) is None:
            continue
        cards.append(_row_to_card(ws, row))
    wb.close()
    return cards


def load_card(card_id):
    wb = load_workbook(EXCEL_PATH)
    ws = wb["Inventar"]
    r = _find_row(ws, card_id)
    card = _row_to_card(ws, r) if r else None
    wb.close()
    return card


def load_sales():
    wb = load_workbook(EXCEL_PATH)
    ws = wb["Sales"]
    sales = []
    for r in _data_rows(ws):
        c = lambda col: ws.cell(row=r, column=col).value
        card_id = _to_int(c(1))
        if card_id is None:
            continue
        price, fees, ship, cost = _num(c(5)), _num(c(6)), _num(c(7)), _num(c(9))
        qty = _to_int(c(13))
        sales.append({
            "row": r, "id": card_id, "name": str(c(2) or ""), "date": c(3),
            "platform": c(4) or "", "price": price, "fees": fees, "shipping": ship,
            "cost": cost, "net": price - fees - ship, "profit": price - fees - ship - cost,
            "quantity": qty, "buyer": c(11) or "", "tracking": c(12) or "",
        })
    wb.close()
    return sales


def _update_card_cells(card_id, values):
    wb = load_workbook(EXCEL_PATH)
    ws = wb["Inventar"]
    r = _find_row(ws, card_id)
    if r:
        for col, val in values.items():
            ws.cell(row=r, column=col).value = val
    wb.save(EXCEL_PATH)
    wb.close()


def update_card_status(card_id, status):
    _update_card_cells(card_id, {15: status})


def mark_card_sold(card_id, sale_price, platform, fees, sell_qty, shipping=0.0):
    wb = load_workbook(EXCEL_PATH)
    ws_inv = wb["Inventar"]
    ws_sales = wb["Sales"]

    inv_row = _find_row(ws_inv, card_id)
    if inv_row is None:
        wb.close()
        return
    card_name = ws_inv.cell(row=inv_row, column=2).value or ""
    purchase_price = _num(ws_inv.cell(row=inv_row, column=11).value)
    total_qty = int(_num(ws_inv.cell(row=inv_row, column=14).value, 1))

    remaining = total_qty - sell_qty
    if remaining <= 0:
        ws_inv.cell(row=inv_row, column=14, value=0)
        ws_inv.cell(row=inv_row, column=15, value="Sold")
    else:
        ws_inv.cell(row=inv_row, column=14, value=remaining)

    r = _first_empty_row(ws_sales)
    if not ws_sales.cell(row=1, column=13).value:
        ws_sales.cell(row=1, column=13, value="Quantity")
    eur = '#,##0.00 "EUR"'
    ws_sales.cell(row=r, column=1, value=card_id)
    ws_sales.cell(row=r, column=2, value=card_name)
    ws_sales.cell(row=r, column=3, value=date.today()).number_format = "YYYY-MM-DD"
    ws_sales.cell(row=r, column=4, value=platform)
    ws_sales.cell(row=r, column=5, value=sale_price).number_format = eur
    ws_sales.cell(row=r, column=6, value=fees).number_format = eur
    ws_sales.cell(row=r, column=7, value=shipping or None).number_format = eur
    ws_sales.cell(row=r, column=8).value = f'=IF(E{r}<>"",E{r}-F{r}-G{r},"")'
    ws_sales.cell(row=r, column=8).number_format = eur
    ws_sales.cell(row=r, column=9, value=purchase_price * sell_qty).number_format = eur
    ws_sales.cell(row=r, column=10).value = f'=IF(AND(H{r}<>"",I{r}<>""),H{r}-I{r},"")'
    ws_sales.cell(row=r, column=10).number_format = eur
    ws_sales.cell(row=r, column=13, value=sell_qty)

    wb.save(EXCEL_PATH)
    wb.close()


# Gleiche Produkte aus mehreren Käufen (z. B. zwei ETBs zu verschiedenen Preisen) bleiben in
# der Excel je eine Zeile, damit jeder Einkaufspreis exakt erhalten bleibt. In der App werden
# sie zu einem Eintrag gruppiert. Anführer (lead) ist der älteste Kauf, verkauft wird FIFO.
_lot_ids = {}   # lead-ID -> IDs aller Käufe der Gruppe (für Fotos/Vorschaubilder)


def group_key(c):
    return (c["name"].strip().lower(), c["set"].strip().lower(), str(c["number"]).strip().lower(),
            c["language"], c["condition"], c["grading_status"], c["grading_service"], str(c["grade"]),
            c["cert"], c["status"], bool(c["vormerkung"]))


def group_cards(cards):
    """Einträge für die Anzeige: Einzelkarten unverändert, Mehrfachkäufe als ein Eintrag mit "lots"."""
    groups = {}
    for c in cards:
        key = ("sold", c["id"]) if c["status"] == "Sold" else group_key(c)
        groups.setdefault(key, []).append(c)
    entries = []
    for lots in groups.values():
        if len(lots) == 1:
            entries.append(lots[0])
            continue
        lots = sorted(lots, key=lambda c: c["id"])
        qty = sum(c["quantity"] for c in lots)
        e = dict(lots[0])
        e["lots"] = lots
        e["quantity"] = qty
        e["price"] = sum(c["price"] * c["quantity"] for c in lots) / qty if qty else lots[0]["price"]
        e["sort_id"] = lots[-1]["id"]
        entries.append(e)
    global _lot_ids
    _lot_ids = {e["id"]: [c["id"] for c in e["lots"]] for e in entries if "lots" in e}  # in einem Schritt, Server-Thread liest mit
    entries.sort(key=lambda e: e["id"])
    return entries


def entry_photo_paths(card_id):
    return [p for i in _lot_ids.get(card_id, [card_id]) for p in get_photo_paths(i)]


def fifo_plan(entry, qty):
    """[(Kauf, Stückzahl)] für einen Verkauf von qty Stück, älteste Käufe zuerst."""
    plan = []
    for lot in entry.get("lots", [entry]):
        if qty <= 0:
            break
        n = min(qty, max(1, lot["quantity"]))
        plan.append((lot, n))
        qty -= n
    return plan


def sell_lots(plan, sale_price, platform, fees, shipping=0.0):
    """Verkauf über mehrere Käufe: Erlös, Gebühren und Versand anteilig nach Stückzahl."""
    total = sum(n for _lot, n in plan)
    left = [sale_price, fees, shipping]
    for i, (lot, n) in enumerate(plan):
        if i == len(plan) - 1:
            part = left
        else:
            part = [round(v * n / total, 2) for v in (sale_price, fees, shipping)]
            left = [a - b for a, b in zip(left, part)]
        mark_card_sold(lot["id"], round(part[0], 2), platform, round(part[1], 2), n, round(part[2], 2))


def undo_sale(sale_row):
    """Verkauf zuruecknehmen: Sales-Zeile leeren, Menge zurueck ins Inventar."""
    wb = load_workbook(EXCEL_PATH)
    ws_inv = wb["Inventar"]
    ws_sales = wb["Sales"]
    card_id = _to_int(ws_sales.cell(row=sale_row, column=1).value)
    qty = _to_int(ws_sales.cell(row=sale_row, column=13).value)
    inv_row = _find_row(ws_inv, card_id) if card_id is not None else None
    if inv_row:
        if qty is None:
            # Alte Verkaeufe ohne Mengen-Spalte: aus Einkaufswert zurueckrechnen.
            unit = _num(ws_inv.cell(row=inv_row, column=11).value)
            cost = _num(ws_sales.cell(row=sale_row, column=9).value)
            qty = max(1, round(cost / unit)) if unit > 0 else 1
        cur = int(_num(ws_inv.cell(row=inv_row, column=14).value, 0))
        ws_inv.cell(row=inv_row, column=14, value=cur + qty)
        if ws_inv.cell(row=inv_row, column=15).value == "Sold":
            ws_inv.cell(row=inv_row, column=15, value="Collection")
    for col in range(1, 14):
        ws_sales.cell(row=sale_row, column=col).value = None
    wb.save(EXCEL_PATH)
    wb.close()


def submit_for_grading(card_ids, service):
    if isinstance(card_ids, int):
        card_ids = [card_ids]
    wb = load_workbook(EXCEL_PATH)
    ws = wb["Inventar"]
    for card_id in card_ids:
        row = _find_row(ws, card_id)
        if not row:
            continue
        if ws.cell(row=row, column=7).value == "Graded":
            ws.cell(row=row, column=18, value=ws.cell(row=row, column=8).value)
            ws.cell(row=row, column=19, value=ws.cell(row=row, column=9).value)
            ws.cell(row=row, column=20, value=ws.cell(row=row, column=10).value)
        ws.cell(row=row, column=7, value="Submitted")
        ws.cell(row=row, column=8, value=service)
        ws.cell(row=row, column=16).value = None
        ws.cell(row=row, column=17).value = None
    wb.save(EXCEL_PATH)
    wb.close()


def toggle_grading_vormerkung(card_id, notiz=""):
    wb = load_workbook(EXCEL_PATH)
    ws = wb["Inventar"]
    row = _find_row(ws, card_id)
    if row:
        if ws.cell(row=row, column=16).value == "Ja":
            ws.cell(row=row, column=16).value = None
            ws.cell(row=row, column=17).value = None
        else:
            ws.cell(row=row, column=16, value="Ja")
            ws.cell(row=row, column=17, value=notiz)
    wb.save(EXCEL_PATH)
    wb.close()


def grading_returned(card_id, grade, cert_number):
    _update_card_cells(card_id, {7: "Graded", 9: grade, 10: cert_number})


def cancel_grading(card_id):
    """Einreichung zuruecknehmen (z.B. versehentlich eingereicht)."""
    wb = load_workbook(EXCEL_PATH)
    ws = wb["Inventar"]
    row = _find_row(ws, card_id)
    if row:
        prev = ws.cell(row=row, column=18).value
        if prev:
            ws.cell(row=row, column=7, value="Graded")
            ws.cell(row=row, column=8, value=prev)
            for col in (18, 19, 20):
                ws.cell(row=row, column=col).value = None
        else:
            ws.cell(row=row, column=7, value="Not Graded")
            ws.cell(row=row, column=8).value = None
    wb.save(EXCEL_PATH)
    wb.close()


def compute_statistics(cards, sales, meta):
    """Kennzahlen aus bereits geladenen Daten.

    Investiert = Einkaufswert des Bestands + Einkaufswert aller verkauften
    Stuecke (frueher wurde bei ausverkauften Karten nur 1 Stueck gezaehlt).
    Tausch-Zugaenge (Quelle Trade) und Tausch-Abgaenge zaehlen nicht als Geld.
    """
    by_id = {c["id"]: c for c in cards}
    s = {
        "entries": 0, "units": 0, "sealed_units": 0, "at_grading": 0, "graded": 0,
        "queued": 0, "for_sale": 0, "stock_value": 0.0, "total_spent": 0.0,
        "market_value": 0.0, "market_cost": 0.0, "market_units": 0,
        "sales_count": 0, "units_sold": 0, "revenue": 0.0, "fees": 0.0, "profit": 0.0,
    }
    for c in cards:
        if c["status"] == "Sold":
            continue
        q = max(0, int(c["quantity"]))
        s["entries"] += 1
        s["units"] += q
        if is_sealed_name(c["name"], c["number"]):
            s["sealed_units"] += q
        if c["grading_status"] == "Submitted":
            s["at_grading"] += q
        if c["grading_status"] == "Graded":
            s["graded"] += q
        if c["vormerkung"]:
            s["queued"] += 1
        if c["status"] == "For Sale":
            s["for_sale"] += q
        if c["source"] != "Trade":
            s["stock_value"] += c["price"] * q
        trend = market_price(meta, c["id"]) if c["grading_status"] != "Graded" else None
        if trend is not None:
            s["market_value"] += trend * q
            s["market_cost"] += c["price"] * q
            s["market_units"] += q
    s["total_spent"] = s["stock_value"]
    for sale in sales:
        card = by_id.get(sale["id"])
        if not card or card["source"] != "Trade":
            s["total_spent"] += sale["cost"]
        if sale["platform"] == "Trade":
            continue
        s["sales_count"] += 1
        s["units_sold"] += sale["quantity"] or 1
        s["revenue"] += sale["price"]
        s["fees"] += sale["fees"] + sale["shipping"]
        s["profit"] += sale["profit"]
    return s


def edit_card(card_id, name, set_name, card_number, language, condition, price, source,
              quantity, grading_service=None, grade=None, cert=None):
    values = {2: name, 3: set_name, 4: card_number, 5: language, 6: condition,
              11: price, 13: source, 14: quantity}
    if grading_service is not None:
        values[8] = grading_service or None
        values[9] = grade if grade not in ("", None) else None
        values[10] = cert or None
    wb = load_workbook(EXCEL_PATH)
    ws = wb["Inventar"]
    r = _find_row(ws, card_id)
    if r:
        for col, val in values.items():
            ws.cell(row=r, column=col).value = val
        ws.cell(row=r, column=11).number_format = '#,##0.00 "EUR"'
        if ws.cell(row=r, column=15).value == "Sold" and quantity > 0:
            ws.cell(row=r, column=15, value="Collection")
    wb.save(EXCEL_PATH)
    wb.close()


def delete_card(card_id):
    wb = load_workbook(EXCEL_PATH)
    ws = wb["Inventar"]
    r = _find_row(ws, card_id)
    if r:
        for col in range(1, 21):
            ws.cell(row=r, column=col).value = None
    ws_sales = wb["Sales"]
    for row in _data_rows(ws_sales):
        if _to_int(ws_sales.cell(row=row, column=1).value) == card_id:
            for col in range(1, 14):
                ws_sales.cell(row=row, column=col).value = None
    wb.save(EXCEL_PATH)
    wb.close()
    photo_dir = os.path.join(PHOTO_DIR, str(card_id))
    if os.path.isdir(photo_dir):
        shutil.rmtree(photo_dir)
    for ext in (".webp", ".png", ".jpg"):
        p = os.path.join(IMG_DIR, f"{card_id}{ext}")
        if os.path.exists(p):
            os.remove(p)


def _load_json(path):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def _save_json(path, data):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f)
    os.replace(tmp, path)


def _load_config():
    return _load_json(CONFIG_PATH)


def _save_config(config):
    _save_json(CONFIG_PATH, config)


# ============================================================
# PHOTO OPERATIONS
# ============================================================

PHOTO_DIR = os.path.join(BASE_DIR, "Fotos")
for _d in (PHOTO_DIR, CACHE_DIR):
    os.makedirs(_d, exist_ok=True)

URLS_PATH = os.path.join(BASE_DIR, ".card_urls.json")


def _load_urls():
    return _load_json(URLS_PATH)


def _save_url(card_id, url):
    urls = _load_urls()
    if url:
        urls[str(card_id)] = url
    else:
        urls.pop(str(card_id), None)
    _save_json(URLS_PATH, urls)


def get_card_url(card_id):
    return _load_urls().get(str(card_id))


# ============================================================
# EBAY API
# ============================================================

def _load_ebay_config():
    try:
        with open(EBAY_CONFIG_PATH) as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def _save_ebay_config(cfg):
    with open(EBAY_CONFIG_PATH, "w") as f:
        json.dump(cfg, f)


def _ebay_base(cfg):
    return "https://api.sandbox.ebay.com" if cfg.get("sandbox") else "https://api.ebay.com"


def _ebay_auth_base(cfg):
    return "https://auth.sandbox.ebay.com" if cfg.get("sandbox") else "https://auth.ebay.com"


def _ebay_exchange_code(auth_code):
    cfg = _load_ebay_config()
    creds = base64.b64encode(f"{cfg['client_id']}:{cfg['client_secret']}".encode()).decode()
    data = urllib.parse.urlencode({
        "grant_type": "authorization_code",
        "code": auth_code,
        "redirect_uri": cfg["ru_name"],
    }).encode()
    req = urllib.request.Request(
        f"{_ebay_base(cfg)}/identity/v1/oauth2/token",
        data=data,
        headers={"Authorization": f"Basic {creds}",
                 "Content-Type": "application/x-www-form-urlencoded"},
    )
    with urllib.request.urlopen(req) as resp:
        result = json.loads(resp.read())
    cfg["access_token"] = result["access_token"]
    cfg["refresh_token"] = result.get("refresh_token", "")
    cfg["token_expiry"] = _time.time() + result.get("expires_in", 7200)
    _save_ebay_config(cfg)


def _ebay_refresh():
    cfg = _load_ebay_config()
    creds = base64.b64encode(f"{cfg['client_id']}:{cfg['client_secret']}".encode()).decode()
    data = urllib.parse.urlencode({
        "grant_type": "refresh_token",
        "refresh_token": cfg["refresh_token"],
        "scope": EBAY_SCOPES,
    }).encode()
    req = urllib.request.Request(
        f"{_ebay_base(cfg)}/identity/v1/oauth2/token",
        data=data,
        headers={"Authorization": f"Basic {creds}",
                 "Content-Type": "application/x-www-form-urlencoded"},
    )
    with urllib.request.urlopen(req) as resp:
        result = json.loads(resp.read())
    cfg["access_token"] = result["access_token"]
    cfg["token_expiry"] = _time.time() + result.get("expires_in", 7200)
    _save_ebay_config(cfg)


def _ebay_token():
    cfg = _load_ebay_config()
    if not cfg.get("access_token"):
        return None
    if _time.time() > cfg.get("token_expiry", 0) - 60:
        try:
            _ebay_refresh()
            cfg = _load_ebay_config()
        except Exception:
            return None
    return cfg["access_token"]


def _ebay_api(method, path, body=None):
    cfg = _load_ebay_config()
    token = _ebay_token()
    if not token:
        raise Exception("Not connected to eBay. Go to the Settings tab first.")
    url = f"{_ebay_base(cfg)}{path}"
    payload = json.dumps(body).encode() if body else None
    req = urllib.request.Request(url, data=payload, method=method, headers={
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
        "Accept": "application/json",
        "Content-Language": "de-DE",
    })
    try:
        with urllib.request.urlopen(req) as resp:
            raw = resp.read()
            return json.loads(raw) if raw else {}
    except urllib.error.HTTPError as e:
        err = e.read().decode()
        raise Exception(f"eBay API error ({e.code}): {err}")


def _ebay_upload_image(image_path):
    token = _ebay_token()
    if not token:
        raise Exception("Not connected to eBay.")
    cfg = _load_ebay_config()
    url = "https://api.ebay.com/ws/api.dll"
    if cfg.get("sandbox"):
        url = "https://api.sandbox.ebay.com/ws/api.dll"

    boundary = "----FormBoundary7MA4YWxkTrZu0gW"
    xml_part = (
        '<?xml version="1.0" encoding="utf-8"?>'
        '<UploadSiteHostedPicturesRequest xmlns="urn:ebay:apis:eBLBaseComponents">'
        f'<RequesterCredentials><eBayAuthToken>{token}</eBayAuthToken></RequesterCredentials>'
        '<PictureName>card</PictureName>'
        '</UploadSiteHostedPicturesRequest>'
    )

    with open(image_path, "rb") as f:
        image_data = f.read()

    ext = os.path.splitext(image_path)[1].lower()
    mime = "image/jpeg" if ext in (".jpg", ".jpeg") else "image/png"

    body = (
        f"--{boundary}\r\n"
        f'Content-Disposition: form-data; name="XML Payload"\r\n'
        f"Content-Type: text/xml\r\n\r\n"
        f"{xml_part}\r\n"
        f"--{boundary}\r\n"
        f'Content-Disposition: form-data; name="file"; filename="card{ext}"\r\n'
        f"Content-Type: {mime}\r\n\r\n"
    ).encode() + image_data + f"\r\n--{boundary}--\r\n".encode()

    req = urllib.request.Request(url, data=body, method="POST", headers={
        "X-EBAY-API-SITEID": "77",
        "X-EBAY-API-CALL-NAME": "UploadSiteHostedPictures",
        "X-EBAY-API-COMPATIBILITY-LEVEL": "967",
        "Content-Type": f"multipart/form-data; boundary={boundary}",
    })
    with urllib.request.urlopen(req) as resp:
        raw = resp.read().decode()
    match = re.search(r"<FullURL>(.+?)</FullURL>", raw)
    if match:
        return match.group(1)
    raise Exception(f"Image upload failed: {raw[:300]}")


def _ebay_fetch_policies():
    cfg = _load_ebay_config()
    mp = cfg.get("marketplace", "EBAY_DE")
    for kind in ("fulfillment", "payment", "return"):
        try:
            result = _ebay_api("GET", f"/sell/account/v1/{kind}_policy?marketplace_id={mp}")
            policies = result.get(f"{kind}Policies", [])
            if policies:
                cfg[f"{kind}_policy_id"] = policies[0][f"{kind}PolicyId"]
        except Exception:
            pass
    _save_ebay_config(cfg)
    try:
        _ebay_no_return_policy()
    except Exception:
        pass


def _ebay_no_return_policy():
    """ID einer Rücknahme-Richtlinie ohne Rücknahme; Marlon verkauft immer ohne Rücknahme.
    Nimmt eine vorhandene aus dem eBay-Konto, legt sonst „Keine Rücknahme“ an."""
    cfg = _load_ebay_config()
    if cfg.get("return_policy_no_returns") == _ebay_base(cfg) and cfg.get("return_policy_id"):
        return cfg["return_policy_id"]
    mp = cfg.get("marketplace", "EBAY_DE")
    result = _ebay_api("GET", f"/sell/account/v1/return_policy?marketplace_id={mp}")
    pid = next((p["returnPolicyId"] for p in result.get("returnPolicies", [])
                if not p.get("returnsAccepted")), None)
    if not pid:
        created = _ebay_api("POST", "/sell/account/v1/return_policy", {
            "name": "Keine Rücknahme", "marketplaceId": mp,
            "categoryTypes": [{"name": "ALL_EXCLUDING_MOTORS_VEHICLES"}],
            "returnsAccepted": False,
        })
        pid = created["returnPolicyId"]
    cfg = _load_ebay_config()
    cfg["return_policy_id"] = pid
    cfg["return_policy_no_returns"] = _ebay_base(cfg)  # gilt nur für diese Umgebung (Sandbox/Prod)
    _save_ebay_config(cfg)
    return pid


def _ebay_update_shipping(de_cost, eu_cost, intl_cost, handling_days=2):
    cfg = _load_ebay_config()
    policy_id = cfg.get("fulfillment_policy_id")
    if not policy_id:
        raise Exception("No fulfillment policy found. Connect to eBay first.")
    _ebay_api("PUT", f"/sell/account/v1/fulfillment_policy/{policy_id}", {
        "name": "Standard Versand",
        "marketplaceId": cfg.get("marketplace", "EBAY_DE"),
        "handlingTime": {"value": handling_days, "unit": "DAY"},
        "globalShipping": False,
        "shippingOptions": [
            {
                "optionType": "DOMESTIC",
                "costType": "FLAT_RATE",
                "shippingServices": [{
                    "sortOrder": 1,
                    "shippingCarrierCode": "DHL",
                    "shippingServiceCode": "DE_DHLPaket",
                    "shippingCost": {"value": f"{de_cost:.2f}", "currency": "EUR"},
                    "freeShipping": False,
                }],
            },
            {
                "optionType": "INTERNATIONAL",
                "costType": "FLAT_RATE",
                "shipToLocations": {"regionIncluded": [{"regionName": "WORLDWIDE"}]},
                "shippingServices": [
                    {
                        "sortOrder": 1,
                        "shippingCarrierCode": "DHL",
                        "shippingServiceCode": "DE_DHLPaketInternational",
                        "shippingCost": {"value": f"{eu_cost:.2f}", "currency": "EUR"},
                        "freeShipping": False,
                        "shipToLocations": {"regionIncluded": [{"regionName": "EUROPE"}]},
                    },
                    {
                        "sortOrder": 2,
                        "shippingCarrierCode": "DHL",
                        "shippingServiceCode": "DE_DHLPaketInternational",
                        "shippingCost": {"value": f"{intl_cost:.2f}", "currency": "EUR"},
                        "freeShipping": False,
                        "shipToLocations": {"regionIncluded": [{"regionName": "WORLDWIDE"}]},
                    },
                ],
            },
        ],
    })
    cfg["shipping_de"] = de_cost
    cfg["shipping_eu"] = eu_cost
    cfg["shipping_intl"] = intl_cost
    cfg["handling_days"] = handling_days
    _save_ebay_config(cfg)


def ebay_create_draft(card_id, price, quantity, listing_format="FIXED_PRICE",
                      auction_start_price=None, auction_reserve_price=None,
                      auction_duration=None, buy_it_now_price=None):
    card = load_card(card_id)
    if not card:
        raise Exception(f"Card #{card_id} not found")

    condition_raw = card["condition"] or "Near Mint"
    graded = card["grading_status"] == "Graded"
    grading_service = card["grading_service"] if graded else ""
    # 10.0 aus Excel -> "10", sonst greift GRADE_TO_EBAY nicht.
    grade = fmt_grade(card["grade"]) if graded else ""

    title_parts = [card["name"]]
    if card["set"]:
        title_parts.append(card["set"])
    if card["number"]:
        title_parts.append(card["number"])
    if grading_service and grade:
        title_parts.append(f"{grading_service} {grade}")
    title_parts.append("Pokemon TCG")
    title = " ".join(title_parts)[:80]

    desc_lines = [f"<h2>{card['name']}</h2>",
                  f"<p><b>Set:</b> {card['set']}</p>"]
    if card["number"]:
        desc_lines.append(f"<p><b>Card Number:</b> {card['number']}</p>")
    desc_lines.append(f"<p><b>Language:</b> {card['language']}</p>")
    desc_lines.append(f"<p><b>Condition:</b> {condition_raw}</p>")
    if grading_service:
        desc_lines.append(f"<p><b>Graded:</b> {grading_service} {grade}</p>")
    description = "\n".join(desc_lines)

    sku = f"PKM-{card_id:04d}"

    if grading_service:
        svc_id = GRADING_SERVICE_TO_EBAY.get(grading_service.upper(), "2750123")
        grade_str = str(grade).strip()
        grade_id = GRADE_TO_EBAY.get(grade_str, "275020")
        cond_enum = "LIKE_NEW"
        cond_descriptors = [
            {"name": "27501", "values": [svc_id]},
            {"name": "27502", "values": [grade_id]},
        ]
    else:
        card_cond_id = CONDITION_TO_EBAY_CARD.get(condition_raw, "400010")
        cond_enum = "USED_VERY_GOOD"
        cond_descriptors = [
            {"name": "40001", "values": [card_cond_id]},
        ]

    image_urls = []
    for photo_path in get_photo_paths(card_id):
        img_url = _ebay_upload_image(photo_path)
        image_urls.append(img_url)

    inv_item = {
        "product": {
            "title": title,
            "description": description,
            "imageUrls": image_urls,
            "aspects": {
                "Kartenname": [card["name"]],
                "Set/Erweiterung": [card["set"]],
                "Spiel": ["Pokémon TCG"],
                "Sprache": [card["language"]],
            },
        },
        "condition": cond_enum,
        "conditionDescriptors": cond_descriptors,
        "conditionDescription": f"{grading_service} {grade}" if grading_service else condition_raw,
        "availability": {
            "shipToLocationAvailability": {"quantity": quantity}
        },
    }
    _ebay_api("PUT", f"/sell/inventory/v1/inventory_item/{sku}", inv_item)

    loc_key = _ebay_ensure_location()
    cfg = _load_ebay_config()
    mp = cfg.get("marketplace", "EBAY_DE")

    if listing_format == "AUCTION":
        pricing = {
            "auctionStartPrice": {"currency": "EUR", "value": f"{auction_start_price:.2f}"},
        }
        if auction_reserve_price:
            pricing["auctionReservePrice"] = {"currency": "EUR", "value": f"{auction_reserve_price:.2f}"}
        if buy_it_now_price:
            pricing["price"] = {"currency": "EUR", "value": f"{buy_it_now_price:.2f}"}
    else:
        pricing = {
            "price": {"currency": "EUR", "value": f"{price:.2f}"}
        }

    offer = {
        "sku": sku,
        "marketplaceId": mp,
        "format": listing_format,
        "listingDescription": description,
        **({"availableQuantity": quantity} if listing_format != "AUCTION" else {}),
        "categoryId": "183454",
        "merchantLocationKey": loc_key,
        "pricingSummary": pricing,
        "listingPolicies": {},
    }
    if listing_format == "AUCTION" and auction_duration:
        offer["listingDuration"] = auction_duration
    if listing_format != "AUCTION":
        offer["listingPolicies"]["bestOfferTerms"] = {"bestOfferEnabled": True}
    for kind in ("fulfillment", "payment"):
        pid = cfg.get(f"{kind}_policy_id")
        if pid:
            offer["listingPolicies"][f"{kind}PolicyId"] = pid
    offer["listingPolicies"]["returnPolicyId"] = _ebay_no_return_policy()

    existing_offer_id = None
    try:
        offers_result = _ebay_api("GET", f"/sell/inventory/v1/offer?sku={urllib.parse.quote(sku)}")
        for o in offers_result.get("offers", []):
            if o.get("sku") == sku:
                if o.get("format") == listing_format:
                    existing_offer_id = o.get("offerId")
                else:
                    _ebay_api("DELETE", f"/sell/inventory/v1/offer/{o.get('offerId')}")
                break
    except Exception:
        pass

    if existing_offer_id:
        _ebay_api("PUT", f"/sell/inventory/v1/offer/{existing_offer_id}", offer)
        return existing_offer_id, None
    else:
        result = _ebay_api("POST", "/sell/inventory/v1/offer", offer)
        return result.get("offerId", "created"), None


def _ebay_ensure_location():
    cfg = _load_ebay_config()
    loc_key = "default_location"
    mp = cfg.get("marketplace", "EBAY_DE")
    mp_to_country = {
        "EBAY_DE": "DE", "EBAY_US": "US", "EBAY_GB": "GB",
        "EBAY_FR": "FR", "EBAY_IT": "IT", "EBAY_ES": "ES",
    }
    country = mp_to_country.get(mp, "DE")
    try:
        _ebay_api("GET", f"/sell/inventory/v1/location/{loc_key}")
        return loc_key
    except Exception:
        pass
    body = {
        "location": {
            "address": {
                "postalCode": "46487",
                "country": country,
            }
        },
        "name": "Default Location",
        "merchantLocationStatus": "ENABLED",
        "locationTypes": ["WAREHOUSE"],
    }
    _ebay_api("POST", f"/sell/inventory/v1/location/{loc_key}", body)
    cfg["location_key"] = loc_key
    _save_ebay_config(cfg)
    return loc_key


def ebay_publish_offer(offer_id):
    result = _ebay_api("POST", f"/sell/inventory/v1/offer/{offer_id}/publish")
    return result.get("listingId", "")


def _extract_auth_code(text):
    text = text.strip()
    if "code=" in text:
        qs = text.split("?", 1)[1] if "?" in text else text
        params = parse_qs(qs)
        code = params.get("code", [None])[0]
        if code:
            return code
    return text



def _migrate_photos():
    """Move legacy Fotos/{id}.ext files into Fotos/{id}/ folders."""
    for fname in os.listdir(PHOTO_DIR):
        fpath = os.path.join(PHOTO_DIR, fname)
        if not os.path.isfile(fpath):
            continue
        name, ext = os.path.splitext(fname)
        if not name.isdigit():
            continue
        card_dir = os.path.join(PHOTO_DIR, name)
        os.makedirs(card_dir, exist_ok=True)
        shutil.move(fpath, os.path.join(card_dir, f"1{ext}"))

_migrate_photos()


def _card_photo_dir(card_id):
    d = os.path.join(PHOTO_DIR, str(card_id))
    os.makedirs(d, exist_ok=True)
    return d


def get_photo_paths(card_id):
    d = os.path.join(PHOTO_DIR, str(card_id))
    if not os.path.isdir(d):
        return []
    files = []
    for f in sorted(os.listdir(d)):
        if f.lower().endswith((".jpg", ".jpeg", ".png")):
            files.append(os.path.join(d, f))
    return files


def get_photo_path(card_id):
    paths = get_photo_paths(card_id)
    return paths[0] if paths else None


def _next_photo_number(card_id):
    d = os.path.join(PHOTO_DIR, str(card_id))
    if not os.path.isdir(d):
        return 1
    nums = []
    for f in os.listdir(d):
        name = os.path.splitext(f)[0]
        if name.isdigit():
            nums.append(int(name))
    return max(nums, default=0) + 1


def save_photo(card_id, source_path):
    d = _card_photo_dir(card_id)
    n = _next_photo_number(card_id)
    ext = os.path.splitext(source_path)[1].lower()
    if ext in (".heic", ".heif"):
        img = Image.open(source_path)
        dest = os.path.join(d, f"{n}.jpg")
        img.save(dest, "JPEG", quality=90)
    else:
        dest = os.path.join(d, f"{n}{ext}")
        shutil.copy2(source_path, dest)
    return dest


def save_photo_bytes(card_id, data, content_type="image/jpeg"):
    d = _card_photo_dir(card_id)
    n = _next_photo_number(card_id)
    ext = ".jpg"
    if "png" in content_type:
        ext = ".png"
    dest = os.path.join(d, f"{n}{ext}")
    with open(dest, "wb") as f:
        f.write(data)
    return dest


def _get_local_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"



# ============================================================
# KARTENBILDER & MARKTPREISE (TCGdex)
# ============================================================
# TCGdex (api.tcgdex.net) ist frei, ohne API-Key, liefert Kartenbilder in
# mehreren Sprachen und Cardmarket-Preise (Trend/30-Tage-Schnitt).
# Zuordnung: Set-Name -> TCGdex-Set, dann Kartennummer (+ Name). Ergebnis wird
# in .card_meta.json gemerkt, Bilder liegen in .bilder/<id>.<ext>.

TCGDEX_API = "https://api.tcgdex.net/v2"
SUBSET_SUFFIXES = {"trainer gallery", "galarian gallery", "shiny vault", "radiant collection"}
TCG_LANG = {"EN": "en", "DE": "de", "FR": "fr", "IT": "it", "ES": "es", "PT": "pt"}
SET_CACHE_DAYS = 7
PRICE_MAX_AGE = 20 * 3600
_IMG_EXT = "webp"
try:
    from PIL import features as _pil_features
    if not _pil_features.check("webp"):
        _IMG_EXT = "png"
except Exception:
    _IMG_EXT = "png"

_tcg_lock = threading.Lock()
_tcg_sets = None
_tcg_set_cards = {}


def _http_get(url, timeout=20):
    req = urllib.request.Request(url, headers={"User-Agent": f"PokeInv/{VERSION}"})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.read()


def _http_json(url):
    return json.loads(_http_get(url))


def _api_url(path, **params):
    q = urllib.parse.urlencode(params) if params else ""
    return f"{TCGDEX_API}{path}" + (f"?{q}" if q else "")


def _norm(text):
    text = unicodedata.normalize("NFKD", str(text or ""))
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.lower().replace("'", "").replace("’", "").replace("&", " ")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return " ".join(w for w in text.split() if w != "and")


def _cached_json(name, url, max_age_days):
    path = os.path.join(CACHE_DIR, name)
    try:
        if _time.time() - os.path.getmtime(path) < max_age_days * 86400:
            with open(path, encoding="utf-8") as f:
                return json.load(f)
    except OSError:
        pass
    try:
        data = _http_json(url)
        _save_json(path, data)
        return data
    except Exception:
        # Offline: lieber veraltete Daten als gar keine.
        if os.path.exists(path):
            with open(path, encoding="utf-8") as f:
                return json.load(f)
        raise


def tcg_sets():
    """Alle Sets ausser TCG Pocket (Handyspiel, keine echten Karten)."""
    global _tcg_sets
    with _tcg_lock:
        if _tcg_sets is None:
            sets = _cached_json("tcg_sets.json", _api_url("/en/sets"), SET_CACHE_DAYS)
            try:
                pocket = _cached_json("tcg_pocket.json", _api_url("/en/series/tcgp"), 30)
                pocket_ids = {s["id"] for s in pocket.get("sets", [])}
            except Exception:
                pocket_ids = set()
            _tcg_sets = [s for s in sets if s["id"] not in pocket_ids
                         and "/tcgp/" not in (s.get("logo") or "")]
        return _tcg_sets


def tcg_set_names():
    try:
        return {s["id"]: s["name"] for s in tcg_sets()}
    except Exception:
        return {}


def tcg_find_set(set_name):
    if not set_name or set_name in ("Promo", "Other"):
        return None
    target = _norm(set_name)
    variants = [target]
    for prefix in ("ex ", "hs "):
        if target.startswith(prefix):
            variants.append(target[len(prefix):])
    sets = tcg_sets()
    for v in variants:
        for s in sets:
            if _norm(s["name"]) == v:
                return s
    return None


def tcg_set_cards(set_id):
    with _tcg_lock:
        if set_id in _tcg_set_cards:
            return _tcg_set_cards[set_id]
    data = _cached_json(f"tcg_set_{set_id}.json", _api_url(f"/en/sets/{set_id}"), SET_CACHE_DAYS)
    cards = data.get("cards", [])
    for c in cards:
        c["set_id"] = set_id
        c["set_name"] = data.get("name", "")
    with _tcg_lock:
        _tcg_set_cards[set_id] = cards
    return cards


def _split_code(code):
    """'OBF125' -> ('OBF', 125); 'TG05' -> ('TG', 5); '125' -> ('', 125)."""
    m = re.search(r"([A-Za-z]*)[-_ ]?0*(\d+)[a-z]?$", str(code or "").strip())
    if not m:
        return None
    return m.group(1).upper(), int(m.group(2))


def _number_matches(local_id, code):
    a, b = _split_code(local_id), _split_code(code)
    if not a or not b or a[1] != b[1]:
        return -1
    # Kartenmarkt-Codes tragen das Set-Kuerzel vorne (OBF125); TCGdex nur die
    # Nummer (125) oder einen Unter-Praefix (TG05). Laengster Praefix gewinnt.
    return len(a[0]) if b[0].endswith(a[0]) else -1


def _with_set(cands):
    names = tcg_set_names()
    for c in cands:
        if "set_id" not in c:
            c["set_id"] = c["id"].rsplit("-", 1)[0]
        c.setdefault("set_name", names.get(c["set_id"], c["set_id"]))
    return cands


def tcg_search(name, limit=150):
    """Karten nach Name suchen (ohne TCG Pocket)."""
    name = (name or "").strip()
    if not name:
        return []
    results = []
    for params in ({"name": f"eq:{name}"}, {"name": name}):
        try:
            results = _http_json(_api_url("/en/cards", **params))
        except Exception:
            results = []
        results = [r for r in results if r.get("image") and "/tcgp/" not in r["image"]]
        if results:
            break
    if not results and " " in name:
        # "Mr Mime" findet "Mr. Mime" nicht: erstes Wort suchen, dann filtern.
        try:
            first = _http_json(_api_url("/en/cards", name=name.split()[0]))
        except Exception:
            first = []
        want = _norm(name)
        results = [r for r in first if r.get("image") and "/tcgp/" not in r["image"]
                   and _norm(r.get("name")) == want]
    return _with_set(results[:limit])


def resolve_card_image(name, set_name, number):
    """Sucht das passende TCGdex-Bild.

    Rueckgabe: {"status": "ok", "match": {...}} oder {"status": "multi"|"none",
    "candidates": [...]}. match["kind"] ist "card" oder "logo" (Sealed).
    """
    want = _norm(name)
    tset = None
    try:
        tset = tcg_find_set(set_name)
    except Exception:
        pass

    if is_sealed_name(name, number):
        if tset and tset.get("logo"):
            return {"status": "ok", "match": {
                "kind": "logo", "id": tset["id"], "image": tset["logo"],
                "name": tset["name"], "set_name": tset["name"], "localId": ""}}
        return {"status": "none", "candidates": []}

    want_words = set(want.split())

    def pick(cands):
        """(sicherer Treffer, Kandidaten). Nummer zuerst, Name als Bestaetigung."""
        if number:
            scored = [(_number_matches(c.get("localId"), number), c) for c in cands]
            scored = [x for x in scored if x[0] >= 0]
            if scored:
                best = max(s for s, _ in scored)
                top = [c for s, c in scored if s == best]
                named = [c for c in top if _norm(c.get("name")) == want]
                if len(named) == 1:
                    return named[0], []
                # Nummer passt, Name nur ungefaehr ("Professors Research" vs.
                # "Professor's Research"): reicht, wenn ein Wort uebereinstimmt.
                loose = [c for c in top if want_words & set(_norm(c.get("name")).split())]
                if not named and len(loose) == 1:
                    return loose[0], []
                if named or loose:
                    return None, named or loose
        named = [c for c in cands if _norm(c.get("name")) == want]
        if len(named) == 1:
            return named[0], []
        return None, named

    rest_set = []
    if tset:
        try:
            # Unter-Sets mitnehmen: "Crown Zenith Galarian Gallery" (GG30),
            # "Brilliant Stars Trainer Gallery" (TG05), "... Shiny Vault".
            base = _norm(tset["name"])
            cands = list(tcg_set_cards(tset["id"]))
            for s in tcg_sets():
                n = _norm(s["name"])
                if s["id"] != tset["id"] and n.startswith(base + " ") and                         n[len(base) + 1:] in SUBSET_SUFFIXES:
                    cands += tcg_set_cards(s["id"])
            match, rest_set = pick(cands)
            if match:
                return _result(match)
        except Exception:
            pass

    cands = tcg_search(name)
    # Set bekannt, aber dort nicht gefunden: Treffer aus fremden Sets nur
    # vorschlagen, nie automatisch nehmen (sonst landet "Pikachu 5" aus
    # Celebrations bei einer Brilliant-Stars-Karte).
    match, rest = pick(cands) if not tset else (None, None)
    if match:
        return _result(match)
    if len(cands) == 1 and not number:
        return _result(cands[0])
    seen, out = set(), []
    for c in _with_set(rest_set) + (rest or cands):
        if c["id"] not in seen and c.get("image"):
            seen.add(c["id"])
            out.append(c)
    return {"status": "multi" if out else "none", "candidates": out}


def _result(match):
    match = dict(match, kind="card")
    # Manche Karten (z. B. Crown Zenith Galarian Gallery) hat TCGdex ohne Bild:
    # Zuordnung trotzdem merken, damit der Marktpreis funktioniert.
    return {"status": "ok" if match.get("image") else "noimage", "match": match}


def download_preview(match, size="low"):
    """Kleines Vorschaubild fuer Auswahl-Dialoge (gecacht)."""
    key = re.sub(r"[^A-Za-z0-9.\-]", "_", match["id"])
    d = os.path.join(CACHE_DIR, "preview")
    os.makedirs(d, exist_ok=True)
    url = (match["image"] + ".png") if match.get("kind") == "logo" else \
        f"{match['image']}/{size}.{_IMG_EXT}"
    path = os.path.join(d, f"{key}_{size}.{url.rsplit('.', 1)[-1]}")
    if not os.path.exists(path):
        data = _http_get(url)
        with open(path, "wb") as f:
            f.write(data)
    return path


def fetch_card_price(tcg_id):
    data = _http_json(_api_url(f"/en/cards/{urllib.parse.quote(tcg_id)}"))
    cm = (data.get("pricing") or {}).get("cardmarket") or {}
    if not cm:
        return {"updated": _time.time(), "trend": None}
    return {
        "updated": _time.time(),
        "trend": cm.get("trend") or cm.get("avg"),
        "avg30": cm.get("avg30"),
        "low": cm.get("low"),
        "rarity": data.get("rarity"),
    }


def market_price(meta, card_id):
    m = meta.get(str(card_id)) or {}
    p = m.get("price") or {}
    t = p.get("trend")
    return float(t) if t else None


def card_image_path(meta, card_id):
    """(Pfad, Art) des Titelbilds: nur eigene Fotos (bei Mehrfachkäufen aus allen Käufen)."""
    photos = entry_photo_paths(card_id)
    if photos:
        return photos[0], "photo"
    return None, None


# ============================================================
# QR-CODE (eigener Encoder, Byte-Modus, Fehlerkorrektur M)
# ============================================================
# Eigene Umsetzung, weil das Auto-Update nur diese Datei lädt und
# keine neuen Pakete nachinstallieren kann.

_QR_ECC_PER_BLOCK = (None, 10, 16, 26, 18, 24, 16, 18, 22, 22)   # Version 1–9, Stufe M
_QR_NUM_BLOCKS = (None, 1, 1, 1, 2, 2, 4, 4, 4, 5)


def _gf_mul(x, y):
    z = 0
    for i in reversed(range(8)):
        z = (z << 1) ^ ((z >> 7) * 0x11D)
        z ^= ((y >> i) & 1) * x
    return z


def _rs_divisor(degree):
    result = [0] * (degree - 1) + [1]
    root = 1
    for _ in range(degree):
        for j in range(degree):
            result[j] = _gf_mul(result[j], root)
            if j + 1 < degree:
                result[j] ^= result[j + 1]
        root = _gf_mul(root, 0x02)
    return result


def _rs_remainder(data, divisor):
    result = [0] * len(divisor)
    for b in data:
        factor = b ^ result.pop(0)
        result.append(0)
        for i, coef in enumerate(divisor):
            result[i] ^= _gf_mul(coef, factor)
    return result


def _qr_raw_modules(ver):
    result = (16 * ver + 128) * ver + 64
    if ver >= 2:
        n = ver // 7 + 2
        result -= (25 * n - 10) * n - 55
        if ver >= 7:
            result -= 36
    return result


def _qr_data_codewords(ver):
    return _qr_raw_modules(ver) // 8 - _QR_ECC_PER_BLOCK[ver] * _QR_NUM_BLOCKS[ver]


def _qr_align_positions(ver, size):
    if ver == 1:
        return []
    n = ver // 7 + 2
    step = (ver * 8 + n * 3 + 5) // (n * 4 - 4) * 2
    return [6] + sorted(size - 7 - i * step for i in range(n - 1))


_QR_MASKS = (
    lambda x, y: (x + y) % 2 == 0,
    lambda x, y: y % 2 == 0,
    lambda x, y: x % 3 == 0,
    lambda x, y: (x + y) % 3 == 0,
    lambda x, y: (x // 3 + y // 2) % 2 == 0,
    lambda x, y: x * y % 2 + x * y % 3 == 0,
    lambda x, y: (x * y % 2 + x * y % 3) % 2 == 0,
    lambda x, y: ((x + y) % 2 + x * y % 3) % 2 == 0,
)


def _qr_penalty(m):
    size = len(m)
    score = 0
    lines = [row for row in m] + [[m[y][x] for y in range(size)] for x in range(size)]
    finder = ([1, 0, 1, 1, 1, 0, 1, 0, 0, 0, 0], [0, 0, 0, 0, 1, 0, 1, 1, 1, 0, 1])
    for line in lines:
        run, prev = 0, None
        for v in line:
            if v == prev:
                run += 1
            else:
                if run >= 5:
                    score += run - 2
                run, prev = 1, v
        if run >= 5:
            score += run - 2
        for i in range(size - 10):
            if line[i:i + 11] in finder:
                score += 40
    for y in range(size - 1):
        for x in range(size - 1):
            if m[y][x] == m[y][x + 1] == m[y + 1][x] == m[y + 1][x + 1]:
                score += 3
    dark = sum(sum(row) for row in m)
    total = size * size
    score += (abs(dark * 20 - total * 10) + total - 1) // total * 10 - 10
    return score


def qr_matrix(text):
    """QR-Code als Liste von Zeilen (1 = dunkel), ohne Ruhezone."""
    data = text.encode("utf-8")
    ver = next((v for v in range(1, 10) if 12 + 8 * len(data) <= _qr_data_codewords(v) * 8), None)
    if ver is None:
        raise ValueError("Text zu lang für QR-Code")
    cap = _qr_data_codewords(ver) * 8

    bits = [0, 1, 0, 0] + [(len(data) >> i) & 1 for i in range(7, -1, -1)]
    for b in data:
        bits += [(b >> i) & 1 for i in range(7, -1, -1)]
    bits += [0] * min(4, cap - len(bits))
    bits += [0] * (-len(bits) % 8)
    codewords = [int("".join(map(str, bits[i:i + 8])), 2) for i in range(0, len(bits), 8)]
    pad = 0xEC
    while len(codewords) < cap // 8:
        codewords.append(pad)
        pad ^= 0xEC ^ 0x11

    nblocks, ecclen = _QR_NUM_BLOCKS[ver], _QR_ECC_PER_BLOCK[ver]
    raw = _qr_raw_modules(ver) // 8
    nshort = nblocks - raw % nblocks
    shortlen = raw // nblocks
    divisor = _rs_divisor(ecclen)
    blocks, k = [], 0
    for i in range(nblocks):
        dat = codewords[k:k + shortlen - ecclen + (0 if i < nshort else 1)]
        k += len(dat)
        ecc = _rs_remainder(dat, divisor)
        if i < nshort:
            dat.append(0)
        blocks.append(dat + ecc)
    final = [blk[i] for i in range(len(blocks[0])) for j, blk in enumerate(blocks)
             if i != shortlen - ecclen or j >= nshort]

    size = ver * 4 + 17
    mods = [[0] * size for _ in range(size)]
    func = [[False] * size for _ in range(size)]

    def setf(x, y, dark):
        mods[y][x] = 1 if dark else 0
        func[y][x] = True

    for i in range(size):
        setf(6, i, i % 2 == 0)
        setf(i, 6, i % 2 == 0)
    for cx, cy in ((3, 3), (size - 4, 3), (3, size - 4)):
        for dy in range(-4, 5):
            for dx in range(-4, 5):
                if 0 <= cx + dx < size and 0 <= cy + dy < size:
                    setf(cx + dx, cy + dy, max(abs(dx), abs(dy)) not in (2, 4))
    pos = _qr_align_positions(ver, size)
    for i, ax in enumerate(pos):
        for j, ay in enumerate(pos):
            if (i, j) in ((0, 0), (0, len(pos) - 1), (len(pos) - 1, 0)):
                continue
            for dy in range(-2, 3):
                for dx in range(-2, 3):
                    setf(ax + dx, ay + dy, max(abs(dx), abs(dy)) != 1)

    def draw_format(mask):
        d = mask  # Stufe M hat die Formatbits 00
        rem = d
        for _ in range(10):
            rem = (rem << 1) ^ ((rem >> 9) * 0x537)
        fb = (d << 10 | rem) ^ 0x5412
        bit = lambda i: (fb >> i) & 1
        for i in range(6):
            setf(8, i, bit(i))
        setf(8, 7, bit(6))
        setf(8, 8, bit(7))
        setf(7, 8, bit(8))
        for i in range(9, 15):
            setf(14 - i, 8, bit(i))
        for i in range(8):
            setf(size - 1 - i, 8, bit(i))
        for i in range(8, 15):
            setf(8, size - 15 + i, bit(i))
        setf(8, size - 8, True)

    draw_format(0)
    if ver >= 7:
        rem = ver
        for _ in range(12):
            rem = (rem << 1) ^ ((rem >> 11) * 0x1F25)
        vb = ver << 12 | rem
        for i in range(18):
            a, b = size - 11 + i % 3, i // 3
            setf(a, b, (vb >> i) & 1)
            setf(b, a, (vb >> i) & 1)

    i = 0
    right = size - 1
    while right >= 1:
        if right == 6:
            right = 5
        for vert in range(size):
            for j in range(2):
                x = right - j
                y = size - 1 - vert if (right + 1) & 2 == 0 else vert
                if not func[y][x] and i < len(final) * 8:
                    mods[y][x] = (final[i >> 3] >> (7 - (i & 7))) & 1
                    i += 1
        right -= 2

    base = [row[:] for row in mods]
    best = None
    for mask, fn in enumerate(_QR_MASKS):
        mods = [row[:] for row in base]
        for y in range(size):
            for x in range(size):
                if not func[y][x] and fn(x, y):
                    mods[y][x] ^= 1
        draw_format(mask)
        score = _qr_penalty(mods)
        if best is None or score < best[0]:
            best = (score, mods)
    return best[1]


def qr_image(text, px=150, border=3, dark="#0f1115", light="#ffffff"):
    """QR-Code als PIL-Bild, ganzzahlig skaliert (scharfe Kanten)."""
    m = qr_matrix(text)
    n = len(m) + 2 * border
    scale = max(1, px // n)
    img = Image.new("RGB", (n * scale, n * scale), light)
    draw = ImageDraw.Draw(img)
    for y, row in enumerate(m):
        for x, v in enumerate(row):
            if v:
                x0, y0 = (x + border) * scale, (y + border) * scale
                draw.rectangle([x0, y0, x0 + scale - 1, y0 + scale - 1], fill=dark)
    return img


# ============================================================
# PHOTO UPLOAD WEB SERVER
# ============================================================

UPLOAD_PORT = 8080

# Sortierung der Sammlung in der App; die Handy-Seite übernimmt sie live.
upload_sort = "new"
# Karten-IDs frisch hochgeladener Handy-Fotos; die App holt sie im Tk-Thread ab.
upload_events = queue.Queue()


def sort_cards(cards, sort, meta):
    keys = {
        "new": (lambda c: -c.get("sort_id", c["id"])), "old": (lambda c: c["id"]),
        "name": (lambda c: (c["name"].lower(), c["id"])),
        "set": (lambda c: (c["set"].lower(), _split_code(c["number"]) or ("", 0), c["id"])),
        "price": (lambda c: -c["price"]),
        "market": (lambda c: -(market_price(meta, c["id"]) or -1)),
        "delta": (lambda c: -((market_price(meta, c["id"]) or c["price"]) - c["price"])),
    }
    return sorted(cards, key=keys.get(sort, keys["new"]))

_HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Poké Inventar · Foto-Upload</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: -apple-system, system-ui, sans-serif; background: #101217; color: #e9ebf1;
         padding: 16px 14px 40px; }
  h1 { font-size: 20px; margin-bottom: 4px; }
  .sub { color: #8b93a7; font-size: 13px; margin-bottom: 14px; }
  #q { width: 100%; padding: 12px 14px; border-radius: 12px; border: 1px solid #2c313c;
       background: #1a1d24; color: #e9ebf1; font-size: 16px; margin-bottom: 14px; }
  .card { display: flex; gap: 12px; align-items: center; background: #1a1d24;
          border: 1px solid #262a33; border-radius: 14px; padding: 10px; margin-bottom: 8px; }
  .thumb { width: 52px; height: 72px; border-radius: 6px; object-fit: cover; background: #262a33; flex: none; }
  .info { flex: 1; min-width: 0; }
  .name { font-weight: 600; font-size: 15px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
  .meta { color: #8b93a7; font-size: 13px; margin-top: 2px; }
  .count { font-size: 12px; margin-top: 4px; color: #f87171; }
  .count.ok { color: #34d399; }
  button { background: #ffcb05; color: #15171c; border: none; border-radius: 10px;
           padding: 12px 14px; font-weight: 700; font-size: 14px; flex: none; }
</style>
</head>
<body>
<h1>Foto-Upload</h1>
<p class="sub">Karte antippen, Foto aufnehmen. Sortiert wie in der App.</p>
<input id="q" type="search" placeholder="Suchen: Name, Set, Nummer" oninput="filter()">
<div id="cards">{{CARDS}}</div>
<script>
// Alle 3 s nachsehen, ob sich am Mac etwas geändert hat (neue Karte, Foto, Sortierung).
let state = '{{STATE}}';
setInterval(async () => {
  if (document.hidden) return;
  try {
    const s = await (await fetch('/state', { cache: 'no-store' })).text();
    if (s === state) return;
    const html = await (await fetch('/', { cache: 'no-store' })).text();
    const doc = new DOMParser().parseFromString(html, 'text/html');
    document.getElementById('cards').innerHTML = doc.getElementById('cards').innerHTML;
    state = s;
    filter();
  } catch (e) {}
}, 3000);
function filter() {
  const q = document.getElementById('q').value.toLowerCase();
  document.querySelectorAll('.card').forEach(c => {
    c.style.display = c.dataset.s.includes(q) ? '' : 'none';
  });
}
function upload(cardId, btn) {
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = 'image/*';
  input.capture = 'environment';
  input.onchange = async () => {
    const file = input.files[0];
    if (!file) return;
    const orig = btn.textContent;
    btn.textContent = '…';
    const form = new FormData();
    form.append('photo', file);
    try {
      const resp = await fetch('/upload/' + cardId, { method: 'POST', body: form });
      if (resp.ok) {
        const data = await resp.json();
        btn.textContent = orig;
        const cnt = btn.closest('.card').querySelector('.count');
        cnt.textContent = data.count + (data.count > 1 ? ' Fotos' : ' Foto');
        cnt.classList.add('ok');
        const img = btn.closest('.card').querySelector('.thumb');
        img.style.visibility = '';
        img.src = '/img/' + cardId + '?t=' + Date.now();
      } else {
        btn.textContent = 'Fehler';
      }
    } catch(e) {
      btn.textContent = 'Fehler';
    }
  };
  input.click();
}
</script>
</body>
</html>"""


def upload_state():
    """Kurzer Fingerabdruck von Inventar, Fotos und Sortierung für das Auto-Neuladen am Handy."""
    parts = [upload_sort]
    try:
        parts.append(str(os.path.getmtime(EXCEL_PATH)))
    except OSError:
        pass
    try:
        for d in sorted(os.listdir(PHOTO_DIR)):
            sub = os.path.join(PHOTO_DIR, d)
            if os.path.isdir(sub):
                parts.append(f"{d}:{len(os.listdir(sub))}")
    except OSError:
        pass
    return hashlib.md5("|".join(parts).encode()).hexdigest()[:16]


class PhotoUploadHandler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        pass

    def _send(self, code, body=b"", ctype="text/plain"):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        # Fotos ändern sich laufend, das Handy soll nie eine alte Fassung zeigen.
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self.path in ("/", ""):
            cards = group_cards([c for c in load_all_cards() if c["status"] != "Sold"])
            counts = {c["id"]: len(entry_photo_paths(c["id"])) for c in cards}
            cards = sort_cards(cards, upload_sort, _load_json(META_PATH))
            rows = []
            for c in cards:
                count = counts[c["id"]]
                e = html.escape
                nr = f" · {e(c['number'])}" if c["number"] else ""
                if "lots" in c:
                    nr += f" · {c['quantity']}×"
                search = e(f"{c['name']} {c['set']} {c['number']} {c['id']}".lower())
                count_txt = f"{count} Foto{'s' if count != 1 else ''}" if count else "noch kein Foto"
                rows.append(
                    f'<div class="card" data-s="{search}">'
                    f'<img class="thumb" loading="lazy" src="/img/{c["id"]}" alt="" '
                    f'onerror="this.style.visibility=\'hidden\'">'
                    f'<div class="info"><div class="name">{e(c["name"])}</div>'
                    f'<div class="meta">#{c["id"]:03d} · {e(c["set"])}{nr}</div>'
                    f'<div class="count{" ok" if count else ""}">{count_txt}</div></div>'
                    f'<button onclick="upload({c["id"]}, this)">Foto</button></div>'
                )
            page = _HTML_TEMPLATE.replace("{{CARDS}}", "".join(rows)).replace("{{STATE}}", upload_state())
            self._send(200, page.encode(), "text/html; charset=utf-8")
        elif self.path.split("?")[0] == "/state":
            self._send(200, upload_state().encode())
        elif self.path.startswith("/img/"):
            card_id = _to_int(self.path.split("?")[0].split("/")[-1])
            path, _kind = card_image_path(_load_json(META_PATH), card_id) if card_id else (None, None)
            if not path:
                self._send(404)
                return
            try:
                img = ImageOps.exif_transpose(Image.open(path)).convert("RGB")
                img.thumbnail((160, 224))
                buf = io.BytesIO()
                img.save(buf, "JPEG", quality=80)
                self._send(200, buf.getvalue(), "image/jpeg")
            except Exception:
                self._send(404)
        else:
            self._send(404)

    def do_POST(self):
        if self.path.startswith("/upload/"):
            try:
                card_id = int(self.path.split("/")[-1])
            except ValueError:
                self._send(400)
                return

            content_length = int(self.headers.get("Content-Length", 0))
            content_type = self.headers.get("Content-Type", "")

            if "multipart/form-data" in content_type:
                boundary = content_type.split("boundary=")[-1].encode()
                body = self.rfile.read(content_length)
                parts = body.split(b"--" + boundary)
                for part in parts:
                    if b"filename=" in part:
                        header_end = part.find(b"\r\n\r\n")
                        if header_end == -1:
                            continue
                        file_data = part[header_end + 4:]
                        if file_data.endswith(b"\r\n"):
                            file_data = file_data[:-2]
                        ct = "image/jpeg"
                        if b"image/png" in part[:header_end]:
                            ct = "image/png"
                        save_photo_bytes(card_id, file_data, ct)
                        upload_events.put(card_id)
                        count = len(get_photo_paths(card_id))
                        self._send(200, json.dumps({"ok": True, "count": count}).encode(),
                                   "application/json")
                        return

            self._send(400)
        else:
            self._send(404)


class _ReusableHTTPServer(HTTPServer):
    allow_reuse_address = True

    def server_bind(self):
        self.socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        self.socket.bind(self.server_address)
        self.server_address = self.socket.getsockname()
        self.server_name = "localhost"
        self.server_port = self.server_address[1]


def start_photo_server():
    try:
        server = _ReusableHTTPServer(("0.0.0.0", UPLOAD_PORT), PhotoUploadHandler)
    except OSError:
        return None
    server.daemon_threads = True
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    return server

# ============================================================
# AUTO-UPDATE
# ============================================================

SCRIPT_PATH = os.path.join(BASE_DIR, "pokemon_card_manager.py")


def _version_tuple(v):
    return tuple(int(x) for x in re.findall(r"\d+", v or ""))


def check_for_update():
    # Die API liefert die Datei frisch; raw.githubusercontent.com hängt nach einem Push
    # teils minutenlang hinterher und bleibt nur Ausweichweg (z. B. bei API-Limit).
    try:
        req = urllib.request.Request(GITHUB_API_URL, headers={"User-Agent": "PokemonCardManager",
                                                              "Accept": "application/vnd.github.raw"})
        with urllib.request.urlopen(req, timeout=15) as resp:
            remote_content = resp.read()
    except Exception:
        req = urllib.request.Request(GITHUB_RAW_URL, headers={"User-Agent": "PokemonCardManager"})
        with urllib.request.urlopen(req, timeout=15) as resp:
            remote_content = resp.read()
    remote_text = remote_content.decode("utf-8", errors="replace")
    match = re.search(r'^VERSION\s*=\s*"([^"]+)"', remote_text, re.MULTILINE)
    remote_version = match.group(1) if match else "unknown"
    # Nur echte neuere Versionen anbieten, sonst würde eine veraltete Kopie zurückstufen.
    if _version_tuple(remote_version) <= _version_tuple(VERSION):
        return False, VERSION, None
    return True, remote_version, remote_content


def apply_update(content):
    with open(SCRIPT_PATH, "wb") as f:
        f.write(content)


def restart_app():
    os.execv(sys.executable, [sys.executable, SCRIPT_PATH])



# ============================================================
# GUI: THEME & BAUSTEINE
# ============================================================

BG = "#0f1115"
SIDEBAR = "#0a0c0f"
SURFACE = "#171a20"
SURFACE2 = "#1f232b"
SURFACE3 = "#2a2f39"
SURFACE4 = "#363c48"
BORDER = "#272b34"
TEXT = "#eceef3"
MUTED = "#8e96a8"
FAINT = "#5d6475"
ACCENT = "#ffcb05"
ACCENT_HOVER = "#ffd84a"
ON_ACCENT = "#17181c"
ACCENT_SOFT = "#3a3214"
GREEN, GREEN_SOFT = "#34d399", "#12321f"
RED, RED_SOFT = "#f87171", "#3a1719"
AMBER, AMBER_SOFT = "#fbbf24", "#3a2c0c"
VIOLET, VIOLET_SOFT = "#a78bfa", "#29224a"
SKY, SKY_SOFT = "#38bdf8", "#0f2c3d"

MOD = "Command" if IS_MAC else "Control"
MOD_LABEL = "⌘" if IS_MAC else "Strg+"

STATUS_LABEL = {"Collection": "In Sammlung", "For Sale": "Zum Verkauf", "Sold": "Verkauft"}
CONDITION_LABEL = {"Mint": "Mint", "Near Mint": "Near Mint", "Excellent": "Excellent", "Good": "Good",
                   "Light Played": "Light Played", "Played": "Played", "Poor": "Poor"}

F = {}


def init_fonts(root):
    family = tkfont.nametofont("TkDefaultFont", root=root).actual("family")
    mk = lambda size, weight="normal": tkfont.Font(root=root, family=family, size=-size, weight=weight)
    F.update({
        "xs": mk(11), "xs_b": mk(11, "bold"), "sm": mk(12), "sm_b": mk(12, "bold"),
        "body": mk(13), "body_b": mk(13, "bold"), "md": mk(15), "md_b": mk(15, "bold"),
        "lg_b": mk(19, "bold"), "xl_b": mk(26, "bold"), "num_b": mk(22, "bold"),
    })


def card_state(c):
    """(Label, Farbe, Hintergrund) fuer den wichtigsten Status einer Karte."""
    if c["status"] == "Sold":
        return "Verkauft", RED, RED_SOFT
    if c["grading_status"] == "Submitted":
        return "Beim Grading", AMBER, AMBER_SOFT
    if c["status"] == "For Sale":
        return "Zum Verkauf", SKY, SKY_SOFT
    return "In Sammlung", GREEN, GREEN_SOFT


def grade_label(c):
    if c["grading_status"] == "Graded" and c["grading_service"]:
        g = fmt_grade(c["grade"])
        return f"{c['grading_service']} {g}".strip()
    return ""


def truncate(text, font, width):
    text = str(text)
    if font.measure(text) <= width:
        return text
    while text and font.measure(text + "…") > width:
        text = text[:-1]
    return text.rstrip() + "…"


def round_rect(canvas, x1, y1, x2, y2, r, **kw):
    r = min(r, (x2 - x1) / 2, (y2 - y1) / 2)
    pts = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2,
           x2 - r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]
    return canvas.create_polygon(pts, smooth=True, splinesteps=12, **kw)


def pill(canvas, x, y, text, fg, bg, font=None, anchor="nw", tags=()):
    """Kleines abgerundetes Label direkt auf einem Canvas."""
    font = font or F["xs_b"]
    w = font.measure(text) + 14
    h = font.metrics("linespace") + 6
    if anchor == "ne":
        x -= w
    elif anchor == "sw":
        y -= h
    elif anchor == "se":
        x, y = x - w, y - h
    round_rect(canvas, x, y, x + w, y + h, h / 2, fill=bg, outline="", tags=tags)
    canvas.create_text(x + w / 2, y + h / 2, text=text, fill=fg, font=font, tags=tags)
    return w, h


def setup_style(root):
    root.configure(bg=BG)
    style = ttk.Style(root)
    style.theme_use("clam")
    style.configure(".", background=BG, foreground=TEXT, fieldbackground=SURFACE2,
                    bordercolor=BORDER, lightcolor=SURFACE2, darkcolor=SURFACE2,
                    troughcolor=BG, selectbackground=SURFACE4, selectforeground=TEXT,
                    font=F["body"], focuscolor=ACCENT)
    style.configure("Treeview", background=SURFACE, fieldbackground=SURFACE, foreground=TEXT,
                    rowheight=58, borderwidth=0, font=F["body"])
    style.map("Treeview", background=[("selected", SURFACE3)], foreground=[("selected", TEXT)])
    style.configure("Treeview.Heading", background=SURFACE2, foreground=MUTED, relief="flat",
                    font=F["sm_b"], padding=(8, 7), borderwidth=0)
    style.map("Treeview.Heading", background=[("active", SURFACE3)])
    style.layout("Treeview", [("Treeview.treearea", {"sticky": "nswe"})])
    for orient in ("Vertical", "Horizontal"):
        style.layout(f"Slim.{orient}.TScrollbar", [
            (f"{orient}.Scrollbar.trough", {"sticky": "nswe", "children": [
                (f"{orient}.Scrollbar.thumb", {"expand": "1", "sticky": "nswe"})]})])
        style.configure(f"Slim.{orient}.TScrollbar", troughcolor=BG, background=SURFACE3,
                        bordercolor=BG, lightcolor=SURFACE3, darkcolor=SURFACE3,
                        arrowsize=9, gripcount=0, relief="flat")
        style.map(f"Slim.{orient}.TScrollbar", background=[("active", SURFACE4)])
    style.configure("TCombobox", fieldbackground=SURFACE2, background=SURFACE2, foreground=TEXT,
                    arrowcolor=MUTED, bordercolor=BORDER, lightcolor=SURFACE2, darkcolor=SURFACE2,
                    padding=(8, 6), insertcolor=TEXT)
    style.map("TCombobox", fieldbackground=[("readonly", SURFACE2), ("focus", SURFACE2)],
              foreground=[("readonly", TEXT)], selectbackground=[("readonly", SURFACE2)],
              selectforeground=[("readonly", TEXT)], bordercolor=[("focus", ACCENT)])
    root.option_add("*TCombobox*Listbox.background", SURFACE2)
    root.option_add("*TCombobox*Listbox.foreground", TEXT)
    root.option_add("*TCombobox*Listbox.selectBackground", SURFACE4)
    root.option_add("*TCombobox*Listbox.selectForeground", TEXT)
    root.option_add("*TCombobox*Listbox.font", F["body"])
    root.option_add("*Menu.background", SURFACE2)
    root.option_add("*Menu.foreground", TEXT)
    root.option_add("*Menu.activeBackground", SURFACE4)
    root.option_add("*Menu.activeForeground", TEXT)
    root.option_add("*Menu.font", F["body"])
    return style


class Btn(tk.Label):
    """Flacher Button (tk.Button ignoriert auf macOS Hintergrundfarben)."""

    KINDS = {
        "primary": (ACCENT, ACCENT_HOVER, ON_ACCENT),
        "secondary": (SURFACE3, SURFACE4, TEXT),
        "subtle": (SURFACE2, SURFACE3, TEXT),
        "danger": (SURFACE2, RED_SOFT, RED),
        "ghost": (None, SURFACE3, MUTED),
    }

    def __init__(self, parent, text, command=None, kind="secondary", small=False, bg=None, **kw):
        base, hover, fg = self.KINDS[kind]
        self._base = base or bg or parent.cget("bg")
        self._hover = hover
        self._fg = fg
        self._command = command
        self._enabled = True
        super().__init__(parent, text=text, bg=self._base, fg=fg, cursor="hand2",
                         font=F["sm_b"] if small else F["body_b"],
                         padx=10 if small else 14, pady=4 if small else 7, **kw)
        self.bind("<Enter>", lambda e: self._enabled and self.config(bg=self._hover))
        self.bind("<Leave>", lambda e: self.config(bg=self._base))
        self.bind("<ButtonRelease-1>", self._click)

    def _click(self, event):
        if not self._enabled or not self._command:
            return
        if 0 <= event.x <= self.winfo_width() and 0 <= event.y <= self.winfo_height():
            self._command()

    def set_enabled(self, enabled):
        self._enabled = enabled
        self.config(fg=self._fg if enabled else FAINT, cursor="hand2" if enabled else "arrow")

    def set_text(self, text):
        self.config(text=text)


class Field(tk.Frame):
    """Eingabefeld mit Innenabstand, Fokusrahmen und Platzhalter."""

    def __init__(self, parent, width=20, value="", placeholder="", textvariable=None,
                 show=None, suffix=None, font=None, bg=SURFACE2):
        super().__init__(parent, bg=bg, highlightthickness=1, highlightbackground=BORDER,
                         highlightcolor=ACCENT)
        self.var = textvariable or tk.StringVar(value=str(value) if value is not None else "")
        self.entry = tk.Entry(self, textvariable=self.var, width=width, bg=bg, fg=TEXT,
                              insertbackground=TEXT, relief="flat", bd=0, highlightthickness=0,
                              selectbackground=SURFACE4, selectforeground=TEXT,
                              font=font or F["body"], show=show or "",
                              disabledbackground=bg, disabledforeground=MUTED)
        self.entry.pack(side="left", fill="x", expand=True, padx=(10, 4 if suffix else 10), pady=7)
        if suffix:
            tk.Label(self, text=suffix, bg=bg, fg=MUTED, font=F["body"]).pack(side="right", padx=(0, 10))
        self._ph = placeholder
        self._ph_on = False
        if placeholder:
            self.entry.bind("<FocusIn>", self._ph_hide, add="+")
            self.entry.bind("<FocusOut>", self._ph_show, add="+")
            self._ph_show()
        self.entry.bind("<FocusIn>", lambda e: self.config(highlightbackground=ACCENT), add="+")
        self.entry.bind("<FocusOut>", lambda e: self.config(highlightbackground=BORDER), add="+")
        self.bind("<Button-1>", lambda e: self.entry.focus_set())

    def _ph_show(self, _e=None):
        if not self.var.get():
            self._ph_on = True
            self.entry.config(fg=FAINT)
            self.var.set(self._ph)

    def _ph_hide(self, _e=None):
        if self._ph_on:
            self._ph_on = False
            self.var.set("")
            self.entry.config(fg=TEXT)

    def get(self):
        return "" if self._ph_on else self.var.get().strip()

    def set(self, value):
        if self._ph_on:
            self._ph_hide()
        self.var.set("" if value is None else str(value))
        if not self.var.get() and self._ph and self.focus_get() is not self.entry:
            self._ph_show()

    def focus(self):
        self.entry.focus_set()
        self.entry.select_range(0, "end")


class Chips(tk.Frame):
    """Auswahl als Chips (statt Dropdown): alle Optionen sofort sichtbar."""

    def __init__(self, parent, options, value=None, command=None, bg=BG, labels=None, small=False):
        super().__init__(parent, bg=bg)
        self.options = list(options)
        self.value = value if value in self.options else (self.options[0] if self.options else None)
        self.command = command
        self.labels = labels or {}
        self._items = {}
        self._width = 0
        for opt in self.options:
            lbl = tk.Label(self, text=self.labels.get(opt, opt), font=F["sm_b"] if small else F["sm"],
                           padx=9 if small else 11, pady=3 if small else 5, cursor="hand2")
            lbl.bind("<Button-1>", lambda e, o=opt: self.set(o, notify=True))
            lbl.bind("<Enter>", lambda e, o=opt: self._hover(o, True))
            lbl.bind("<Leave>", lambda e, o=opt: self._hover(o, False))
            self._items[opt] = lbl
        self._paint()
        self.bind("<Configure>", self._reflow)
        self._layout(10_000)

    def _hover(self, opt, on):
        if opt != self.value:
            self._items[opt].config(bg=SURFACE3 if on else SURFACE2)

    def _paint(self):
        for opt, lbl in self._items.items():
            sel = opt == self.value
            lbl.config(bg=ACCENT_SOFT if sel else SURFACE2, fg=ACCENT if sel else MUTED)

    def _layout(self, width):
        x = row = col = 0
        for opt in self.options:
            lbl = self._items[opt]
            w = lbl.winfo_reqwidth() + 6
            if x and x + w > width:
                row += 1
                col = x = 0
            lbl.grid(row=row, column=col, padx=(0, 6), pady=(0, 6), sticky="w")
            col += 1
            x += w

    def _reflow(self, event):
        if abs(event.width - self._width) > 4:
            self._width = event.width
            for lbl in self._items.values():
                lbl.grid_forget()
            self._layout(event.width)

    def relabel(self, mapping):
        for k, text in mapping.items():
            if k in self._items:
                self._items[k].config(text=text)
        for lbl in self._items.values():
            lbl.grid_forget()
        self._layout(self._width or 10_000)

    def get(self):
        return self.value

    def set(self, value, notify=False):
        if value not in self._items:
            return
        self.value = value
        self._paint()
        if notify and self.command:
            self.command(value)


class ScrollFrame(tk.Frame):
    """Vertikal scrollbarer Bereich; Mausrad wird global verteilt (siehe App)."""

    def __init__(self, parent, bg=BG, padx=0):
        super().__init__(parent, bg=bg)
        self.canvas = tk.Canvas(self, bg=bg, highlightthickness=0, bd=0)
        self.sb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview,
                                style="Slim.Vertical.TScrollbar")
        self.inner = tk.Frame(self.canvas, bg=bg)
        self._win = self.canvas.create_window((padx, 0), window=self.inner, anchor="nw")
        self._padx = padx
        self.canvas.configure(yscrollcommand=self._on_scroll)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(
            self._win, width=max(1, e.width - 2 * self._padx)))
        self.canvas._scroll_target = self.canvas

    def _on_scroll(self, first, last):
        self.sb.set(first, last)
        if float(first) <= 0 and float(last) >= 1:
            self.sb.pack_forget()
        elif not self.sb.winfo_ismapped():
            self.sb.pack(side="right", fill="y")

    def to_top(self):
        self.canvas.yview_moveto(0)


class Toast:
    def __init__(self, root):
        self.root = root
        self.lbl = None
        self.job = None

    def show(self, text, kind="ok", ms=2600):
        if self.lbl:
            self.lbl.destroy()
        if self.job:
            self.root.after_cancel(self.job)
        fg = {"ok": GREEN, "err": RED, "info": TEXT}[kind]
        self.lbl = tk.Label(self.root, text=text, bg=SURFACE3, fg=fg, font=F["body_b"],
                            padx=18, pady=10)
        self.lbl.place(relx=0.5, rely=1.0, y=-24, anchor="s")
        self.lbl.lift()
        self.job = self.root.after(ms, self.hide)

    def hide(self):
        if self.lbl:
            self.lbl.destroy()
            self.lbl = None


class Worker:
    """Hintergrund-Threads fuer Netz & Bildbearbeitung. Ergebnisse kommen ueber
    eine Queue zurueck und werden im Tk-Thread verarbeitet (Tk ist nicht threadsicher)."""

    def __init__(self, root, threads=4):
        self.root = root
        self.tasks = queue.Queue()
        self.results = queue.Queue()
        self.pending = set()
        self.lock = threading.Lock()
        for _ in range(threads):
            threading.Thread(target=self._run, daemon=True).start()
        self._poll()

    def submit(self, key, fn, args=(), callback=None, front=False):
        with self.lock:
            if key in self.pending:
                return False
            self.pending.add(key)
        self.tasks.put((key, fn, args, callback))
        return True

    def busy(self, prefix=""):
        with self.lock:
            return sum(1 for k in self.pending if k.startswith(prefix))

    def _run(self):
        while True:
            key, fn, args, cb = self.tasks.get()
            try:
                res, err = fn(*args), None
            except Exception as e:
                res, err = None, e
            with self.lock:
                self.pending.discard(key)
            self.results.put((cb, res, err))

    def _poll(self):
        try:
            for _ in range(60):
                cb, res, err = self.results.get_nowait()
                if cb:
                    try:
                        cb(res, err)
                    except tk.TclError:
                        pass
        except queue.Empty:
            pass
        self.root.after(80, self._poll)


THUMB_DIR = os.path.join(CACHE_DIR, "thumbs")
os.makedirs(THUMB_DIR, exist_ok=True)


def make_thumb(src, w, h, kind, dim, radius):
    """Erzeugt (im Hintergrund) eine abgerundete Vorschau als PNG im Cache."""
    st = os.stat(src)
    key = hashlib.md5(f"{src}|{st.st_mtime}|{st.st_size}|{w}x{h}|{kind}|{dim}|{radius}|v2".encode()).hexdigest()
    out = os.path.join(THUMB_DIR, key + ".png")
    if os.path.exists(out):
        return out
    img = ImageOps.exif_transpose(Image.open(src)).convert("RGBA")
    if kind == "logo":
        canvas = Image.new("RGBA", (w, h), SURFACE2)
        logo = ImageOps.contain(img, (int(w * 0.84), int(h * 0.6)), Image.LANCZOS)
        canvas.alpha_composite(logo, ((w - logo.width) // 2, (h - logo.height) // 2))
        img = canvas
    else:
        img = ImageOps.fit(img, (w, h), Image.LANCZOS, centering=(0.5, 0.45))
    if dim:
        rgb = ImageEnhance.Color(img.convert("RGB")).enhance(0.15)
        rgb = ImageEnhance.Brightness(rgb).enhance(0.45)
        img = rgb.convert("RGBA")
    if radius:
        mask = Image.new("L", (w * 3, h * 3), 0)
        ImageDraw.Draw(mask).rounded_rectangle((0, 0, w * 3 - 1, h * 3 - 1), radius * 3, fill=255)
        img.putalpha(mask.resize((w, h), Image.LANCZOS))
    img.save(out)
    return out


class ThumbCache:
    """Haelt Tk-Bilder im Speicher; erzeugt fehlende Thumbs per Worker."""

    def __init__(self, worker):
        self.worker = worker
        self.mem = {}
        self.waiting = {}
        self.blanks = {}

    def blank(self, w, h):
        """Leeres Bild fester Groesse: Labels mit Bild messen in Pixeln statt Zeichen."""
        if (w, h) not in self.blanks:
            self.blanks[(w, h)] = tk.PhotoImage(width=w, height=h)
        return self.blanks[(w, h)]

    def get(self, src, w, h, kind="card", dim=False, radius=8, on_ready=None):
        if not src:
            return None
        try:
            mt = os.path.getmtime(src)
        except OSError:
            return None
        key = (src, mt, w, h, kind, dim, radius)
        if key in self.mem:
            return self.mem[key]
        if key in self.waiting:
            if on_ready:
                self.waiting[key].append(on_ready)
            return None
        self.waiting[key] = [on_ready] if on_ready else []

        def done(path, err):
            callbacks = self.waiting.pop(key, [])
            if err or not path:
                return
            try:
                self.mem[key] = ImageTk.PhotoImage(Image.open(path))
            except Exception:
                return
            for cb in callbacks:
                cb(self.mem[key])

        self.worker.submit(f"thumb:{key}", make_thumb, (src, w, h, kind, dim, radius), done)
        return None


class Dialog(tk.Toplevel):
    def __init__(self, app, title, subtitle="", width=460):
        super().__init__(app.root, bg=SURFACE)
        self.app = app
        self.title(title)
        self.resizable(False, False)
        self.transient(app.root)
        self._primary = None
        head = tk.Frame(self, bg=SURFACE)
        head.pack(fill="x", padx=24, pady=(22, 0))
        tk.Label(head, text=title, bg=SURFACE, fg=TEXT, font=F["lg_b"], anchor="w").pack(fill="x")
        if subtitle:
            tk.Label(head, text=subtitle, bg=SURFACE, fg=MUTED, font=F["body"], anchor="w",
                     justify="left", wraplength=width - 48).pack(fill="x", pady=(4, 0))
        self.body = tk.Frame(self, bg=SURFACE, width=width)
        self.body.pack(fill="both", expand=True, padx=24, pady=16)
        self.footer = tk.Frame(self, bg=SURFACE)
        self.footer.pack(fill="x", padx=24, pady=(0, 22))
        self.bind("<Escape>", lambda e: self.close())
        self.bind("<Return>", self._enter)
        self.protocol("WM_DELETE_WINDOW", self.close)

    def _enter(self, event):
        if self._primary and not isinstance(event.widget, tk.Text):
            self._primary()

    def button(self, text, command, kind="secondary"):
        b = Btn(self.footer, text, command, kind=kind, bg=SURFACE)
        b.pack(side="right", padx=(8, 0))
        if kind == "primary":
            self._primary = command
        return b

    def label(self, parent, text, row, col=0, pady=(10, 4)):
        tk.Label(parent, text=text, bg=SURFACE, fg=MUTED, font=F["sm_b"], anchor="w").grid(
            row=row, column=col, sticky="w", pady=pady)

    def show(self, focus=None):
        self.update_idletasks()
        r = self.app.root
        x = r.winfo_rootx() + (r.winfo_width() - self.winfo_reqwidth()) // 2
        y = r.winfo_rooty() + max(40, (r.winfo_height() - self.winfo_reqheight()) // 3)
        self.geometry(f"+{max(0, x)}+{max(0, y)}")
        try:
            self.grab_set()
        except tk.TclError:
            pass
        if focus is not None:
            focus.focus()
        else:
            self.focus_set()

    def close(self):
        try:
            self.grab_release()
        except tk.TclError:
            pass
        self.destroy()


# ============================================================
# GUI: HAUPTFENSTER
# ============================================================

NAV = [
    ("collection", "Sammlung"),
    ("add", "Hinzufügen"),
    ("grading", "Grading"),
    ("sales", "Verkäufe"),
    ("stats", "Statistik"),
    ("settings", "Einstellungen"),
]


def _slim_match(match):
    return {k: match.get(k) for k in ("id", "image", "kind", "name", "localId", "set_id", "set_name")}


class App:
    def __init__(self, root, photo_server=None):
        self.root = root
        init_fonts(root)
        setup_style(root)
        root.title("Poké Inventar")
        self.config = _load_config()
        root.geometry(self.config.get("geometry_v2", "1360x860"))
        root.minsize(1060, 680)
        root.protocol("WM_DELETE_WINDOW", self._on_close)
        if IS_MAC:
            root.createcommand("::tk::mac::Quit", self._on_close)

        self.meta = _load_json(META_PATH)
        self._meta_job = None
        self._soft_job = None
        self.worker = Worker(root, threads=5)
        self.thumbs = ThumbCache(self.worker)
        self.toast = Toast(root)
        self.cards, self.sales, self.by_id = [], [], {}
        self.entries, self.entry_by_id, self.lead_of = [], {}, {}
        self.current = None
        self.photo_server = photo_server

        self.sidebar = tk.Frame(root, bg=SIDEBAR, width=212)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)
        self.content = tk.Frame(root, bg=BG)
        self.content.pack(side="left", fill="both", expand=True)
        self._build_sidebar()

        self.views = {
            "collection": CollectionView(self.content, self),
            "add": AddView(self.content, self),
            "grading": GradingView(self.content, self),
            "sales": SalesView(self.content, self),
            "stats": StatsView(self.content, self),
            "settings": SettingsView(self.content, self),
        }

        self._touch_target, self._touch_acc = None, 0
        root.bind_all("<MouseWheel>", self._on_wheel)
        root.bind_all("<Shift-MouseWheel>", lambda e: None)
        try:
            # Ab Tk 8.7 meldet das Mac-Trackpad <TouchpadScroll> statt <MouseWheel>.
            root.bind_all("<TouchpadScroll>", self._on_touchpad)
        except tk.TclError:
            pass
        for i, (key, _label) in enumerate(NAV, 1):
            root.bind_all(f"<{MOD}-Key-{i}>", lambda e, k=key: self.show(k))
        root.bind_all(f"<{MOD}-n>", lambda e: self.show("add"))
        root.bind_all(f"<{MOD}-f>", lambda e: self._focus_search())
        root.bind_all(f"<{MOD}-r>", lambda e: self.reload())

        self.reload()
        self.show(self.config.get("view", "collection") if self.config.get("view") in self.views else "collection")
        root.after(700, self.ensure_images)
        root.after(2500, self.refresh_prices)
        root.after(1500, self._auto_check_update)

    # ---------- Layout ----------

    def _build_sidebar(self):
        sb = self.sidebar
        brand = tk.Frame(sb, bg=SIDEBAR)
        brand.pack(fill="x", padx=18, pady=(22, 26))
        ball = tk.Canvas(brand, width=30, height=30, bg=SIDEBAR, highlightthickness=0)
        ball.create_arc(2, 2, 28, 28, start=0, extent=180, fill="#ef4444", outline="")
        ball.create_arc(2, 2, 28, 28, start=180, extent=180, fill="#e5e7eb", outline="")
        ball.create_rectangle(2, 14, 28, 16, fill=SIDEBAR, outline="")
        ball.create_oval(10, 10, 20, 20, fill="#e5e7eb", outline=SIDEBAR, width=3)
        ball.pack(side="left")
        txt = tk.Frame(brand, bg=SIDEBAR)
        txt.pack(side="left", padx=(10, 0))
        tk.Label(txt, text="Poké Inventar", bg=SIDEBAR, fg=TEXT, font=F["md_b"]).pack(anchor="w")
        tk.Label(txt, text=f"Version {VERSION}", bg=SIDEBAR, fg=FAINT, font=F["xs"]).pack(anchor="w")

        self.nav = {}
        for i, (key, label) in enumerate(NAV, 1):
            row = tk.Frame(sb, bg=SIDEBAR, cursor="hand2")
            row.pack(fill="x", padx=10, pady=1)
            bar = tk.Frame(row, bg=SIDEBAR, width=3)
            bar.pack(side="left", fill="y")
            name = tk.Label(row, text=label, bg=SIDEBAR, fg=MUTED, font=F["body_b"], anchor="w",
                            padx=12, pady=9)
            name.pack(side="left", fill="x", expand=True)
            count = tk.Label(row, text="", bg=SIDEBAR, fg=FAINT, font=F["xs_b"], padx=10)
            count.pack(side="right")
            for w in (row, name, count, bar):
                w.bind("<Button-1>", lambda e, k=key: self.show(k))
                w.bind("<Enter>", lambda e, k=key: self._nav_hover(k, True))
                w.bind("<Leave>", lambda e, k=key: self._nav_hover(k, False))
            self.nav[key] = (row, bar, name, count)

        bottom = tk.Frame(sb, bg=SIDEBAR)
        bottom.pack(side="bottom", fill="x", padx=18, pady=18)
        tk.Label(bottom, text="HANDY-UPLOAD", bg=SIDEBAR, fg=FAINT, font=F["xs_b"]).pack(anchor="w")
        if self.photo_server:
            url = f"http://{_get_local_ip()}:{UPLOAD_PORT}"
            try:
                self._qr_photo = ImageTk.PhotoImage(qr_image(url, px=160))
                qr = tk.Label(bottom, image=self._qr_photo, bg=SIDEBAR, bd=0, cursor="hand2")
                qr.pack(anchor="w", pady=(8, 6))
                qr.bind("<Button-1>", lambda e: self.copy(url, "Upload-Adresse kopiert"))
            except Exception:
                pass
            link = tk.Label(bottom, text=url, bg=SIDEBAR, fg=ACCENT, font=F["sm_b"], cursor="hand2")
            link.pack(anchor="w", pady=(0, 2))
            link.bind("<Button-1>", lambda e: self.copy(url, "Upload-Adresse kopiert"))
            tk.Label(bottom, text="Mit der Handy-Kamera scannen,\nim selben WLAN.", bg=SIDEBAR,
                     fg=FAINT, font=F["xs"], justify="left").pack(anchor="w")
        else:
            tk.Label(bottom, text=f"Port {UPLOAD_PORT} ist belegt.\nLäuft die App doppelt?",
                     bg=SIDEBAR, fg=RED, font=F["xs"], justify="left").pack(anchor="w", pady=(4, 0))
        self.activity = tk.Label(sb, text="", bg=SIDEBAR, fg=FAINT, font=F["xs"], anchor="w")
        self.activity.pack(side="bottom", fill="x", padx=18)
        self._activity_tick()
        self._upload_tick()

    def _upload_tick(self):
        ids = []
        while True:
            try:
                ids.append(upload_events.get_nowait())
            except queue.Empty:
                break
        for cid in dict.fromkeys(ids):
            self.card_image_changed(cid)
        if ids:
            c = self.by_id.get(ids[-1])
            name = c["name"] if c and c["name"] else f"#{ids[-1]:03d}"
            self.toast.show(f"Foto vom Handy: {name}" if len(ids) == 1
                            else f"{len(ids)} Fotos vom Handy", "ok")
        self.root.after(500, self._upload_tick)

    def _activity_tick(self):
        n_img = self.worker.busy("resolve:")
        n_price = self.worker.busy("price:")
        parts = []
        if n_img:
            parts.append(f"Zuordnung: {n_img} offen")
        if n_price:
            parts.append(f"Preise: {n_price} offen")
        self.activity.config(text=" · ".join(parts))
        self.root.after(700, self._activity_tick)

    def _nav_hover(self, key, on):
        if key == self.current:
            return
        row, bar, name, count = self.nav[key]
        bg = SURFACE if on else SIDEBAR
        for w in (row, name, count):
            w.config(bg=bg)

    def _paint_nav(self):
        for key, (row, bar, name, count) in self.nav.items():
            sel = key == self.current
            bg = SURFACE2 if sel else SIDEBAR
            for w in (row, name, count):
                w.config(bg=bg)
            bar.config(bg=ACCENT if sel else SIDEBAR)
            name.config(fg=TEXT if sel else MUTED)

    def _update_counts(self):
        stock = sum(1 for c in self.entries if c["status"] != "Sold")
        grading = sum(1 for c in self.cards if c["status"] != "Sold"
                      and (c["vormerkung"] or c["grading_status"] == "Submitted"))
        sales = sum(1 for s in self.sales)
        for key, n in (("collection", stock), ("grading", grading), ("sales", sales)):
            self.nav[key][3].config(text=str(n) if n else "")

    def show(self, name, **kw):
        if self.current == name and not kw:
            return
        if self.current:
            self.views[self.current].pack_forget()
        self.current = name
        view = self.views[name]
        view.pack(fill="both", expand=True)
        if view.dirty:
            view.refresh()
            view.dirty = False
        view.on_show(**kw)
        self._paint_nav()

    def show_card(self, card_id):
        self.show("collection")
        self.views["collection"].reveal(card_id)

    def _focus_search(self):
        v = self.views.get(self.current)
        if hasattr(v, "search"):
            v.search.focus()
        else:
            self.show("collection")
            self.views["collection"].search.focus()

    # ---------- Daten ----------

    def reload(self):
        try:
            self.cards = load_all_cards()
            self.sales = load_sales()
        except PermissionError:
            messagebox.showerror("Excel gesperrt", "Die Excel-Datei ist gerade geöffnet oder gesperrt.\n"
                                 "Bitte in Excel schließen und erneut versuchen.")
        except Exception as e:
            messagebox.showerror("Fehler", f"Excel konnte nicht gelesen werden:\n{e}")
        self.by_id = {c["id"]: c for c in self.cards}
        self.entries = group_cards(self.cards)
        self.entry_by_id = {e["id"]: e for e in self.entries}
        self.lead_of = {c["id"]: e["id"] for e in self.entries for c in e.get("lots", [])}
        self._update_counts()
        for v in self.views.values():
            v.dirty = True
        if self.current:
            self.views[self.current].refresh()
            self.views[self.current].dirty = False

    def mutate(self, fn, *args, ok=None, select=None):
        """Excel-Aenderung mit Fehlerbehandlung, danach neu laden."""
        try:
            fn(*args)
        except PermissionError:
            messagebox.showerror("Excel gesperrt", "Die Excel-Datei ist gerade geöffnet.\n"
                                 "Bitte in Excel schließen und erneut versuchen.")
            return False
        except Exception as e:
            messagebox.showerror("Fehler", str(e))
            return False
        self.reload()
        if ok:
            self.toast.show(ok)
        if select is not None and self.current == "collection":
            self.views["collection"].select(select)
        return True

    def copy(self, text, msg="Kopiert"):
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.toast.show(msg, "info")

    def image_for(self, card_id):
        return card_image_path(self.meta, card_id)

    def save_meta_later(self):
        if self._meta_job:
            self.root.after_cancel(self._meta_job)
        self._meta_job = self.root.after(800, self._save_meta)

    def _save_meta(self):
        self._meta_job = None
        try:
            _save_json(META_PATH, self.meta)
        except OSError:
            pass

    def soft_refresh_later(self):
        """Preise/Bilder kamen rein: aktuelle Ansicht gesammelt neu zeichnen."""
        if self._soft_job:
            self.root.after_cancel(self._soft_job)
        self._soft_job = self.root.after(900, self._soft_refresh)

    def _soft_refresh(self):
        self._soft_job = None
        for key, v in self.views.items():
            if key == self.current:
                v.soft_refresh()
            else:
                v.dirty = True

    # ---------- Bilder & Preise ----------

    def ensure_images(self, cards=None, force=False):
        now = _time.time()
        for c in cards or self.cards:
            k = str(c["id"])
            m = self.meta.get(k, {})
            if m.get("status") == "hidden":
                continue
            if m.get("match"):
                if m["match"].get("kind") == "card" and not m.get("price"):
                    self._fetch_price(c["id"], m["match"]["id"])
                continue
            if m.get("status") == "noimage" and not force:
                continue
            if not force and m.get("checked", 0) > now - 3 * 86400:
                continue
            self.worker.submit(f"resolve:{c['id']}", resolve_card_image,
                               (c["name"], c["set"], c["number"]),
                               lambda res, err, cid=c["id"]: self._on_resolved(cid, res, err))

    def _on_resolved(self, card_id, res, err):
        if err or not res:
            return
        m = self.meta.setdefault(str(card_id), {})
        m["checked"] = _time.time()
        m["status"] = res["status"]
        if res["status"] in ("ok", "noimage"):
            m["match"] = _slim_match(res["match"])
            if m["match"].get("kind") == "card":
                self._fetch_price(card_id, m["match"]["id"])
            self.card_image_changed(card_id)
        else:
            m["candidates"] = len(res.get("candidates", []))
        self.save_meta_later()

    def set_card_match(self, card_id, match, manual=True):
        """TCGdex-Karte zuordnen (fuer den Marktpreis), aus Auswahl-Dialog oder beim Anlegen."""
        m = self.meta.setdefault(str(card_id), {})
        m.pop("file", None)
        m.pop("price", None)
        m.update(status="manual" if manual else "ok", match=_slim_match(match), checked=_time.time())
        if m["match"].get("kind") == "card":
            self._fetch_price(card_id, m["match"]["id"])
        self.save_meta_later()
        self.card_image_changed(card_id)

    def hide_card_image(self, card_id):
        m = self.meta.setdefault(str(card_id), {})
        m.clear()
        m.update(status="hidden", checked=_time.time())
        self.save_meta_later()
        self.card_image_changed(card_id)

    def card_image_changed(self, card_id):
        card_id = self.lead_of.get(card_id, card_id)
        v = self.views.get(self.current)
        if v:
            v.image_changed(card_id)
        for key, other in self.views.items():
            if key != self.current:
                other.dirty = True

    def refresh_prices(self, force=False):
        if not force and not self.config.get("auto_prices", True):
            return 0
        now = _time.time()
        n = 0
        for c in self.cards:
            if c["status"] == "Sold":
                continue
            m = self.meta.get(str(c["id"]), {})
            match = m.get("match")
            if not match or match.get("kind") != "card":
                continue
            p = m.get("price")
            if not force and p and now - p.get("updated", 0) < PRICE_MAX_AGE:
                continue
            n += self._fetch_price(c["id"], match["id"])
        return n

    def _fetch_price(self, card_id, tcg_id):
        def done(res, err):
            if err or not res:
                return
            self.meta.setdefault(str(card_id), {})["price"] = res
            self.save_meta_later()
            self.soft_refresh_later()
        return 1 if self.worker.submit(f"price:{card_id}", fetch_card_price, (tcg_id,), done) else 0

    def price_info(self, card_id):
        return (self.meta.get(str(card_id)) or {}).get("price") or {}

    # ---------- Aktionen ----------

    def open_cardmarket(self, card_id):
        url = get_card_url(card_id)
        if url:
            webbrowser.open(url)
            return
        c = self.by_id.get(card_id)
        if c:
            q = urllib.parse.quote(f"{c['name']} {c['number']}".strip())
            webbrowser.open(f"https://www.cardmarket.com/de/Pokemon/Products/Search?searchString={q}")

    def entry(self, card_id):
        """Anzeige-Eintrag (bei Mehrfachkäufen die Gruppe) zu einer Karten-ID."""
        return self.entry_by_id.get(self.lead_of.get(card_id, card_id)) or self.by_id.get(card_id)

    def edit(self, card_id):
        e = self.entry_by_id.get(card_id)
        if e and "lots" in e:
            self.show_card(card_id)
            self.toast.show("Mehrere Käufe: im Detail unter „Käufe“ einzeln bearbeiten", "info")
            return
        EditDialog(self, self.by_id[card_id])

    def sell(self, card_id):
        SellDialog(self, self.entry(card_id))

    def ebay(self, card_id):
        cfg = _load_ebay_config()
        if not cfg.get("access_token"):
            if messagebox.askyesno("eBay nicht verbunden",
                                   "Die App ist noch nicht mit eBay verbunden.\n"
                                   "Jetzt zu den Einstellungen wechseln?"):
                self.show("settings")
            return
        EbayDialog(self, self.entry(card_id))

    def toggle_queue(self, card_id):
        c = self.by_id[card_id]
        if c["vormerkung"]:
            self.mutate(toggle_grading_vormerkung, card_id, ok="Vormerkung entfernt", select=card_id)
        else:
            QueueDialog(self, c)

    def submit_grading(self, card_ids):
        SubmitDialog(self, [self.by_id[i] for i in card_ids if i in self.by_id])

    def grading_back(self, card_id):
        ReturnDialog(self, self.by_id[card_id])

    def cancel_grading(self, card_id):
        if messagebox.askyesno("Einreichung zurücknehmen",
                               "Die Karte wieder als nicht eingereicht markieren?"):
            self.mutate(cancel_grading, card_id, ok="Einreichung zurückgenommen", select=card_id)

    def set_for_sale(self, card_id, on):
        ids = [c["id"] for c in self.entry(card_id).get("lots", [])] or [card_id]
        status = "For Sale" if on else "Collection"
        self.mutate(lambda: [update_card_status(i, status) for i in ids],
                    ok="Als „Zum Verkauf“ markiert" if on else "Wieder in der Sammlung", select=card_id)

    def add_photos(self, card_id):
        paths = filedialog.askopenfilenames(
            parent=self.root, title="Fotos auswählen",
            filetypes=[("Bilder", "*.jpg *.jpeg *.png *.heic *.heif *.JPG *.JPEG *.PNG *.HEIC")])
        if not paths:
            return
        failed = 0
        for p in paths:
            try:
                save_photo(card_id, p)
            except Exception:
                failed += 1
        self.card_image_changed(card_id)
        self.toast.show(f"{len(paths) - failed} Foto(s) hinzugefügt" +
                        (f", {failed} fehlgeschlagen (HEIC?)" if failed else ""),
                        "err" if failed else "ok")

    def view_images(self, card_id, start=0):
        c = self.entry(card_id)
        if not c:
            return
        items = [(p, "photo") for p in entry_photo_paths(c["id"])]
        if not items:
            self.toast.show("Noch kein Foto vorhanden", "info")
            return
        Viewer(self, c, items, start)

    def pick_image(self, card_id):
        c = self.by_id[card_id]
        ImagePicker(self, c["name"], c["set"], on_pick=lambda match: self._picked(card_id, match),
                    on_none=lambda: self.hide_card_image(card_id))

    def _picked(self, card_id, match):
        self.set_card_match(card_id, match)
        self.toast.show("Karte zugeordnet, Marktpreis wird geladen …", "info")

    def delete(self, card_id):
        e = self.entry_by_id.get(card_id)
        if e and "lots" in e:
            n = len(e["lots"])
            if not messagebox.askyesno("Eintrag löschen",
                                       f"{e['name']}: alle {n} Käufe ({e['quantity']} Stück) löschen?\n\n"
                                       "Fotos und Verkaufseinträge werden ebenfalls entfernt. "
                                       "Das lässt sich nicht rückgängig machen.\n\n"
                                       "Einzelne Käufe löschst du im Detail unter „Käufe“."):
                return
            ids = [c["id"] for c in e["lots"]]
            if self.mutate(lambda: [delete_card(i) for i in ids], ok=f"{n} Käufe gelöscht"):
                for i in ids:
                    self.meta.pop(str(i), None)
                self.save_meta_later()
            return
        c = self.by_id[card_id]
        if not messagebox.askyesno("Karte löschen",
                                   f"#{card_id:03d} {c['name']} wirklich löschen?\n\n"
                                   "Fotos und Verkaufseinträge dieser Karte werden ebenfalls entfernt. "
                                   "Das lässt sich nicht rückgängig machen."):
            return
        if self.mutate(delete_card, card_id, ok=f"#{card_id:03d} gelöscht"):
            self.meta.pop(str(card_id), None)
            self.save_meta_later()

    def delete_lot(self, card_id):
        c = self.by_id[card_id]
        if not messagebox.askyesno("Kauf löschen",
                                   f"Kauf #{card_id:03d} ({fmt_eur(c['price'])}, {fmt_date(c['purchase_date'])}) "
                                   "löschen?\n\nFotos und Verkaufseinträge dieses Kaufs werden ebenfalls entfernt."):
            return
        if self.mutate(delete_card, card_id, ok=f"Kauf #{card_id:03d} gelöscht"):
            self.meta.pop(str(card_id), None)
            self.save_meta_later()

    def card_menu(self, card_id, event):
        c = self.entry(card_id)
        if not c:
            return
        menu = tk.Menu(self.root, tearoff=0)
        sold = c["status"] == "Sold"
        menu.add_command(label="Fotos ansehen", command=lambda: self.view_images(card_id))
        menu.add_command(label="Auf Cardmarket öffnen", command=lambda: self.open_cardmarket(card_id))
        menu.add_separator()
        if not sold:
            menu.add_command(label="Verkauf eintragen …", command=lambda: self.sell(card_id))
            menu.add_command(label="Auf eBay listen …", command=lambda: self.ebay(card_id))
            if c["status"] == "For Sale":
                menu.add_command(label="Nicht mehr zum Verkauf", command=lambda: self.set_for_sale(card_id, False))
            else:
                menu.add_command(label="Als „Zum Verkauf“ markieren", command=lambda: self.set_for_sale(card_id, True))
            menu.add_separator()
            if c["grading_status"] == "Submitted":
                menu.add_command(label="Grading zurück erhalten …", command=lambda: self.grading_back(card_id))
            else:
                menu.add_command(label="Vormerkung entfernen" if c["vormerkung"] else "Für Grading vormerken …",
                                 command=lambda: self.toggle_queue(card_id))
                menu.add_command(label="Zum Grading einreichen …", command=lambda: self.submit_grading([card_id]))
            menu.add_separator()
        menu.add_command(label="Fotos hinzufügen …", command=lambda: self.add_photos(card_id))
        menu.add_command(label="Für Marktpreis zuordnen …", command=lambda: self.pick_image(card_id))
        menu.add_command(label="Bearbeiten …", command=lambda: self.edit(card_id))
        menu.add_separator()
        menu.add_command(label="Löschen …", command=lambda: self.delete(card_id))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    # ---------- Fenster ----------

    def _wheel_target(self, e):
        try:
            w = self.root.winfo_containing(e.x_root, e.y_root)
        except (KeyError, tk.TclError):
            return None
        while w is not None:
            if isinstance(w, (ttk.Treeview, tk.Listbox, tk.Text)):
                return None
            target = getattr(w, "_scroll_target", None)
            if target is not None:
                return target
            w = getattr(w, "master", None)
        return None

    def _on_wheel(self, e):
        target = self._wheel_target(e)
        if target is None:
            return
        # Windows und Tk 9 liefern Vielfache von 120, Tk 8.6 auf dem
        # Mac kleine Werte pro Schritt.
        if abs(e.delta) >= 120:
            steps = -int(e.delta / 120) * 3
        else:
            steps = -e.delta
        if steps:
            target.yview_scroll(steps, "units")

    def _on_touchpad(self, e):
        target = self._wheel_target(e)
        if target is None:
            return
        # delta packt dx (obere 16 Bit) und dy (untere 16 Bit) in Pixeln.
        dy = e.delta & 0xFFFF
        if dy >= 0x8000:
            dy -= 0x10000
        if not dy:
            return
        # Mit yscrollincrement rundet die Canvas jede Position auf dieses
        # Raster; die wenigen Pixel pro Ereignis würden einfach verpuffen.
        inc = int(float(target.cget("yscrollincrement")))
        if inc > 0:
            if target is not self._touch_target:
                self._touch_target, self._touch_acc = target, 0
            self._touch_acc -= dy
            steps = int(self._touch_acc / inc)
            if steps:
                self._touch_acc -= steps * inc
                target.yview_scroll(steps, "units")
            return
        try:
            x0, y0, x1, y1 = (float(v) for v in target.cget("scrollregion").split())
        except ValueError:
            return
        if y1 - y0 > 0:
            target.yview_moveto(target.yview()[0] - dy / (y1 - y0))

    def _save_state(self):
        cfg = _load_config()
        cfg.update(self.config)
        cfg["geometry_v2"] = self.root.geometry()
        cfg["view"] = self.current
        for v in self.views.values():
            v.save_state(cfg)
        _save_config(cfg)
        if self._meta_job:
            self._save_meta()

    def _on_close(self):
        try:
            self._save_state()
        finally:
            self.root.destroy()

    def _auto_check_update(self):
        def check():
            return check_for_update()

        def done(res, err):
            if err or not res or not res[0]:
                self.update_state = "aktuell" if res and not res[0] else "unbekannt"
                return
            self._offer_update(res[1], res[2])
        self.worker.submit("update", check, (), done)

    def _offer_update(self, remote_version, content):
        if not messagebox.askyesno(
            "Update verfügbar",
            f"Version {remote_version} ist verfügbar (installiert: {VERSION}).\n\n"
            "Jetzt herunterladen und installieren?\nDeine Daten bleiben unverändert.",
        ):
            return
        try:
            apply_update(content)
            messagebox.showinfo("Aktualisiert", f"Auf Version {remote_version} aktualisiert.\n"
                                "Die App startet jetzt neu.")
            self._save_state()
            restart_app()
        except Exception as e:
            messagebox.showerror("Update-Fehler", f"Update konnte nicht installiert werden:\n{e}")

    def check_update_now(self):
        try:
            has_update, remote_version, content = check_for_update()
        except Exception as e:
            messagebox.showerror("Update-Fehler", f"Konnte nicht nach Updates suchen:\n{e}")
            return
        if not has_update:
            messagebox.showinfo("Aktuell", f"Du nutzt die neueste Version ({VERSION}).")
            return
        self._offer_update(remote_version, content)


class View(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=BG)
        self.app = app
        self.dirty = True

    def header(self, parent, title, subtitle=""):
        h = tk.Frame(parent, bg=BG)
        h.pack(fill="x", padx=28, pady=(24, 0))
        left = tk.Frame(h, bg=BG)
        left.pack(side="left", fill="x", expand=True)
        tk.Label(left, text=title, bg=BG, fg=TEXT, font=F["xl_b"], anchor="w").pack(anchor="w")
        sub = tk.Label(left, text=subtitle, bg=BG, fg=MUTED, font=F["body"], anchor="w")
        sub.pack(anchor="w", pady=(2, 0))
        right = tk.Frame(h, bg=BG)
        right.pack(side="right", anchor="s")
        return sub, right

    def refresh(self):
        pass

    def soft_refresh(self):
        self.refresh()

    def on_show(self, **kw):
        pass

    def image_changed(self, card_id):
        self.dirty = True

    def save_state(self, cfg):
        pass


# ============================================================
# GUI: SAMMLUNG
# ============================================================

def card_matches(c, terms):
    hay = " ".join(str(x) for x in (
        c["name"], c["set"], c["number"], c["language"], f"#{c['id']:03d}", c["id"],
        c["grading_service"], fmt_grade(c["grade"]), c["vormerkung_notiz"], c["cert"],
        c["condition"], c["source"])).lower()
    return all(t in hay for t in terms)


class Gallery(tk.Frame):
    IMG_W, IMG_H = 150, 209
    PAD = 10
    TILE_W = IMG_W + 2 * PAD
    TILE_H = IMG_H + PAD + 80
    GAP = 16

    def __init__(self, parent, app, on_select, on_open):
        super().__init__(parent, bg=BG)
        self.app = app
        self.on_select = on_select
        self.on_open = on_open
        self.canvas = tk.Canvas(self, bg=BG, highlightthickness=0, bd=0, yscrollincrement=20,
                                takefocus=1)
        sb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview,
                           style="Slim.Vertical.TScrollbar")
        self.canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas._scroll_target = self.canvas
        self.cards = []
        self.pos = {}
        self.selected = None
        self.hover = None
        self.cols = 0
        self.empty_text = ""
        self.canvas.bind("<Configure>", self._on_resize)
        self.canvas.bind("<Button-1>", self._click)
        self.canvas.bind("<Double-Button-1>", self._double)
        self.canvas.bind("<Motion>", self._motion)
        self.canvas.bind("<Leave>", lambda e: self._set_hover(None))
        for seq in ("<Button-3>",) + (("<Button-2>", "<Control-Button-1>") if IS_MAC else ()):
            self.canvas.bind(seq, self._menu)

    def _cols_for(self, width):
        return max(1, (width - 8 + self.GAP) // (self.TILE_W + self.GAP))

    def _on_resize(self, e):
        cols = self._cols_for(e.width)
        if cols != self.cols:
            self.cols = cols
            self.render()

    def set_items(self, cards, empty_text=""):
        self.cards = cards
        self.empty_text = empty_text
        self.render()

    def render(self):
        cv = self.canvas
        cv.delete("all")
        self.pos.clear()
        width = max(cv.winfo_width(), 300)
        self.cols = self._cols_for(width)
        if not self.cards:
            cv.create_text(width / 2, 120, text=self.empty_text or "Keine Karten", fill=MUTED,
                           font=F["md"], justify="center")
            cv.configure(scrollregion=(0, 0, width, 200))
            return
        # Kacheln zentrieren, damit rechts kein unruhiger Restrand bleibt.
        used = self.cols * self.TILE_W + (self.cols - 1) * self.GAP
        x0 = max(4, (width - used) // 2)
        for i, c in enumerate(self.cards):
            r, col = divmod(i, self.cols)
            x = x0 + col * (self.TILE_W + self.GAP)
            y = 6 + r * (self.TILE_H + self.GAP)
            self.pos[c["id"]] = (x, y)
            self._draw_tile(c, x, y)
        rows = (len(self.cards) + self.cols - 1) // self.cols
        cv.configure(scrollregion=(0, 0, width, 12 + rows * (self.TILE_H + self.GAP)))

    def _draw_tile(self, c, x, y):
        cv = self.canvas
        cid = c["id"]
        t = f"t{cid}"
        sel = cid == self.selected
        round_rect(cv, x, y, x + self.TILE_W, y + self.TILE_H, 14,
                   fill=SURFACE2 if cid == self.hover else SURFACE,
                   outline=ACCENT if sel else BORDER, width=2 if sel else 1, tags=(t, f"bg{cid}"))
        ix, iy = x + self.PAD, y + self.PAD
        path, kind = self.app.image_for(cid)
        sold = c["status"] == "Sold"
        img = self.app.thumbs.get(path, self.IMG_W, self.IMG_H, kind or "card", sold, 9,
                                  on_ready=lambda im, cid=cid: self._set_image(cid, im))
        if img:
            cv.create_image(ix, iy, image=img, anchor="nw", tags=(t, f"img{cid}"))
        else:
            self._placeholder(c, ix, iy, t)

        # Status oben links, Menge oben rechts, Grade unten links auf dem Bild.
        label, fg, bg = card_state(c)
        by = iy + 8
        if label != "In Sammlung":
            _w, h = pill(cv, ix + 8, by, label, fg, bg, tags=(t, f"badge{cid}"))
            by += h + 4
        if c["vormerkung"] and not sold:
            pill(cv, ix + 8, by, "Vorgemerkt", VIOLET, VIOLET_SOFT, tags=(t, f"badge{cid}"))
        if c["quantity"] > 1 and not sold:
            pill(cv, ix + self.IMG_W - 8, iy + 8, f"×{c['quantity']}", TEXT, SIDEBAR,
                 anchor="ne", tags=(t, f"badge{cid}"))
        g = grade_label(c)
        if g:
            pill(cv, ix + 8, iy + self.IMG_H - 8, g, ON_ACCENT, ACCENT, anchor="sw",
                 tags=(t, f"badge{cid}"))

        tw = self.TILE_W - 2 * self.PAD
        ty = iy + self.IMG_H + 10
        cv.create_text(ix, ty, text=truncate(c["name"] or "Ohne Namen", F["body_b"], tw), anchor="nw",
                       fill=FAINT if sold else TEXT, font=F["body_b"], tags=(t,))
        sub = c["set"] + (f" · {c['number']}" if c["number"] else "")
        cv.create_text(ix, ty + 19, text=truncate(sub, F["sm"], tw), anchor="nw", fill=MUTED,
                       font=F["sm"], tags=(t,))
        py = ty + 44
        cv.create_text(ix, py, text=fmt_eur(c["price"]), anchor="nw", fill=MUTED if not sold else FAINT,
                       font=F["sm"], tags=(t,))
        trend = market_price(self.app.meta, cid) if c["grading_status"] != "Graded" else None
        if trend is not None and not sold:
            col = GREEN if trend >= c["price"] else RED
            arrow = "▲" if trend >= c["price"] else "▼"
            cv.create_text(ix + tw, py, text=f"{arrow} {fmt_eur(trend)}", anchor="ne", fill=col,
                           font=F["sm_b"], tags=(t,))
        elif c["language"]:
            cv.create_text(ix + tw, py, text=c["language"], anchor="ne", fill=FAINT,
                           font=F["xs_b"], tags=(t,))

    def _placeholder(self, c, x, y, t):
        cv = self.canvas
        cid = c["id"]
        round_rect(cv, x, y, x + self.IMG_W, y + self.IMG_H, 9, fill=SURFACE3, outline="",
                   tags=(t, f"ph{cid}"))
        cx, cy = x + self.IMG_W / 2, y + self.IMG_H / 2 - 10
        cv.create_oval(cx - 22, cy - 22, cx + 22, cy + 22, outline=SURFACE4, width=3, tags=(t, f"ph{cid}"))
        cv.create_line(cx - 22, cy, cx + 22, cy, fill=SURFACE4, width=3, tags=(t, f"ph{cid}"))
        cv.create_oval(cx - 7, cy - 7, cx + 7, cy + 7, fill=SURFACE3, outline=SURFACE4, width=3,
                       tags=(t, f"ph{cid}"))
        hint = "Noch kein Foto"
        if hint:
            cv.create_text(cx, cy + 46, text=hint, fill=FAINT, font=F["xs"], justify="center",
                           tags=(t, f"ph{cid}"))

    def _set_image(self, cid, img):
        if cid not in self.pos:
            return
        x, y = self.pos[cid]
        cv = self.canvas
        cv.delete(f"ph{cid}")
        cv.delete(f"img{cid}")
        item = cv.create_image(x + self.PAD, y + self.PAD, image=img, anchor="nw",
                               tags=(f"t{cid}", f"img{cid}"))
        cv.tag_raise(f"badge{cid}", item)

    def redraw_card(self, cid):
        c = next((x for x in self.cards if x["id"] == cid), None)
        if c is None or cid not in self.pos:
            return
        self.canvas.delete(f"t{cid}")
        x, y = self.pos[cid]
        self._draw_tile(c, x, y)

    def _id_at(self, event):
        cv = self.canvas
        x, y = cv.canvasx(event.x), cv.canvasy(event.y)
        for cid, (tx, ty) in self.pos.items():
            if tx <= x <= tx + self.TILE_W and ty <= y <= ty + self.TILE_H:
                return cid
        return None

    def _click(self, e):
        self.canvas.focus_set()
        cid = self._id_at(e)
        if cid is not None:
            self.select(cid)
            self.on_select(cid)

    def _double(self, e):
        cid = self._id_at(e)
        if cid is not None:
            self.on_open(cid)

    def _menu(self, e):
        cid = self._id_at(e)
        if cid is not None:
            self.select(cid)
            self.on_select(cid)
            self.app.card_menu(cid, e)

    def _motion(self, e):
        self._set_hover(self._id_at(e))

    def _set_hover(self, cid):
        if cid == self.hover:
            return
        old, self.hover = self.hover, cid
        if old is not None:
            self.canvas.itemconfigure(f"bg{old}", fill=SURFACE)
        if cid is not None:
            self.canvas.itemconfigure(f"bg{cid}", fill=SURFACE2)
        self.canvas.config(cursor="hand2" if cid is not None else "")

    def select(self, cid, scroll=True):
        old, self.selected = self.selected, cid
        if old is not None:
            self.canvas.itemconfigure(f"bg{old}", outline=BORDER, width=1)
        if cid is not None and cid in self.pos:
            self.canvas.itemconfigure(f"bg{cid}", outline=ACCENT, width=2)
            if scroll:
                self._ensure_visible(cid)

    def _ensure_visible(self, cid):
        cv = self.canvas
        if cv.winfo_height() < 50 or cid not in self.pos:
            # Gerade erst eingeblendet: Hoehe noch unbekannt, sonst springt
            # die Ansicht ans Ende.
            self.after(60, lambda: cid == self.selected and cid in self.pos and self._ensure_visible(cid))
            return
        _x, y = self.pos[cid]
        sr = cv.cget("scrollregion").split()
        total = float(sr[3]) if len(sr) == 4 else 1
        top, bottom = cv.canvasy(0), cv.canvasy(cv.winfo_height())
        if y < top:
            cv.yview_moveto(max(0, (y - 8) / total))
        elif y + self.TILE_H > bottom:
            cv.yview_moveto(max(0, (y + self.TILE_H + 8 - cv.winfo_height()) / total))

    def move(self, dx, dy):
        if not self.cards:
            return None
        ids = [c["id"] for c in self.cards]
        if self.selected not in ids:
            new = 0
        else:
            new = ids.index(self.selected) + dx + dy * self.cols
        new = max(0, min(len(ids) - 1, new))
        self.select(ids[new])
        return ids[new]


class CardList(tk.Frame):
    """Tabellenansicht mit Mini-Bildern; Spaltenkopf klicken sortiert."""

    COLS = [("name", "Karte", 150, "w"), ("set", "Set", 120, "w"), ("number", "Nr.", 64, "w"),
            ("lang", "Spr.", 40, "center"), ("qty", "Anz.", 40, "center"),
            ("price", "Kaufpreis", 80, "e"), ("market", "Marktwert", 84, "e"),
            ("status", "Status", 120, "w")]

    def __init__(self, parent, app, on_select, on_open):
        super().__init__(parent, bg=BG)
        self.app = app
        self.on_select = on_select
        self.on_open = on_open
        self.tree = ttk.Treeview(self, columns=[c[0] for c in self.COLS], show="tree headings",
                                 selectmode="browse")
        self.tree.column("#0", width=62, stretch=False, anchor="center")
        self.tree.heading("#0", text="")
        for key, label, w, anchor in self.COLS:
            self.tree.heading(key, text=label, anchor=anchor, command=lambda k=key: self._sort_by(k))
            self.tree.column(key, width=w, anchor=anchor, stretch=key in ("name", "set"))
        sb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview, style="Slim.Vertical.TScrollbar")
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.tag_configure("sold", foreground=FAINT)
        self.tree.bind("<<TreeviewSelect>>", self._sel)
        self.tree.bind("<Double-1>", lambda e: self.on_open(self._sel_id()) if self._sel_id() else None)
        for seq in ("<Button-3>",) + (("<Button-2>", "<Control-Button-1>") if IS_MAC else ()):
            self.tree.bind(seq, self._menu)
        self.cards = []
        self.sort_key = None
        self.sort_rev = False
        self._suppress = False

    def _sel_id(self):
        s = self.tree.selection()
        return int(s[0]) if s else None

    def _sel(self, _e):
        if not self._suppress and self._sel_id() is not None:
            self.on_select(self._sel_id())

    def _menu(self, e):
        iid = self.tree.identify_row(e.y)
        if iid:
            self.tree.selection_set(iid)
            self.app.card_menu(int(iid), e)

    def _sort_by(self, key):
        self.sort_rev = not self.sort_rev if self.sort_key == key else False
        self.sort_key = key
        self.set_items(self.cards)

    def _value(self, c, key):
        meta = self.app.meta
        return {
            "name": c["name"].lower(), "set": c["set"].lower(), "number": c["number"],
            "lang": c["language"], "qty": c["quantity"], "price": c["price"],
            "market": market_price(meta, c["id"]) or -1,
            "status": (card_state(c)[0], grade_label(c)),
        }[key]

    def set_items(self, cards, selected=None):
        self.cards = cards
        items = list(cards)
        if self.sort_key:
            items.sort(key=lambda c: self._value(c, self.sort_key), reverse=self.sort_rev)
        self._suppress = True
        self.tree.delete(*self.tree.get_children())
        for c in items:
            iid = str(c["id"])
            trend = market_price(self.app.meta, c["id"]) if c["grading_status"] != "Graded" else None
            status = card_state(c)[0]
            if c["grading_status"] == "Submitted":
                status += f" ({c['grading_service']})"
            extra = [x for x in (grade_label(c), "vorgemerkt" if c["vormerkung"] else "") if x]
            if extra:
                status = " · ".join(extra) if status == "In Sammlung" else " · ".join([status] + extra)
            path, kind = self.app.image_for(c["id"])
            sold = c["status"] == "Sold"
            img = self.app.thumbs.get(path, 36, 50, kind or "card", sold, 4,
                                      on_ready=lambda im, i=iid: self.tree.exists(i) and self.tree.item(i, image=im))
            self.tree.insert("", "end", iid=iid, image=img or "", values=(
                c["name"], c["set"], c["number"], c["language"],
                c["quantity"] if not sold else "–", fmt_eur(c["price"]),
                fmt_eur(trend) if trend is not None else "", status,
            ), tags=("sold",) if sold else ())
        if selected and self.tree.exists(str(selected)):
            self.tree.selection_set(str(selected))
            self.tree.see(str(selected))
        self._suppress = False

    def select(self, cid):
        if cid is not None and self.tree.exists(str(cid)):
            self._suppress = True
            self.tree.selection_set(str(cid))
            self.tree.see(str(cid))
            self._suppress = False

    def redraw_card(self, cid):
        if not self.tree.exists(str(cid)):
            return
        path, kind = self.app.image_for(cid)
        c = self.app.entry(cid)
        if not c:
            return
        img = self.app.thumbs.get(path, 36, 50, kind or "card", c["status"] == "Sold", 4,
                                  on_ready=lambda im, i=str(cid): self.tree.exists(i) and self.tree.item(i, image=im))
        if img:
            self.tree.item(str(cid), image=img)


class DetailPanel(tk.Frame):
    WIDTH = 372

    def __init__(self, parent, app):
        super().__init__(parent, bg=SURFACE, width=self.WIDTH)
        self.pack_propagate(False)
        self.app = app
        self.card_id = None
        self.scroll = ScrollFrame(self, bg=SURFACE)
        self.scroll.pack(fill="both", expand=True)
        self.inner = self.scroll.inner
        self._img_label = None
        self.show(None)

    def _clear(self):
        for w in self.inner.winfo_children():
            w.destroy()

    def show(self, card_id):
        self.card_id = card_id
        self._clear()
        self.scroll.to_top()
        c = self.app.entry_by_id.get(card_id) if card_id is not None else None
        if not c:
            box = tk.Frame(self.inner, bg=SURFACE)
            box.pack(fill="x", padx=28, pady=120)
            tk.Label(box, text="Keine Karte ausgewählt", bg=SURFACE, fg=TEXT, font=F["md_b"]).pack()
            tk.Label(box, text="Karte anklicken für Details.\nDoppelklick zeigt die Bilder groß,\n"
                     "Rechtsklick öffnet alle Aktionen.\n\nPfeiltasten wechseln die Karte.",
                     bg=SURFACE, fg=MUTED, font=F["sm"], justify="center").pack(pady=(8, 0))
            return
        self._build(c)

    def refresh_image(self):
        if self.card_id is not None:
            self.show(self.card_id)

    def _build(self, c):
        p = tk.Frame(self.inner, bg=SURFACE)
        p.pack(fill="both", expand=True, padx=22, pady=20)
        cid = c["id"]
        sold = c["status"] == "Sold"

        # Bild
        iw, ih = 236, 329
        holder = tk.Frame(p, bg=SURFACE, width=iw, height=ih)
        holder.pack()
        holder.pack_propagate(False)
        lbl = tk.Label(holder, bg=SURFACE, cursor="hand2")
        lbl.pack(fill="both", expand=True)
        lbl.bind("<Button-1>", lambda e: self.app.view_images(cid))
        path, kind = self.app.image_for(cid)
        if path:
            img = self.app.thumbs.get(path, iw, ih, kind or "card", False, 12,
                                      on_ready=lambda im, l=lbl: l.winfo_exists() and l.config(image=im))
            if img:
                lbl.config(image=img)
        else:
            lbl.config(text="Noch kein Foto.\nKlick: Foto hinzufügen\noder per Handy hochladen",
                       fg=MUTED, font=F["sm"], bg=SURFACE2)
            lbl.bind("<Button-1>", lambda e: self.app.add_photos(cid))

        links = tk.Frame(p, bg=SURFACE)
        links.pack(pady=(10, 0))
        Btn(links, "Cardmarket ↗", lambda: self.app.open_cardmarket(cid), kind="ghost", small=True,
            bg=SURFACE).pack(side="left")
        Btn(links, "+ Foto", lambda: self.app.add_photos(cid), kind="ghost", small=True,
            bg=SURFACE).pack(side="left", padx=(4, 0))

        photos = entry_photo_paths(cid)
        lots = c.get("lots")
        if photos:
            strip = tk.Frame(p, bg=SURFACE)
            strip.pack(fill="x", pady=(10, 0))
            tk.Label(strip, text=f"EIGENE FOTOS · {len(photos)}", bg=SURFACE, fg=FAINT,
                     font=F["xs_b"]).pack(anchor="w", pady=(0, 6))
            row = tk.Frame(strip, bg=SURFACE)
            row.pack(anchor="w")
            for i, ph in enumerate(photos[:5]):
                t = tk.Label(row, bg=SURFACE2, image=self.app.thumbs.blank(52, 52), cursor="hand2")
                t.pack(side="left", padx=(0, 6))
                im = self.app.thumbs.get(ph, 52, 52, "photo", False, 6,
                                         on_ready=lambda im, l=t: l.winfo_exists() and l.config(image=im))
                if im:
                    t.config(image=im)
                t.bind("<Button-1>", lambda e, i=i: self.app.view_images(cid, i))
            if len(photos) > 5:
                tk.Label(row, text=f"+{len(photos) - 5}", bg=SURFACE, fg=MUTED, font=F["sm_b"]).pack(side="left")

        # Titel
        tk.Label(p, text=c["name"] or "Ohne Namen", bg=SURFACE, fg=TEXT, font=F["lg_b"], anchor="w",
                 justify="left", wraplength=self.WIDTH - 50).pack(fill="x", pady=(16, 2))
        ref = f"{len(lots)} Käufe" if lots else f"#{cid:03d}"
        sub = " · ".join(x for x in (ref, c["set"], c["number"], c["language"]) if x)
        tk.Label(p, text=sub, bg=SURFACE, fg=MUTED, font=F["sm"], anchor="w", justify="left",
                 wraplength=self.WIDTH - 50).pack(fill="x")

        chips = tk.Frame(p, bg=SURFACE)
        chips.pack(fill="x", pady=(10, 0))
        label, fg, bg = card_state(c)
        items = [(label, fg, bg)]
        g = grade_label(c)
        if g:
            items.append((g, ON_ACCENT, ACCENT))
        if c["vormerkung"]:
            items.append(("Vorgemerkt", VIOLET, VIOLET_SOFT))
        if not g:
            items.append((c["condition"], MUTED, SURFACE2))
        for text, f, b in items:
            tk.Label(chips, text=text, bg=b, fg=f, font=F["xs_b"], padx=8, pady=3).pack(side="left", padx=(0, 6))

        # Preise
        box = tk.Frame(p, bg=SURFACE2)
        box.pack(fill="x", pady=(16, 0))
        qty = c["quantity"]
        raw_trend = market_price(self.app.meta, cid)
        trend = raw_trend if c["grading_status"] != "Graded" else None
        cells = [("Ø Kaufpreis / Stk" if lots else "Kaufpreis / Stk", fmt_eur(c["price"]), TEXT),
                 ("Anzahl", "verkauft" if sold else str(qty), TEXT)]
        if not sold:
            cells.append(("Einkauf gesamt", fmt_eur(c["price"] * qty), TEXT))
            if trend is not None:
                diff = (trend - c["price"]) * qty
                cells.append(("Marktwert gesamt", fmt_eur(trend * qty), GREEN if diff >= 0 else RED))
            else:
                cells.append(("Marktwert", "–", MUTED))
        for i, (k, v, col) in enumerate(cells):
            cell = tk.Frame(box, bg=SURFACE2)
            cell.grid(row=i // 2, column=i % 2, sticky="nw", padx=14, pady=(10 if i < 2 else 2, 10))
            tk.Label(cell, text=k, bg=SURFACE2, fg=MUTED, font=F["xs"]).pack(anchor="w")
            tk.Label(cell, text=v, bg=SURFACE2, fg=col, font=F["md_b"]).pack(anchor="w")
        box.columnconfigure(0, weight=1)
        box.columnconfigure(1, weight=1)
        pi = self.app.price_info(cid)
        if trend is not None and not sold:
            diff = trend - c["price"]
            pct = f" ({diff / c['price'] * 100:+.0f} %)" if c["price"] else ""
            note = f"Cardmarket-Trend {fmt_eur(trend)} pro Stück, {fmt_eur(diff, True)}{pct}"
            if pi.get("avg30"):
                note += f"\n30-Tage-Schnitt {fmt_eur(pi['avg30'])}"
            if pi.get("updated"):
                note += f" · Stand {datetime.fromtimestamp(pi['updated']).strftime('%d.%m. %H:%M')}"
            tk.Label(p, text=note, bg=SURFACE, fg=MUTED, font=F["xs"], justify="left", anchor="w",
                     wraplength=self.WIDTH - 50).pack(fill="x", pady=(6, 0))

        if raw_trend is not None and trend is None and not sold:
            tk.Label(p, text=f"Rohkarte (ungegradet) laut Cardmarket-Trend: {fmt_eur(raw_trend)}. "
                     "Für Slabs gibt es keinen automatischen Preis.", bg=SURFACE, fg=MUTED, font=F["xs"],
                     justify="left", anchor="w", wraplength=self.WIDTH - 50).pack(fill="x", pady=(6, 0))

        # Details
        rows = [("Zustand", c["condition"])]
        if not lots:
            rows += [("Quelle", c["source"] or "–"), ("Gekauft am", fmt_date(c["purchase_date"]))]
        if c["grading_status"] == "Submitted":
            rows.append(("Grading", f"eingereicht bei {c['grading_service']}"))
        if c["cert"]:
            rows.append(("Cert-Nr.", c["cert"]))
        if c["prev_service"]:
            rows.append(("Vorher", f"{c['prev_service']} {fmt_grade(c['prev_grade'])}".strip()))
        if c["vormerkung_notiz"]:
            rows.append(("Notiz", c["vormerkung_notiz"]))
        det = tk.Frame(p, bg=SURFACE)
        det.pack(fill="x", pady=(14, 0))
        for i, (k, v) in enumerate(rows):
            tk.Label(det, text=k, bg=SURFACE, fg=MUTED, font=F["sm"], anchor="w").grid(
                row=i, column=0, sticky="nw", pady=2)
            tk.Label(det, text=v, bg=SURFACE, fg=TEXT, font=F["sm"], anchor="w", justify="left",
                     wraplength=200).grid(row=i, column=1, sticky="nw", padx=(14, 0), pady=2)

        if lots:
            self._lots(p, lots)

        # Aktionen
        act = tk.Frame(p, bg=SURFACE)
        act.pack(fill="x", pady=(20, 0))

        def add(text, cmd, kind="subtle"):
            b = Btn(act, text, cmd, kind=kind, bg=SURFACE)
            b.pack(fill="x", pady=(0, 6))
            return b

        if not sold:
            add("Verkauf eintragen", lambda: self.app.sell(cid), "primary")
            add("Auf eBay listen", lambda: self.app.ebay(cid))
            if c["grading_status"] == "Submitted":
                add("Grading zurück erhalten", lambda: self.app.grading_back(cid))
                add("Einreichung zurücknehmen", lambda: self.app.cancel_grading(cid))
            else:
                add("Vormerkung entfernen" if c["vormerkung"] else "Für Grading vormerken",
                    lambda: self.app.toggle_queue(cid))
                add("Zum Grading einreichen", lambda: self.app.submit_grading([cid]))
            if c["status"] == "For Sale":
                add("Nicht mehr zum Verkauf", lambda: self.app.set_for_sale(cid, False))
            else:
                add("Als „Zum Verkauf“ markieren", lambda: self.app.set_for_sale(cid, True))
        else:
            add("Verkäufe dieser Karte", lambda: self.app.show("sales", search=f"#{cid:03d}"))
        two = tk.Frame(act, bg=SURFACE)
        two.pack(fill="x", pady=(6, 0))
        if not lots:
            Btn(two, "Bearbeiten", lambda: self.app.edit(cid), kind="subtle", bg=SURFACE).pack(
                side="left", fill="x", expand=True, padx=(0, 3))
        Btn(two, "Alle Käufe löschen" if lots else "Löschen", lambda: self.app.delete(cid),
            kind="danger", bg=SURFACE).pack(side="left", fill="x", expand=True, padx=(0 if lots else 3, 0))

    def _lots(self, p, lots):
        """Liste der einzelnen Käufe eines Mehrfach-Eintrags, älteste zuerst (wird zuerst verkauft)."""
        tk.Label(p, text=f"KÄUFE · {len(lots)}", bg=SURFACE, fg=FAINT, font=F["xs_b"]).pack(
            anchor="w", pady=(18, 6))
        box = tk.Frame(p, bg=SURFACE2)
        box.pack(fill="x")
        for i, lot in enumerate(lots):
            row = tk.Frame(box, bg=SURFACE2)
            row.pack(fill="x", padx=12, pady=(10 if i == 0 else 4, 10 if i == len(lots) - 1 else 4))
            qty = f"{lot['quantity']}× " if lot["quantity"] != 1 else ""
            tk.Label(row, text=f"{qty}{fmt_eur(lot['price'])}", bg=SURFACE2, fg=TEXT,
                     font=F["body_b"]).pack(side="left")
            info = " · ".join(x for x in (fmt_date(lot["purchase_date"]), lot["source"], f"#{lot['id']:03d}") if x)
            tk.Label(row, text=info, bg=SURFACE2, fg=MUTED, font=F["xs"]).pack(side="left", padx=(10, 0))
            Btn(row, "✕", lambda i=lot["id"]: self.app.delete_lot(i), kind="ghost", small=True,
                bg=SURFACE2).pack(side="right")
            Btn(row, "Bearbeiten", lambda l=lot: EditDialog(self.app, l), kind="ghost", small=True,
                bg=SURFACE2).pack(side="right")
        tk.Label(p, text="Beim Verkauf geht der oberste (älteste) Kauf zuerst raus.", bg=SURFACE, fg=FAINT,
                 font=F["xs"], anchor="w").pack(fill="x", pady=(6, 0))


class CollectionView(View):
    FILTERS = [("stock", "Bestand"), ("for_sale", "Zum Verkauf"), ("grading", "Beim Grading"),
               ("queued", "Vorgemerkt"), ("graded", "Gegradet"), ("sold", "Verkauft"), ("all", "Alle")]
    SORTS = [("Neueste zuerst", "new"), ("Älteste zuerst", "old"), ("Name A–Z", "name"),
             ("Set", "set"), ("Kaufpreis ↓", "price"), ("Marktwert ↓", "market"),
             ("Wertentwicklung ↓", "delta")]
    OLD_FILTERS = {"All": "all", "Collection": "stock", "At Grading": "grading",
                   "Grading-Queue": "queued", "Sold": "sold"}

    def __init__(self, parent, app):
        super().__init__(parent, app)
        cfg = app.config
        self.summary, right = self.header(self, "Sammlung")
        Btn(right, "+ Karte hinzufügen", lambda: app.show("add"), kind="primary", bg=BG).pack()

        bar = tk.Frame(self, bg=BG)
        bar.pack(fill="x", padx=28, pady=(18, 0))
        self.search = Field(bar, width=36, placeholder="Suchen: Name, Set, Nummer, #ID")
        self.search.pack(side="left")
        self.search.var.trace_add("write", lambda *_: self._search_changed())
        self.search.entry.bind("<Escape>", lambda e: (self.search.set(""), self.gallery.canvas.focus_set()))
        self.search.entry.bind("<Down>", lambda e: self._key_move(0, 0))

        self.view_mode = tk.StringVar(value=cfg.get("view_mode", "gallery"))
        self.mode_chips = Chips(bar, ["gallery", "list"], self.view_mode.get(), self._set_mode,
                                labels={"gallery": "Galerie", "list": "Liste"}, small=True)
        self.mode_chips.pack(side="right", pady=(6, 0))
        sort_names = [s[0] for s in self.SORTS]
        cur_sort = cfg.get("sort", "new")
        self.sort_cb = ttk.Combobox(bar, values=sort_names, state="readonly", width=17, font=F["sm"])
        self.sort_cb.set(next((n for n, k in self.SORTS if k == cur_sort), sort_names[0]))
        self.sort_cb.bind("<<ComboboxSelected>>", lambda e: self.refresh())
        self.sort_cb.pack(side="right", padx=(0, 14))
        tk.Label(bar, text="Sortierung", bg=BG, fg=FAINT, font=F["sm"]).pack(side="right", padx=(0, 6))

        fbar = tk.Frame(self, bg=BG)
        fbar.pack(fill="x", padx=28, pady=(12, 10))
        flt = cfg.get("filter_v2") or self.OLD_FILTERS.get(cfg.get("filter"), "stock")
        self.filter_chips = Chips(fbar, [k for k, _ in self.FILTERS], flt, lambda v: self.refresh(),
                                  labels=dict(self.FILTERS), small=True)
        self.filter_chips.pack(side="left", fill="x", expand=True)

        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=(22, 0), pady=(0, 0))
        self.detail = DetailPanel(body, app)
        self.detail.pack(side="right", fill="y")
        self.stage = tk.Frame(body, bg=BG)
        self.stage.pack(side="left", fill="both", expand=True, padx=(0, 16))
        self.gallery = Gallery(self.stage, app, self._on_select, self._on_open)
        self.list = CardList(self.stage, app, self._on_select, self._on_open)
        (self.gallery if self.view_mode.get() == "gallery" else self.list).pack(fill="both", expand=True)

        self.selected = None
        self.items = []
        self._search_job = None
        for key, (dx, dy) in {"Left": (-1, 0), "Right": (1, 0), "Up": (0, -1), "Down": (0, 1)}.items():
            self.gallery.canvas.bind(f"<{key}>", lambda e, d=(dx, dy): self._key_move(*d))
        self.gallery.canvas.bind("<space>", lambda e: self.selected and self.app.view_images(self.selected))
        self.gallery.canvas.bind("<Return>", lambda e: self.selected and self.app.edit(self.selected))
        self.gallery.canvas.bind("<BackSpace>", lambda e: self.selected and self.app.delete(self.selected))

    def _search_changed(self):
        if self._search_job:
            self.after_cancel(self._search_job)
        self._search_job = self.after(140, self.refresh)

    def _set_mode(self, mode):
        self.view_mode.set(mode)
        self.gallery.pack_forget()
        self.list.pack_forget()
        (self.gallery if mode == "gallery" else self.list).pack(fill="both", expand=True)
        self.refresh()

    def _filtered(self):
        f = self.filter_chips.get()
        terms = [t for t in self.search.get().lower().split() if t]
        out = []
        for c in self.app.entries:
            sold = c["status"] == "Sold"
            if f == "stock" and sold:
                continue
            if f == "for_sale" and c["status"] != "For Sale":
                continue
            if f == "grading" and (sold or c["grading_status"] != "Submitted"):
                continue
            if f == "queued" and (sold or not c["vormerkung"]):
                continue
            if f == "graded" and (sold or c["grading_status"] != "Graded"):
                continue
            if f == "sold" and not sold:
                continue
            if terms and not card_matches(c, terms):
                continue
            out.append(c)
        global upload_sort
        upload_sort = dict(self.SORTS).get(self.sort_cb.get(), "new")
        return sort_cards(out, upload_sort, self.app.meta)

    def _update_filter_counts(self):
        counts = {k: 0 for k, _ in self.FILTERS}
        for c in self.app.entries:
            sold = c["status"] == "Sold"
            counts["all"] += 1
            if sold:
                counts["sold"] += 1
                continue
            counts["stock"] += 1
            counts["for_sale"] += c["status"] == "For Sale"
            counts["grading"] += c["grading_status"] == "Submitted"
            counts["queued"] += bool(c["vormerkung"])
            counts["graded"] += c["grading_status"] == "Graded"
        self.filter_chips.relabel({k: f"{label}  {counts[k]}" for k, label in self.FILTERS})

    def refresh(self):
        self.items = self._filtered()
        self._update_filter_counts()
        units = sum(c["quantity"] for c in self.items if c["status"] != "Sold")
        cost = sum(c["price"] * c["quantity"] for c in self.items if c["status"] != "Sold")
        market = [(market_price(self.app.meta, c["id"]), c) for c in self.items
                  if c["status"] != "Sold" and c["grading_status"] != "Graded"]
        mval = sum(t * c["quantity"] for t, c in market if t is not None)
        parts = [f"{len(self.items)} Einträge", f"{units} Stück", f"Einkauf {fmt_eur(cost)}"]
        if mval:
            parts.append(f"Marktwert {fmt_eur(mval)}")
        self.summary.config(text="  ·  ".join(parts))
        if self.view_mode.get() == "gallery":
            empty = "Keine Karten in dieser Ansicht." if self.app.cards else \
                "Noch keine Karten.\nÜber „+ Karte hinzufügen“ oder " + MOD_LABEL + "N geht es los."
            if self.search.get():
                empty = f"Nichts gefunden für „{self.search.get()}“."
            self.gallery.selected = self.selected
            self.gallery.set_items(self.items, empty)
        else:
            self.list.set_items(self.items, self.selected)
        if self.selected is not None and self.selected not in self.app.entry_by_id:
            self.selected = None
        self.detail.show(self.selected)

    def soft_refresh(self):
        y = self.gallery.canvas.yview()[0]
        self.refresh()
        self.gallery.canvas.yview_moveto(y)

    def image_changed(self, card_id):
        if self.view_mode.get() == "gallery":
            self.gallery.redraw_card(card_id)
        else:
            self.list.redraw_card(card_id)
        if card_id == self.selected:
            self.detail.show(card_id)

    def _on_select(self, cid):
        self.selected = cid
        self.detail.show(cid)

    def _on_open(self, cid):
        self.app.view_images(cid)

    def select(self, cid):
        self.selected = cid
        self.gallery.select(cid)
        self.list.select(cid)
        self.detail.show(cid)

    def reveal(self, cid):
        """Karte anzeigen, notfalls Filter/Suche zuruecksetzen."""
        cid = self.app.lead_of.get(cid, cid)
        if cid not in {c["id"] for c in self.items}:
            self.search.set("")
            self.filter_chips.set("all")
            self.refresh()
        self.select(cid)

    def _key_move(self, dx, dy):
        if self.view_mode.get() != "gallery":
            return
        self.gallery.canvas.focus_set()
        cid = self.gallery.move(dx, dy)
        if cid is not None:
            self._on_select(cid)
        return "break"

    def on_show(self, **kw):
        if self.view_mode.get() == "gallery":
            self.gallery.canvas.focus_set()

    def save_state(self, cfg):
        cfg["filter_v2"] = self.filter_chips.get()
        cfg["sort"] = dict(self.SORTS).get(self.sort_cb.get(), "new")
        cfg["view_mode"] = self.view_mode.get()


# ============================================================
# GUI: HINZUFUEGEN
# ============================================================

class AddView(View):
    PREV_W, PREV_H = 250, 349

    def __init__(self, parent, app):
        super().__init__(parent, app)
        cfg = app.config
        self.header(self, "Karte hinzufügen",
                    "Cardmarket-Link einfügen: Name, Set und Nummer werden erkannt, das Bild wird automatisch gesucht.")
        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=28, pady=(18, 0))

        right = tk.Frame(body, bg=BG, width=self.PREV_W + 40)
        right.pack(side="right", fill="y", padx=(24, 0))
        right.pack_propagate(False)
        scroll = ScrollFrame(body, bg=BG)
        scroll.pack(side="left", fill="both", expand=True)
        form = tk.Frame(scroll.inner, bg=BG)
        form.pack(fill="x", anchor="nw")
        form.columnconfigure(0, weight=3)
        form.columnconfigure(1, weight=2)
        form.columnconfigure(2, weight=1)

        def section(text, row, col=0, span=3, top=16):
            tk.Label(form, text=text, bg=BG, fg=MUTED, font=F["sm_b"], anchor="w").grid(
                row=row, column=col, columnspan=span, sticky="w", pady=(top, 6))

        section("CARDMARKET-LINK", 0, top=0)
        link_row = tk.Frame(form, bg=BG)
        link_row.grid(row=1, column=0, columnspan=3, sticky="ew")
        self.url = Field(link_row, width=50, placeholder="https://www.cardmarket.com/de/Pokemon/Products/Singles/…")
        self.url.pack(side="left", fill="x", expand=True)
        Btn(link_row, "Einfügen", self._paste, kind="secondary", bg=BG).pack(side="left", padx=(8, 0))
        self.url.var.trace_add("write", lambda *_: self._debounce("url", 250, self._parse_url))
        self.url.entry.bind("<Return>", lambda e: (self._parse_url(), self.price.focus(), "break")[-1])

        section("NAME", 2, 0, 1)
        section("SET", 2, 1, 1)
        section("NUMMER", 2, 2, 1)
        self.name = Field(form, width=26)
        self.name.grid(row=3, column=0, sticky="ew", padx=(0, 10))
        self.set_var = tk.StringVar()
        self.set_cb = ttk.Combobox(form, textvariable=self.set_var, values=SETS, font=F["body"])
        self.set_cb.grid(row=3, column=1, sticky="ew", padx=(0, 10), ipady=3)
        self.set_cb.bind("<KeyRelease>", self._filter_sets)
        self.number = Field(form, width=10)
        self.number.grid(row=3, column=2, sticky="ew")
        for var in (self.name.var, self.set_var, self.number.var):
            var.trace_add("write", lambda *_: self._fields_changed())

        section("SPRACHE", 4)
        self.lang = Chips(form, LANGUAGES, cfg.get("add_lang", "EN"), small=True)
        self.lang.grid(row=5, column=0, columnspan=3, sticky="ew")
        section("ZUSTAND", 6)
        self.cond = Chips(form, CONDITIONS, cfg.get("add_cond", "Near Mint"), small=True)
        self.cond.grid(row=7, column=0, columnspan=3, sticky="ew")
        section("QUELLE", 8)
        self.source = Chips(form, SOURCES, cfg.get("add_source", "Cardmarket"), small=True)
        self.source.grid(row=9, column=0, columnspan=3, sticky="ew")

        section("ANZAHL", 10, 0, 1)
        section("PREIS PRO STÜCK", 10, 1, 2)
        qrow = tk.Frame(form, bg=BG)
        qrow.grid(row=11, column=0, sticky="w")
        Btn(qrow, "−", lambda: self._step(-1), kind="subtle", bg=BG).pack(side="left")
        self.qty = Field(qrow, width=4, value="1")
        self.qty.entry.config(justify="center")
        self.qty.pack(side="left", padx=6)
        Btn(qrow, "+", lambda: self._step(1), kind="subtle", bg=BG).pack(side="left")
        prow = tk.Frame(form, bg=BG)
        prow.grid(row=11, column=1, columnspan=2, sticky="ew")
        self.price = Field(prow, width=10, suffix="€")
        self.price.pack(side="left")
        self.total = tk.Label(prow, text="", bg=BG, fg=MUTED, font=F["sm"])
        self.total.pack(side="left", padx=(12, 0))
        for v in (self.qty.var, self.price.var):
            v.trace_add("write", lambda *_: self._update_total())

        section("GRADING", 12)
        self.graded = Chips(form, ["no", "yes"], "no", self._toggle_grading,
                            labels={"no": "Ungegradet", "yes": "Gegradet (Slab)"}, small=True)
        self.graded.grid(row=13, column=0, columnspan=3, sticky="w")
        self.grade_box = tk.Frame(form, bg=BG)
        self.grade_box.grid(row=14, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        self.service = Chips(self.grade_box, GRADING_SERVICES, "PSA", small=True)
        self.service.pack(fill="x")
        gb2 = tk.Frame(self.grade_box, bg=BG)
        gb2.pack(fill="x", pady=(4, 0))
        tk.Label(gb2, text="Note", bg=BG, fg=MUTED, font=F["sm"]).pack(side="left")
        self.grade = Field(gb2, width=6, placeholder="10")
        self.grade.pack(side="left", padx=(8, 18))
        tk.Label(gb2, text="Cert-Nr.", bg=BG, fg=MUTED, font=F["sm"]).pack(side="left")
        self.cert = Field(gb2, width=16)
        self.cert.pack(side="left", padx=(8, 0))
        self.grade_box.grid_remove()

        actions = tk.Frame(form, bg=BG)
        actions.grid(row=15, column=0, columnspan=3, sticky="ew", pady=(26, 0))
        Btn(actions, "Zur Sammlung hinzufügen  ↵", self.save, kind="primary", bg=BG).pack(side="left")
        Btn(actions, "Leeren", self.reset, kind="ghost", bg=BG).pack(side="left", padx=(8, 0))
        self.status = tk.Label(actions, text="", bg=BG, fg=MUTED, font=F["sm"])
        self.status.pack(side="left", padx=(14, 0))

        self.recent_box = tk.Frame(form, bg=BG)
        self.recent_box.grid(row=16, column=0, columnspan=3, sticky="ew", pady=(30, 20))

        for w in (self.name.entry, self.number.entry, self.price.entry, self.qty.entry,
                  self.grade.entry, self.cert.entry, self.set_cb):
            w.bind("<Return>", lambda e: (self.save(), "break")[-1])

        # Vorschau
        tk.Label(right, text="VORSCHAU", bg=BG, fg=MUTED, font=F["sm_b"]).pack(anchor="w")
        self.prev_img = tk.Label(right, bg=SURFACE, image=app.thumbs.blank(self.PREV_W, self.PREV_H))
        self.prev_img.pack(anchor="w", pady=(6, 0))
        self.prev_name = tk.Label(right, text="", bg=BG, fg=TEXT, font=F["md_b"], anchor="w",
                                  justify="left", wraplength=self.PREV_W)
        self.prev_name.pack(anchor="w", pady=(12, 0), fill="x")
        self.prev_sub = tk.Label(right, text="", bg=BG, fg=MUTED, font=F["sm"], anchor="w",
                                 justify="left", wraplength=self.PREV_W)
        self.prev_sub.pack(anchor="w", fill="x")
        self.prev_price = tk.Label(right, text="", bg=BG, fg=TEXT, font=F["sm_b"], anchor="w")
        self.prev_price.pack(anchor="w", pady=(6, 0))
        self.prev_state = tk.Label(right, text="", bg=BG, fg=FAINT, font=F["sm"], anchor="w",
                                   justify="left", wraplength=self.PREV_W)
        self.prev_state.pack(anchor="w", pady=(6, 0), fill="x")
        self.pick_btn = Btn(right, "Anderes Bild wählen", self._pick, kind="subtle", bg=BG)
        self.pick_btn.pack(anchor="w", pady=(10, 0))

        self._jobs = {}
        self._filling = False
        self._lookup_key = None
        self._seq = 0
        self.match = None
        self.candidates = []
        self._preview_image = None
        self._show_preview()

    # ----- Eingabe -----

    def _debounce(self, name, ms, fn):
        if name in self._jobs:
            self.after_cancel(self._jobs[name])
        self._jobs[name] = self.after(ms, fn)

    def _paste(self):
        try:
            text = self.app.root.clipboard_get()
        except tk.TclError:
            text = ""
        self.url.set(text.strip())
        self._parse_url()

    def _parse_url(self):
        url = self.url.get()
        if not url:
            return
        res = parse_cardmarket_url(url)
        if not res:
            self.status.config(text="Link nicht erkannt. Felder einfach von Hand ausfüllen.", fg=AMBER)
            return
        self._filling = True
        self.name.set(res["name"])
        self.set_var.set(res["set"])
        self.number.set(res["number"])
        self._filling = False
        self.status.config(text="Link erkannt", fg=GREEN)
        self._lookup()

    def _fields_changed(self):
        if not self._filling:
            self._debounce("lookup", 700, self._lookup)

    def _filter_sets(self, event):
        if event.keysym in ("Up", "Down", "Return", "Escape", "Tab"):
            return
        q = self.set_var.get().lower()
        self.set_cb.configure(values=[s for s in SETS if q in s.lower()] or SETS)

    def _step(self, d):
        try:
            v = int(self.qty.get() or "1")
        except ValueError:
            v = 1
        self.qty.set(max(1, v + d))

    def _update_total(self):
        try:
            q = int(self.qty.get() or "1")
            p = parse_eur(self.price.get())
            self.total.config(text=f"= {fmt_eur(p * q)} gesamt" if q > 1 else "")
        except ValueError:
            self.total.config(text="")

    def _toggle_grading(self, v):
        if v == "yes":
            self.grade_box.grid()
        else:
            self.grade_box.grid_remove()

    # ----- Bildsuche -----

    def _lookup(self):
        name = self.name.get()
        key = (name, self.set_var.get().strip(), self.number.get())
        if not name:
            self.match, self.candidates, self._lookup_key = None, [], None
            self._show_preview()
            return
        if key == self._lookup_key:
            return
        self._lookup_key = key
        self._seq += 1
        seq = self._seq
        self.match, self.candidates = None, []
        self._show_preview(state="Suche Bild …")
        self.app.worker.submit(f"addlookup:{seq}", resolve_card_image, key,
                               lambda res, err: self._on_lookup(seq, res, err))

    def _on_lookup(self, seq, res, err):
        if seq != self._seq:
            return
        if err or not res:
            self._show_preview(state="Bildsuche fehlgeschlagen (offline?). Speichern geht trotzdem.")
            return
        if res["status"] == "ok":
            self._set_match(res["match"])
        else:
            self.candidates = res.get("candidates", [])
            n = len(self.candidates)
            self._show_preview(state=f"{n} mögliche Karten gefunden. Bitte das richtige Bild wählen."
                               if n else "Kein Bild gefunden. Du kannst eines suchen oder ohne speichern.")

    def _set_match(self, match):
        self.match = match
        seq = self._seq
        self._show_preview(state="Bild gefunden")

        def loaded(path, err):
            if seq == self._seq and path and not err:
                self._show_preview(path=path, state="Bild gefunden")
        self.app.worker.submit(f"addprev:{match['id']}:{seq}", download_preview, (match, "high"), loaded)
        if match.get("kind") == "card":
            def priced(res, err):
                if seq == self._seq and res and res.get("trend"):
                    self.prev_price.config(text=f"Cardmarket-Trend: {fmt_eur(res['trend'])}")
            self.app.worker.submit(f"addprice:{match['id']}:{seq}", fetch_card_price, (match["id"],), priced)

    def _show_preview(self, path=None, state=""):
        name = self.name.get()
        self.prev_name.config(text=(self.match or {}).get("name") or name or "Noch keine Karte")
        if self.match:
            sub = self.match.get("set_name", "")
            if self.match.get("localId"):
                sub += f" · {self.match['localId']}"
        else:
            sub = " · ".join(x for x in (self.set_var.get().strip(), self.number.get()) if x)
        self.prev_sub.config(text=sub)
        if not path:
            self.prev_price.config(text="")
        self.prev_state.config(text=state, fg=GREEN if state == "Bild gefunden" else FAINT)
        if path:
            kind = (self.match or {}).get("kind", "card")
            img = self.app.thumbs.get(path, self.PREV_W, self.PREV_H, kind, False, 12,
                                      on_ready=lambda im: self.prev_img.config(image=im, width=self.PREV_W, height=self.PREV_H))
            if img:
                self.prev_img.config(image=img, width=self.PREV_W, height=self.PREV_H)
        else:
            if self._preview_image is None:
                ph = Image.new("RGBA", (self.PREV_W, self.PREV_H), SURFACE)
                d = ImageDraw.Draw(ph)
                d.rounded_rectangle((0, 0, self.PREV_W - 1, self.PREV_H - 1), 12, fill=SURFACE2)
                cx, cy = self.PREV_W // 2, self.PREV_H // 2
                d.ellipse((cx - 34, cy - 34, cx + 34, cy + 34), outline=SURFACE4, width=4)
                d.line((cx - 34, cy, cx + 34, cy), fill=SURFACE4, width=4)
                d.ellipse((cx - 11, cy - 11, cx + 11, cy + 11), fill=SURFACE2, outline=SURFACE4, width=4)
                self._preview_image = ImageTk.PhotoImage(ph)
            self.prev_img.config(image=self._preview_image, width=self.PREV_W, height=self.PREV_H)
        self.pick_btn.set_text("Bild wählen" if not self.match else "Anderes Bild wählen")

    def _pick(self):
        name = self.name.get()
        if not name:
            self.status.config(text="Erst einen Namen eingeben.", fg=AMBER)
            return
        ImagePicker(self.app, name, self.set_var.get().strip(), candidates=self.candidates,
                    on_pick=self._picked)

    def _picked(self, match):
        self._seq += 1
        self._set_match(match)

    # ----- Speichern -----

    def save(self):
        name = self.name.get()
        set_name = self.set_var.get().strip()
        if not name:
            self.status.config(text="Bitte einen Namen eingeben oder einen Link einfügen.", fg=RED)
            self.name.focus()
            return
        try:
            price = parse_eur(self.price.get(), 0.0)
        except ValueError:
            self.status.config(text="Preis muss eine Zahl sein, z. B. 12,50.", fg=RED)
            self.price.focus()
            return
        try:
            qty = int(self.qty.get() or "1")
            if qty < 1:
                raise ValueError
        except ValueError:
            self.status.config(text="Anzahl muss eine ganze Zahl ab 1 sein.", fg=RED)
            return
        g_status, service, grade, cert = "Not Graded", "", None, ""
        if self.graded.get() == "yes":
            g_status, service = "Graded", self.service.get()
            if self.grade.get():
                try:
                    grade = parse_grade(self.grade.get())
                except ValueError:
                    self.status.config(text="Note: Zahl wie 9.5 oder P10 / B10.", fg=RED)
                    return
            cert = self.cert.get()
        try:
            card_id = save_card(name, set_name or "Other", self.number.get(), self.lang.get(),
                                self.cond.get(), price, self.source.get(), qty, g_status,
                                service, grade, cert, url=self.url.get())
        except PermissionError:
            messagebox.showerror("Excel gesperrt", "Die Excel-Datei ist gerade geöffnet.\n"
                                 "Bitte in Excel schließen und erneut speichern.")
            return
        except Exception as e:
            messagebox.showerror("Fehler", f"Speichern fehlgeschlagen:\n{e}")
            return
        self.app.reload()
        if self.match:
            self.app.set_card_match(card_id, self.match, manual=False)
        else:
            c = self.app.by_id.get(card_id)
            if c:
                self.app.ensure_images([c], force=True)
        self.app.toast.show(f"#{card_id:03d} {name} gespeichert")
        self.reset(keep_status=True)
        self.status.config(text=f"#{card_id:03d} gespeichert. Nächste Karte?", fg=GREEN)

    def reset(self, keep_status=False):
        self._filling = True
        for f in (self.url, self.name, self.number, self.price, self.grade, self.cert):
            f.set("")
        self.set_var.set("")
        self.qty.set("1")
        self._filling = False
        self.graded.set("no")
        self._toggle_grading("no")
        self.match, self.candidates, self._lookup_key = None, [], None
        self._seq += 1
        self._show_preview()
        if not keep_status:
            self.status.config(text="")
        self._render_recent()
        self.url.focus()

    def _render_recent(self):
        for w in self.recent_box.winfo_children():
            w.destroy()
        recent = sorted(self.app.cards, key=lambda c: -c["id"])[:6]
        if not recent:
            return
        tk.Label(self.recent_box, text="ZULETZT HINZUGEFÜGT", bg=BG, fg=MUTED, font=F["sm_b"]).pack(anchor="w")
        row = tk.Frame(self.recent_box, bg=BG)
        row.pack(anchor="w", pady=(8, 0))
        for c in recent:
            cell = tk.Frame(row, bg=BG, cursor="hand2")
            cell.pack(side="left", padx=(0, 12))
            path, kind = self.app.image_for(c["id"])
            lbl = tk.Label(cell, bg=SURFACE2, image=self.app.thumbs.blank(72, 100))
            lbl.pack()
            img = self.app.thumbs.get(path, 72, 100, kind or "card", False, 6,
                                      on_ready=lambda im, l=lbl: l.winfo_exists() and l.config(image=im))
            if img:
                lbl.config(image=img)
            tk.Label(cell, text=truncate(c["name"], F["xs"], 76), bg=BG, fg=MUTED, font=F["xs"]).pack(anchor="w")
            for w in (cell, lbl):
                w.bind("<Button-1>", lambda e, cid=c["id"]: self.app.show_card(cid))

    def refresh(self):
        self._render_recent()

    def image_changed(self, card_id):
        self._render_recent()

    def on_show(self, **kw):
        # Cardmarket-Link aus der Zwischenablage direkt uebernehmen.
        if not self.url.get():
            try:
                clip = self.app.root.clipboard_get().strip()
            except tk.TclError:
                clip = ""
            if "cardmarket.com" in clip and "/Products/" in clip and clip != self.app.config.get("_last_clip"):
                self.app.config["_last_clip"] = clip
                self.url.set(clip)
                self._parse_url()
                self.price.focus()
                self.status.config(text="Link aus der Zwischenablage übernommen", fg=GREEN)
                return
        self.url.focus()

    def save_state(self, cfg):
        cfg["add_lang"] = self.lang.get()
        cfg["add_cond"] = self.cond.get()
        cfg["add_source"] = self.source.get()
        cfg.pop("_last_clip", None)


# ============================================================
# GUI: GRADING
# ============================================================

class GradingView(View):
    def __init__(self, parent, app):
        super().__init__(parent, app)
        self.header(self, "Grading", "Vormerken, einreichen, Ergebnis eintragen. Klick auf eine Karte öffnet sie.")
        board = tk.Frame(self, bg=BG)
        board.pack(fill="both", expand=True, padx=28, pady=(18, 22))
        self.cols = {}
        spec = [("queued", "Vorgemerkt", VIOLET), ("submitted", "Eingereicht", AMBER),
                ("graded", "Gegradet", ACCENT)]
        for i, (key, title, color) in enumerate(spec):
            board.columnconfigure(i, weight=1, uniform="g")
            col = tk.Frame(board, bg=SURFACE)
            col.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 8, 0))
            head = tk.Frame(col, bg=SURFACE)
            head.pack(fill="x", padx=16, pady=(14, 10))
            tk.Frame(head, bg=color, width=8, height=8).pack(side="left", pady=2)
            tk.Label(head, text=title, bg=SURFACE, fg=TEXT, font=F["md_b"]).pack(side="left", padx=(8, 0))
            count = tk.Label(head, text="", bg=SURFACE, fg=MUTED, font=F["sm_b"])
            count.pack(side="left", padx=(8, 0))
            action = tk.Frame(head, bg=SURFACE)
            action.pack(side="right")
            sf = ScrollFrame(col, bg=SURFACE)
            sf.pack(fill="both", expand=True, padx=(10, 4), pady=(0, 10))
            self.cols[key] = (sf, count, action)
        board.rowconfigure(0, weight=1)

    def refresh(self):
        cards = [c for c in self.app.cards if c["status"] != "Sold"]
        groups = {
            "queued": [c for c in cards if c["vormerkung"] and c["grading_status"] != "Submitted"],
            "submitted": [c for c in cards if c["grading_status"] == "Submitted"],
            "graded": sorted([c for c in cards if c["grading_status"] == "Graded"],
                             key=lambda c: -c["id"]),
        }
        for key, items in groups.items():
            sf, count, action = self.cols[key]
            count.config(text=str(len(items)))
            for w in action.winfo_children():
                w.destroy()
            for w in sf.inner.winfo_children():
                w.destroy()
            if key == "queued" and items:
                Btn(action, "Alle einreichen", lambda ids=[c["id"] for c in items]: self.app.submit_grading(ids),
                    kind="primary", small=True, bg=SURFACE).pack()
            if not items:
                msg = {"queued": "Nichts vorgemerkt.\nIn der Sammlung: Rechtsklick,\n„Für Grading vormerken“.",
                       "submitted": "Gerade nichts eingereicht.",
                       "graded": "Noch keine gegradeten Karten."}[key]
                tk.Label(sf.inner, text=msg, bg=SURFACE, fg=FAINT, font=F["sm"], justify="center").pack(pady=40)
            for c in items:
                self._row(sf.inner, c, key)

    def _row(self, parent, c, key):
        cid = c["id"]
        row = tk.Frame(parent, bg=SURFACE2, cursor="hand2")
        row.pack(fill="x", pady=(0, 8), padx=(0, 6))
        path, kind = self.app.image_for(cid)
        lbl = tk.Label(row, bg=SURFACE3, image=self.app.thumbs.blank(50, 70))
        lbl.pack(side="left", padx=10, pady=10)
        img = self.app.thumbs.get(path, 50, 70, kind or "card", False, 5,
                                  on_ready=lambda im, l=lbl: l.winfo_exists() and l.config(image=im))
        if img:
            lbl.config(image=img)
        txt = tk.Frame(row, bg=SURFACE2)
        txt.pack(side="left", fill="both", expand=True, pady=10)
        tk.Label(txt, text=truncate(c["name"], F["body_b"], 230), bg=SURFACE2, fg=TEXT, font=F["body_b"],
                 anchor="w").pack(fill="x")
        sub = c["set"] + (f" · {c['number']}" if c["number"] else "")
        tk.Label(txt, text=truncate(sub, F["sm"], 230), bg=SURFACE2, fg=MUTED, font=F["sm"], anchor="w").pack(fill="x")
        if key == "queued":
            typ = "Regrade" if c["grading_status"] == "Graded" else "Erstgrading"
            extra = typ + (f" · {c['vormerkung_notiz']}" if c["vormerkung_notiz"] else "")
            btn = ("Einreichen", lambda: self.app.submit_grading([cid]))
        elif key == "submitted":
            extra = f"bei {c['grading_service']}" + (
                f" · vorher {c['prev_service']} {fmt_grade(c['prev_grade'])}" if c["prev_service"] else "")
            btn = ("Zurück", lambda: self.app.grading_back(cid))
        else:
            extra = grade_label(c) + (f" · Cert {c['cert']}" if c["cert"] else "")
            btn = None
        tk.Label(txt, text=truncate(extra, F["sm"], 230), bg=SURFACE2,
                 fg={"queued": VIOLET, "submitted": AMBER, "graded": ACCENT}[key],
                 font=F["sm_b"], anchor="w").pack(fill="x", pady=(4, 0))
        if btn:
            Btn(row, btn[0], btn[1], kind="secondary", small=True, bg=SURFACE2).pack(side="right", padx=10)
        for w in (row, lbl, txt) + tuple(txt.winfo_children()):
            w.bind("<Button-1>", lambda e: self.app.show_card(cid))
            for seq in ("<Button-3>",) + (("<Button-2>", "<Control-Button-1>") if IS_MAC else ()):
                w.bind(seq, lambda e: self.app.card_menu(cid, e))


# ============================================================
# GUI: VERKAEUFE
# ============================================================

class SalesView(View):
    COLS = [("date", "Datum", 96, "w"), ("platform", "Plattform", 100, "w"), ("qty", "Anz.", 50, "center"),
            ("price", "Verkaufspreis", 110, "e"), ("fees", "Gebühren+Versand", 130, "e"),
            ("cost", "Einkauf", 100, "e"), ("profit", "Gewinn", 110, "e")]

    def __init__(self, parent, app):
        super().__init__(parent, app)
        self.sub, right = self.header(self, "Verkäufe")
        self.kpis = tk.Frame(self, bg=BG)
        self.kpis.pack(fill="x", padx=28, pady=(18, 0))
        bar = tk.Frame(self, bg=BG)
        bar.pack(fill="x", padx=28, pady=(18, 10))
        self.search = Field(bar, width=30, placeholder="Suchen: Karte, Plattform, #ID")
        self.search.pack(side="left")
        self.search.var.trace_add("write", lambda *_: self.refresh())
        self.undo_btn = Btn(bar, "Verkauf rückgängig", self._undo, kind="subtle", bg=BG)
        self.undo_btn.pack(side="right")
        self.open_btn = Btn(bar, "Karte öffnen", self._open, kind="subtle", bg=BG)
        self.open_btn.pack(side="right", padx=(0, 8))
        wrap = tk.Frame(self, bg=BG)
        wrap.pack(fill="both", expand=True, padx=28, pady=(0, 22))
        self.tree = ttk.Treeview(wrap, columns=[c[0] for c in self.COLS], show="tree headings",
                                 selectmode="browse")
        self.tree.heading("#0", text="Karte", anchor="w")
        self.tree.column("#0", width=300, anchor="w")
        for key, label, w, anchor in self.COLS:
            self.tree.heading(key, text=label, anchor=anchor)
            self.tree.column(key, width=w, anchor=anchor, stretch=False)
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview, style="Slim.Vertical.TScrollbar")
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.tag_configure("trade", foreground=MUTED)
        self.tree.bind("<<TreeviewSelect>>", lambda e: self._sel_changed())
        self.tree.bind("<Double-1>", lambda e: self._open())
        self._sel_changed()

    def _kpi(self, parent, label, value, color=TEXT, sub=""):
        box = tk.Frame(parent, bg=SURFACE)
        box.pack(side="left", fill="x", expand=True, padx=(0, 10))
        tk.Label(box, text=label, bg=SURFACE, fg=MUTED, font=F["sm"]).pack(anchor="w", padx=16, pady=(12, 0))
        tk.Label(box, text=value, bg=SURFACE, fg=color, font=F["num_b"]).pack(anchor="w", padx=16)
        tk.Label(box, text=sub, bg=SURFACE, fg=FAINT, font=F["xs"]).pack(anchor="w", padx=16, pady=(0, 12))

    def refresh(self):
        s = compute_statistics(self.app.cards, self.app.sales, self.app.meta)
        for w in self.kpis.winfo_children():
            w.destroy()
        self._kpi(self.kpis, "Verkäufe", str(s["sales_count"]), sub=f"{s['units_sold']} Stück, ohne Tausch")
        self._kpi(self.kpis, "Umsatz", fmt_eur(s["revenue"]))
        self._kpi(self.kpis, "Gebühren & Versand", fmt_eur(s["fees"]))
        self._kpi(self.kpis, "Gewinn", fmt_eur(s["profit"], True), GREEN if s["profit"] >= 0 else RED,
                  sub="Verkaufspreis − Gebühren − Einkauf")
        self.sub.config(text="Alle eingetragenen Verkäufe, neueste oben. Tausch-Abgänge zählen nicht zum Gewinn.")
        terms = [t for t in self.search.get().lower().split() if t]
        self.tree.delete(*self.tree.get_children())
        sales = sorted(self.app.sales, key=lambda x: (str(x["date"] or ""), x["row"]), reverse=True)
        for sale in sales:
            c = self.app.by_id.get(sale["id"])
            hay = f"{sale['name']} {sale['platform']} #{sale['id']:03d} {c['set'] if c else ''}".lower()
            if terms and not all(t in hay for t in terms):
                continue
            path, kind = self.app.image_for(sale["id"])
            iid = f"s{sale['row']}"
            img = self.app.thumbs.get(path, 36, 50, kind or "card", False, 4,
                                      on_ready=lambda im, i=iid: self.tree.exists(i) and self.tree.item(i, image=im))
            trade = sale["platform"] == "Trade"
            tag = "trade" if trade else "sale"
            self.tree.insert("", "end", iid=iid, text=f"  {sale['name']}", image=img or "", values=(
                fmt_date(sale["date"]), sale["platform"], sale["quantity"] or "–", fmt_eur(sale["price"]),
                fmt_eur(sale["fees"] + sale["shipping"]), fmt_eur(sale["cost"]),
                "Tausch" if trade else fmt_eur(sale["profit"], True)), tags=(tag,))
        self._sel_changed()

    def _selected_sale(self):
        sel = self.tree.selection()
        if not sel:
            return None
        row = int(sel[0][1:])
        return next((s for s in self.app.sales if s["row"] == row), None)

    def _sel_changed(self):
        s = self._selected_sale()
        self.undo_btn.set_enabled(s is not None)
        self.open_btn.set_enabled(s is not None and s["id"] in self.app.by_id)

    def _open(self):
        s = self._selected_sale()
        if s and s["id"] in self.app.by_id:
            self.app.show_card(s["id"])

    def _undo(self):
        s = self._selected_sale()
        if not s:
            return
        if messagebox.askyesno("Verkauf rückgängig",
                               f"Verkauf von „{s['name']}“ vom {fmt_date(s['date'])} löschen\n"
                               "und die Menge zurück in die Sammlung buchen?"):
            self.app.mutate(undo_sale, s["row"], ok="Verkauf rückgängig gemacht")

    def on_show(self, search=None, **kw):
        if search is not None:
            self.search.set(search)
            self.refresh()


# ============================================================
# GUI: STATISTIK
# ============================================================

class StatsView(View):
    def __init__(self, parent, app):
        super().__init__(parent, app)
        self.scroll = ScrollFrame(self, bg=BG)
        self.scroll.pack(fill="both", expand=True)
        self.inner = self.scroll.inner

    def refresh(self):
        for w in self.inner.winfo_children():
            w.destroy()
        app = self.app
        s = compute_statistics(app.cards, app.sales, app.meta)
        sub, _ = self.header(self.inner, "Statistik")
        sub.config(text="Zahlen aus der Excel-Datei; Marktwerte vom Cardmarket-Trend (TCGdex), täglich aktualisiert.")

        grid = tk.Frame(self.inner, bg=BG)
        grid.pack(fill="x", padx=28, pady=(18, 0))
        delta = s["market_value"] - s["market_cost"]
        pct = f"{delta / s['market_cost'] * 100:+.1f} % auf den Einkauf" if s["market_cost"] else "noch keine Preise"
        tiles = [
            ("Karten im Bestand", str(s["units"]), TEXT,
             f"{s['entries']} Einträge" + (f", davon {s['sealed_units']} Sealed" if s["sealed_units"] else "")),
            ("Investiert gesamt", fmt_eur(s["total_spent"]), TEXT, "inkl. verkaufter Karten, ohne Tausch"),
            ("Bestand zum Einkaufspreis", fmt_eur(s["stock_value"]), TEXT, "was aktuell in der Sammlung steckt"),
            ("Marktwert (Cardmarket)", fmt_eur(s["market_value"]) if s["market_units"] else "–", TEXT,
             f"{s['market_units']} ungegradete Stück mit Preis (ohne Slabs)"),
            ("Wertentwicklung", fmt_eur(delta, True) if s["market_units"] else "–",
             GREEN if delta >= 0 else RED, pct),
            ("Realisierter Gewinn", fmt_eur(s["profit"], True), GREEN if s["profit"] >= 0 else RED,
             f"aus {s['sales_count']} Verkäufen"),
            ("Umsatz", fmt_eur(s["revenue"]), TEXT, f"Gebühren & Versand {fmt_eur(s['fees'])}"),
            ("Grading", f"{s['graded']} gegradet", TEXT,
             f"{s['at_grading']} eingereicht, {s['queued']} vorgemerkt"),
        ]
        for i, (label, value, color, note) in enumerate(tiles):
            grid.columnconfigure(i % 4, weight=1, uniform="k")
            box = tk.Frame(grid, bg=SURFACE)
            box.grid(row=i // 4, column=i % 4, sticky="nsew", padx=(0, 10), pady=(0, 10))
            tk.Label(box, text=label, bg=SURFACE, fg=MUTED, font=F["sm"]).pack(anchor="w", padx=16, pady=(14, 0))
            tk.Label(box, text=value, bg=SURFACE, fg=color, font=F["num_b"]).pack(anchor="w", padx=16, pady=(2, 0))
            tk.Label(box, text=note, bg=SURFACE, fg=FAINT, font=F["xs"], wraplength=230, justify="left").pack(
                anchor="w", padx=16, pady=(2, 14))

        charts = tk.Frame(self.inner, bg=BG)
        charts.pack(fill="x", padx=28, pady=(10, 0))
        charts.columnconfigure(0, weight=1, uniform="c")
        charts.columnconfigure(1, weight=1, uniform="c")

        stock = [c for c in app.cards if c["status"] != "Sold"]
        by_set = {}
        for c in stock:
            if c["source"] != "Trade":
                by_set[c["set"] or "Ohne Set"] = by_set.get(c["set"] or "Ohne Set", 0) + c["price"] * c["quantity"]
        top_sets = sorted(by_set.items(), key=lambda x: -x[1])[:8]
        self._hbar_panel(charts, 0, "Einkaufswert im Bestand nach Set", top_sets)

        sold_cost = {}
        for sale in app.sales:
            sold_cost[sale["id"]] = sold_cost.get(sale["id"], 0) + sale["cost"]
        months = []
        today = date.today()
        for k in range(11, -1, -1):
            y, m = today.year, today.month - k
            while m <= 0:
                m += 12
                y -= 1
            months.append((y, m))
        spend = {ym: 0.0 for ym in months}
        for c in app.cards:
            d = c["purchase_date"]
            if isinstance(d, datetime):
                d = d.date()
            if not isinstance(d, date) or c["source"] == "Trade":
                continue
            ym = (d.year, d.month)
            if ym in spend:
                spend[ym] += c["price"] * max(0, c["quantity"]) + sold_cost.get(c["id"], 0)
        mnames = ["Jan", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"]
        self._vbar_panel(charts, 1, "Einkäufe pro Monat",
                         [(mnames[m - 1] + (f" {str(y)[2:]}" if m == 1 else ""), spend[(y, m)]) for y, m in months])

        top = [(market_price(app.meta, c["id"]) * c["quantity"], c) for c in stock
               if market_price(app.meta, c["id"]) is not None and c["grading_status"] != "Graded"]
        top.sort(key=lambda x: -x[0])
        panel = tk.Frame(self.inner, bg=SURFACE)
        panel.pack(fill="x", padx=28, pady=(10, 28))
        tk.Label(panel, text="Wertvollste Karten (Marktwert, ungegradet)", bg=SURFACE, fg=TEXT, font=F["md_b"]).pack(
            anchor="w", padx=18, pady=(16, 10))
        if not top:
            tk.Label(panel, text="Noch keine Marktpreise geladen. Sie kommen automatisch, sobald die Karten bei TCGdex zugeordnet sind.",
                     bg=SURFACE, fg=FAINT, font=F["sm"]).pack(anchor="w", padx=18, pady=(0, 18))
        row = tk.Frame(panel, bg=SURFACE)
        row.pack(fill="x", padx=18, pady=(0, 18))
        for value, c in top[:6]:
            cell = tk.Frame(row, bg=SURFACE, cursor="hand2")
            cell.pack(side="left", padx=(0, 18), anchor="n")
            path, kind = app.image_for(c["id"])
            lbl = tk.Label(cell, bg=SURFACE2, image=app.thumbs.blank(100, 140))
            lbl.pack(anchor="w")
            img = app.thumbs.get(path, 100, 140, kind or "card", False, 7,
                                 on_ready=lambda im, l=lbl: l.winfo_exists() and l.config(image=im))
            if img:
                lbl.config(image=img)
            tk.Label(cell, text=truncate(c["name"], F["sm_b"], 110), bg=SURFACE, fg=TEXT, font=F["sm_b"]).pack(anchor="w", pady=(6, 0))
            diff = value - c["price"] * c["quantity"]
            tk.Label(cell, text=fmt_eur(value), bg=SURFACE, fg=TEXT, font=F["sm"]).pack(anchor="w")
            tk.Label(cell, text=fmt_eur(diff, True), bg=SURFACE, fg=GREEN if diff >= 0 else RED,
                     font=F["xs_b"]).pack(anchor="w")
            for w in (cell, lbl):
                w.bind("<Button-1>", lambda e, cid=c["id"]: app.show_card(cid))

    def _panel(self, parent, col, title):
        p = tk.Frame(parent, bg=SURFACE)
        p.grid(row=0, column=col, sticky="nsew", padx=(0, 10))
        tk.Label(p, text=title, bg=SURFACE, fg=TEXT, font=F["md_b"]).pack(anchor="w", padx=18, pady=(16, 6))
        cv = tk.Canvas(p, bg=SURFACE, height=290, highlightthickness=0)
        cv.pack(fill="x", padx=18, pady=(0, 16))
        return cv

    def _hbar_panel(self, parent, col, title, data):
        cv = self._panel(parent, col, title)

        def draw(_e=None):
            cv.delete("all")
            w = cv.winfo_width()
            if not data:
                cv.create_text(w / 2, 130, text="Noch keine Daten", fill=FAINT, font=F["sm"])
                return
            mx = max(v for _, v in data) or 1
            label_w = 150
            bar_max = max(40, w - label_w - 90)
            for i, (name, v) in enumerate(data):
                y = 8 + i * 35
                cv.create_text(0, y + 11, text=truncate(name, F["sm"], label_w - 12), anchor="w", fill=MUTED, font=F["sm"])
                bw = max(4, v / mx * bar_max)
                round_rect(cv, label_w, y + 2, label_w + bw, y + 20, 4, fill=ACCENT, outline="")
                cv.create_text(label_w + bw + 8, y + 11, text=fmt_eur(v), anchor="w", fill=TEXT, font=F["sm_b"])
        cv.bind("<Configure>", draw)

    def _vbar_panel(self, parent, col, title, data):
        cv = self._panel(parent, col, title)

        def draw(_e=None):
            cv.delete("all")
            w = cv.winfo_width()
            h = 290
            base = h - 26
            mx = max((v for _, v in data), default=0)
            if not mx:
                cv.create_text(w / 2, 130, text="In den letzten 12 Monaten keine Käufe mit Datum",
                               fill=FAINT, font=F["sm"])
                return
            cv.create_line(0, base, w, base, fill=BORDER)
            slot = w / len(data)
            bw = min(34, slot * 0.62)
            peak = max(range(len(data)), key=lambda i: data[i][1])
            for i, (label, v) in enumerate(data):
                x = i * slot + slot / 2
                bh = v / mx * (base - 34)
                if v > 0:
                    item = round_rect(cv, x - bw / 2, base - max(bh, 4), x + bw / 2, base, 4,
                                      fill=ACCENT, outline="", tags=("bar",))
                    cv.tag_bind(item, "<Enter>", lambda e, x=x, v=v, label=label, top=base - bh: show(x, top, label, v))
                    cv.tag_bind(item, "<Leave>", lambda e: cv.delete("tip"))
                if i == peak:
                    cv.create_text(x, base - bh - 10, text=fmt_eur(v), fill=TEXT, font=F["xs_b"])
                cv.create_text(x, base + 13, text=label, fill=MUTED, font=F["xs"])

        def show(x, top, label, v):
            cv.delete("tip")
            text = f"{label}: {fmt_eur(v)}"
            tw = F["sm_b"].measure(text) + 16
            x1 = min(max(0, x - tw / 2), cv.winfo_width() - tw)
            y1 = max(0, top - 36)
            round_rect(cv, x1, y1, x1 + tw, y1 + 24, 6, fill=SURFACE4, outline="", tags=("tip",))
            cv.create_text(x1 + tw / 2, y1 + 12, text=text, fill=TEXT, font=F["sm_b"], tags=("tip",))
        cv.bind("<Configure>", draw)

    def soft_refresh(self):
        y = self.scroll.canvas.yview()[0]
        self.refresh()
        self.scroll.canvas.yview_moveto(y)


# ============================================================
# GUI: EINSTELLUNGEN
# ============================================================

class SettingsView(View):
    def __init__(self, parent, app):
        super().__init__(parent, app)
        scroll = ScrollFrame(self, bg=BG)
        scroll.pack(fill="both", expand=True)
        root = scroll.inner
        self.header(root, "Einstellungen", "eBay, Versand, Bilder & Preise, Daten und Updates.")
        grid = tk.Frame(root, bg=BG)
        grid.pack(fill="x", padx=28, pady=(18, 28))
        grid.columnconfigure(0, weight=1, uniform="s")
        grid.columnconfigure(1, weight=1, uniform="s")
        cfg = _load_ebay_config()

        # eBay
        ebay = self._card(grid, 0, 0, "eBay-Verbindung", "Für Entwürfe und Listings direkt aus der App.")
        self.cfg_client_id = self._field(ebay, "Client ID (App ID)", cfg.get("client_id", ""))
        self.cfg_client_secret = self._field(ebay, "Client Secret (Cert ID)", cfg.get("client_secret", ""), show="•")
        self.cfg_ru_name = self._field(ebay, "RuName", cfg.get("ru_name", ""))
        tk.Label(ebay, text="Marktplatz", bg=SURFACE, fg=MUTED, font=F["sm_b"]).pack(anchor="w", pady=(10, 6))
        self.cfg_marketplace = Chips(ebay, ["EBAY_DE", "EBAY_US", "EBAY_GB", "EBAY_FR", "EBAY_IT", "EBAY_ES"],
                                     cfg.get("marketplace", "EBAY_DE"), bg=SURFACE, small=True,
                                     labels={m: m.split("_")[1] for m in ["EBAY_DE", "EBAY_US", "EBAY_GB", "EBAY_FR", "EBAY_IT", "EBAY_ES"]})
        self.cfg_marketplace.pack(fill="x")
        brow = tk.Frame(ebay, bg=SURFACE)
        brow.pack(fill="x", pady=(14, 0))
        Btn(brow, "Speichern & verbinden", self._ebay_connect, kind="primary", bg=SURFACE).pack(side="left")
        Btn(brow, "Nur speichern", self._ebay_save, kind="subtle", bg=SURFACE).pack(side="left", padx=(8, 0))
        self.ebay_status = tk.Label(ebay, text="", bg=SURFACE, font=F["sm_b"])
        self.ebay_status.pack(anchor="w", pady=(10, 0))
        self._set_ebay_status("Verbunden" if cfg.get("access_token") else "Nicht verbunden",
                              GREEN if cfg.get("access_token") else RED)

        # Versand
        ship = self._card(grid, 0, 1, "Versand (eBay)", "Wird in die eBay-Versandrichtlinie geschrieben.")
        self.cfg_ship_de = self._field(ship, "Deutschland", f"{cfg.get('shipping_de', 6.19):.2f}".replace(".", ","), suffix="€")
        self.cfg_ship_eu = self._field(ship, "EU", f"{cfg.get('shipping_eu', 14.49):.2f}".replace(".", ","), suffix="€")
        self.cfg_ship_intl = self._field(ship, "International", f"{cfg.get('shipping_intl', 26.49):.2f}".replace(".", ","), suffix="€")
        tk.Label(ship, text="Bearbeitungszeit (Tage)", bg=SURFACE, fg=MUTED, font=F["sm_b"]).pack(anchor="w", pady=(10, 6))
        self.cfg_handling = Chips(ship, ["1", "2", "3", "4", "5"], str(cfg.get("handling_days", 2)), bg=SURFACE, small=True)
        self.cfg_handling.pack(fill="x")
        Btn(ship, "Versand speichern", self._save_shipping, kind="primary", bg=SURFACE).pack(anchor="w", pady=(14, 0))
        self.ship_status = tk.Label(ship, text="", bg=SURFACE, fg=MUTED, font=F["sm_b"])
        self.ship_status.pack(anchor="w", pady=(10, 0))

        # Bilder & Preise
        imgs = self._card(grid, 1, 0, "Marktpreise",
                          "Cardmarket-Preise kommen kostenlos von TCGdex. Passt ein Preis nicht: "
                          "Karte rechtsklicken, „Für Marktpreis zuordnen“.")
        self.img_stats = tk.Label(imgs, text="", bg=SURFACE, fg=TEXT, font=F["sm"], justify="left")
        self.img_stats.pack(anchor="w", pady=(4, 10))
        self.auto_prices = tk.BooleanVar(value=app.config.get("auto_prices", True))
        cb = tk.Checkbutton(imgs, text="Preise beim Start automatisch aktualisieren (1× täglich)",
                            variable=self.auto_prices, command=self._toggle_auto, bg=SURFACE, fg=TEXT,
                            selectcolor=SURFACE2, activebackground=SURFACE, activeforeground=TEXT,
                            font=F["sm"], highlightthickness=0, bd=0)
        cb.pack(anchor="w")
        brow = tk.Frame(imgs, bg=SURFACE)
        brow.pack(fill="x", pady=(14, 0))
        Btn(brow, "Fehlende Zuordnungen suchen", self._find_images, kind="subtle", bg=SURFACE).pack(side="left")
        Btn(brow, "Preise jetzt laden", self._load_prices, kind="subtle", bg=SURFACE).pack(side="left", padx=(8, 0))

        # Daten & App
        data = self._card(grid, 1, 1, "Daten & App", f"Alles liegt lokal in {BASE_DIR}")
        brow = tk.Frame(data, bg=SURFACE)
        brow.pack(fill="x", pady=(4, 0))
        Btn(brow, "Datenordner öffnen", lambda: _open_path(BASE_DIR), kind="subtle", bg=SURFACE).pack(side="left")
        Btn(brow, "Excel öffnen", lambda: _open_path(EXCEL_PATH), kind="subtle", bg=SURFACE).pack(side="left", padx=(8, 0))
        tk.Label(data, text=f"Version {VERSION}", bg=SURFACE, fg=TEXT, font=F["body_b"]).pack(anchor="w", pady=(18, 6))
        Btn(data, "Nach Updates suchen", app.check_update_now, kind="subtle", bg=SURFACE).pack(anchor="w")
        keys = (f"{MOD_LABEL}1–6  Bereiche wechseln\n{MOD_LABEL}N  Karte hinzufügen\n{MOD_LABEL}F  Suchen\n"
                f"{MOD_LABEL}R  Neu laden (nach Änderungen in Excel)\n"
                "Pfeiltasten  Karte wechseln · Leertaste  Bild groß · ↵  Bearbeiten")
        tk.Label(data, text="TASTENKÜRZEL", bg=SURFACE, fg=FAINT, font=F["xs_b"]).pack(anchor="w", pady=(18, 4))
        tk.Label(data, text=keys, bg=SURFACE, fg=MUTED, font=F["sm"], justify="left").pack(anchor="w")

    def _card(self, grid, row, col, title, subtitle):
        box = tk.Frame(grid, bg=SURFACE)
        box.grid(row=row, column=col, sticky="nsew", padx=(0, 12), pady=(0, 12))
        inner = tk.Frame(box, bg=SURFACE)
        inner.pack(fill="both", expand=True, padx=22, pady=20)
        tk.Label(inner, text=title, bg=SURFACE, fg=TEXT, font=F["md_b"]).pack(anchor="w")
        tk.Label(inner, text=subtitle, bg=SURFACE, fg=MUTED, font=F["sm"], wraplength=420,
                 justify="left").pack(anchor="w", pady=(2, 8))
        return inner

    def _field(self, parent, label, value, show=None, suffix=None):
        tk.Label(parent, text=label, bg=SURFACE, fg=MUTED, font=F["sm_b"]).pack(anchor="w", pady=(10, 6))
        f = Field(parent, width=40, value=value, show=show, suffix=suffix)
        f.pack(fill="x")
        return f

    def refresh(self):
        stock = [c for c in self.app.cards if c["status"] != "Sold"]
        with_img = sum(1 for c in stock if self.app.image_for(c["id"])[0])
        prices = [self.app.price_info(c["id"]).get("updated") for c in stock]
        prices = [p for p in prices if p]
        stand = datetime.fromtimestamp(max(prices)).strftime("%d.%m.%Y %H:%M") if prices else "noch nie"
        self.img_stats.config(text=f"{with_img} von {len(stock)} Karten im Bestand haben ein Foto\n"
                                   f"{len(prices)} mit Marktpreis · letzter Abruf {stand}")

    def _set_ebay_status(self, text, color):
        self.ebay_status.config(text=f"Status: {text}", fg=color)

    def _ebay_save(self, quiet=False):
        cfg = _load_ebay_config()
        cfg["client_id"] = self.cfg_client_id.get()
        cfg["client_secret"] = self.cfg_client_secret.get()
        cfg["ru_name"] = self.cfg_ru_name.get()
        cfg["marketplace"] = self.cfg_marketplace.get()
        _save_ebay_config(cfg)
        if not quiet:
            self.app.toast.show("eBay-Einstellungen gespeichert")

    def _ebay_connect(self):
        self._ebay_save(quiet=True)
        cfg = _load_ebay_config()
        if not cfg.get("client_id") or not cfg.get("client_secret") or not cfg.get("ru_name"):
            messagebox.showwarning("Felder fehlen", "Bitte Client ID, Client Secret und RuName ausfüllen.")
            return
        auth_url = (
            f"{_ebay_auth_base(cfg)}/oauth2/authorize"
            f"?client_id={urllib.parse.quote(cfg['client_id'])}"
            f"&response_type=code"
            f"&redirect_uri={urllib.parse.quote(cfg['ru_name'])}"
            f"&scope={urllib.parse.quote(EBAY_SCOPES)}"
        )
        webbrowser.open(auth_url)
        self._set_ebay_status("Warte auf Freigabe …", AMBER)
        paste_url = simpledialog.askstring(
            "eBay-Freigabe",
            "1. Im Browser bei eBay anmelden und Zugriff erlauben.\n"
            "2. Danach wirst du weitergeleitet.\n"
            "3. Die KOMPLETTE Adresse aus der Adressleiste kopieren\n"
            "   und hier einfügen:\n",
            parent=self.app.root,
        )
        if not paste_url:
            self._set_ebay_status("Nicht verbunden", RED)
            return
        try:
            _ebay_exchange_code(_extract_auth_code(paste_url))
            _ebay_fetch_policies()
            self._set_ebay_status("Verbunden", GREEN)
            self.app.toast.show("Mit eBay verbunden")
        except Exception as e:
            self._set_ebay_status("Verbindung fehlgeschlagen", RED)
            messagebox.showerror("eBay-Fehler", str(e))

    def _save_shipping(self):
        try:
            de = parse_eur(self.cfg_ship_de.get())
            eu = parse_eur(self.cfg_ship_eu.get())
            intl = parse_eur(self.cfg_ship_intl.get())
            handling = int(self.cfg_handling.get())
        except ValueError:
            messagebox.showwarning("Ungültige Eingabe", "Bitte gültige Beträge eintragen.")
            return
        cfg = _load_ebay_config()
        if not cfg.get("access_token"):
            cfg.update(shipping_de=de, shipping_eu=eu, shipping_intl=intl, handling_days=handling)
            _save_ebay_config(cfg)
            self.ship_status.config(text="Lokal gespeichert (eBay nicht verbunden).", fg=SKY)
            return
        self.ship_status.config(text="Aktualisiere eBay …", fg=AMBER)
        self.update_idletasks()
        try:
            _ebay_update_shipping(de, eu, intl, handling)
            self.ship_status.config(text="Versandkosten aktualisiert.", fg=GREEN)
        except Exception as e:
            if "20403" in str(e) or "identisch" in str(e):
                self.ship_status.config(text="Gespeichert (auf eBay schon aktuell).", fg=GREEN)
            else:
                self.ship_status.config(text="Aktualisierung fehlgeschlagen.", fg=RED)
                messagebox.showerror("eBay-Fehler", str(e))

    def _toggle_auto(self):
        self.app.config["auto_prices"] = self.auto_prices.get()

    def _find_images(self):
        missing = [c for c in self.app.cards if not self.app.meta.get(str(c["id"]), {}).get("match")
                   and self.app.meta.get(str(c["id"]), {}).get("status") != "hidden"]
        for c in missing:
            m = self.app.meta.get(str(c["id"]))
            if m and m.get("status") in ("none", "multi"):
                m["checked"] = 0
        self.app.ensure_images(missing, force=True)
        self.app.toast.show(f"Ordne {len(missing)} Karten zu …" if missing else "Alle Karten sind zugeordnet",
                            "info")

    def _load_prices(self):
        n = self.app.refresh_prices(force=True)
        self.app.toast.show(f"Lade Preise für {n} Karten …" if n else "Keine Karten mit zugeordnetem Bild", "info")


def _open_path(path):
    try:
        if IS_MAC:
            subprocess.Popen(["open", path])
        elif sys.platform.startswith("win"):
            os.startfile(path)
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception as e:
        messagebox.showerror("Fehler", str(e))


# ============================================================
# GUI: DIALOGE
# ============================================================

def _card_title(c):
    return f"#{c['id']:03d} {c['name']}" + (f" · {c['set']}" if c["set"] else "")


class SellDialog(Dialog):
    def __init__(self, app, c):
        super().__init__(app, "Verkauf eintragen", _card_title(c))
        self.c = c
        b = self.body
        b.columnconfigure(0, weight=1)
        b.columnconfigure(1, weight=1)
        total = max(1, c["quantity"])
        self.label(b, f"ANZAHL (von {total})", 0, 0, pady=(0, 4))
        self.label(b, "VERKAUFSPREIS GESAMT", 0, 1, pady=(0, 4))
        self.qty = Field(b, width=6, value=str(total))
        self.qty.grid(row=1, column=0, sticky="w")
        trend = market_price(app.meta, c["id"])
        self.price = Field(b, width=12, suffix="€", placeholder=f"{trend:.2f}".replace(".", ",") if trend else "")
        self.price.grid(row=1, column=1, sticky="ew")
        self.label(b, "PLATTFORM", 2, 0)
        self.platform = Chips(b, PLATFORMS, app.config.get("last_platform", "Cardmarket"),
                              lambda v: self._update(), bg=SURFACE, small=True)
        self.platform.grid(row=3, column=0, columnspan=2, sticky="ew")
        self.label(b, "GEBÜHREN", 4, 0)
        self.label(b, "VERSANDKOSTEN (DEINE)", 4, 1)
        self.fees = Field(b, width=10, value="0", suffix="€")
        self.fees.grid(row=5, column=0, sticky="w", padx=(0, 10))
        self.ship = Field(b, width=10, value="0", suffix="€")
        self.ship.grid(row=5, column=1, sticky="ew")
        self.result = tk.Label(b, text="", bg=SURFACE, fg=MUTED, font=F["body_b"], anchor="w", justify="left")
        self.result.grid(row=6, column=0, columnspan=2, sticky="w", pady=(18, 0))
        for f in (self.qty, self.price, self.fees, self.ship):
            f.var.trace_add("write", lambda *_: self._update())
        self.button("Verkauf speichern", self.confirm, "primary")
        self.button("Abbrechen", self.close)
        self._update()
        self.show(self.price)

    def _values(self):
        qty = int(self.qty.get())
        price = parse_eur(self.price.get())
        fees = parse_eur(self.fees.get(), 0.0)
        ship = parse_eur(self.ship.get(), 0.0)
        return qty, price, fees, ship

    def _update(self):
        try:
            qty, price, fees, ship = self._values()
        except ValueError:
            self.result.config(text="Preis eintragen, dann siehst du hier den Gewinn.", fg=MUTED)
            return
        cost = sum(lot["price"] * n for lot, n in fifo_plan(self.c, qty))
        profit = price - fees - ship - cost
        if self.platform.get() == "Trade":
            self.result.config(text="Tausch: zählt nicht zu Umsatz und Gewinn.", fg=MUTED)
            return
        self.result.config(text=f"Netto {fmt_eur(price - fees - ship)}  ·  Einkauf {fmt_eur(cost)}\n"
                                f"Gewinn {fmt_eur(profit, True)}", fg=GREEN if profit >= 0 else RED)

    def confirm(self):
        try:
            qty, price, fees, ship = self._values()
        except ValueError:
            messagebox.showwarning("Ungültig", "Anzahl ganze Zahl, Beträge als Zahl (z. B. 12,50).", parent=self)
            return
        total = max(1, self.c["quantity"])
        if not 1 <= qty <= total:
            messagebox.showwarning("Ungültige Anzahl", f"Zwischen 1 und {total}.", parent=self)
            return
        self.app.config["last_platform"] = self.platform.get()
        cid = self.c["id"]
        self.close()
        if "lots" in self.c:
            self.app.mutate(sell_lots, fifo_plan(self.c, qty), price, self.platform.get(), fees, ship,
                            ok=f"Verkauf gespeichert: {fmt_eur(price)}", select=cid)
            return
        self.app.mutate(mark_card_sold, cid, price, self.platform.get(), fees, qty, ship,
                        ok=f"Verkauf gespeichert: {fmt_eur(price)}", select=cid)


class EditDialog(Dialog):
    def __init__(self, app, c):
        super().__init__(app, "Karte bearbeiten", f"#{c['id']:03d}", width=560)
        self.c = c
        b = self.body
        b.columnconfigure(0, weight=3)
        b.columnconfigure(1, weight=2)
        b.columnconfigure(2, weight=1)
        self.label(b, "NAME", 0, 0, pady=(0, 4))
        self.label(b, "SET", 0, 1, pady=(0, 4))
        self.label(b, "NUMMER", 0, 2, pady=(0, 4))
        self.name = Field(b, width=24, value=c["name"])
        self.name.grid(row=1, column=0, sticky="ew", padx=(0, 8))
        self.set_cb = ttk.Combobox(b, values=SETS, font=F["body"])
        self.set_cb.set(c["set"])
        self.set_cb.grid(row=1, column=1, sticky="ew", padx=(0, 8), ipady=3)
        self.number = Field(b, width=8, value=c["number"])
        self.number.grid(row=1, column=2, sticky="ew")
        self.label(b, "SPRACHE", 2)
        self.lang = Chips(b, LANGUAGES, c["language"] or "EN", bg=SURFACE, small=True)
        self.lang.grid(row=3, column=0, columnspan=3, sticky="ew")
        self.label(b, "ZUSTAND", 4)
        self.cond = Chips(b, CONDITIONS, c["condition"], bg=SURFACE, small=True)
        self.cond.grid(row=5, column=0, columnspan=3, sticky="ew")
        self.label(b, "QUELLE", 6)
        self.source = Chips(b, SOURCES, c["source"] or "Cardmarket", bg=SURFACE, small=True)
        self.source.grid(row=7, column=0, columnspan=3, sticky="ew")
        self.label(b, "ANZAHL", 8, 0)
        self.label(b, "PREIS PRO STÜCK", 8, 1)
        self.qty = Field(b, width=6, value=str(c["quantity"]))
        self.qty.grid(row=9, column=0, sticky="w")
        self.price = Field(b, width=10, value=f"{c['price']:.2f}".replace(".", ","), suffix="€")
        self.price.grid(row=9, column=1, columnspan=2, sticky="ew")
        self.grading = c["grading_status"] in ("Graded", "Submitted")
        if self.grading:
            self.label(b, "GRADING-SERVICE", 10, 0)
            self.label(b, "NOTE", 10, 1)
            self.label(b, "CERT-NR.", 10, 2)
            self.service = ttk.Combobox(b, values=GRADING_SERVICES, font=F["body"])
            self.service.set(c["grading_service"])
            self.service.grid(row=11, column=0, sticky="ew", padx=(0, 8), ipady=3)
            self.grade = Field(b, width=6, value=fmt_grade(c["grade"]))
            self.grade.grid(row=11, column=1, sticky="ew", padx=(0, 8))
            self.cert = Field(b, width=10, value=c["cert"])
            self.cert.grid(row=11, column=2, sticky="ew")
        self.label(b, "CARDMARKET-LINK", 12)
        self.url = Field(b, width=40, value=get_card_url(c["id"]) or "")
        self.url.grid(row=13, column=0, columnspan=3, sticky="ew")
        self.button("Speichern", self.confirm, "primary")
        self.button("Abbrechen", self.close)
        self.show(self.name)

    def confirm(self):
        try:
            price = parse_eur(self.price.get(), 0.0)
            qty = int(self.qty.get() or "1")
            if qty < 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Ungültig", "Preis als Zahl, Anzahl als ganze Zahl.", parent=self)
            return
        extra = {}
        if self.grading:
            grade = None
            if self.grade.get():
                try:
                    grade = parse_grade(self.grade.get())
                except ValueError:
                    messagebox.showwarning("Ungültig", "Note: Zahl wie 9.5 oder P10 / B10.", parent=self)
                    return
            extra = {"grading_service": self.service.get().strip(), "grade": grade, "cert": self.cert.get()}
        c = self.c
        # Alle Werte lesen, bevor close() die Widgets zerstört (Combobox wirft danach TclError).
        url = self.url.get()
        name, set_name, number = self.name.get(), self.set_cb.get().strip(), self.number.get()
        lang, cond, source = self.lang.get(), self.cond.get(), self.source.get()
        changed_lookup = (name, set_name, number) != (c["name"], c["set"], c["number"])
        self.close()
        ok = self.app.mutate(edit_card, c["id"], name, set_name, number,
                             lang, cond, price, source, qty,
                             extra.get("grading_service"), extra.get("grade"), extra.get("cert"),
                             ok="Änderungen gespeichert", select=c["id"])
        if ok:
            if url != (get_card_url(c["id"]) or ""):
                _save_url(c["id"], url)
            m = self.app.meta.get(str(c["id"]), {})
            if changed_lookup and m.get("status") != "manual":
                # Name/Set/Nummer geaendert: Bild neu zuordnen lassen.
                self.app.meta.pop(str(c["id"]), None)
                self.app.ensure_images([self.app.by_id[c["id"]]], force=True)
            elif lang != c["language"] and m.get("match"):
                self.app.set_card_match(c["id"], m["match"], manual=m.get("status") == "manual")


class QueueDialog(Dialog):
    def __init__(self, app, c):
        typ = "Regrade" if c["grading_status"] == "Graded" else "Erstgrading"
        super().__init__(app, "Für Grading vormerken", f"{_card_title(c)}\n{typ}")
        self.c = c
        self.label(self.body, "NOTIZ (OPTIONAL)", 0, pady=(0, 4))
        self.note = Field(self.body, width=40, placeholder="z. B. Zentrierung prüfen, Ziel PSA 10")
        self.note.grid(row=1, column=0, sticky="ew")
        self.body.columnconfigure(0, weight=1)
        self.button("Vormerken", self.confirm, "primary")
        self.button("Abbrechen", self.close)
        self.show(self.note)

    def confirm(self):
        cid = self.c["id"]
        note = self.note.get()
        self.close()
        self.app.mutate(toggle_grading_vormerkung, cid, note, ok="Vorgemerkt", select=cid)


class SubmitDialog(Dialog):
    def __init__(self, app, cards):
        n = len(cards)
        sub = _card_title(cards[0]) if n == 1 else f"{n} Karten: " + ", ".join(c["name"] for c in cards[:4]) + \
            (" …" if n > 4 else "")
        super().__init__(app, "Zum Grading einreichen", sub)
        self.cards = cards
        self.label(self.body, "GRADING-SERVICE", 0, pady=(0, 6))
        self.service = Chips(self.body, GRADING_SERVICES, app.config.get("last_service", "PSA"), bg=SURFACE)
        self.service.grid(row=1, column=0, sticky="ew")
        regrades = [c for c in cards if c["grading_status"] == "Graded"]
        if regrades:
            tk.Label(self.body, text=f"{len(regrades)} davon ist/sind schon gegradet: die alte Note wird als "
                     "„Vorher“ gespeichert.", bg=SURFACE, fg=MUTED, font=F["sm"], wraplength=400,
                     justify="left").grid(row=2, column=0, sticky="w", pady=(12, 0))
        self.button("Einreichen", self.confirm, "primary")
        self.button("Abbrechen", self.close)
        self.show()

    def confirm(self):
        svc = self.service.get()
        self.app.config["last_service"] = svc
        ids = [c["id"] for c in self.cards]
        self.close()
        self.app.mutate(submit_for_grading, ids, svc,
                        ok=f"{len(ids)} Karte(n) bei {svc} eingereicht", select=ids[0] if len(ids) == 1 else None)


class ReturnDialog(Dialog):
    def __init__(self, app, c):
        super().__init__(app, "Grading zurück", f"{_card_title(c)}\nbei {c['grading_service']}")
        self.c = c
        b = self.body
        self.label(b, "NOTE", 0, 0, pady=(0, 4))
        self.label(b, "CERT-NR.", 0, 1, pady=(0, 4))
        self.grade = Field(b, width=8, placeholder="10")
        self.grade.grid(row=1, column=0, sticky="w", padx=(0, 10))
        self.cert = Field(b, width=18)
        self.cert.grid(row=1, column=1, sticky="ew")
        tk.Label(b, text="Auch möglich: P10 (Pristine) oder B10 (Black Label).", bg=SURFACE, fg=FAINT,
                 font=F["xs"]).grid(row=2, column=0, columnspan=2, sticky="w", pady=(8, 0))
        self.button("Speichern", self.confirm, "primary")
        self.button("Abbrechen", self.close)
        self.show(self.grade)

    def confirm(self):
        try:
            grade = parse_grade(self.grade.get()) if self.grade.get() else 0
        except ValueError:
            messagebox.showwarning("Ungültig", "Note: Zahl wie 9.5 oder P10 / B10.", parent=self)
            return
        cid = self.c["id"]
        cert = self.cert.get()
        self.close()
        self.app.mutate(grading_returned, cid, grade, cert, ok="Grading eingetragen", select=cid)


class EbayDialog(Dialog):
    def __init__(self, app, c):
        super().__init__(app, "Auf eBay listen", _card_title(c))
        self.c = c
        b = self.body
        b.columnconfigure(0, weight=1)
        b.columnconfigure(1, weight=1)
        photos = len(get_photo_paths(c["id"]))
        if not photos:
            tk.Label(b, text="Hinweis: Diese Karte hat noch keine eigenen Fotos. eBay-Listings brauchen "
                     "echte Fotos (Handy-Upload oder „+ Foto“).", bg=AMBER_SOFT, fg=AMBER, font=F["sm"],
                     wraplength=420, justify="left", padx=10, pady=8).grid(row=0, column=0, columnspan=2,
                                                                          sticky="ew", pady=(0, 10))
        self.label(b, "FORMAT", 1, pady=(0, 6))
        self.fmt = Chips(b, ["FIXED_PRICE", "AUCTION"], "FIXED_PRICE", self._fmt_changed, bg=SURFACE,
                         labels={"FIXED_PRICE": "Festpreis (Preisvorschlag an)", "AUCTION": "Auktion"})
        self.fmt.grid(row=2, column=0, columnspan=2, sticky="w")
        trend = market_price(app.meta, c["id"])
        self.fixed = tk.Frame(b, bg=SURFACE)
        self.fixed.grid(row=3, column=0, columnspan=2, sticky="ew")
        self.fixed.columnconfigure(0, weight=1)
        self.fixed.columnconfigure(1, weight=1)
        self.label(self.fixed, "PREIS", 0, 0)
        self.label(self.fixed, f"ANZAHL (von {max(1, c['quantity'])})", 0, 1)
        start = trend if trend else c["price"]
        self.price = Field(self.fixed, width=10, value=f"{start:.2f}".replace(".", ",") if start else "", suffix="€")
        self.price.grid(row=1, column=0, sticky="ew", padx=(0, 10))
        self.qty = Field(self.fixed, width=6, value=str(max(1, c["quantity"])))
        self.qty.grid(row=1, column=1, sticky="w")
        if trend:
            tk.Label(self.fixed, text=f"Vorschlag = Cardmarket-Trend ({fmt_eur(trend)}), Einkauf war {fmt_eur(c['price'])}",
                     bg=SURFACE, fg=FAINT, font=F["xs"]).grid(row=2, column=0, columnspan=2, sticky="w", pady=(6, 0))
        self.auction = tk.Frame(b, bg=SURFACE)
        self.auction.columnconfigure(0, weight=1)
        self.auction.columnconfigure(1, weight=1)
        self.label(self.auction, "STARTPREIS", 0, 0)
        self.label(self.auction, "DAUER", 0, 1)
        self.start = Field(self.auction, width=10, value="1,00", suffix="€")
        self.start.grid(row=1, column=0, sticky="ew", padx=(0, 10))
        self.duration = Chips(self.auction, ["DAYS_3", "DAYS_5", "DAYS_7", "DAYS_10"], "DAYS_7", bg=SURFACE,
                              small=True, labels={d: d.split("_")[1] + " T." for d in ["DAYS_3", "DAYS_5", "DAYS_7", "DAYS_10"]})
        self.duration.grid(row=1, column=1, sticky="w")
        self.label(self.auction, "SOFORT-KAUFEN (OPTIONAL)", 2, 0)
        self.label(self.auction, "MINDESTPREIS (OPTIONAL)", 2, 1)
        self.bin = Field(self.auction, width=10, suffix="€")
        self.bin.grid(row=3, column=0, sticky="ew", padx=(0, 10))
        self.reserve = Field(self.auction, width=10, suffix="€")
        self.reserve.grid(row=3, column=1, sticky="ew")
        self.status = tk.Label(b, text="", bg=SURFACE, fg=MUTED, font=F["sm_b"])
        self.status.grid(row=5, column=0, columnspan=2, sticky="w", pady=(14, 0))
        self.button("Jetzt veröffentlichen", lambda: self._create(True), "primary")
        self.button("Als Entwurf", lambda: self._create(False))
        self.button("Abbrechen", self.close)
        self.show(self.price)

    def _fmt_changed(self, v):
        if v == "AUCTION":
            self.fixed.grid_remove()
            self.auction.grid(row=3, column=0, columnspan=2, sticky="ew")
            self.start.focus()
        else:
            self.auction.grid_remove()
            self.fixed.grid()
            self.price.focus()

    def _validate(self):
        total = max(1, self.c["quantity"])
        try:
            if self.fmt.get() == "FIXED_PRICE":
                price = parse_eur(self.price.get())
                qty = int(self.qty.get())
                if price <= 0:
                    raise ValueError("Preis muss größer als 0 sein.")
                if not 1 <= qty <= total:
                    raise ValueError(f"Anzahl zwischen 1 und {total}.")
                return {"price": price, "quantity": qty}
            start = parse_eur(self.start.get())
            if start <= 0:
                raise ValueError("Startpreis muss größer als 0 sein.")
            reserve = parse_eur(self.reserve.get(), 0.0) or None
            bin_price = parse_eur(self.bin.get(), 0.0) or None
            if reserve is not None and reserve <= start:
                raise ValueError("Mindestpreis muss über dem Startpreis liegen.")
            if bin_price is not None and bin_price <= start:
                raise ValueError("Sofort-Kaufen muss über dem Startpreis liegen.")
            return {"auction_start_price": start, "auction_reserve_price": reserve,
                    "buy_it_now_price": bin_price, "auction_duration": self.duration.get()}
        except ValueError as e:
            msg = str(e) if "muss" in str(e) or "zwischen" in str(e) else "Bitte gültige Zahlen eintragen."
            messagebox.showwarning("Ungültig", msg, parent=self)
            return None

    def _create(self, publish):
        params = self._validate()
        if params is None:
            return
        self.status.config(text="Veröffentliche …" if publish else "Speichere Entwurf …", fg=AMBER)
        self.update()
        cid = self.c["id"]
        try:
            if self.fmt.get() == "FIXED_PRICE":
                offer_id, _ = ebay_create_draft(cid, params["price"], params["quantity"])
            else:
                offer_id, _ = ebay_create_draft(cid, 0, 1, listing_format="AUCTION", **params)
            if publish:
                listing_id = ebay_publish_offer(offer_id)
                self.close()
                try:
                    update_card_status(cid, "For Sale")
                except Exception:
                    pass
                self.app.reload()
                if messagebox.askyesno("eBay", f"Online! Listing-ID {listing_id}\n\nAngebot im Browser öffnen?"):
                    webbrowser.open(f"https://www.ebay.de/itm/{listing_id}")
            else:
                self.close()
                messagebox.showinfo("eBay", f"Entwurf gespeichert (Offer-ID {offer_id}).\n"
                                    "Später über „Auf eBay listen“ veröffentlichen.")
        except Exception as e:
            self.status.config(text="Fehlgeschlagen", fg=RED)
            messagebox.showerror("eBay-Fehler", str(e), parent=self)


class ImagePicker(Dialog):
    """Bild aus TCGdex waehlen: hilft, wenn viele Karten gleich heissen."""

    TW, TH = 120, 167

    def __init__(self, app, name, set_name="", candidates=None, on_pick=None, on_none=None):
        super().__init__(app, "Karte zuordnen",
                         "Für den Marktpreis die passende Karte anklicken. Gleichnamige Karten gibt es oft in vielen Sets.",
                         width=820)
        self.resizable(True, True)
        self.on_pick = on_pick
        bar = tk.Frame(self.body, bg=SURFACE)
        bar.pack(fill="x")
        self.query = Field(bar, width=30, value=name)
        self.query.pack(side="left")
        self.query.entry.bind("<Return>", lambda e: (self.search(), "break")[-1])
        Btn(bar, "Suchen", self.search, kind="secondary", bg=SURFACE).pack(side="left", padx=(8, 0))
        self.only_set = tk.BooleanVar(value=bool(set_name and tcg_find_set_safe(set_name)))
        self.set_name = set_name
        if set_name:
            tk.Checkbutton(bar, text=f"Nur Set „{set_name}“", variable=self.only_set, command=self._render,
                           bg=SURFACE, fg=TEXT, selectcolor=SURFACE2, activebackground=SURFACE,
                           activeforeground=TEXT, font=F["sm"], highlightthickness=0, bd=0).pack(side="left", padx=(14, 0))
        self.info = tk.Label(self.body, text="", bg=SURFACE, fg=MUTED, font=F["sm"], anchor="w")
        self.info.pack(fill="x", pady=(10, 6))
        self.grid_sf = ScrollFrame(self.body, bg=SURFACE)
        self.grid_sf.pack(fill="both", expand=True)
        self.grid_sf.configure(width=780, height=470)
        self.grid_sf.canvas.configure(width=780, height=470)
        self.results = list(candidates or [])
        if on_none:
            self.button("Nicht zuordnen", lambda: (self.close(), on_none()), "subtle")
        self.button("Abbrechen", self.close)
        self.unbind("<Return>")
        self.show(self.query)
        if self.results:
            self._render()
        else:
            self.search()

    def search(self):
        q = self.query.get()
        if not q:
            return
        self.info.config(text="Suche …")
        self.app.worker.submit(f"pick:{q}:{id(self)}", self._do_search, (q,), self._found)

    def _do_search(self, q):
        res = tcg_search(q)
        # Sealed / Set-Logo als Option, wenn der Begriff nach Produkt klingt.
        tset = tcg_find_set_safe(self.set_name)
        if tset and tset.get("logo") and is_sealed_name(q):
            res.insert(0, {"id": tset["id"], "image": tset["logo"], "kind": "logo", "name": "Set-Logo",
                           "localId": "", "set_id": tset["id"], "set_name": tset["name"]})
        return res

    def _found(self, res, err):
        if not self.winfo_exists():
            return
        if err:
            self.info.config(text="Suche fehlgeschlagen. Internetverbindung prüfen.")
            return
        self.results = res or []
        self._render()

    def _render(self):
        for w in self.grid_sf.inner.winfo_children():
            w.destroy()
        items = [r for r in self.results if r.get("image")]
        tset = tcg_find_set_safe(self.set_name) if self.only_set.get() else None
        if tset:
            items = [r for r in items if r.get("set_id") == tset["id"]] or items
        self.info.config(text=f"{len(items)} Treffer" + (" (neueste Sets zuerst)" if len(items) > 1 else "")
                         if items else "Keine Treffer. Anderen Namen probieren (englischer Kartenname).")
        order = {s["id"]: i for i, s in enumerate(tcg_sets_safe())}
        items = sorted(items, key=lambda r: -order.get(r.get("set_id"), -1))
        cols = 6
        for i, r in enumerate(items[:120]):
            cell = tk.Frame(self.grid_sf.inner, bg=SURFACE, cursor="hand2")
            cell.grid(row=i // cols, column=i % cols, padx=(0, 10), pady=(0, 12), sticky="n")
            lbl = tk.Label(cell, bg=SURFACE2, image=self.app.thumbs.blank(self.TW, self.TH))
            lbl.pack()
            tk.Label(cell, text=truncate(r.get("set_name", ""), F["xs_b"], self.TW), bg=SURFACE, fg=TEXT,
                     font=F["xs_b"]).pack(anchor="w", pady=(4, 0))
            tk.Label(cell, text=truncate(f"{r.get('name', '')} · {r.get('localId', '')}", F["xs"], self.TW),
                     bg=SURFACE, fg=MUTED, font=F["xs"]).pack(anchor="w")
            for w in cell.winfo_children() + [cell]:
                w.bind("<Button-1>", lambda e, r=r: self._choose(r))
            self._load_thumb(r, lbl)

    def _load_thumb(self, r, lbl):
        def done(path, err):
            if err or not path or not lbl.winfo_exists():
                return
            img = self.app.thumbs.get(path, self.TW, self.TH, r.get("kind", "card"), False, 6,
                                      on_ready=lambda im: lbl.winfo_exists() and lbl.config(image=im))
            if img:
                lbl.config(image=img)
        self.app.worker.submit(f"pv:{r['id']}:{r.get('kind')}:{id(lbl)}", download_preview, (dict(r, kind=r.get("kind", "card")),), done)

    def _choose(self, r):
        self.close()
        if self.on_pick:
            self.on_pick(dict(r, kind=r.get("kind", "card")))


def tcg_find_set_safe(name):
    try:
        return tcg_find_set(name)
    except Exception:
        return None


def tcg_sets_safe():
    try:
        return tcg_sets()
    except Exception:
        return []


class Viewer(tk.Toplevel):
    """Eigene Fotos gross ansehen, Pfeiltasten blaettern."""

    def __init__(self, app, c, items, start=0):
        super().__init__(app.root, bg=SIDEBAR)
        self.app = app
        self.c = c
        self.items = items
        self.idx = max(0, min(start, len(items) - 1))
        self.title(_card_title(c))
        self.transient(app.root)
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        # Platz für Menüleiste, Titelleiste, Dock und die Knopfleiste lassen,
        # sonst schneidet macOS das Fenster unten ab und die Knöpfe fehlen.
        self.max_w, self.max_h = min(760, sw - 120), min(900, sh - 320)
        # Knopfleiste zuerst packen, damit sie bei Platzmangel nicht verdrängt wird.
        bar = tk.Frame(self, bg=SIDEBAR)
        bar.pack(side="bottom", fill="x", padx=24, pady=(0, 18))
        self.img_lbl = tk.Label(self, bg=SIDEBAR)
        self.img_lbl.pack(padx=24, pady=(20, 10))
        self.prev_b = Btn(bar, "◀", lambda: self.go(-1), kind="secondary", bg=SIDEBAR)
        self.prev_b.pack(side="left")
        self.caption = tk.Label(bar, text="", bg=SIDEBAR, fg=MUTED, font=F["sm"])
        self.caption.pack(side="left", padx=12)
        self.next_b = Btn(bar, "▶", lambda: self.go(1), kind="secondary", bg=SIDEBAR)
        self.next_b.pack(side="left")
        Btn(bar, "Schließen", self.destroy, kind="secondary", bg=SIDEBAR).pack(side="right")
        self.del_b = Btn(bar, "Foto löschen", self.delete_photo, kind="danger", bg=SIDEBAR)
        self.bind("<Left>", lambda e: self.go(-1))
        self.bind("<Right>", lambda e: self.go(1))
        self.bind("<Escape>", lambda e: self.destroy())
        self.bind("<space>", lambda e: self.destroy())
        self.show()
        self.focus_set()

    def show(self):
        path, kind = self.items[self.idx]
        try:
            img = ImageOps.exif_transpose(Image.open(path))
            img.thumbnail((self.max_w, self.max_h), Image.LANCZOS)
            self._tk = ImageTk.PhotoImage(img)
            self.img_lbl.config(image=self._tk, text="")
        except Exception:
            self.img_lbl.config(image="", text="Bild kann nicht geladen werden", fg=MUTED, font=F["body"],
                                width=40, height=10)
        label = "Eigenes Foto"
        self.caption.config(text=f"{label}  ·  {self.idx + 1} / {len(self.items)}")
        multi = len(self.items) > 1
        self.prev_b.set_enabled(multi)
        self.next_b.set_enabled(multi)
        if kind == "photo":
            self.del_b.pack(side="right", padx=(0, 8))
        else:
            self.del_b.pack_forget()

    def go(self, d):
        if len(self.items) > 1:
            self.idx = (self.idx + d) % len(self.items)
            self.show()

    def delete_photo(self):
        path, kind = self.items[self.idx]
        if kind != "photo" or not messagebox.askyesno("Foto löschen", "Dieses Foto löschen?", parent=self):
            return
        try:
            os.remove(path)
        except OSError as e:
            messagebox.showerror("Fehler", str(e), parent=self)
            return
        self.items.pop(self.idx)
        self.app.card_image_changed(self.c["id"])
        if not self.items:
            self.destroy()
            return
        self.idx = min(self.idx, len(self.items) - 1)
        self.show()


# ============================================================
# START
# ============================================================

def _activate_window():
    if not IS_MAC:
        return
    subprocess.Popen([
        "osascript", "-e",
        'tell application "System Events" to set frontmost of first process '
        'whose unix id is ' + str(os.getpid()) + ' to true'
    ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def _set_dock_icon(root):
    """The launcher execs Python, so the Dock would show the Python rocket; use the bundle icon."""
    for app in ("/Applications/Poke Inv.app", os.path.expanduser("~/Desktop/Poke Inv.app")):
        icns = os.path.join(app, "Contents", "Resources", "AppIcon.icns")
        if os.path.exists(icns):
            try:
                with Image.open(icns) as im:
                    root._dock_icon = ImageTk.PhotoImage(im.convert("RGBA"))
                root.iconphoto(True, root._dock_icon)
            except Exception:
                pass
            return


if __name__ == "__main__":
    photo_server = start_photo_server()

    root = tk.Tk()
    _set_dock_icon(root)
    app = App(root, photo_server)
    root.after(300, _activate_window)
    root.mainloop()
