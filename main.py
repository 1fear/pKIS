import os
import re
import tempfile
import sys
import json
import logging
import subprocess
import hashlib
import html
import threading
from datetime import datetime, timedelta

import tkinter as tk
from tkinter import ttk, messagebox

import gspread
from oauth2client.service_account import ServiceAccountCredentials

from PIL import Image, ImageDraw, ImageFont

SPREADSHEET_ID = "1soHrN7Iqd3jk9iLGdUGK9APxVfRBwWXHxoI8x2Hsh1o"
SHEET_NAME = "data"

def get_app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

APP_DIR = get_app_dir()
CREDENTIALS_FILE = os.path.join(APP_DIR, "credentials.json")
LOG_FILE = os.path.join(APP_DIR, "pKIS.log")
BACKUP_DIR = os.path.join(APP_DIR, "scan_backups")
REPORTS_DIR = os.path.join(APP_DIR, "reports")
PENDING_PRINTS_FILE = os.path.join(APP_DIR, "pending_prints.json")

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    encoding="utf-8"
)

REQUIRED_COLUMNS = [
    "Дата получения заказа",
    "Тип оплаты",
    "Клиент",
    "Адрес",
    "Торговый представитель",
    "Товары",
    "Кол-во ШТ",
    "Кол-во блок",
    "Отсканированные коды",
]

LABEL_WIDTH_MM = 100
LABEL_HEIGHT_MM = 100
LABEL_DPI = 203
KIZ_MIN_LENGTH = 20
KIZ_MAX_LENGTH = 120

BG_MAIN = "#f5f7fa"
BG_CARD = "#ffffff"
FG_TEXT = "#1a1f2e"
FG_MUTED = "#6b7280"
ACCENT = "#4f46e5"
SUCCESS = "#10b981"
INFO = "#3b82f6"
WARNING = "#f59e0b"
DANGER = "#ef4444"
ERROR_BG = "#fee2e2"
ERROR_FG = "#dc2626"
BORDER = "#e5e7eb"

def format_exception_message(title, exc):
    return (
        f"{title}\n\n"
        f"Причина: {exc}\n\n"
        f"Подробности записаны в лог:\n{LOG_FILE}"
    )

def show_exception_message(title, exc):
    logging.exception(title)
    try:
        messagebox.showerror("Ошибка", format_exception_message(title, exc))
    except Exception:
        pass

def global_exception_handler(exc_type, exc_value, exc_traceback):
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return
    logging.error(
        "Неперехваченная ошибка",
        exc_info=(exc_type, exc_value, exc_traceback)
    )
    try:
        messagebox.showerror(
            "Критическая ошибка",
            format_exception_message("Неперехваченная ошибка", exc_value)
        )
    except Exception:
        pass

sys.excepthook = global_exception_handler

def get_google_client():
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
    return gspread.authorize(creds)

def clean_date_value(date_value):
    if date_value is None:
        return None
    date_str = str(date_value).strip()
    date_str = re.sub(r'^[\'"]+|[\'"]+$', '', date_str)
    if ' ' in date_str:
        date_str = date_str.split()[0]
    return date_str

def parse_date_to_standard(date_value):
    cleaned = clean_date_value(date_value)
    if not cleaned:
        return None
    formats = ["%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d.%m.%y", "%Y.%m.%d"]
    for fmt in formats:
        try:
            parsed = datetime.strptime(cleaned, fmt)
            return parsed.strftime("%d.%m.%Y")
        except:
            continue
    return cleaned

def normalize_text(value):
    return str(value or "").strip()

def normalize_header_name(value):
    return normalize_text(value).replace("\ufeff", "")

def get_header_index(header):
    return {normalize_header_name(col): idx for idx, col in enumerate(header) if normalize_header_name(col)}

def get_cell(row, idx):
    if idx is None or idx >= len(row):
        return ""
    return normalize_text(row[idx])

def split_codes(codes_str):
    if not codes_str:
        return []
    codes = []
    for line in str(codes_str).splitlines():
        code = line.strip()
        if code:
            codes.append(code)
    return codes

def normalize_payment_type(value):
    payment = normalize_text(value).lower().replace("ё", "е")
    if "терминал" in payment:
        return "terminal"
    if "перечис" in payment or "безнал" in payment:
        return "transfer"
    return "unknown"

def parse_int_value(value):
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return int(value)
    value_str = normalize_text(value).replace(" ", "").replace(",", ".")
    if not value_str:
        return 0
    try:
        return int(float(value_str))
    except ValueError:
        return 0

def get_plan_blocks(order):
    plan_blocks = parse_int_value(order.get("Кол-во блок", 0))
    if plan_blocks == 0:
        plan_blocks = parse_int_value(order.get("Кол-во блоков", 0))
    return plan_blocks

def is_order_completed(order):
    plan_blocks = get_plan_blocks(order)
    scanned_count = len(split_codes(order.get("Отсканированные коды")))
    return plan_blocks > 0 and scanned_count >= plan_blocks

def order_group_key(order):
    client = normalize_text(order.get("Клиент")) or "Клиент не указан"
    address = normalize_text(order.get("Адрес")) or "Адрес не указан"
    return (
        client,
        address,
    )

def row_matches_order(row, header_idx, order):
    checks = [
        ("Дата получения заказа", parse_date_to_standard(get_cell(row, header_idx.get("Дата получения заказа"))) == parse_date_to_standard(order.get("Дата получения заказа"))),
        ("Клиент", get_cell(row, header_idx.get("Клиент")) == normalize_text(order.get("Клиент"))),
        ("Адрес", get_cell(row, header_idx.get("Адрес")) == normalize_text(order.get("Адрес"))),
        ("Товары", get_cell(row, header_idx.get("Товары")) == normalize_text(order.get("Товары"))),
    ]
    return all(result for _, result in checks)

def validate_sheet_header(header):
    header_idx = get_header_index(header)
    missing = [col for col in REQUIRED_COLUMNS if col not in header_idx]
    return header_idx, missing

def get_all_existing_codes(sheet):
    try:
        all_rows = sheet.get_all_values()
        if not all_rows:
            return set()
        header_idx = get_header_index(all_rows[0])
        codes_idx = header_idx.get("Отсканированные коды")
        if codes_idx is None:
            logging.warning("Колонка 'Отсканированные коды' не найдена")
            return set()

        all_codes = set()
        for row in all_rows[1:]:
            for code in split_codes(get_cell(row, codes_idx)):
                all_codes.add(code)
        return all_codes
    except Exception as e:
        logging.exception("Не удалось загрузить существующие коды")
        return set()

def get_today_orders():
    try:
        client = get_google_client()
        sheet = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)
        all_rows = sheet.get_all_values()
        if not all_rows:
            raise ValueError("Лист Google Sheets пустой")

        header = [normalize_header_name(col) for col in all_rows[0]]
        header_idx, missing = validate_sheet_header(header)
        if missing:
            raise ValueError("В таблице не найдены обязательные колонки: " + ", ".join(missing))

        today_str = datetime.now().strftime("%d.%m.%Y")
        today_orders = []

        for row_number, row in enumerate(all_rows[1:], start=2):
            if not any(normalize_text(cell) for cell in row):
                continue

            record = {}
            for col_name, idx in header_idx.items():
                record[col_name] = get_cell(row, idx)

            normalized_date = parse_date_to_standard(record.get("Дата получения заказа"))
            scanned_codes = split_codes(record.get("Отсканированные коды"))

            if normalized_date == today_str and not is_order_completed(record):
                record["_row_number"] = row_number
                record["_normalized_date"] = normalized_date
                record["_existing_scanned_codes"] = scanned_codes
                today_orders.append(record)
        
        return today_orders, sheet
        
    except Exception as e:
        logging.exception("Не удалось загрузить данные из Google Sheets")
        raise

def fetch_sheet_data():
    today_orders, sheet = get_today_orders()
    all_existing_codes = get_all_existing_codes(sheet) if sheet else set()
    return today_orders, sheet, all_existing_codes

def update_scanned_codes_to_gsheet(sheet, order, scanned_codes):
    try:
        if not scanned_codes:
            return False, "Нет отсканированных кодов для записи"

        if len(scanned_codes) != len(set(scanned_codes)):
            return False, "В текущей позиции есть повторяющиеся коды"

        all_rows = sheet.get_all_values()
        if not all_rows:
            return False, "Лист Google Sheets пустой"

        header_idx, missing = validate_sheet_header(all_rows[0])
        if missing:
            return False, "В таблице не найдены обязательные колонки: " + ", ".join(missing)

        codes_idx = header_idx["Отсканированные коды"]
        target_row_number = parse_int_value(order.get("_row_number"))

        target_row = None
        if 2 <= target_row_number <= len(all_rows):
            candidate = all_rows[target_row_number - 1]
            if row_matches_order(candidate, header_idx, order):
                target_row = target_row_number

        if target_row is None:
            for row_number, row in enumerate(all_rows[1:], start=2):
                if row_matches_order(row, header_idx, order):
                    target_row = row_number
                    break

        if target_row is None:
            return False, "Не найдена строка заказа для записи кодов"

        existing_codes = split_codes(get_cell(all_rows[target_row - 1], codes_idx))
        if existing_codes:
            existing_set = set(existing_codes)
            scanned_set = set(scanned_codes)
            if not existing_set.issubset(scanned_set):
                return False, "В строке заказа уже есть другие отсканированные коды"

        duplicate_codes = []
        scanned_set = set(scanned_codes)
        for row_number, row in enumerate(all_rows[1:], start=2):
            if row_number == target_row:
                continue
            row_codes = set(split_codes(get_cell(row, codes_idx)))
            duplicates = scanned_set.intersection(row_codes)
            duplicate_codes.extend(sorted(duplicates))

        if duplicate_codes:
            return False, "Коды уже есть в другой строке Google Sheets: " + ", ".join(duplicate_codes[:3])

        sheet.update_cell(target_row, codes_idx + 1, "\n".join(scanned_codes))
        return True, "Коды записаны в Google Sheets"
    except Exception as e:
        logging.exception("Не удалось записать коды в Google Sheets")
        return False, str(e)

def build_day_report_rows_from_gsheet(sheet):
    all_rows = sheet.get_all_values()
    if not all_rows:
        return {"terminal": [], "transfer": [], "unknown": []}

    header_idx, missing = validate_sheet_header(all_rows[0])
    if missing:
        raise ValueError("В таблице не найдены обязательные колонки: " + ", ".join(missing))

    today_str = datetime.now().strftime("%d.%m.%Y")
    report_rows = {"terminal": [], "transfer": [], "unknown": []}

    for row in all_rows[1:]:
        if parse_date_to_standard(get_cell(row, header_idx.get("Дата получения заказа"))) != today_str:
            continue

        codes = split_codes(get_cell(row, header_idx.get("Отсканированные коды")))
        if not codes:
            continue

        payment_type = get_cell(row, header_idx.get("Тип оплаты"))
        payment_group = normalize_payment_type(payment_type)
        rows = report_rows[payment_group]

        for code in codes:
            rows.append({
                "Клиент": get_cell(row, header_idx.get("Клиент")),
                "Торговый представитель": get_cell(row, header_idx.get("Торговый представитель")),
                "Адрес": get_cell(row, header_idx.get("Адрес")),
                "Товар": get_cell(row, header_idx.get("Товары")),
                "Тип оплаты": payment_type,
                "Кол-во ШТ в блоке": 10,
                "Кол-во блок": 1,
                "Итого ШТ": 10,
                "Код": code
            })

    return report_rows

def build_summary_products_from_gsheet(sheet, group_key):
    all_rows = sheet.get_all_values()
    if not all_rows:
        return []

    header_idx, missing = validate_sheet_header(all_rows[0])
    if missing:
        raise ValueError("В таблице не найдены обязательные колонки: " + ", ".join(missing))

    today_str = datetime.now().strftime("%d.%m.%Y")
    products = []
    for row in all_rows[1:]:
        row_date = parse_date_to_standard(get_cell(row, header_idx.get("Дата получения заказа")))
        if row_date != today_str:
            continue

        row_group = (
            get_cell(row, header_idx.get("Клиент")) or "Клиент не указан",
            get_cell(row, header_idx.get("Адрес")) or "Адрес не указан",
        )
        if row_group != group_key:
            continue

        codes = split_codes(get_cell(row, header_idx.get("Отсканированные коды")))
        if not codes:
            continue

        products.append({
            "Клиент": get_cell(row, header_idx.get("Клиент")),
            "Адрес": get_cell(row, header_idx.get("Адрес")),
            "Торговый представитель": get_cell(row, header_idx.get("Торговый представитель")),
            "Товары": get_cell(row, header_idx.get("Товары")),
            "Тип оплаты": get_cell(row, header_idx.get("Тип оплаты")),
            "План": parse_int_value(get_cell(row, header_idx.get("Кол-во блок"))),
            "Отсканировано": len(codes),
            "Коды": codes,
        })

    return products

def load_font(candidates, size):
    for path in candidates:
        try:
            return ImageFont.truetype(path, size=size)
        except:
            continue
    return ImageFont.load_default()

def wrap_text(text, max_chars):
    if len(text) <= max_chars:
        return [text]
    words = text.split()
    lines = []
    current_line = ""
    for word in words:
        if len(current_line) + len(word) + 1 <= max_chars:
            if current_line:
                current_line += " " + word
            else:
                current_line = word
        else:
            if current_line:
                lines.append(current_line)
            current_line = word
    if current_line:
        lines.append(current_line)
    return lines

def mm_to_px(mm_value):
    return int(mm_value / 25.4 * LABEL_DPI)

def create_print_html(image_path):
    html_file = tempfile.NamedTemporaryFile(suffix=".html", delete=False, mode="w", encoding="utf-8")
    image_url = "file:///" + os.path.abspath(image_path).replace("\\", "/")
    html_file.write(f"""<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>Сводный лист 100x100</title>
<style>
@page {{
  size: {LABEL_WIDTH_MM}mm {LABEL_HEIGHT_MM}mm;
  margin: 0;
}}
html, body {{
  width: {LABEL_WIDTH_MM}mm;
  height: {LABEL_HEIGHT_MM}mm;
  margin: 0;
  padding: 0;
}}
img {{
  display: block;
  width: {LABEL_WIDTH_MM}mm;
  height: {LABEL_HEIGHT_MM}mm;
}}
</style>
</head>
<body>
<img src="{html.escape(image_url)}" alt="Сводный лист">
<script>
window.onload = function() {{
  window.focus();
  window.print();
}};
</script>
</body>
</html>""")
    html_file.close()
    return html_file.name

def send_file_to_printer(file_path):
    try:
        print_file = create_print_html(file_path)
        if os.name == 'nt':
            os.startfile(print_file)
            return True

        if sys.platform == "darwin":
            open_cmd = "open"
        else:
            open_cmd = "xdg-open"

        subprocess.run([open_cmd, print_file], check=True, timeout=10)
        return True
    except Exception:
        logging.exception("Не удалось отправить сводку на печать")
        return False

def print_summary(address, all_products):
    try:
        if not all_products:
            return None

        width_px = mm_to_px(LABEL_WIDTH_MM)
        height_px = mm_to_px(LABEL_HEIGHT_MM)
        scale = LABEL_DPI / 96

        def s(value):
            return int(value * scale)
        
        products_per_page = 3
        pages = []
        for i in range(0, len(all_products), products_per_page):
            pages.append(all_products[i:i + products_per_page])
        
        printed_files = []
        
        for page_idx, page_products in enumerate(pages):
            img = Image.new("RGB", (width_px, height_px), "white")
            draw = ImageDraw.Draw(img)
            
            margin = s(10)
            inner_width = width_px - (margin * 2)
            inner_height = height_px - (margin * 2)
            
            draw.rectangle([margin, margin, width_px - margin, height_px - margin], outline="#333333", width=max(1, s(2)))
            
            font_title = load_font(["arialbd.ttf", "Arial Bold.ttf"], s(14))
            font_text = load_font(["arial.ttf", "Arial.ttf"], s(11))
            font_small = load_font(["arial.ttf", "Arial.ttf"], s(9))
            font_bold = load_font(["arialbd.ttf", "Arial Bold.ttf"], s(12))
            
            y = margin + s(8)
            x = margin + s(8)
            line_height = s(18)
            
            draw.text((x, y), f"СВОДНЫЙ ОТЧЁТ ПО АДРЕСУ", fill="black", font=font_title)
            y += line_height + s(5)
            
            addr_lines = wrap_text(address, 34)
            for line in addr_lines:
                draw.text((x, y), line, fill="black", font=font_small)
                y += line_height - s(2)
            y += s(5)
            
            if all_products:
                client = all_products[0].get('Клиент', '')
                draw.text((x, y), f"Клиент: {client[:35]}", fill="black", font=font_text)
                y += line_height
            
            if all_products:
                rep = all_products[0].get('Торговый представитель', '')
                draw.text((x, y), f"Торг.пред: {rep[:35]}", fill="black", font=font_text)
                y += line_height + s(5)
            
            draw.line([(x, y), (x + inner_width - s(16), y)], fill="#cccccc", width=max(1, s(1)))
            y += s(10)
            
            page_text = f"Стр. {page_idx + 1} из {len(pages)}"
            draw.text((width_px - margin - s(60), margin + s(8)), page_text, fill="gray", font=font_small)
            
            draw.text((x, y), "№", fill="black", font=font_bold)
            draw.text((x + s(25), y), "Товар", fill="black", font=font_bold)
            draw.text((x + s(200), y), "Блоков", fill="black", font=font_bold)
            draw.text((x + s(260), y), "ШТ", fill="black", font=font_bold)
            y += line_height + s(3)
            draw.line([(x, y - s(5)), (x + inner_width - s(16), y - s(5))], fill="#000000", width=max(1, s(1)))
            
            total_blocks = 0
            total_shields = 0
            
            for idx, product in enumerate(page_products):
                product_name = product.get('Товары', '')[:22]
                blocks = product.get('Отсканировано', 0)
                shields = blocks * 10
                total_blocks += blocks
                total_shields += shields
                
                draw.text((x, y), f"{idx + 1 + (page_idx * products_per_page)}", fill="black", font=font_text)
                draw.text((x + s(25), y), product_name, fill="black", font=font_text)
                draw.text((x + s(205), y), str(blocks), fill="black", font=font_text)
                draw.text((x + s(265), y), str(shields), fill="black", font=font_text)
                y += line_height
                
                if y > height_px - s(60) and idx < len(page_products) - 1:
                    draw.text((x, height_px - margin - s(12)), datetime.now().strftime("%d.%m.%Y %H:%M"), fill="gray", font=font_small)
                    break
            
            draw.line([(x, y), (x + inner_width - s(16), y)], fill="#cccccc", width=max(1, s(1)))
            y += s(8)
            
            draw.text((x, y), f"ИТОГО:", fill="black", font=font_bold)
            draw.text((x + s(205), y), str(total_blocks), fill="black", font=font_bold)
            draw.text((x + s(265), y), str(total_shields), fill="black", font=font_bold)
            
            draw.text((x, height_px - margin - s(12)), datetime.now().strftime("%d.%m.%Y %H:%M"), fill="gray", font=font_small)
            
            temp_file = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
            img.save(temp_file.name, dpi=(LABEL_DPI, LABEL_DPI))
            temp_file.close()
            printed_files.append(temp_file.name)

            if not send_file_to_printer(temp_file.name):
                return None
        
        return printed_files
    except Exception as e:
        logging.exception("Ошибка печати сводного листа")
        return None

def write_scan_backup(action, order, code=None, codes=None):
    try:
        os.makedirs(BACKUP_DIR, exist_ok=True)
        filename = os.path.join(BACKUP_DIR, f"scan_backup_{datetime.now().strftime('%d.%m.%Y')}.jsonl")
        payload = {
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "action": action,
            "row_number": order.get("_row_number"),
            "date": order.get("Дата получения заказа", ""),
            "client": order.get("Клиент", ""),
            "address": order.get("Адрес", ""),
            "product": order.get("Товары", ""),
            "payment_type": order.get("Тип оплаты", ""),
            "code": code,
            "codes": codes or [],
        }
        with open(filename, "a", encoding="utf-8") as backup_file:
            backup_file.write(json.dumps(payload, ensure_ascii=False) + "\n")
        return True
    except Exception:
        logging.exception("Не удалось записать локальный backup")
        return False

def load_pending_prints():
    try:
        if not os.path.exists(PENDING_PRINTS_FILE):
            return []
        with open(PENDING_PRINTS_FILE, "r", encoding="utf-8") as pending_file:
            data = json.load(pending_file)
        if isinstance(data, list):
            return data
        return []
    except Exception:
        logging.exception("Не удалось загрузить очередь печати")
        return []

def save_pending_prints(items):
    try:
        with open(PENDING_PRINTS_FILE, "w", encoding="utf-8") as pending_file:
            json.dump(items, pending_file, ensure_ascii=False, indent=2)
        return True
    except Exception:
        logging.exception("Не удалось сохранить очередь печати")
        return False

def make_pending_print_id(address, products):
    payload = {
        "address": address,
        "products": [
            {
                "client": product.get("Клиент", ""),
                "address": product.get("Адрес", ""),
                "product": product.get("Товары", ""),
                "codes": product.get("Коды", []),
            }
            for product in products
        ],
    }
    raw = json.dumps(payload, ensure_ascii=False, sort_keys=True)
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()

def add_pending_print(address, products):
    pending = load_pending_prints()
    pending_id = make_pending_print_id(address, products)
    for item in pending:
        if item.get("id") == pending_id:
            return pending_id

    pending.append({
        "id": pending_id,
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "address": address,
        "products": products,
    })
    save_pending_prints(pending)
    return pending_id

def remove_pending_print(pending_id):
    pending = load_pending_prints()
    new_pending = [item for item in pending if item.get("id") != pending_id]
    if len(new_pending) != len(pending):
        save_pending_prints(new_pending)

class ScanningApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("📦 Система учёта сканирования блоков")
        self.configure(bg=BG_MAIN)
        self.geometry("1150x750")
        
        self.today_orders = []
        self.sheet = None
        self.all_existing_codes = set()
        self.current_legal_entity = None
        self.current_legal_entity_orders = []
        self.current_product_idx = 0
        self.current_order = None
        self.scanned_codes = []
        self.saved_codes_count = 0
        self.completed_orders = []
        self.current_legal_entity_products = []
        self.error_timer = None
        self.visible_order_groups = []
        self.current_group_key = None
        self.operation_in_progress = False
        os.makedirs(BACKUP_DIR, exist_ok=True)
        os.makedirs(REPORTS_DIR, exist_ok=True)
        
        self._build_ui()
        self.center_window()
        self.after(100, lambda: self.scan_entry.focus_set())
        self.after(150, lambda: self.refresh_from_sheet(initial=True))
        self.after(500, self.check_pending_prints)
        self.protocol("WM_DELETE_WINDOW", self.on_close)
    
    def load_data(self, show_empty_warning=True):
        self.today_orders, self.sheet, self.all_existing_codes = fetch_sheet_data()
        if show_empty_warning and not self.today_orders:
            messagebox.showwarning("Нет заданий", 
                f"На сегодня ({datetime.now().strftime('%d.%m.%Y')}) нет заданий.\n\n"
                f"Проверьте:\n"
                f"1. В таблице есть строка с этой датой\n"
                f"2. Столбец называется 'Дата получения заказа'\n"
                f"3. Колонка 'Отсканированные коды' пустая у активных заказов")

    def apply_loaded_data(self, result, show_empty_warning):
        self.today_orders, self.sheet, self.all_existing_codes = result
        if show_empty_warning and not self.today_orders:
            messagebox.showwarning(
                "Нет заданий",
                f"На сегодня ({datetime.now().strftime('%d.%m.%Y')}) нет заданий.\n\n"
                f"Проверьте:\n"
                f"1. В таблице есть строка с этой датой\n"
                f"2. Столбец называется 'Дата получения заказа'\n"
                f"3. Колонка 'Отсканированные коды' пустая или заполнена не полностью у активных заказов"
            )

    def run_background(self, title, work, on_success=None, on_error=None, on_finally=None):
        def worker():
            try:
                result = work()
            except Exception as exc:
                logging.exception(title)

                def fail(exc=exc):
                    try:
                        if on_error:
                            on_error(exc)
                        else:
                            self.show_critical_error(title, exc)
                    finally:
                        if on_finally:
                            on_finally()

                try:
                    self.after(0, fail)
                except tk.TclError:
                    pass
                return

            def done():
                try:
                    if on_success:
                        on_success(result)
                finally:
                    if on_finally:
                        on_finally()

            try:
                self.after(0, done)
            except tk.TclError:
                pass

        threading.Thread(target=worker, daemon=True).start()

    def set_busy(self, message):
        self.operation_in_progress = True
        self.status_var.set(message)
        self.status_label.config(bg=BG_MAIN, fg=FG_MUTED)

    def clear_busy(self):
        self.operation_in_progress = False
    
    def center_window(self):
        self.update_idletasks()
        x = (self.winfo_screenwidth() - self.winfo_width()) // 2
        y = (self.winfo_screenheight() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")
    
    def show_error(self, message, popup=True):
        logging.warning("Ошибка для пользователя: %s", message)
        self.status_var.set(f"❌ {message}")
        self.status_label.config(bg=ERROR_BG, fg=ERROR_FG)
        if self.error_timer:
            self.after_cancel(self.error_timer)
        self.error_timer = self.after(3000, self.clear_error)
        if popup:
            messagebox.showerror(
                "Ошибка",
                f"Причина: {message}\n\nЕсли ошибка повторяется, подробности будут в логе:\n{LOG_FILE}"
            )

    def show_critical_error(self, title, exc_or_message):
        if isinstance(exc_or_message, BaseException):
            message = str(exc_or_message)
            logging.error(
                title,
                exc_info=(type(exc_or_message), exc_or_message, exc_or_message.__traceback__)
            )
            detail = format_exception_message(title, exc_or_message)
        else:
            message = str(exc_or_message)
            logging.error("%s: %s", title, message)
            detail = f"{title}\n\nПричина: {message}\n\nПодробности записаны в лог:\n{LOG_FILE}"

        self.show_error(message, popup=False)
        messagebox.showerror("Ошибка", detail)

    def report_callback_exception(self, exc_type, exc_value, exc_traceback):
        logging.error(
            "Ошибка в интерфейсе",
            exc_info=(exc_type, exc_value, exc_traceback)
        )
        try:
            self.show_error(str(exc_value), popup=False)
            messagebox.showerror(
                "Ошибка",
                format_exception_message("Ошибка в интерфейсе", exc_value)
            )
        except Exception:
            pass
    
    def clear_error(self):
        self.status_var.set("✅ Готов к работе")
        self.status_label.config(bg=BG_MAIN, fg=FG_MUTED)
        self.error_timer = None
    
    def validate_code(self, code):
        if not code:
            return False, "Код пустой"
        if not code.startswith('01'):
            return False, "КИЗ должен начинаться с 01"
        if len(code) < KIZ_MIN_LENGTH:
            return False, f"Код слишком короткий для КИЗа (минимум {KIZ_MIN_LENGTH} символов)"
        if len(code) > KIZ_MAX_LENGTH:
            return False, f"Код слишком длинный для КИЗа (максимум {KIZ_MAX_LENGTH} символов)"
        if re.search(r'[а-яА-ЯёЁ]', code):
            return False, "Код содержит русские буквы! Используйте только латиницу"
        if re.search(r'\s', code):
            return False, "Код содержит пробелы или переносы"
        if not re.fullmatch(r'[\x1d\x21-\x7E]+', code):
            return False, "Код содержит недопустимые символы"
        return True, ""
    
    def undo_last_scan(self):
        if self.operation_in_progress:
            self.show_error("Дождитесь завершения текущей операции")
            return

        if not self.current_order:
            self.show_error("Нет активной позиции")
            return
        
        if not self.scanned_codes:
            self.show_error("Нет кодов для отмены")
            return

        if len(self.scanned_codes) <= self.saved_codes_count:
            self.show_error("Нельзя отменить коды, уже записанные в Google Sheets")
            return
        
        removed_code = self.scanned_codes.pop()
        self.all_existing_codes.discard(removed_code)
        write_scan_backup("undo_scan", self.current_order, code=removed_code, codes=self.scanned_codes.copy())
        
        plan_blocks = get_plan_blocks(self.current_order)
        
        scanned_count = len(self.scanned_codes)
        self.progress_label.config(text=f"{scanned_count} / {plan_blocks}")
        self.last_code_label.config(text=f"Отменён код: {removed_code[:40]}...")
        self.status_var.set(f"↩️ Отменён последний код ({scanned_count}/{plan_blocks})")
        
        if scanned_count < plan_blocks:
            self.next_product_btn.config(state="disabled")
            self.finish_btn.config(state="normal")
        
        self.scan_entry.focus_set()
    
    def _build_ui(self):
        main = tk.Frame(self, bg=BG_MAIN)
        main.pack(fill="both", expand=True, padx=25, pady=20)
        
        title = tk.Label(main, text="📦 УЧЁТ СКАНИРОВАНИЯ БЛОКОВ", 
                        bg=BG_MAIN, fg=FG_TEXT, font=("Segoe UI", 24, "bold"))
        title.pack(pady=(0, 5))
        
        date_label = tk.Label(main, text=f"Дата: {datetime.now().strftime('%d.%m.%Y')}", 
                             bg=BG_MAIN, fg=FG_MUTED, font=("Segoe UI", 12))
        date_label.pack(pady=(0, 20))
        
        content = tk.Frame(main, bg=BG_MAIN)
        content.pack(fill="both", expand=True)
        
        left_panel = tk.Frame(content, bg=BG_MAIN)
        left_panel.pack(side="left", fill="both", expand=True, padx=(0, 15))
        
        list_card = tk.Frame(left_panel, bg=BG_CARD, relief="flat", bd=1, highlightbackground=BORDER)
        list_card.pack(fill="both", expand=True)
        
        list_header = tk.Frame(list_card, bg=BG_CARD)
        list_header.pack(fill="x", padx=20, pady=(15, 10))

        tk.Label(list_header, text="🏢 Заказы на сегодня",
                bg=BG_CARD, fg=FG_TEXT, font=("Segoe UI", 14, "bold")).pack(side="left")

        self.refresh_btn = tk.Button(list_header, text="🔄 ОБНОВИТЬ",
                                     bg=INFO, fg="white", font=("Segoe UI", 9, "bold"),
                                     command=self.refresh_from_sheet, relief="flat", cursor="hand2")
        self.refresh_btn.pack(side="right")

        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self.refresh_legal_list())
        self.search_entry = tk.Entry(list_card, textvariable=self.search_var, bg=BG_MAIN, fg=FG_TEXT,
                                     font=("Segoe UI", 11), relief="flat", bd=1,
                                     highlightbackground=BORDER)
        self.search_entry.pack(fill="x", padx=15, pady=(0, 10))
        
        list_container = tk.Frame(list_card, bg=BG_CARD)
        list_container.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        
        self.legal_listbox = tk.Listbox(list_container, bg=BG_CARD, fg=FG_TEXT, 
                                        font=("Segoe UI", 11), selectmode=tk.SINGLE,
                                        relief="flat", selectbackground=ACCENT, selectforeground="white")
        self.legal_listbox.pack(side="left", fill="both", expand=True)
        
        scrollbar = tk.Scrollbar(list_container, orient="vertical", command=self.legal_listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self.legal_listbox.config(yscrollcommand=scrollbar.set)
        
        self.refresh_legal_list()
        
        self.select_btn = tk.Button(left_panel, text="✅ ВЫБРАТЬ ЗАКАЗ", 
                                   bg=ACCENT, fg="white", font=("Segoe UI", 12, "bold"),
                                   command=self.select_legal_entity, relief="flat", pady=12,
                                   cursor="hand2")
        self.select_btn.pack(pady=(15, 0), fill="x")
        
        right_panel = tk.Frame(content, bg=BG_MAIN)
        right_panel.pack(side="right", fill="both", expand=True, padx=(15, 0))
        
        info_card = tk.Frame(right_panel, bg=BG_CARD, relief="flat", bd=1, highlightbackground=BORDER)
        info_card.pack(fill="x", pady=(0, 15))
        
        tk.Label(info_card, text="📋 ТЕКУЩАЯ ПОЗИЦИЯ", 
                bg=BG_CARD, fg=ACCENT, font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=20, pady=(15, 10))
        
        self.current_info = tk.Label(info_card, text="Не выбрано", 
                                    bg=BG_CARD, fg=FG_TEXT, font=("Segoe UI", 11),
                                    wraplength=400, justify="left")
        self.current_info.pack(anchor="w", padx=20, pady=(0, 10))
        
        positions_frame = tk.Frame(info_card, bg=BG_CARD)
        positions_frame.pack(fill="x", padx=20, pady=(0, 10))
        
        self.position_label = tk.Label(positions_frame, text="", bg=BG_CARD, fg=WARNING, font=("Segoe UI", 11, "bold"))
        self.position_label.pack(side="left")
        
        progress_frame = tk.Frame(info_card, bg=BG_CARD)
        progress_frame.pack(fill="x", padx=20, pady=(0, 15))
        
        tk.Label(progress_frame, text="Сканирование:", bg=BG_CARD, fg=FG_MUTED, font=("Segoe UI", 11)).pack(side="left")
        self.progress_label = tk.Label(progress_frame, text="0 / 0", bg=BG_CARD, fg=SUCCESS, font=("Segoe UI", 14, "bold"))
        self.progress_label.pack(side="left", padx=(10, 0))
        
        scan_card = tk.Frame(right_panel, bg=BG_CARD, relief="flat", bd=1, highlightbackground=BORDER)
        scan_card.pack(fill="x", pady=(0, 15))
        
        tk.Label(scan_card, text="🔍 СКАНИРОВАНИЕ КОДА", 
                bg=BG_CARD, fg=ACCENT, font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=20, pady=(15, 10))
        
        self.scan_entry = tk.Entry(scan_card, bg=BG_MAIN, fg=FG_TEXT, font=("Segoe UI", 14),
                                   relief="flat", bd=1, highlightbackground=BORDER)
        self.scan_entry.pack(fill="x", padx=20, pady=(0, 10))
        self.scan_entry.bind("<Return>", self.on_scan)
        
        self.last_code_label = tk.Label(scan_card, text="", bg=BG_CARD, fg=SUCCESS, font=("Segoe UI", 10))
        self.last_code_label.pack(anchor="w", padx=20, pady=(5, 5))
        
        actions_frame = tk.Frame(right_panel, bg=BG_MAIN)
        actions_frame.pack(fill="x", pady=(0, 15))
        
        self.undo_btn = tk.Button(actions_frame, text="↩️ ОТМЕНИТЬ ПОСЛЕДНИЙ КОД", 
                                 bg=DANGER, fg="white", font=("Segoe UI", 10, "bold"),
                                 command=self.undo_last_scan, relief="flat", state="disabled",
                                 cursor="hand2")
        self.undo_btn.pack(side="left", fill="x", expand=True, padx=(0, 10), pady=5)
        
        self.next_product_btn = tk.Button(actions_frame, text="➡️ СЛЕДУЮЩАЯ ПОЗИЦИЯ", 
                                         bg=WARNING, fg="white", font=("Segoe UI", 11, "bold"),
                                         command=self.next_product, relief="flat", state="disabled",
                                         cursor="hand2")
        self.next_product_btn.pack(side="left", fill="x", expand=True, padx=(0, 10), pady=5)
        
        self.finish_btn = tk.Button(actions_frame, text="🏁 ЗАВЕРШИТЬ ЗАКАЗ", 
                                   bg=SUCCESS, fg="white", font=("Segoe UI", 11, "bold"),
                                   command=self.finish_legal_entity, relief="flat", state="disabled",
                                   cursor="hand2")
        self.finish_btn.pack(side="right", fill="x", expand=True, padx=(10, 0), pady=5)
        
        stats_card = tk.Frame(right_panel, bg=BG_CARD, relief="flat", bd=1, highlightbackground=BORDER)
        stats_card.pack(fill="x")
        
        tk.Label(stats_card, text="📊 СТАТИСТИКА", 
                bg=BG_CARD, fg=ACCENT, font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=20, pady=(15, 10))
        
        stats_frame = tk.Frame(stats_card, bg=BG_CARD)
        stats_frame.pack(fill="x", padx=20, pady=(0, 15))
        
        self.completed_count_label = tk.Label(stats_frame, text="Выполнено: 0", bg=BG_CARD, fg=FG_TEXT, font=("Segoe UI", 11))
        self.completed_count_label.pack(side="left", padx=(0, 20))
        
        self.total_blocks_label = tk.Label(stats_frame, text="Блоков: 0", bg=BG_CARD, fg=FG_TEXT, font=("Segoe UI", 11))
        self.total_blocks_label.pack(side="left")
        
        self.report_btn = tk.Button(right_panel, text="📊 ЗАВЕРШИТЬ ДЕНЬ (ОТЧЁТ)", 
                                   bg=INFO, fg="white", font=("Segoe UI", 11, "bold"),
                                   command=self.end_day, relief="flat", pady=10,
                                   cursor="hand2")
        self.report_btn.pack(fill="x", pady=(10, 0))
        
        status_frame = tk.Frame(main, bg=BG_MAIN)
        status_frame.pack(fill="x", pady=(20, 0))
        
        self.status_var = tk.StringVar(value="✅ Готов к работе")
        self.status_label = tk.Label(status_frame, textvariable=self.status_var, 
                                     bg=BG_MAIN, fg=FG_MUTED, font=("Segoe UI", 10))
        self.status_label.pack()
    
    def refresh_legal_list(self):
        self.legal_listbox.delete(0, tk.END)
        self.visible_order_groups = []
        grouped_orders = {}
        search_text = normalize_text(self.search_var.get()).lower() if hasattr(self, "search_var") else ""

        for order in self.today_orders:
            key = order_group_key(order)
            client, address = key
            client = client or "Клиент не указан"
            address = address or "Адрес не указан"
            if search_text and search_text not in client.lower():
                continue
            grouped_orders.setdefault((client, address), []).append(order)
        
        for key in sorted(grouped_orders.keys(), key=lambda item: (item[0].lower(), item[1].lower())):
            client, address = key
            count = len(grouped_orders[key])
            self.visible_order_groups.append(key)
            self.legal_listbox.insert(tk.END, f"{client} | {address} | {count} поз.")

    def reset_current_selection(self):
        self.current_legal_entity = None
        self.current_group_key = None
        self.current_legal_entity_orders = []
        self.current_product_idx = 0
        self.current_order = None
        self.scanned_codes = []
        self.saved_codes_count = 0
        self.current_legal_entity_products = []
        self.current_info.config(text="Не выбрано")
        self.position_label.config(text="")
        self.progress_label.config(text="0 / 0")
        self.next_product_btn.config(state="disabled")
        self.finish_btn.config(state="disabled")
        self.undo_btn.config(state="disabled")
        self.last_code_label.config(text="")

    def refresh_from_sheet(self, initial=False):
        if self.operation_in_progress:
            self.show_error("Дождитесь завершения текущей операции")
            return

        if self.current_order and not initial:
            if not messagebox.askyesno(
                "Обновить список?",
                "Есть выбранный заказ. Обновление сбросит текущий выбор и несохранённые сканы.\n\nПродолжить?"
            ):
                return

        self.set_busy("⏳ Обновляю список заказов...")
        self.refresh_btn.config(state="disabled")

        def on_success(result):
            self.apply_loaded_data(result, show_empty_warning=initial)
            self.reset_current_selection()
            self.refresh_legal_list()
            self.status_var.set("✅ Список заказов обновлён")
            self.status_label.config(bg=BG_MAIN, fg=FG_MUTED)

        def on_error(exc):
            self.show_critical_error("Не удалось обновить список заказов", exc)

        def on_finally():
            self.refresh_btn.config(state="normal")
            self.clear_busy()

        self.run_background(
            "Не удалось обновить список заказов",
            fetch_sheet_data,
            on_success=on_success,
            on_error=on_error,
            on_finally=on_finally
        )
    
    def select_legal_entity(self):
        if self.operation_in_progress:
            self.show_error("Дождитесь завершения текущей операции")
            return

        if not self.today_orders:
            messagebox.showwarning("Ошибка", "Нет доступных юридических лиц!")
            return
        
        selection = self.legal_listbox.curselection()
        if not selection:
            messagebox.showwarning("Ошибка", "Выберите заказ из списка")
            return
        
        if selection[0] >= len(self.visible_order_groups):
            messagebox.showwarning("Ошибка", "Выбранный заказ не найден в списке")
            return

        selected_group = self.visible_order_groups[selection[0]]
        legal_entity, address = selected_group

        self.current_legal_entity = legal_entity
        self.current_group_key = selected_group
        self.current_legal_entity_orders = [
            o for o in self.today_orders
            if order_group_key(o) == selected_group
        ]
        self.current_legal_entity_orders.sort(key=lambda order: parse_int_value(order.get("_row_number")))
        self.current_product_idx = 0
        self.scanned_codes = []
        self.current_legal_entity_products = []
        
        self.load_current_product()
        
        self.status_var.set(f"✅ Выбран заказ: {legal_entity} | {address}")
        self.scan_entry.focus_set()
    
    def load_current_product(self):
        if self.current_product_idx >= len(self.current_legal_entity_orders):
            return
        
        self.current_order = self.current_legal_entity_orders[self.current_product_idx]
        
        plan_blocks = get_plan_blocks(self.current_order)
        
        info_text = f"""🏢 Юр.лицо: {self.current_order.get('Клиент', '')}
👤 Торг.пред: {self.current_order.get('Торговый представитель', '')}
📍 Адрес: {self.current_order.get('Адрес', 'Адрес не указан')}
📦 Товар: {self.current_order.get('Товары', '')}
💳 Тип оплаты: {self.current_order.get('Тип оплаты', '')}
📦 План: {plan_blocks} блоков (1 блок = 10 ШТ)"""
        
        self.current_info.config(text=info_text)
        
        total_products = len(self.current_legal_entity_orders)
        self.position_label.config(text=f"Позиция {self.current_product_idx + 1} из {total_products}")
        
        existing_codes = self.current_order.get("_existing_scanned_codes", [])
        self.scanned_codes = existing_codes.copy()
        self.saved_codes_count = len(existing_codes)
        self.progress_label.config(text=f"{len(self.scanned_codes)} / {plan_blocks}")
        self.next_product_btn.config(state="disabled")
        self.finish_btn.config(state="normal")
        self.undo_btn.config(state="normal")
        self.scan_entry.delete(0, tk.END)
        if existing_codes:
            self.last_code_label.config(text=f"Уже записано в таблице: {len(existing_codes)} кодов")
        else:
            self.last_code_label.config(text="")
        if plan_blocks > 0 and len(self.scanned_codes) >= plan_blocks:
            self.next_product_btn.config(state="normal")
            self.finish_btn.config(state="disabled")
        self.scan_entry.focus_set()
    
    def on_scan(self, event=None):
        if self.operation_in_progress:
            self.show_error("Дождитесь завершения текущей операции")
            self.scan_entry.delete(0, tk.END)
            return

        if not self.current_order:
            messagebox.showwarning("Ошибка", "Сначала выберите заказ")
            return
        
        code = self.scan_entry.get().strip()
        if not code:
            return
        
        is_valid, error_msg = self.validate_code(code)
        if not is_valid:
            self.show_error(error_msg)
            self.scan_entry.delete(0, tk.END)
            return
        
        plan_blocks = get_plan_blocks(self.current_order)
        if plan_blocks <= 0:
            self.show_error("В заказе не указано корректное 'Кол-во блок'")
            self.scan_entry.delete(0, tk.END)
            return
        
        if len(self.scanned_codes) >= plan_blocks:
            self.show_error(f"План выполнен! Нельзя сканировать больше {plan_blocks} блоков")
            self.scan_entry.delete(0, tk.END)
            return
        
        if code in self.scanned_codes:
            self.show_error("Код уже отсканирован в этой позиции")
            self.scan_entry.delete(0, tk.END)
            return
        
        if code in self.all_existing_codes:
            self.show_error(f"Код {code[:20]}... уже существует в Google Sheets!")
            self.scan_entry.delete(0, tk.END)
            return
        
        for completed in self.completed_orders:
            if code in completed.get("Коды", []):
                self.show_error("Код уже использован в другом задании сегодня")
                self.scan_entry.delete(0, tk.END)
                return

        if not write_scan_backup("scan", self.current_order, code=code, codes=self.scanned_codes + [code]):
            self.show_error("Не удалось сохранить локальный backup. Код не принят")
            self.scan_entry.delete(0, tk.END)
            return
        
        self.scanned_codes.append(code)
        self.all_existing_codes.add(code)
        scanned_count = len(self.scanned_codes)
        
        self.progress_label.config(text=f"{scanned_count} / {plan_blocks}")
        self.last_code_label.config(text=f"Последний код: {code[:40]}...")
        self.status_var.set(f"✅ Отсканирован код ({scanned_count}/{plan_blocks})")
        self.status_label.config(bg=BG_MAIN, fg=FG_MUTED)
        self.scan_entry.delete(0, tk.END)
        
        if scanned_count >= plan_blocks:
            self.status_var.set(f"🎯 Позиция выполнена! Нажмите 'Следующая позиция'")
            self.next_product_btn.config(state="normal")
            self.finish_btn.config(state="disabled")
        
        self.scan_entry.focus_set()
    
    def next_product(self):
        if self.operation_in_progress:
            self.show_error("Дождитесь завершения текущей операции")
            return

        if not self.current_order:
            return
        
        plan_blocks = get_plan_blocks(self.current_order)
	        
        scanned_count = len(self.scanned_codes)
	        
        if scanned_count != plan_blocks:
            self.show_error(f"Отсканировано {scanned_count} из {plan_blocks} блоков. Завершите позицию!")
            return

        if not self.sheet:
            self.show_error("Нет подключения к Google Sheets. Позиция не завершена")
            return

        order = self.current_order
        scanned_codes = self.scanned_codes.copy()
        self.set_busy("⏳ Сохраняю КИЗы в Google Sheets...")
        self.next_product_btn.config(state="disabled")
        self.finish_btn.config(state="disabled")

        def work():
            ok, message = update_scanned_codes_to_gsheet(self.sheet, order, scanned_codes)
            if not ok:
                raise RuntimeError(message)
            if not write_scan_backup("position_saved", order, codes=scanned_codes):
                raise RuntimeError("Коды записаны в Google Sheets, но локальный backup позиции не создан")
            return True

        def on_success(_):
            product_result = {
                "Клиент": order.get('Клиент', ''),
                "Адрес": order.get('Адрес', ''),
                "Торговый представитель": order.get('Торговый представитель', ''),
                "Товары": order.get('Товары', ''),
                "Тип оплаты": order.get('Тип оплаты', ''),
                "План": plan_blocks,
                "Отсканировано": scanned_count,
                "Коды": scanned_codes.copy()
            }
            self.current_legal_entity_products.append(product_result)

            completed_result = product_result.copy()
            completed_result["Кол-во ШТ в блоке"] = 10
            completed_result["План блоков"] = plan_blocks
            self.completed_orders.append(completed_result)

            self.current_product_idx += 1
            self.clear_busy()

            if self.current_product_idx < len(self.current_legal_entity_orders):
                self.load_current_product()
                self.status_var.set("✅ Позиция сохранена")
                self.status_label.config(bg=BG_MAIN, fg=FG_MUTED)
            else:
                self.finish_legal_entity(from_next_product=True)

        def on_error(exc):
            self.show_critical_error("КИЗы не записаны", exc)
            self.next_product_btn.config(state="normal")
            self.clear_busy()

        self.run_background(
            "Не удалось сохранить позицию",
            work,
            on_success=on_success,
            on_error=on_error
        )
    
    def finish_legal_entity(self, from_next_product=False):
        if self.operation_in_progress:
            self.show_error("Дождитесь завершения текущей операции")
            return

        if not self.current_legal_entity:
            return
        
        if self.current_product_idx < len(self.current_legal_entity_orders):
            self.show_error("Сначала завершите все позиции по заказу!")
            return
        
        if not self.current_legal_entity_products:
            self.show_error("Нет завершённых позиций по заказу!")
            return
        
        if not self.confirm_print_settings():
            self.show_error("Печать сводного листа отменена")
            self.finish_btn.config(state="normal")
            return

        group_key = self.current_group_key
        current_products = [product.copy() for product in self.current_legal_entity_products]
        self.set_busy("⏳ Готовлю и печатаю сводный лист...")
        self.finish_btn.config(state="disabled")
        self.next_product_btn.config(state="disabled")

        def work():
            first_product = current_products[0]
            address = first_product.get('Адрес', 'Адрес не указан')
            summary_products = current_products

            if self.sheet:
                sheet_products = build_summary_products_from_gsheet(
                    self.sheet,
                    group_key or order_group_key(first_product)
                )
                if sheet_products:
                    summary_products = sheet_products
                    first_product = summary_products[0]
                    address = first_product.get('Адрес', address)

            pending_print_id = add_pending_print(address, summary_products)

            printed_files = print_summary(address, summary_products)
            if not printed_files:
                raise RuntimeError("Сводочный лист не создан или не отправлен на печать")

            remove_pending_print(pending_print_id)

            if not write_scan_backup(
                "address_finished",
                first_product,
                codes=[code for product in summary_products for code in product.get("Коды", [])]
            ):
                raise RuntimeError("Сводка напечатана, но backup завершения заказа не создан")

            return {
                "first_product": first_product,
                "summary_products": summary_products,
                "finished_group": group_key or order_group_key(first_product),
            }

        def on_success(result):
            self.update_stats_display()

            finished_group = result["finished_group"]
            self.today_orders = [o for o in self.today_orders if order_group_key(o) != finished_group]
            self.refresh_legal_list()

            self.reset_current_selection()
            self.status_var.set("✅ Заказ завершён! Сводка отправлена на печать")
            self.status_label.config(bg=BG_MAIN, fg=FG_MUTED)

            if self.legal_listbox.size() > 0:
                self.legal_listbox.selection_set(0)

        def on_error(exc):
            self.show_critical_error("Не удалось завершить заказ", exc)
            self.finish_btn.config(state="normal")

        def on_finally():
            self.clear_busy()

        self.run_background(
            "Не удалось завершить заказ",
            work,
            on_success=on_success,
            on_error=on_error,
            on_finally=on_finally
        )
	    
    def confirm_print_settings(self):
        result = {"print": False}
        dialog = tk.Toplevel(self)
        dialog.title("Параметры печати")
        dialog.configure(bg=BG_MAIN)
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()

        container = tk.Frame(dialog, bg=BG_CARD, padx=24, pady=20)
        container.pack(fill="both", expand=True, padx=16, pady=16)

        tk.Label(
            container,
            text="Печать сводного листа",
            bg=BG_CARD,
            fg=FG_TEXT,
            font=("Segoe UI", 14, "bold")
        ).pack(anchor="w", pady=(0, 12))

        rows = [
            ("Принтер", "Термопринтер"),
            ("Размер листа", f"{LABEL_WIDTH_MM} x {LABEL_HEIGHT_MM} мм"),
            ("Масштаб", "100%"),
            ("Ориентация", "Квадратная этикетка"),
            ("Разрешение макета", f"{LABEL_DPI} DPI"),
        ]

        for label, value in rows:
            row = tk.Frame(container, bg=BG_CARD)
            row.pack(fill="x", pady=3)
            tk.Label(row, text=label + ":", bg=BG_CARD, fg=FG_MUTED, font=("Segoe UI", 10), width=18, anchor="w").pack(side="left")
            tk.Label(row, text=value, bg=BG_CARD, fg=FG_TEXT, font=("Segoe UI", 10, "bold"), anchor="w").pack(side="left")

        actions = tk.Frame(container, bg=BG_CARD)
        actions.pack(fill="x", pady=(18, 0))

        def confirm():
            result["print"] = True
            dialog.destroy()

        def cancel():
            dialog.destroy()

        tk.Button(
            actions,
            text="ПЕЧАТАТЬ",
            bg=SUCCESS,
            fg="white",
            font=("Segoe UI", 10, "bold"),
            relief="flat",
            padx=18,
            pady=8,
            command=confirm,
            cursor="hand2"
        ).pack(side="right", padx=(8, 0))

        tk.Button(
            actions,
            text="ОТМЕНА",
            bg=FG_MUTED,
            fg="white",
            font=("Segoe UI", 10, "bold"),
            relief="flat",
            padx=18,
            pady=8,
            command=cancel,
            cursor="hand2"
        ).pack(side="right")

        dialog.protocol("WM_DELETE_WINDOW", cancel)
        self.update_idletasks()
        dialog.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - dialog.winfo_width()) // 2
        y = self.winfo_y() + (self.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f"+{x}+{y}")
        self.wait_window(dialog)
        return result["print"]

    def check_pending_prints(self):
        if self.operation_in_progress:
            self.after(1000, self.check_pending_prints)
            return

        pending = load_pending_prints()
        if not pending:
            return

        if not messagebox.askyesno(
            "Непечатанные сводки",
            f"Найдено непечатанных сводок: {len(pending)}.\n\nНапечатать сейчас?"
        ):
            return

        if not self.confirm_print_settings():
            return

        self.set_busy("⏳ Печатаю сводки из очереди...")

        def work():
            printed_count = 0
            for item in pending[:]:
                printed_files = print_summary(item.get("address", "Адрес не указан"), item.get("products", []))
                if not printed_files:
                    raise RuntimeError("Не удалось напечатать сводку из очереди")
                remove_pending_print(item.get("id"))
                printed_count += 1
            return printed_count

        def on_success(printed_count):
            self.status_var.set(f"✅ Напечатано сводок из очереди: {printed_count}")
            self.status_label.config(bg=BG_MAIN, fg=FG_MUTED)

        def on_error(exc):
            self.show_critical_error("Не удалось напечатать сводки из очереди", exc)

        def on_finally():
            self.clear_busy()

        self.run_background(
            "Не удалось напечатать сводки из очереди",
            work,
            on_success=on_success,
            on_error=on_error,
            on_finally=on_finally
        )

    def update_stats_display(self):
        completed = len(self.completed_orders)
        total_blocks = sum(o.get("Отсканировано", 0) for o in self.completed_orders)
        self.completed_count_label.config(text=f"Выполнено: {completed}")
        self.total_blocks_label.config(text=f"Блоков: {total_blocks}")
    
    def end_day(self):
        if self.operation_in_progress:
            self.show_error("Дождитесь завершения текущей операции")
            return

        if self.current_legal_entity:
            if not messagebox.askyesno("Внимание", "У вас есть незавершённый заказ!\n\nЗавершить день без сохранения текущего заказа?"):
                return
        
        self.set_busy("⏳ Формирую Excel-отчёт за день...")
        self.report_btn.config(state="disabled")

        def work():
            import pandas as pd

            sheet = self.sheet
            if not sheet:
                client = get_google_client()
                sheet = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)

            report_rows = build_day_report_rows_from_gsheet(sheet)
            terminal_rows = report_rows["terminal"]
            transfer_rows = report_rows["transfer"]
            unknown_rows = report_rows["unknown"]
            total_report_rows = len(terminal_rows) + len(transfer_rows) + len(unknown_rows)

            if total_report_rows == 0:
                return {
                    "empty": True,
                    "sheet": sheet,
                }
            
            filename = os.path.join(REPORTS_DIR, f"scan_report_{datetime.now().strftime('%d.%m.%Y_%H%M%S')}.xlsx")
            
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                if terminal_rows:
                    pd.DataFrame(terminal_rows).to_excel(writer, sheet_name="Терминал", index=False)
                else:
                    pd.DataFrame({"Сообщение": ["Нет данных"]}).to_excel(writer, sheet_name="Терминал", index=False)
                
                if transfer_rows:
                    pd.DataFrame(transfer_rows).to_excel(writer, sheet_name="Перечисление", index=False)
                else:
                    pd.DataFrame({"Сообщение": ["Нет данных"]}).to_excel(writer, sheet_name="Перечисление", index=False)

                if unknown_rows:
                    pd.DataFrame(unknown_rows).to_excel(writer, sheet_name="Не распознано", index=False)
            
            return {
                "empty": False,
                "sheet": sheet,
                "filename": filename,
                "total_report_rows": total_report_rows,
                "terminal_count": len(terminal_rows),
                "transfer_count": len(transfer_rows),
                "unknown_count": len(unknown_rows),
            }

        def on_success(result):
            self.sheet = result.get("sheet") or self.sheet
            if result.get("empty"):
                messagebox.showwarning("Нет данных", "За сегодня в Google Sheets нет отсканированных КИЗов для отчёта")
                return

            total_report_rows = result["total_report_rows"]
            messagebox.showinfo("Отчёт сохранён", 
                f"📊 Отчёт сохранён: {result['filename']}\n\n"
                f"━━━━━━━━━━━━━━━━━━━━\n"
                f"✅ Строк КИЗов: {total_report_rows}\n"
                f"📦 Блоков: {total_report_rows}\n"
                f"🔢 Кодов: {total_report_rows}\n"
                f"├─ Терминал: {result['terminal_count']} кодов\n"
                f"├─ Перечисление: {result['transfer_count']} кодов\n"
                f"└─ Не распознано: {result['unknown_count']} кодов\n"
                f"━━━━━━━━━━━━━━━━━━━━")
            
            self.on_close()

        def on_error(exc):
            if isinstance(exc, ImportError):
                self.show_critical_error("Не установлены зависимости для Excel-отчёта", "Установите pandas и openpyxl:\npip install pandas openpyxl")
            else:
                self.show_critical_error("Не удалось сохранить Excel-отчёт", exc)

        def on_finally():
            try:
                if self.winfo_exists():
                    self.report_btn.config(state="normal")
                    self.clear_busy()
            except tk.TclError:
                pass

        self.run_background(
            "Не удалось сохранить Excel-отчёт",
            work,
            on_success=on_success,
            on_error=on_error,
            on_finally=on_finally
        )
    
    def on_close(self):
        if self.current_order and len(self.scanned_codes) > self.saved_codes_count:
            if not messagebox.askyesno(
                "Закрыть программу?",
                "Есть несохранённые сканы по текущей позиции.\n\nЗакрыть программу без завершения позиции?"
            ):
                return
        self.destroy()

if __name__ == "__main__":
    if not os.path.exists(CREDENTIALS_FILE):
        messagebox.showerror("Ошибка", 
            f"Файл {CREDENTIALS_FILE} не найден!\n\n"
            f"Положите файл с учётными данными Google Sheets в папку с программой")
    else:
        app = ScanningApp()
        app.mainloop()
