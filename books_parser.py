"""
Парсер каталога books.toscrape.com (учебный сайт для скрапинга).

Собирает книги — название, цена, ссылка, рейтинг, наличие — по всем страницам
каталога или выбранной категории и сохраняет в Excel (.xlsx) и CSV.

Примеры:
    python books_parser.py                                # весь каталог → books.xlsx + books.csv
    python books_parser.py --category travel              # одна категория
    python books_parser.py -c "Science Fiction" -p 2 -o out/scifi
    python books_parser.py --list-categories
"""

from __future__ import annotations

import argparse
import csv
import logging
import sys
import time
from dataclasses import asdict, dataclass, fields
from pathlib import Path
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

BASE_URL = "https://books.toscrape.com/"
RATINGS = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}
HEADERS = {"User-Agent": "Mozilla/5.0 (demo-books-parser; +https://books.toscrape.com)"}

log = logging.getLogger("books_parser")


@dataclass
class Book:
    title: str
    price: float
    currency: str
    rating: int
    in_stock: bool
    url: str


COLUMNS = {
    "title": "Название",
    "price": "Цена",
    "currency": "Валюта",
    "rating": "Рейтинг",
    "in_stock": "В наличии",
    "url": "Ссылка",
}


# ───────────── сеть ─────────────

def make_session(retries: int = 4, backoff: float = 1.0) -> requests.Session:
    """Сессия с повтором запросов при обрывах и ответах 429/5xx (1 с, 2 с, 4 с…)."""
    session = requests.Session()
    session.headers.update(HEADERS)
    retry = Retry(
        total=retries,
        backoff_factor=backoff,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=("GET",),
        respect_retry_after_header=True,
    )
    session.mount("https://", HTTPAdapter(max_retries=retry))
    session.mount("http://", HTTPAdapter(max_retries=retry))
    return session


def get_soup(session: requests.Session, url: str) -> BeautifulSoup:
    log.debug("GET %s", url)
    response = session.get(url, timeout=20)
    response.raise_for_status()
    response.encoding = "utf-8"  # сайт отдаёт UTF-8, но не всегда пишет это в заголовке
    return BeautifulSoup(response.text, "html.parser")


# ───────────── разбор страниц ─────────────

def parse_categories(soup: BeautifulSoup) -> dict[str, str]:
    """{'travel': 'https://…/travel_2/index.html', …} — ключи в нижнем регистре."""
    result = {}
    for link in soup.select(".side_categories ul li ul li a"):
        name = " ".join(link.get_text().split())
        result[name.lower()] = urljoin(BASE_URL, link["href"])
    return result


def parse_price(text: str) -> tuple[float, str]:
    """'£51.77' → (51.77, 'GBP')."""
    text = text.strip().replace("Â", "")
    currency = {"£": "GBP", "$": "USD", "€": "EUR"}.get(text[:1], "")
    number = text[1:] if currency else text
    return float(number.replace(",", "")), currency


def parse_books(soup: BeautifulSoup, page_url: str) -> list[Book]:
    books = []
    for card in soup.select("article.product_pod"):
        link = card.select_one("h3 a")
        price, currency = parse_price(card.select_one(".price_color").get_text())
        rating_classes = card.select_one("p.star-rating")["class"]
        rating = next((RATINGS[c] for c in rating_classes if c in RATINGS), 0)
        in_stock = "in stock" in card.select_one(".availability").get_text().lower()
        books.append(Book(
            title=link.get("title") or link.get_text(strip=True),
            price=price,
            currency=currency,
            rating=rating,
            in_stock=in_stock,
            url=urljoin(page_url, link["href"]),
        ))
    return books


def next_page(soup: BeautifulSoup, page_url: str) -> str | None:
    link = soup.select_one("li.next a")
    return urljoin(page_url, link["href"]) if link else None


def scrape(session: requests.Session, start_url: str, max_pages: int = 0, delay: float = 0.3) -> list[Book]:
    """Идёт по кнопке «next», пока страницы не кончатся или не достигнут лимит."""
    books: list[Book] = []
    url: str | None = start_url
    page = 0
    while url and (not max_pages or page < max_pages):
        page += 1
        try:
            soup = get_soup(session, url)
        except requests.RequestException as error:
            log.error("Страница %d не загрузилась после повторов: %s — останавливаюсь", page, error)
            break
        found = parse_books(soup, url)
        books.extend(found)
        log.info("Страница %d: %d книг (всего %d)", page, len(found), len(books))
        url = next_page(soup, url)
        if url and delay:
            time.sleep(delay)  # вежливая пауза, чтобы не нагружать сайт
    return books


# ───────────── сохранение ─────────────

def save_csv(books: list[Book], path: Path) -> None:
    with path.open("w", newline="", encoding="utf-8-sig") as f:  # utf-8-sig — чтобы Excel открыл без кракозябр
        writer = csv.DictWriter(f, fieldnames=list(COLUMNS), delimiter=";")
        writer.writerow(COLUMNS)
        for book in books:
            writer.writerow(asdict(book))


def save_xlsx(books: list[Book], path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Книги"
    ws.append(list(COLUMNS.values()))
    for book in books:
        row = asdict(book)
        row["in_stock"] = "да" if book.in_stock else "нет"
        ws.append([row[key] for key in COLUMNS])

    header_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 22

    keys = [f.name for f in fields(Book)]
    price_col = keys.index("price") + 1
    url_col = keys.index("url") + 1
    for row in ws.iter_rows(min_row=2):
        row[price_col - 1].number_format = "0.00"
        link_cell = row[url_col - 1]
        link_cell.hyperlink = link_cell.value
        link_cell.font = Font(color="0563C1", underline="single")

    widths = {"title": 60, "price": 10, "currency": 9, "rating": 10, "in_stock": 11, "url": 70}
    for i, key in enumerate(keys, start=1):
        ws.column_dimensions[get_column_letter(i)].width = widths[key]
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(path)


# ───────────── запуск ─────────────

def setup_logging(log_file: Path, verbose: bool) -> None:
    log_file.parent.mkdir(parents=True, exist_ok=True)
    fmt = logging.Formatter("%(asctime)s %(levelname)-7s %(message)s", "%Y-%m-%d %H:%M:%S")
    console = logging.StreamHandler(sys.stdout)
    console.setFormatter(fmt)
    file = logging.FileHandler(log_file, encoding="utf-8")
    file.setFormatter(fmt)
    log.setLevel(logging.DEBUG if verbose else logging.INFO)
    log.addHandler(console)
    log.addHandler(file)


def build_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Парсер книг с books.toscrape.com → Excel + CSV")
    p.add_argument("-c", "--category", help="категория, например travel или \"Science Fiction\" (по умолчанию весь каталог)")
    p.add_argument("-p", "--pages", type=int, default=0, help="лимит страниц (0 — все)")
    p.add_argument("-o", "--output", default="books", help="имя файла без расширения (по умолчанию books)")
    p.add_argument("--delay", type=float, default=0.3, help="пауза между страницами, сек")
    p.add_argument("--log", default="logs/parser.log", help="файл лога")
    p.add_argument("--list-categories", action="store_true", help="показать категории и выйти")
    p.add_argument("-v", "--verbose", action="store_true", help="подробный лог")
    args = p.parse_args(argv)
    if args.pages < 0:
        p.error("--pages не может быть отрицательным")
    return args


def main(argv: list[str] | None = None) -> int:
    args = build_args(argv)
    setup_logging(Path(args.log), args.verbose)
    session = make_session()

    try:
        home = get_soup(session, BASE_URL)
    except requests.RequestException as error:
        log.error("Сайт недоступен: %s", error)
        return 1
    categories = parse_categories(home)

    if args.list_categories:
        print("\n".join(sorted(categories)))
        return 0

    start_url = urljoin(BASE_URL, "catalogue/page-1.html")
    if args.category:
        start_url = categories.get(args.category.strip().lower())
        if not start_url:
            log.error("Категория «%s» не найдена. Список: --list-categories", args.category)
            return 2

    log.info("Старт: %s, лимит страниц: %s", args.category or "весь каталог", args.pages or "нет")
    started = time.monotonic()
    books = scrape(session, start_url, args.pages, args.delay)
    if not books:
        log.error("Ничего не собрано")
        return 1

    out = Path(args.output).with_suffix("")
    out.parent.mkdir(parents=True, exist_ok=True)
    save_xlsx(books, out.with_suffix(".xlsx"))
    save_csv(books, out.with_suffix(".csv"))
    log.info("Готово: %d книг за %.1f с → %s, %s", len(books), time.monotonic() - started,
             out.with_suffix(".xlsx"), out.with_suffix(".csv"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
