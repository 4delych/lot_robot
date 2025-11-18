import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import requests
from bs4 import BeautifulSoup
import pandas as pd
import threading
import logging
import time
from urllib.parse import urljoin, urlparse, parse_qs
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import re
import json

# Configuration
CONFIG = {
    "BASE_URL": "https://zakupki.gov.ru",
    "RESULTS_PER_PAGE": 20,
    "REQUEST_TIMEOUT": 30,  # Увеличил таймаут
    "REQUEST_DELAY": 1,
    "MAX_RETRIES": 3,
    "USER_AGENT": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
}

# Search filter options
PURCHASE_STAGES = {
    "": "Все этапы",
    "SUBMISSION": "Подача заявок",
    "EVALUATION": "Работа комиссии",
    "COMPLETED": "Закупка завершена",
    "CANCELLED": "Закупка отменена",
}

LAWS = {"": "Все законы", "44FZ": "44-ФЗ", "223FZ": "223-ФЗ", "PP615": "ПП РФ 615"}

# Setup logging
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)


class ProcurementSearcher:
    """Handle web scraping operations with proper error handling and session management."""

    def __init__(self):
        self.session = self._create_session()

    def _create_session(self):
        """Create a requests session with retry strategy and proper headers."""
        session = requests.Session()

        # Setup retry strategy
        retry_strategy = Retry(
            total=CONFIG["MAX_RETRIES"],
            backoff_factor=1,
            status_forcelist=[429, 500, 502, 503, 504],
        )
        adapter = HTTPAdapter(max_retries=retry_strategy)
        session.mount("http://", adapter)
        session.mount("https://", adapter)

        # Set headers
        session.headers.update(
            {
                "User-Agent": CONFIG["USER_AGENT"],
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
                "Accept-Language": "ru-RU,ru;q=0.9,en;q=0.8",
                "Accept-Encoding": "gzip, deflate, br",
                "Connection": "keep-alive",
                "Upgrade-Insecure-Requests": "1",
                "Sec-Fetch-Dest": "document",
                "Sec-Fetch-Mode": "navigate",
                "Sec-Fetch-Site": "none",
            }
        )

        return session

    def search_procurements(
        self,
        keyword,
        min_price=None,
        max_price=None,
        purchase_stage=None,
        law=None,
        progress_callback=None,
    ):
        """
        Search for procurements with improved error handling and robustness.
        """
        if not keyword or not keyword.strip():
            raise ValueError("Ключевое слово не может быть пустым")

        url = f"{CONFIG['BASE_URL']}/epz/order/extendedsearch/results.html"
        params = {
            "searchString": keyword.strip(),
            "pageNumber": "1",
            "recordsPerPage": f"_{CONFIG['RESULTS_PER_PAGE']}",
        }

        # Add purchase stage filter
        if purchase_stage and purchase_stage in PURCHASE_STAGES:
            stage_params = self._get_stage_params(purchase_stage)
            params.update(stage_params)

        # Add law filter
        if law and law in LAWS:
            law_params = self._get_law_params(law)
            params.update(law_params)

        if progress_callback:
            progress_callback("Отправка запроса...")

        try:
            logger.info(
                f"Searching for: {keyword} with filters - stage: {purchase_stage}, law: {law}"
            )
            response = self.session.get(
                url, params=params, timeout=CONFIG["REQUEST_TIMEOUT"]
            )
            response.raise_for_status()

            if progress_callback:
                progress_callback("Обработка результатов...")

            time.sleep(CONFIG["REQUEST_DELAY"])

            return self._parse_results(
                response.text, min_price, max_price, progress_callback
            )

        except requests.exceptions.Timeout:
            logger.error("Request timeout")
            raise Exception("Превышено время ожидания ответа от сервера")
        except requests.exceptions.ConnectionError:
            logger.error("Connection error")
            raise Exception("Ошибка подключения к серверу")
        except requests.exceptions.HTTPError as e:
            logger.error(f"HTTP error: {e}")
            raise Exception(f"Ошибка HTTP: {e.response.status_code}")
        except Exception as e:
            logger.error(f"Unexpected error during search: {e}")
            raise Exception(f"Неожиданная ошибка: {str(e)}")

    def _get_stage_params(self, purchase_stage):
        """Get URL parameters for purchase stage filter."""
        stage_mapping = {
            "SUBMISSION": {"purchaseStage": "SUBMISSION_OF_APPLICATIONS"},
            "EVALUATION": {"purchaseStage": "COMMISSION_WORK"},
            "COMPLETED": {"purchaseStage": "PURCHASE_COMPLETED"},
            "CANCELLED": {"purchaseStage": "PURCHASE_CANCELLED"},
        }
        return stage_mapping.get(purchase_stage, {})

    def _get_law_params(self, law):
        """Get URL parameters for law filter."""
        law_mapping = {
            "44FZ": {"fz44": "on"},
            "223FZ": {"fz223": "on"},
            "PP615": {"ppRf615": "on"},
        }
        return law_mapping.get(law, {})

    def _parse_results(
        self, html_content, min_price=None, max_price=None, progress_callback=None
    ):
        """Parse HTML content and extract procurement data."""
        try:
            soup = BeautifulSoup(html_content, "html.parser")

            # Multiple selector strategies for robustness
            selectors = [
                ".search-registry-entry-block",
                ".registry-entry",
                "[class*='registry-entry']",
                ".search-result-item",
            ]

            lots = []
            for selector in selectors:
                lots = soup.select(selector)
                if lots:
                    logger.info(f"Found {len(lots)} results using selector: {selector}")
                    break

            if not lots:
                logger.warning("No results found with any selector")
                return []

            results = []
            total_lots = len(lots)

            for i, lot in enumerate(lots):
                if progress_callback:
                    progress_callback(f"Обработка результата {i+1} из {total_lots}...")

                try:
                    result = self._parse_single_lot(lot)
                    if result and self._passes_price_filter(
                        result["Цена"], min_price, max_price
                    ):
                        results.append(result)
                except Exception as e:
                    logger.warning(f"Failed to parse lot {i+1}: {e}")
                    continue

            logger.info(f"Successfully parsed {len(results)} results")
            return results

        except Exception as e:
            logger.error(f"Failed to parse HTML: {e}")
            raise Exception("Ошибка обработки данных с сервера")

    def _passes_price_filter(self, price, min_price, max_price):
        """Check if price passes the min/max price filters."""
        if min_price is not None and price < min_price:
            return False
        if max_price is not None and price > max_price:
            return False
        return True

    def _parse_single_lot(self, lot):
        """Parse a single procurement lot with multiple fallback strategies."""
        title = self._extract_title(lot)
        price = self._extract_price(lot)
        link = self._extract_link(lot)

        if not title:
            return None

        return {"Название": title, "Цена": price, "Ссылка": link}

    def _extract_title(self, lot):
        """Extract title with multiple fallback selectors."""
        title_selectors = [
            ".registry-entry__header-top__title",
            ".registry-entry__title",
            "[class*='title']",
            "h3",
            "h2",
            ".search-result-title",
        ]

        for selector in title_selectors:
            title_el = lot.select_one(selector)
            if title_el:
                title = title_el.get_text(strip=True)
                if title and len(title) > 5:
                    return title

        return "Название не найдено"

    def _extract_price(self, lot):
        """Extract and validate price with multiple fallback selectors."""
        price_selectors = [
            ".price-block__value",
            ".price",
            "[class*='price']",
            ".cost",
            "[class*='cost']",
        ]

        for selector in price_selectors:
            price_el = lot.select_one(selector)
            if price_el:
                try:
                    price_text = (
                        price_el.get_text(strip=True)
                        .replace("\xa0", "")
                        .replace("₽", "")
                        .replace("руб", "")
                        .replace(" ", "")
                        .replace(",", ".")
                    )

                    numbers = re.findall(r"\d+\.?\d*", price_text)
                    if numbers:
                        price = float(numbers[0])
                        if price >= 0:
                            return price
                except (ValueError, AttributeError, IndexError) as e:
                    logger.debug(f"Price parsing failed for selector {selector}: {e}")
                    continue

        return 0.0

    def _extract_link(self, lot):
        """Extract and validate link."""
        link_el = lot.select_one("a")
        if link_el and link_el.get("href"):
            href = link_el["href"]
            if href.startswith("http"):
                return href
            else:
                return urljoin(CONFIG["BASE_URL"], href)
        return "Ссылка не найдена"

    # УЛУЧШЕННЫЕ МЕТОДЫ ДЛЯ РАБОТЫ С ДОКУМЕНТАМИ
    def _extract_notice_info_id(self, url):
        """Извлекает noticeInfoId из URL."""
        try:
            parsed_url = urlparse(url)
            query_params = parse_qs(parsed_url.query)
            notice_info_id = query_params.get("noticeInfoId", [None])[0]
            return notice_info_id
        except Exception as e:
            logger.error(f"Error extracting noticeInfoId from {url}: {e}")
            return None

    def download_documents(self, lot_url, progress_callback=None):
        """Скачивает документы из лота с улучшенным парсингом."""
        try:
            # Извлекаем noticeInfoId из URL
            notice_info_id = self._extract_notice_info_id(lot_url)
            if not notice_info_id:
                logger.error(f"Could not extract noticeInfoId from URL: {lot_url}")
                return []

            # Формируем URL страницы документов
            documents_url = (
                f"{CONFIG['BASE_URL']}/epz/order/notice/notice223/documents.html"
            )
            params = {
                "noticeInfoId": notice_info_id,
                "backUrl": "/epz/order/notice/notice223/search.html",
            }

            if progress_callback:
                progress_callback(f"Загрузка страницы документов...")

            logger.info(
                f"Downloading documents from: {documents_url}?noticeInfoId={notice_info_id}"
            )

            response = self.session.get(
                documents_url, params=params, timeout=CONFIG["REQUEST_TIMEOUT"]
            )
            response.raise_for_status()

            soup = BeautifulSoup(response.text, "html.parser")

            # Сохраняем HTML для отладки
            with open("debug_documents_page.html", "w", encoding="utf-8") as f:
                f.write(soup.prettify())

            # Ищем документы - несколько стратегий
            document_links = []

            # Стратегия 1: Ищем по классам, которые обычно используются для документов
            document_selectors = [
                "a[href*='download']",
                "a[href*='.pdf']",
                "a[href*='.doc']",
                "a[href*='.docx']",
                "a[href*='.xls']",
                "a[href*='.xlsx']",
                "a[href*='.zip']",
                "a[href*='.rar']",
                ".document-link",
                ".file-link",
                "[class*='document'] a",
                "[class*='file'] a",
                ".cardFile",
                ".file-download",
                "tr td a",  # Ссылки в таблицах
                ".table-block a",  # Ссылки в табличных блоках
            ]

            for selector in document_selectors:
                links = soup.select(selector)
                for link in links:
                    href = link.get("href")
                    if href:
                        full_url = urljoin(CONFIG["BASE_URL"], href)

                        # Проверяем, что это документ
                        if self._is_document_link(full_url, link):
                            name = self._get_document_name(link)
                            document_links.append({"name": name, "url": full_url})

            # Стратегия 2: Ищем все ссылки и фильтруем по расширениям
            all_links = soup.find_all("a", href=True)
            for link in all_links:
                href = link.get("href")
                if href:
                    full_url = urljoin(CONFIG["BASE_URL"], href)

                    # Пропускаем, если уже есть в списке
                    if any(doc["url"] == full_url for doc in document_links):
                        continue

                    if self._is_document_link(full_url, link):
                        name = self._get_document_name(link)
                        document_links.append({"name": name, "url": full_url})

            # Убираем дубликаты
            unique_links = []
            seen_urls = set()
            for doc in document_links:
                if doc["url"] not in seen_urls:
                    unique_links.append(doc)
                    seen_urls.add(doc["url"])

            logger.info(f"Found {len(unique_links)} unique document links")

            # Скачиваем документы
            downloaded_docs = []
            for i, doc in enumerate(unique_links):
                if progress_callback:
                    progress_callback(
                        f"Скачивание {i+1}/{len(unique_links)}: {doc['name'][:30]}..."
                    )

                try:
                    doc_response = self.session.get(
                        doc["url"], timeout=CONFIG["REQUEST_TIMEOUT"], stream=True
                    )
                    doc_response.raise_for_status()

                    content = doc_response.content

                    downloaded_docs.append(
                        {
                            "name": doc["name"],
                            "content": content,
                            "size": len(content),
                            "url": doc["url"],
                            "content_type": doc_response.headers.get(
                                "content-type", ""
                            ),
                        }
                    )

                    logger.info(
                        f"Successfully downloaded: {doc['name']} ({len(content)} bytes)"
                    )

                    time.sleep(0.5)  # Задержка между запросами

                except Exception as e:
                    logger.warning(f"Failed to download document {doc['name']}: {e}")
                    continue

            logger.info(f"Successfully downloaded {len(downloaded_docs)} documents")
            return downloaded_docs

        except Exception as e:
            logger.error(f"Error downloading documents: {e}")
            import traceback

            logger.error(traceback.format_exc())
            return []

    def _is_document_link(self, url, link_element):
        """Проверяет, является ли ссылка документом."""
        # Проверяем расширения файлов
        document_extensions = [
            ".pdf",
            ".doc",
            ".docx",
            ".xls",
            ".xlsx",
            ".zip",
            ".rar",
            ".txt",
            ".rtf",
            ".odt",
            ".ods",
        ]
        if any(url.lower().endswith(ext) for ext in document_extensions):
            return True

        # Проверяем ключевые слова в URL
        url_keywords = [
            "/download/",
            "/file/",
            "document",
            "file",
            "download",
            "getFile",
        ]
        if any(keyword in url.lower() for keyword in url_keywords):
            return True

        # Проверяем классы элемента
        element_classes = link_element.get("class", [])
        class_keywords = ["document", "file", "download", "cardFile", "file-download"]
        if any(
            any(keyword in str(cls).lower() for keyword in class_keywords)
            for cls in element_classes
        ):
            return True

        # Проверяем текст ссылки
        link_text = link_element.get_text(strip=True).lower()
        text_keywords = ["скачать", "документ", "файл", "doc", "pdf", "xls", "zip"]
        if any(keyword in link_text for keyword in text_keywords):
            return True

        return False

    def _get_document_name(self, link_element):
        """Извлекает имя документа из элемента ссылки."""
        # Пробуем получить имя из текста ссылки
        name = link_element.get_text(strip=True)
        if name and len(name) > 2:
            return name

        # Пробуем получить из атрибута title
        name = link_element.get("title", "")
        if name and len(name) > 2:
            return name

        # Пробуем получить из data-атрибутов
        for attr in link_element.attrs:
            if "name" in attr.lower() or "title" in attr.lower():
                name = link_element.get(attr, "")
                if name and len(name) > 2:
                    return name

        # Если ничего не нашли, генерируем имя из URL
        href = link_element.get("href", "")
        if href:
            # Извлекаем имя файла из URL
            filename = href.split("/")[-1]
            if "?" in filename:
                filename = filename.split("?")[0]
            if filename and "." in filename:
                return filename

        return f"document_{int(time.time())}"

    def search_in_documents(self, documents, keywords):
        """Ищет ключевые слова в документах."""
        results = []

        for doc in documents:
            doc_results = {
                "document_name": doc["name"],
                "size": doc["size"],
                "url": doc["url"],
                "matches": [],
                "match_count": 0,
            }

            try:
                content_text = self._extract_text_from_content(
                    doc["content"], doc["name"], doc.get("content_type", "")
                )

                # Ищем ключевые слова
                for keyword in keywords:
                    keyword_lower = keyword.lower().strip()
                    if keyword_lower and keyword_lower in content_text.lower():
                        doc_results["matches"].append(keyword)
                        doc_results["match_count"] += 1

                # Добавляем контекст для первых совпадений
                if doc_results["matches"]:
                    doc_results["sample_context"] = self._get_keyword_context(
                        content_text, keywords, max_contexts=2
                    )

                results.append(doc_results)

            except Exception as e:
                logger.warning(f"Error analyzing document {doc['name']}: {e}")
                results.append(doc_results)
                continue

        # Сортируем по количеству совпадений
        results.sort(key=lambda x: x["match_count"], reverse=True)
        return results

    def _extract_text_from_content(self, content, filename, content_type):
        """Извлекает текст из содержимого документа."""
        try:
            # Для текстовых файлов
            if filename.lower().endswith(".txt") or "text/" in content_type:
                return content.decode("utf-8", errors="ignore")

            # Для PDF, DOC, DOCX и других бинарных форматов
            # Используем простой метод извлечения текста
            text = content.decode("latin-1", errors="ignore")

            # Очищаем текст от непечатаемых символов
            text = re.sub(r"[^\x20-\x7E\xC0-\xFF\n\r\t]", " ", text)
            text = re.sub(r"\s+", " ", text)

            return text

        except Exception as e:
            logger.warning(f"Could not extract text from {filename}: {e}")
            return ""

    def _get_keyword_context(self, text, keywords, max_contexts=2, context_length=100):
        """Извлекает контекст вокруг найденных ключевых слов."""
        contexts = []
        text_lower = text.lower()

        for keyword in keywords:
            keyword_lower = keyword.lower().strip()
            if not keyword_lower:
                continue

            start = 0
            while len(contexts) < max_contexts:
                pos = text_lower.find(keyword_lower, start)
                if pos == -1:
                    break

                start_context = max(0, pos - context_length)
                end_context = min(len(text), pos + len(keyword_lower) + context_length)

                context = text[start_context:end_context]
                context = " ".join(context.split())
                contexts.append(f"...{context}...")

                start = pos + len(keyword_lower)

        return contexts[:max_contexts]


class ProcurementApp:
    """Main application class with improved UI and threading."""

    def __init__(self, root):
        self.root = root
        self.root.title("Поиск закупок - Улучшенная версия с анализом документов")
        self.root.geometry("900x700")
        self.results = []
        self.searcher = ProcurementSearcher()
        self.search_thread = None
        self.analysis_results = []

        self._setup_ui()

        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)

    def _setup_ui(self):
        """Setup the user interface."""
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(6, weight=1)

        # Input fields
        ttk.Label(main_frame, text="Ключевое слово:").grid(
            row=0, column=0, sticky="w", pady=2
        )
        self.keyword_entry = ttk.Entry(main_frame, width=50)
        self.keyword_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=2)
        self.keyword_entry.bind("<Return>", lambda e: self.search())

        # Price filters frame
        price_frame = ttk.LabelFrame(
            main_frame, text="Фильтр по цене (руб)", padding="5"
        )
        price_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        price_frame.columnconfigure(1, weight=1)
        price_frame.columnconfigure(3, weight=1)

        ttk.Label(price_frame, text="От:").grid(
            row=0, column=0, sticky="w", padx=(0, 5)
        )
        self.min_price_entry = ttk.Entry(price_frame, width=15)
        self.min_price_entry.grid(row=0, column=1, sticky="w", padx=(0, 20))

        ttk.Label(price_frame, text="До:").grid(
            row=0, column=2, sticky="w", padx=(0, 5)
        )
        self.max_price_entry = ttk.Entry(price_frame, width=15)
        self.max_price_entry.grid(row=0, column=3, sticky="w")

        # Additional filters frame
        filters_frame = ttk.LabelFrame(
            main_frame, text="Дополнительные фильтры", padding="5"
        )
        filters_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        filters_frame.columnconfigure(1, weight=1)
        filters_frame.columnconfigure(3, weight=1)

        # Purchase stage filter
        ttk.Label(filters_frame, text="Этап закупки:").grid(
            row=0, column=0, sticky="w", padx=(0, 5)
        )
        self.stage_var = tk.StringVar(value="")
        self.stage_combo = ttk.Combobox(
            filters_frame,
            textvariable=self.stage_var,
            values=list(PURCHASE_STAGES.values()),
            state="readonly",
            width=20,
        )
        self.stage_combo.grid(row=0, column=1, sticky="w", padx=(0, 20))

        # Law filter
        ttk.Label(filters_frame, text="Закон:").grid(
            row=0, column=2, sticky="w", padx=(0, 5)
        )
        self.law_var = tk.StringVar(value="")
        self.law_combo = ttk.Combobox(
            filters_frame,
            textvariable=self.law_var,
            values=list(LAWS.values()),
            state="readonly",
            width=15,
        )
        self.law_combo.grid(row=0, column=3, sticky="w")

        # Buttons frame
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.grid(row=3, column=0, columnspan=2, pady=10)

        self.search_btn = ttk.Button(
            buttons_frame, text="🔍 Поиск", command=self.search
        )
        self.search_btn.pack(side=tk.LEFT, padx=5)

        self.export_btn = ttk.Button(
            buttons_frame,
            text="📊 Сохранить в Excel",
            command=self.save_to_excel,
            state="disabled",
        )
        self.export_btn.pack(side=tk.LEFT, padx=5)

        self.clear_btn = ttk.Button(
            buttons_frame, text="🗑️ Очистить", command=self.clear_results
        )
        self.clear_btn.pack(side=tk.LEFT, padx=5)

        # Progress bar
        self.progress_var = tk.StringVar(value="Готов к поиску")
        progress_frame = ttk.Frame(main_frame)
        progress_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        progress_frame.columnconfigure(0, weight=1)

        self.progress_bar = ttk.Progressbar(progress_frame, mode="indeterminate")
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E))

        self.status_label = ttk.Label(progress_frame, textvariable=self.progress_var)
        self.status_label.grid(row=1, column=0, sticky="w")

        # Документы frame
        doc_frame = ttk.LabelFrame(
            main_frame, text="Поиск в документах ТЗ", padding="5"
        )
        doc_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        doc_frame.columnconfigure(0, weight=1)

        ttk.Label(doc_frame, text="Ключевые слова для поиска (через запятую):").grid(
            row=0, column=0, sticky="w", pady=2
        )
        self.keywords_entry = ttk.Entry(doc_frame, width=50)
        self.keywords_entry.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=2)

        doc_buttons_frame = ttk.Frame(doc_frame)
        doc_buttons_frame.grid(row=2, column=0, pady=5)

        self.analyze_btn = ttk.Button(
            doc_buttons_frame,
            text="🔍 Анализировать документы выбранного лота",
            command=self.analyze_documents,
        )
        self.analyze_btn.pack(side=tk.LEFT, padx=5)

        self.export_docs_btn = ttk.Button(
            doc_buttons_frame,
            text="📋 Экспорт результатов анализа",
            command=self.export_analysis,
            state="disabled",
        )
        self.export_docs_btn.pack(side=tk.LEFT, padx=5)

        # Results table
        self._setup_results_table(main_frame)

    def _setup_results_table(self, parent):
        """Setup the results table with scrollbars."""
        table_frame = ttk.Frame(parent)
        table_frame.grid(
            row=6, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10
        )
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)

        columns = ("Название", "Цена", "Ссылка")
        self.tree = ttk.Treeview(
            table_frame, columns=columns, show="headings", height=15
        )

        self.tree.heading("Название", text="Название", anchor="w")
        self.tree.heading("Цена", text="Цена (руб)", anchor="e")
        self.tree.heading("Ссылка", text="Ссылка", anchor="w")

        self.tree.column("Название", width=400, anchor="w")
        self.tree.column("Цена", width=120, anchor="e")
        self.tree.column("Ссылка", width=350, anchor="w")

        v_scrollbar = ttk.Scrollbar(
            table_frame, orient="vertical", command=self.tree.yview
        )
        h_scrollbar = ttk.Scrollbar(
            table_frame, orient="horizontal", command=self.tree.xview
        )
        self.tree.configure(
            yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set
        )

        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))

        self.tree.bind("<Double-1>", self._on_item_double_click)

    def _on_item_double_click(self, event):
        """Handle double-click on table item to open link."""
        item = self.tree.selection()[0] if self.tree.selection() else None
        if item:
            values = self.tree.item(item, "values")
            if len(values) >= 3 and values[2] != "Ссылка не найдена":
                import webbrowser

                try:
                    webbrowser.open(values[2])
                except Exception as e:
                    messagebox.showerror("Ошибка", f"Не удалось открыть ссылку:\n{e}")

    def search(self):
        """Start search in a separate thread."""
        if self.search_thread and self.search_thread.is_alive():
            messagebox.showwarning("Поиск", "Поиск уже выполняется...")
            return

        keyword = self.keyword_entry.get().strip()
        if not keyword:
            messagebox.showerror("Ошибка", "Введите ключевое слово для поиска")
            self.keyword_entry.focus()
            return

        # Validate minimum price
        min_price = None
        min_price_text = self.min_price_entry.get().strip()
        if min_price_text:
            try:
                min_price = float(min_price_text.replace(",", "."))
                if min_price < 0:
                    raise ValueError("Отрицательная цена")
            except ValueError:
                messagebox.showerror(
                    "Ошибка", "Минимальная цена должна быть положительным числом"
                )
                self.min_price_entry.focus()
                return

        # Validate maximum price
        max_price = None
        max_price_text = self.max_price_entry.get().strip()
        if max_price_text:
            try:
                max_price = float(max_price_text.replace(",", "."))
                if max_price < 0:
                    raise ValueError("Отрицательная цена")
            except ValueError:
                messagebox.showerror(
                    "Ошибка", "Максимальная цена должна быть положительным числом"
                )
                self.max_price_entry.focus()
                return

        # Validate price range
        if min_price is not None and max_price is not None and min_price > max_price:
            messagebox.showerror(
                "Ошибка", "Минимальная цена не может быть больше максимальной"
            )
            self.min_price_entry.focus()
            return

        # Get filter values
        purchase_stage = self._get_stage_key(self.stage_var.get())
        law = self._get_law_key(self.law_var.get())

        # Start search thread
        self.search_thread = threading.Thread(
            target=self._perform_search,
            args=(keyword, min_price, max_price, purchase_stage, law),
            daemon=True,
        )
        self.search_thread.start()

    def _get_stage_key(self, stage_value):
        """Get stage key from display value."""
        for key, value in PURCHASE_STAGES.items():
            if value == stage_value:
                return key
        return None

    def _get_law_key(self, law_value):
        """Get law key from display value."""
        for key, value in LAWS.items():
            if value == law_value:
                return key
        return None

    def _perform_search(self, keyword, min_price, max_price, purchase_stage, law):
        """Perform search in background thread."""

        def update_progress(message):
            self.root.after(0, lambda: self.progress_var.set(message))

        def update_ui_start():
            self.search_btn.config(state="disabled")
            self.export_btn.config(state="disabled")
            self.progress_bar.start()
            self.tree.delete(*self.tree.get_children())
            self.results = []

        def update_ui_finish(results, error=None):
            self.progress_bar.stop()
            self.search_btn.config(state="normal")

            if error:
                self.progress_var.set(f"Ошибка: {error}")
                messagebox.showerror("Ошибка поиска", error)
                return

            self.results = results

            # Update table
            for result in results:
                price_display = (
                    f"{result['Цена']:,.2f}" if result["Цена"] > 0 else "Не указана"
                )
                self.tree.insert(
                    "",
                    "end",
                    values=(
                        (
                            result["Название"][:100] + "..."
                            if len(result["Название"]) > 100
                            else result["Название"]
                        ),
                        price_display,
                        result["Ссылка"],
                    ),
                )

            # Update status
            if results:
                self.progress_var.set(f"Найдено результатов: {len(results)}")
                self.export_btn.config(state="normal")
            else:
                self.progress_var.set("Результаты не найдены")
                messagebox.showinfo("Результат", "По вашему запросу ничего не найдено")

        # Start UI updates
        self.root.after(0, update_ui_start)

        try:
            results = self.searcher.search_procurements(
                keyword, min_price, max_price, purchase_stage, law, update_progress
            )
            self.root.after(0, lambda: update_ui_finish(results))
        except Exception as e:
            logger.error(f"Search failed: {e}")
            self.root.after(0, lambda: update_ui_finish([], str(e)))

    def save_to_excel(self):
        """Save results to Excel file."""
        if not self.results:
            messagebox.showwarning("Нет данных", "Сначала выполните поиск")
            return

        try:
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel файлы", "*.xlsx"), ("Все файлы", "*.*")],
                title="Сохранить результаты поиска",
            )

            if not file_path:
                return

            # Create DataFrame with formatted data
            df_data = []
            for result in self.results:
                df_data.append(
                    {
                        "Название": result["Название"],
                        "Цена (руб)": (
                            result["Цена"] if result["Цена"] > 0 else "Не указана"
                        ),
                        "Ссылка": result["Ссылка"],
                    }
                )

            df = pd.DataFrame(df_data)

            # Save with formatting
            with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="Результаты поиска", index=False)

                # Auto-adjust column widths
                worksheet = writer.sheets["Результаты поиска"]
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width

            messagebox.showinfo("Сохранено", f"Файл успешно сохранён:\n{file_path}")
            logger.info(f"Results exported to: {file_path}")

        except Exception as e:
            logger.error(f"Failed to save Excel file: {e}")
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл:\n{e}")

    def clear_results(self):
        """Clear search results and reset filters."""
        self.tree.delete(*self.tree.get_children())
        self.results = []
        self.export_btn.config(state="disabled")
        self.progress_var.set("Результаты очищены")

        # Reset filters
        self.keyword_entry.delete(0, tk.END)
        self.min_price_entry.delete(0, tk.END)
        self.max_price_entry.delete(0, tk.END)
        self.stage_var.set("")
        self.law_var.set("")

    def _on_closing(self):
        """Handle application closing."""
        if self.search_thread and self.search_thread.is_alive():
            if messagebox.askokcancel(
                "Выход", "Поиск ещё выполняется. Закрыть приложение?"
            ):
                self.root.destroy()
        else:
            self.root.destroy()

    # МЕТОДЫ ДЛЯ РАБОТЫ С ДОКУМЕНТАМИ
    def analyze_documents(self):
        """Анализирует документы выбранного лота."""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Предупреждение", "Выберите лот для анализа")
            return

        item = selection[0]
        values = self.tree.item(item, "values")
        lot_url = values[2] if len(values) >= 3 else None

        if not lot_url or lot_url == "Ссылка не найдена":
            messagebox.showerror("Ошибка", "Неверная ссылка на лот")
            return

        keywords_text = self.keywords_entry.get().strip()
        if not keywords_text:
            messagebox.showerror("Ошибка", "Введите ключевые слова для поиска")
            return

        keywords = [k.strip() for k in keywords_text.split(",") if k.strip()]

        # Запускаем в отдельном потоке
        thread = threading.Thread(
            target=self._perform_document_analysis,
            args=(lot_url, keywords),
            daemon=True,
        )
        thread.start()

    def _perform_document_analysis(self, lot_url, keywords):
        """Выполняет анализ документов в фоновом режиме."""

        def update_progress(message):
            self.root.after(0, lambda: self.progress_var.set(message))

        def show_results(documents, analysis_results):
            self.progress_bar.stop()
            self.export_docs_btn.config(
                state="normal" if analysis_results else "disabled"
            )

            self._show_analysis_results(documents, analysis_results, keywords)

        self.root.after(0, lambda: self.progress_bar.start())

        try:
            update_progress("Поиск и скачивание документов...")
            documents = self.searcher.download_documents(lot_url, update_progress)

            if not documents:
                self.root.after(
                    0,
                    lambda: messagebox.showinfo(
                        "Результат",
                        f"Документы не найдены в выбранном лоте.\n\n"
                        f"URL: {lot_url}\n"
                        f"Проверьте:\n"
                        f"1. Доступность страницы в браузере\n"
                        f"2. Наличие раздела 'Документы'\n"
                        f"3. Права доступа к документам",
                    ),
                )
                self.progress_bar.stop()
                return

            update_progress(f"Анализ {len(documents)} документов...")
            analysis_results = self.searcher.search_in_documents(documents, keywords)

            self.root.after(0, lambda: show_results(documents, analysis_results))

        except Exception as e:
            logger.error(f"Document analysis failed: {e}")
            self.root.after(
                0, lambda: messagebox.showerror("Ошибка", f"Ошибка анализа: {str(e)}")
            )
            self.progress_bar.stop()

    def _show_analysis_results(self, documents, analysis_results, keywords):
        """Показывает результаты анализа."""
        result_window = tk.Toplevel(self.root)
        result_window.title("Результаты анализа документов")
        result_window.geometry("900x600")

        text_frame = ttk.Frame(result_window)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        text_widget = scrolledtext.ScrolledText(text_frame, wrap=tk.WORD)
        text_widget.pack(fill=tk.BOTH, expand=True)

        # Заголовок
        text_widget.insert(tk.END, "РЕЗУЛЬТАТЫ АНАЛИЗА ДОКУМЕНТОВ\n", "title")
        text_widget.insert(tk.END, "=" * 50 + "\n\n")

        text_widget.insert(
            tk.END, f"Ключевые слова для поиска: {', '.join(keywords)}\n"
        )
        text_widget.insert(tk.END, f"Всего документов: {len(documents)}\n")

        documents_with_matches = [r for r in analysis_results if r["matches"]]
        text_widget.insert(
            tk.END, f"Документы с совпадениями: {len(documents_with_matches)}\n\n"
        )

        # Документы с совпадениями
        if documents_with_matches:
            text_widget.insert(
                tk.END, "ДОКУМЕНТЫ С НАЙДЕННЫМИ КЛЮЧЕВЫМИ СЛОВАМИ:\n", "subtitle"
            )
            text_widget.insert(tk.END, "=" * 50 + "\n\n")

            for result in documents_with_matches:
                text_widget.insert(
                    tk.END, f"📄 {result['document_name']}\n", "document_name"
                )
                text_widget.insert(tk.END, f"   Размер: {result['size']} bytes\n")
                text_widget.insert(
                    tk.END, f"   Найдено совпадений: {result['match_count']}\n"
                )
                text_widget.insert(
                    tk.END, f"   Ключевые слова: {', '.join(result['matches'])}\n"
                )

                if "sample_context" in result and result["sample_context"]:
                    text_widget.insert(tk.END, f"   Контекст:\n")
                    for i, context in enumerate(result["sample_context"][:2]):
                        text_widget.insert(tk.END, f"     {i+1}. {context}\n")

                text_widget.insert(tk.END, f"   Ссылка: {result['url']}\n\n")
        else:
            text_widget.insert(tk.END, "Ключевые слова не найдены в документах.\n\n")

        # Все документы
        text_widget.insert(tk.END, "ВСЕ НАЙДЕННЫЕ ДОКУМЕНТЫ:\n", "subtitle")
        text_widget.insert(tk.END, "=" * 50 + "\n\n")

        for doc in documents:
            text_widget.insert(tk.END, f"📄 {doc['name']} ({doc['size']} bytes)\n")

        # Настройка стилей текста
        text_widget.tag_configure("title", font=("Arial", 12, "bold"))
        text_widget.tag_configure("subtitle", font=("Arial", 10, "bold"))
        text_widget.tag_configure("document_name", font=("Arial", 9, "bold"))

        text_widget.config(state=tk.DISABLED)

        self.analysis_results = analysis_results

    def export_analysis(self):
        """Экспортирует результаты анализа."""
        if not self.analysis_results:
            messagebox.showwarning("Нет данных", "Сначала выполните анализ документов")
            return

        try:
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel файлы", "*.xlsx"), ("Все файлы", "*.*")],
                title="Сохранить результаты анализа",
            )

            if not file_path:
                return

            # Создаем DataFrame с результатами анализа
            df_data = []
            for result in self.analysis_results:
                df_data.append(
                    {
                        "Документ": result["document_name"],
                        "Размер (bytes)": result["size"],
                        "Найденные ключевые слова": ", ".join(result["matches"]),
                        "Количество совпадений": result["match_count"],
                        "Ссылка на документ": result["url"],
                    }
                )

            df = pd.DataFrame(df_data)

            # Сохраняем в Excel
            with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="Результаты анализа", index=False)

                worksheet = writer.sheets["Результаты анализа"]
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width

            messagebox.showinfo(
                "Сохранено", f"Результаты анализа сохранены:\n{file_path}"
            )

        except Exception as e:
            logger.error(f"Failed to save analysis results: {e}")
            messagebox.showerror(
                "Ошибка", f"Не удалось сохранить результаты анализа:\n{e}"
            )


def main():
    """Main application entry point."""
    try:
        root = tk.Tk()

        try:
            root.state("zoomed")
        except:
            try:
                root.attributes("-zoomed", True)
            except:
                pass

        app = ProcurementApp(root)

        root.update_idletasks()
        x = (root.winfo_screenwidth() // 2) - (900 // 2)
        y = (root.winfo_screenheight() // 2) - (700 // 2)
        root.geometry(f"900x700+{x}+{y}")

        root.mainloop()

    except Exception as e:
        logger.error(f"Application failed to start: {e}")
        messagebox.showerror(
            "Критическая ошибка", f"Не удалось запустить приложение:\n{e}"
        )


if __name__ == "__main__":
    main()
