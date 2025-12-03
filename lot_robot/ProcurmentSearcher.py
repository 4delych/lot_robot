import os
import sys
import time
import re
import logging
import zipfile
import html
import tempfile
import locale
from urllib.parse import urljoin, urlparse, parse_qs, unquote

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from bs4 import BeautifulSoup
from io import BytesIO

try:
    from docx import Document
except ImportError:
    Document = None

try:
    from openpyxl import load_workbook
except ImportError:
    load_workbook = None

try:
    import xlrd
except ImportError:
    xlrd = None

try:
    import pythoncom  # type: ignore
except ImportError:
    pythoncom = None  # type: ignore

try:
    import win32com.client as win32_client  # type: ignore
    from win32com.client import constants as win32_constants  # type: ignore
except ImportError:
    win32_client = None  # type: ignore
    win32_constants = None

from config import CONFIG, PURCHASE_STAGES, LAWS

logger = logging.getLogger(__name__)


class ProcurementSearcher:
    """Handle web scraping operations with proper error handling and session management."""

    def __init__(self):
        self.session = self._create_session()

    # создание сессии
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

        if min_price is not None and min_price is not None :
            self._add_price_params(params, min_price, max_price)

        if progress_callback:
            progress_callback("Отправка запроса...")

        try:
            logger.info(
                f"Searching for: {keyword!r} with filters - "
                f"stage: {purchase_stage}, law: {law}, "
                f"min_price: {min_price}, max_price: {max_price}"
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
            "SUBMISSION": {"af": "on"},  # Подача заявок
            "EVALUATION": {"ca": "on"},  # Работа комиссии
            "COMPLETED": {"pc": "on"},  # Закупка завершена
            "CANCELLED": {"pa": "on"},  # Закупка отменена
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

    def _add_price_params(self, params: dict, min_price, max_price) -> None:
        """
        Добавляет в параметры запроса фильтр по цене так,
        как это делает сайт zakupki.gov.ru:

        priceFromGeneral — минимальная цена
        priceToGeneral   — максимальная цена
        currencyIdGeneral = -1 (любая валюта / по умолчанию)
        """

        def _to_int_or_none(value):
            if value is None:
                return None
            if isinstance(value, (int, float)):
                return int(value)

            # Строка из UI: убираем пробелы/неразрывные пробелы, меняем запятую на точку
            s = str(value).strip()
            if not s:
                return None
            s = s.replace("\u00a0", "").replace(" ", "")
            s = s.replace(",", ".")
            try:
                return int(float(s))
            except ValueError:
                logger.warning(f"Не удалось преобразовать цену '{value}' в число")
                return None

        min_val = _to_int_or_none(min_price)
        max_val = _to_int_or_none(max_price)

        if min_val is not None:
            params["priceFromGeneral"] = str(min_val)
        if max_val is not None:
            params["priceToGeneral"] = str(max_val)

        # как в примере запроса с сайта: -1 = любая валюта
        if "currencyIdGeneral" not in params:
            params["currencyIdGeneral"] = "-1"

        logger.info(
            "Фильтр по цене в запросе: priceFromGeneral=%s, priceToGeneral=%s, currencyIdGeneral=%s",
            params.get("priceFromGeneral"),
            params.get("priceToGeneral"),
            params.get("currencyIdGeneral"),
        )

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
                    progress_callback(f"Обработка результата {i + 1} из {total_lots}...")

                try:
                    result = self._parse_single_lot(lot)
                    if result and self._passes_price_filter(
                            result["Цена"], min_price, max_price
                    ):
                        results.append(result)
                except Exception as e:
                    logger.warning(f"Failed to parse lot {i + 1}: {e}")
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
        header_text = None
        object_text = None

        # 1. Заголовок сверху: "44-ФЗ Электронный аукцион"
        try:
            header_el = lot.select_one(".registry-entry__header-top__title")
            if header_el:
                header_text = header_el.get_text(" ", strip=True)
                if header_text:
                    header_text = " ".join(header_text.split())
        except Exception as e:
            logger.debug(f"Failed to extract header title: {e}")

        # 2. "Объект закупки" из body
        try:
            for body_block in lot.select(".registry-entry__body-block"):
                title_el = body_block.select_one(".registry-entry__body-title")
                if not title_el:
                    continue

                title_label = title_el.get_text(strip=True).lower()
                if "объект закупки" in title_label:
                    value_el = body_block.select_one(".registry-entry__body-value")
                    if value_el:
                        object_text = value_el.get_text(" ", strip=True)
                        if object_text:
                            object_text = " ".join(object_text.split())
                    break
        except Exception as e:
            logger.debug(f"Failed to extract object text: {e}")

        # 3. Комбинируем
        if header_text and object_text:
            return f"{header_text} — {object_text}"
        if object_text:
            return object_text
        if header_text:
            return header_text

        # 4. Фолбэк — старые селекторы на всякий случай
        title_selectors = [
            ".registry-entry__title",
            "[class*='title']",
            "h3",
            "h2",
            ".search-result-title",
        ]

        for selector in title_selectors:
            title_el = lot.select_one(selector)
            if title_el:
                title = title_el.get_text(" ", strip=True)
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
        base_url = CONFIG["BASE_URL"]

        link_el = lot.select_one(".registry-entry__header-mid__number a[href]")
        if link_el:
            href = link_el["href"].strip()
            if href:
                return urljoin(base_url, href)

        # 2. Альтернативный вариант — если структура чуть отличается
        candidates = lot.select("a[href*='notice/']")
        for link in candidates:
            href = (link.get("href") or "").strip()
            if not href:
                continue
            # пропускаем ссылки на printForm
            if "printForm" in href.lower():
                continue
            if "view" in href.lower() and "regNumber" in href:
                # ТУТ тоже нужно делать urljoin, иначе получаешь только хвост
                return urljoin(base_url, href)

        # 3. Если всё совсем плохо — явное "не найдено"
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
        """
        Скачивает документы из лота с учётом разных схем:
        - 44-ФЗ: /epz/order/notice/ea20|ea44|.../view/common-info.html?regNumber=...
                 -> /epz/order/notice/.../view/documents.html?regNumber=...
        - 223-ФЗ: /epz/order/notice/notice223/... + noticeInfoId
        """
        try:
            parsed_url = urlparse(lot_url)
            path = parsed_url.path or ""
            query_dict = parse_qs(parsed_url.query)  # значения = списки

            documents_url = None
            params = {}

            # Ветка 44-ФЗ и родственных (ea20, ea44, zk20 и т.п.)
            # Пример:
            # https://zakupki.gov.ru/epz/order/notice/ea20/view/common-info.html?regNumber=0117...
            if "/epz/order/notice/" in path and "view/" in path:
                # Заменяем common-info.html -> documents.html
                if "common-info.html" in path:
                    docs_path = path.replace("common-info.html", "documents.html")
                else:
                    # вдруг нам уже дали documents.html
                    docs_path = path

                documents_url = urljoin(CONFIG["BASE_URL"], docs_path)

                # parse_qs возвращает {key: [value]}; сплющим до {key: value}
                params = {k: v[0] for k, v in query_dict.items() if v}

            else:
                # Ветка 223-ФЗ: старая схема через notice223 и noticeInfoId
                notice_info_id = query_dict.get("noticeInfoId", [None])[0]
                if not notice_info_id:
                    notice_info_id = self._extract_notice_info_id(lot_url)

                if not notice_info_id:
                    logger.error(f"Could not extract noticeInfoId from URL: {lot_url}")
                    return []

                documents_url = (
                    f"{CONFIG['BASE_URL']}/epz/order/notice/notice223/documents.html"
                )
                params = {
                    "noticeInfoId": notice_info_id,
                    "backUrl": "/epz/order/notice/notice223/search.html",
                }

            if progress_callback:
                progress_callback("Загрузка страницы документов.")

            logger.info(
                f"Downloading documents from: {documents_url} with params={params}"
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
                "tr td a",        # Ссылки в таблицах
                ".table-block a", # Ссылки в табличных блоках
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
                        f"Скачивание {i + 1}/{len(unique_links)}: {doc['name'][:30]}..."
                    )

                try:
                    doc_response = self.session.get(
                        doc["url"], timeout=CONFIG["REQUEST_TIMEOUT"], stream=True
                    )
                    doc_response.raise_for_status()

                    content = doc_response.content
                    display_name = self._normalize_filename(doc.get("name"))

                    # Пытаемся получить реальное имя файла из заголовков ответа или URL
                    real_name = self._guess_real_filename(
                        doc_response, display_name, doc["url"]
                    )

                    # Имя, которое показываем в UI
                    final_name = real_name or display_name
                    if not final_name:
                        final_name = self._normalize_filename(
                            os.path.basename(urlparse(doc["url"]).path)
                        )
                    if not final_name:
                        final_name = "document"

                    downloaded_docs.append(
                        {
                            "name": final_name,
                            "filename": real_name or final_name,
                            "content": content,
                            "size": len(content),
                            "url": doc["url"],
                            "content_type": doc_response.headers.get(
                                "content-type", ""
                            ),
                        }
                    )

                    logger.info(
                        "Successfully downloaded: display=%r, real=%r, final=%r, size=%s bytes",
                        display_name,
                        real_name,
                        final_name,
                        len(content),
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

    def search_in_documents(self, documents, keywords, progress_callback=None):
        """Ищет ключевые слова в документах."""
        results = []
        total = len(documents)

        logger.info(
            "Начат анализ документов: всего %s, ключевые слова: %s",
            total,
            ", ".join(map(str, keywords)),
        )

        for idx, doc in enumerate(documents, start=1):
            filename = self._determine_document_filename(doc)
            filename_l = filename.lower()
            doc_results = {
                "document_name": doc["name"],
                "size": doc["size"],
                "url": doc["url"],
                "matches": [],
                "match_count": 0,
            }

            if progress_callback:
                try:
                    progress_callback(
                        f"Анализ документа {idx}/{total}: {doc['name'][:60]}..."
                    )
                except Exception:
                    # Не даём GUI-ошибкам завалить анализ
                    pass

            logger.info(
                "Анализ документа %s/%s: name=%r, url=%r, filename_for_parser=%r",
                idx,
                total,
                doc.get("name"),
                doc.get("url"),
                filename,
            )

            try:
                content_text = self._extract_text_from_content(
                    doc["content"], filename, doc.get("content_type", "")
                )

                if idx == 1:
                    # для отладки: покажем фрагмент извлечённого текста первого документа
                    logger.info(
                        "Пример извлечённого текста (первые 500 символов) из %r: %r",
                        doc.get("name"),
                        content_text[:500],
                    )

                logger.debug(
                    "Извлечено %s символов текста из документа %r",
                    len(content_text),
                    doc.get("name"),
                )

                # Для Word‑документов используем более точный поиск по словам
                if filename_l.endswith(".docx") or filename_l.endswith(".doc"):
                    matched, match_count = self._find_word_matches_in_text(
                        content_text, keywords
                    )
                    logger.info(
                        "Совпадения в Word‑документе %r: найдено %s, ключевые слова: %s",
                        doc.get("name"),
                        match_count,
                        matched,
                    )
                    doc_results["matches"].extend(matched)
                    doc_results["match_count"] += match_count
                else:
                    # Для остальных форматов оставляем прежний простой поиск подстроки
                    lower_text = content_text.lower()
                    for keyword in keywords:
                        keyword_lower = keyword.lower().strip()
                        if keyword_lower and keyword_lower in lower_text:
                            doc_results["matches"].append(keyword)
                            doc_results["match_count"] += 1

                    if doc_results["match_count"]:
                        logger.info(
                            "Совпадения в документе %r: найдено %s, ключевые слова: %s",
                            doc.get("name"),
                            doc_results["match_count"],
                            doc_results["matches"],
                        )

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

        logger.info(
            "Анализ документов завершён. Документов: %s, с совпадениями: %s",
            len(results),
            sum(1 for r in results if r["matches"]),
        )

        # Сортируем по количеству совпадений
        results.sort(key=lambda x: x["match_count"], reverse=True)
        return results

    def _determine_document_filename(self, doc: dict) -> str:
        """
        Пытается определить реальное имя файла:
        - сначала используем то, что вывели в UI (doc['name'])
        - если расширения нет, пробуем взять basename из URL
        """
        name = (doc.get("filename")
                or doc.get("name")
                or "").strip()
        if "." in os.path.basename(name):
            return name

        url = doc.get("url") or ""
        try:
            path = urlparse(url).path
            candidate = os.path.basename(path)
            if candidate:
                return candidate
        except Exception:
            pass

        # Фолбэк — оставляем оригинальное имя или генерируем
        return name or f"document_{int(time.time())}"

    def _find_word_matches_in_text(self, text: str, keywords, min_len: int = 2):
        """
        Более аккуратный поиск по словам в тексте, в первую очередь для Word‑документов.

        - Использует регулярные выражения с границами слов (\b) и учётом юникода.
        - Игнорирует регистр.
        - Для каждого ключевого слова увеличивает счётчик на количество найденных вхождений.
        """
        if not text:
            return [], 0

        text_normalized = " ".join(str(text).split())
        total_count = 0
        matched_keywords = []

        for raw_kw in keywords:
            kw = (raw_kw or "").strip()
            if not kw or len(kw) < min_len:
                continue

            # Экранируем спецсимволы, чтобы искать именно текст, а не паттерн regex
            pattern = r"\b" + re.escape(kw) + r"\b"
            matches = re.findall(pattern, text_normalized, flags=re.IGNORECASE)
            if matches:
                matched_keywords.append(raw_kw)
                total_count += len(matches)

        return matched_keywords, total_count

    def _extract_text_from_content(self, content, filename, content_type):
        """Извлекает текст из содержимого документа с учётом формата."""
        filename_l = (filename or "").lower()
        ctype = (content_type or "").lower()

        try:
            # 0. Если это ZIP‑контейнер (OOXML: docx/xlsx и т.п.) — пробуем как DOCX
            #    Независимо от расширения файла, так как на сайте часто путают .doc и .docx.
            if content[:2] == b"PK":
                text = self._extract_text_from_docx(content)
                if text:
                    return text

            # 1. Обычный текст
            if filename_l.endswith(".txt") or "text/" in ctype:
                return content.decode("utf-8", errors="ignore")

            # 2. DOCX
            if (
                    filename_l.endswith(".docx")
                    or "officedocument.wordprocessingml" in ctype
            ):
                text = self._extract_text_from_docx(content)
                if text:
                    return text

            # 2.1. Старые DOC (часто RTF или HTML внутри .doc)
            if (
                    filename_l.endswith(".doc")
                    and not filename_l.endswith(".docx")
            ) or "msword" in ctype:
                text = self._extract_text_from_doc(content)
                if text:
                    return text

            # 3. XLSX / XLSM
            if filename_l.endswith((".xlsx", ".xlsm", ".xltx", ".xltm")) or "spreadsheetml" in ctype:
                text = self._extract_text_from_xlsx(content)
                if text:
                    return text

            # 4. Старые XLS
            if filename_l.endswith(".xls") or "ms-excel" in ctype:
                text = self._extract_text_from_xls(content)
                if text:
                    return text

            # 5. Всё остальное (PDF, DOC, RTF и т.п.) — грубый универсальный метод
            # Сначала пробуем UTF-8, чтобы не сломать кириллицу, даже если формат распознан неверно.
            text = content.decode("utf-8", errors="ignore")
            text = re.sub(r"[^\x20-\x7E\xC0-\xFF\n\r\t]", " ", text)
            text = re.sub(r"\s+", " ", text)
            return text

        except Exception as e:
            logger.warning(f"Could not extract text from {filename}: {e}")
            return ""


    def _extract_text_from_docx(self, content: bytes) -> str:
        """Извлекает текст из .docx с помощью python-docx."""
        # 1) Основной путь — через python-docx (если установлен)
        if Document is not None:
            try:
                doc = Document(BytesIO(content))
                parts = []

                # Параграфы
                for p in doc.paragraphs:
                    text = p.text.strip()
                    if text:
                        parts.append(text)

                # Таблицы
                for table in doc.tables:
                    for row in table.rows:
                        row_text = " ".join(
                            (cell.text or "").strip() for cell in row.cells if cell.text
                        )
                        if row_text:
                            parts.append(row_text)

                return "\n".join(parts)
            except Exception as e:
                logger.warning(f"Не удалось извлечь текст из DOCX через python-docx: {e}")

        # 2) Запасной путь — разбираем .docx как ZIP и достаём word/document.xml
        try:
            with zipfile.ZipFile(BytesIO(content)) as zf:
                with zf.open("word/document.xml") as doc_xml:
                    xml_bytes = doc_xml.read()
            xml_text = xml_bytes.decode("utf-8", errors="ignore")
            # Убираем XML/HTML-теги и декодируем сущности (&amp; и т.п.)
            xml_text = html.unescape(re.sub(r"<[^>]+>", " ", xml_text))
            xml_text = re.sub(r"\s+", " ", xml_text)
            return xml_text.strip()
        except Exception as e:
            logger.warning(f"Не удалось извлечь текст из DOCX через zipfile: {e}")
            return ""

    def _guess_real_filename(self, response, display_name: str | None, url: str) -> str:
        """
        Определяет реальное имя файла по следующему приоритету:
        1) filename / filename* в заголовке Content-Disposition
        2) basename из URL (если есть расширение)
        3) display_name (как в ссылке на сайте)
        """
        # 1. Content-Disposition
        cd = response.headers.get("content-disposition", "") or ""
        cd_lower = cd.lower()
        filename = ""

        # filename* (RFC 5987, например: filename*=UTF-8''%D0%90%D0%BA%D1%82.docx)
        if "filename*=" in cd_lower:
            try:
                part = cd.split("filename*=", 1)[1].split(";", 1)[0].strip()
                if part.lower().startswith("utf-8''"):
                    part = part[8:]
                part = part.strip('";')
                filename = unquote(part)
            except Exception:
                filename = ""

        # обычный filename="name.ext"
        if not filename and "filename=" in cd_lower:
            try:
                part = cd.split("filename=", 1)[1].split(";", 1)[0].strip()
                filename = part.strip('";')
            except Exception:
                filename = ""

        # 2. basename из URL, если в нём есть расширение
        if not filename:
            try:
                path = urlparse(url).path
                base = os.path.basename(path)
                if "." in base:
                    filename = base
            except Exception:
                filename = ""

        # 3. display_name как есть
        if not filename:
            filename = (display_name or "").strip()

        return self._normalize_filename(filename or "document")

    def _normalize_filename(self, name: str | None) -> str:
        """Приводит имя файла к удобному виду и исправляет типичный mojibake."""
        if not name:
            return ""

        cleaned = (name or "").strip().strip('"').replace("\\", "/")
        cleaned = cleaned.split("/")[-1]
        if not cleaned:
            return ""

        if self._looks_like_mojibake(cleaned):
            for encoding in ("utf-8", "cp1251"):
                try:
                    cleaned = cleaned.encode("latin-1").decode(encoding)
                    break
                except Exception:
                    continue

        return cleaned.strip()

    def _looks_like_mojibake(self, text: str) -> bool:
        """Пытается определить, искажено ли имя файла (Ð, Ñ, Ã и т.п.)."""
        if not text:
            return False
        mojibake_markers = ("Ð", "Ñ", "Ã", "Ò", "Â")
        return any(marker in text for marker in mojibake_markers)

    def _extract_text_from_doc(self, content: bytes) -> str:
        """
        Пытается извлечь текст из старого .doc.

        Частые варианты на гос-сайтах:
        - RTF, сохранённый как .doc
        - HTML, сохранённый как .doc
        - бинарный DOC, где текстовые фрагменты всё равно можно частично вытащить декодированием.
        """
        header = content[:2048]

        # 0) Если установлен Microsoft Word и доступен win32com — сначала пробуем извлечь текст через него.
        #    Это самый надёжный способ для «настоящих» бинарных DOC, независимо от сигнатуры.
        if self._can_use_win32_word():
            text = self._extract_text_from_doc_with_word(content)
            if text:
                return text

        # 1) RTF внутри .doc
        if b"{\\rtf" in header.lower():
            try:
                # В RTF для русских документов часто используется cp1251
                txt = content.decode("cp1251", errors="ignore")

                # Преобразуем последовательности \'hh в символы cp1251
                def _rtf_hex_to_char(match):
                    try:
                        byte_val = int(match.group(1), 16)
                        return bytes([byte_val]).decode("cp1251", errors="ignore")
                    except Exception:
                        return ""

                txt = re.sub(r"\\'([0-9a-fA-F]{2})", lambda m: _rtf_hex_to_char(m), txt)

                # Убираем управляющие последовательности RTF и фигурные скобки
                txt = re.sub(r"\\[a-zA-Z]+\d*", " ", txt)  # команды \b, \par, \fs20 и т.п.
                txt = re.sub(r"[{}]", " ", txt)
                txt = html.unescape(txt)
                txt = re.sub(r"\s+", " ", txt)
                return txt.strip()
            except Exception as e:
                logger.warning(f"Не удалось извлечь текст из RTF внутри DOC: {e}")

        # 2) HTML внутри .doc
        if b"<html" in header.lower() or b"<body" in header.lower():
            try:
                # Пробуем cp1251, затем UTF-8
                try:
                    html_text = content.decode("cp1251")
                except UnicodeDecodeError:
                    html_text = content.decode("utf-8", errors="ignore")

                soup = BeautifulSoup(html_text, "html.parser")
                txt = soup.get_text(separator=" ", strip=True)
                txt = re.sub(r"\s+", " ", txt)
                return txt.strip()
            except Exception as e:
                logger.warning(f"Не удалось извлечь HTML-текст из DOC: {e}")

        # 3) Бинарный DOC — грубая попытка вытащить видимый текст
        try:
            # Сначала cp1251, чтобы не сломать кириллицу
            txt = content.decode("cp1251", errors="ignore")
            # Оставляем только печатные символы и пробелы
            txt = re.sub(r"[^\x20-\x7E\xC0-\xFF\n\r\t]", " ", txt)
            txt = re.sub(r"\s+", " ", txt)
            return txt.strip()
        except Exception as e:
            logger.warning(f"Не удалось извлечь текст из бинарного DOC: {e}")
            return ""

    def _can_use_win32_word(self) -> bool:
        return (
                sys.platform.startswith("win")
                and win32_client is not None
                and win32_constants is not None
                and pythoncom is not None
        )

    def _extract_text_from_doc_with_word(self, content: bytes) -> str:
        """Использует установленный Microsoft Word (через win32com) для извлечения текста из .doc."""
        temp_doc = None
        temp_txt = None
        word = None
        try:
            pythoncom.CoInitialize()
            temp = tempfile.NamedTemporaryFile(delete=False, suffix=".doc")
            temp.write(content)
            temp_doc = temp.name
            temp.close()

            temp_txt = temp_doc + ".txt"

            word = win32_client.Dispatch("Word.Application")  # type: ignore
            word.Visible = False

            doc = word.Documents.Open(temp_doc)  # type: ignore
            wd_format_text = getattr(win32_constants, "wdFormatText", 2)
            doc.SaveAs(temp_txt, FileFormat=wd_format_text)  # type: ignore
            doc.Close(False)  # type: ignore

            with open(temp_txt, "rb") as f:
                raw = f.read()

            encodings_to_try = [
                "utf-8",
                locale.getpreferredencoding(False) or "cp1251",
                "cp1251",
            ]
            text = ""
            for enc in encodings_to_try:
                try:
                    text = raw.decode(enc)
                    break
                except Exception:
                    continue
            if not text:
                text = raw.decode("latin-1", errors="ignore")

            return " ".join(text.split())

        except Exception as e:
            logger.warning(f"Не удалось извлечь текст из DOC через Word COM: {e}")
            return ""
        finally:
            if word:
                try:
                    word.Quit()
                except Exception:
                    pass
            if pythoncom is not None:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
            for path in (temp_doc, temp_txt):
                if not path:
                    continue
                try:
                    os.remove(path)
                except Exception:
                    pass

    def _extract_text_from_xlsx(self, content: bytes) -> str:
        """Извлекает текст из .xlsx/.xlsm с помощью openpyxl."""
        if load_workbook is None:
            logger.warning("openpyxl не установлен, использую простой декодер")
            return ""

        try:
            wb = load_workbook(BytesIO(content), data_only=True, read_only=True)
            parts = []

            for ws in wb.worksheets:
                for row in ws.iter_rows():
                    row_vals = []
                    for cell in row:
                        v = cell.value
                        if v is None:
                            continue
                        row_vals.append(str(v))
                    if row_vals:
                        parts.append(" ".join(row_vals))

            return "\n".join(parts)
        except Exception as e:
            logger.warning(f"Не удалось извлечь текст из XLSX: {e}")
            return ""

    def _extract_text_from_xls(self, content: bytes) -> str:
        """Извлекает текст из .xls с помощью xlrd (если установлен)."""
        if xlrd is None:
            logger.warning("xlrd не установлен, использую простой декодер")
            return ""
        try:
            book = xlrd.open_workbook(file_contents=content)
            parts = []
            for sheet in book.sheets():
                for rx in range(sheet.nrows):
                    row_vals = sheet.row_values(rx)
                    row_vals = [str(v) for v in row_vals if v not in ("", None)]
                    if row_vals:
                        parts.append(" ".join(row_vals))
            return "\n".join(parts)
        except Exception as e:
            logger.warning(f"Не удалось извлечь текст из XLS: {e}")
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
