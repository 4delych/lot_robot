import time
import re
import logging
from urllib.parse import urljoin, urlparse, parse_qs

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from bs4 import BeautifulSoup

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
                        f"Скачивание {i + 1}/{len(unique_links)}: {doc['name'][:30]}..."
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
