"""
Модуль для работы с разными источниками данных о закупках.
Каждый источник реализует единый интерфейс для поиска и парсинга.
"""
import re
import logging
from abc import ABC, abstractmethod
from urllib.parse import urljoin, urlparse, quote_plus, urlencode
from bs4 import BeautifulSoup
import requests

logger = logging.getLogger(__name__)


class ProcurementSource(ABC):
    """Базовый класс для источников данных о закупках."""

    @abstractmethod
    def get_name(self) -> str:
        """Возвращает название источника."""
        pass

    @abstractmethod
    def build_search_url(self, keyword, min_price=None, max_price=None, 
                        purchase_stage=None, law=None) -> tuple[str, dict]:
        """
        Строит URL и параметры для поиска.
        Возвращает (url, params_dict).
        """
        pass

    @abstractmethod
    def parse_results(self, html_content: str, progress_callback=None) -> list[dict]:
        """
        Парсит HTML и возвращает список лотов.
        Каждый лот: {"Название": str, "Цена": float, "Ссылка": str}
        """
        pass

    @abstractmethod
    def get_documents_url(self, lot_url: str) -> tuple[str, dict] | None:
        """
        Получает URL страницы с документами для лота.
        Возвращает (url, params_dict) или None, если не поддерживается.
        """
        pass

    def _to_int_or_none(self, value):
        """Преобразует значение в int или None."""
        if value is None:
            return None
        if isinstance(value, (int, float)):
            return int(value)
        s = str(value).strip()
        if not s:
            return None
        s = s.replace("\u00a0", "").replace(" ", "").replace(",", ".")
        try:
            return int(float(s))
        except ValueError:
            return None


class ZakupkiGovSource(ProcurementSource):
    """Источник данных с zakupki.gov.ru (текущая реализация)."""

    def __init__(self, base_url="https://zakupki.gov.ru"):
        self.base_url = base_url

    def get_name(self) -> str:
        return "zakupki.gov.ru"

    def build_search_url(self, keyword, min_price=None, max_price=None,
                        purchase_stage=None, law=None) -> tuple[str, dict]:
        url = f"{self.base_url}/epz/order/extendedsearch/results.html"
        params = {
            "searchString": keyword.strip(),
            "pageNumber": "1",
            "recordsPerPage": "_20",
        }

        # Purchase stage
        stage_mapping = {
            "SUBMISSION": {"af": "on"},
            "EVALUATION": {"ca": "on"},
            "COMPLETED": {"pc": "on"},
            "CANCELLED": {"pa": "on"},
        }
        if purchase_stage and purchase_stage in stage_mapping:
            params.update(stage_mapping[purchase_stage])

        # Law
        law_mapping = {
            "44FZ": {"fz44": "on"},
            "223FZ": {"fz223": "on"},
            "PP615": {"ppRf615": "on"},
        }
        if law and law in law_mapping:
            params.update(law_mapping[law])

        # Price
        if min_price is not None or max_price is not None:
            min_val = self._to_int_or_none(min_price)
            max_val = self._to_int_or_none(max_price)
            if min_val is not None:
                params["priceFromGeneral"] = str(min_val)
            if max_val is not None:
                params["priceToGeneral"] = str(max_val)
            params["currencyIdGeneral"] = "-1"

        return url, params

    def parse_results(self, html_content: str, progress_callback=None) -> list[dict]:
        # Убеждаемся, что контент правильно декодирован
        if isinstance(html_content, bytes):
            html_content = html_content.decode('utf-8', errors='ignore')
        
        soup = BeautifulSoup(html_content, "html.parser")
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
            return []

        results = []
        for i, lot in enumerate(lots):
            if progress_callback:
                progress_callback(f"Обработка результата {i + 1} из {len(lots)}...")

            try:
                title = self._extract_title(lot)
                price = self._extract_price(lot)
                link = self._extract_link(lot)

                if title and title != "Название не найдено":
                    results.append({"Название": title, "Цена": price, "Ссылка": link})
            except Exception as e:
                logger.warning(f"Failed to parse lot {i + 1}: {e}")
                continue

        return results

    def _extract_title(self, lot):
        header_text = None
        object_text = None

        try:
            header_el = lot.select_one(".registry-entry__header-top__title")
            if header_el:
                # Явно указываем кодировку и нормализуем пробелы
                header_text = header_el.get_text(separator=" ", strip=True)
                if header_text:
                    header_text = " ".join(header_text.split())
        except Exception as e:
            logger.debug(f"Error extracting header: {e}")

        try:
            for body_block in lot.select(".registry-entry__body-block"):
                title_el = body_block.select_one(".registry-entry__body-title")
                if title_el:
                    title_label = title_el.get_text(strip=True).lower()
                    if "объект закупки" in title_label:
                        value_el = body_block.select_one(".registry-entry__body-value")
                        if value_el:
                            object_text = value_el.get_text(separator=" ", strip=True)
                            if object_text:
                                object_text = " ".join(object_text.split())
                        break
        except Exception as e:
            logger.debug(f"Error extracting object text: {e}")

        if header_text and object_text:
            return f"{header_text} — {object_text}"
        if object_text:
            return object_text
        if header_text:
            return header_text
        
        # Дополнительные попытки найти название
        try:
            title_els = lot.select("h2, h3, [class*='title']")
            for el in title_els:
                text = el.get_text(separator=" ", strip=True)
                if text and len(text) > 5:
                    return " ".join(text.split())
        except Exception:
            pass

        return "Название не найдено"

    def _extract_price(self, lot):
        price_selectors = [
            ".price-block__value",
            ".price",
            "[class*='price']",
            ".cost",
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
                except (ValueError, AttributeError):
                    continue

        return 0.0

    def _extract_link(self, lot):
        link_el = lot.select_one(".registry-entry__header-mid__number a[href]")
        if link_el:
            href = link_el["href"].strip()
            if href:
                return urljoin(self.base_url, href)

        candidates = lot.select("a[href*='notice/']")
        for link in candidates:
            href = (link.get("href") or "").strip()
            if href and "printForm" not in href.lower() and "view" in href.lower():
                return urljoin(self.base_url, href)

        return "Ссылка не найдена"

    def get_documents_url(self, lot_url: str) -> tuple[str, dict] | None:
        parsed_url = urlparse(lot_url)
        path = parsed_url.path or ""
        query_dict = {}
        if parsed_url.query:
            from urllib.parse import parse_qs
            query_dict = {k: v[0] for k, v in parse_qs(parsed_url.query).items() if v}

        if "/epz/order/notice/" in path and "view/" in path:
            if "common-info.html" in path:
                docs_path = path.replace("common-info.html", "documents.html")
            else:
                docs_path = path
            documents_url = urljoin(self.base_url, docs_path)
            return documents_url, query_dict

        # 223-ФЗ
        notice_info_id = query_dict.get("noticeInfoId")
        if notice_info_id:
            documents_url = f"{self.base_url}/epz/order/notice/notice223/documents.html"
            return documents_url, {
                "noticeInfoId": notice_info_id,
                "backUrl": "/epz/order/notice/notice223/search.html",
            }

        return None


class TektorgSource(ProcurementSource):
    """Источник данных с tektorg.ru."""

    def __init__(self, base_url="https://www.tektorg.ru"):
        self.base_url = base_url

    def get_name(self) -> str:
        return "tektorg.ru"

    def build_search_url(self, keyword, min_price=None, max_price=None,
                        purchase_stage=None, law=None) -> tuple[str, dict]:
        """
        Формирует URL для поиска на tektorg.ru.
        Возвращает (url, params), где params будет использован requests для формирования query string.
        """
        url = f"{self.base_url}/procedures"
        params = {}

        # Ключевое слово
        if keyword and keyword.strip():
            params["name"] = keyword.strip()

        # Price - формат: "200+000" (с + вместо пробелов в URL)
        if min_price is not None:
            min_val = self._to_int_or_none(min_price)
            if min_val is not None:
                # Форматируем с пробелами для разделения тысяч: 200000 -> "200 000"
                # Но в URL нужно использовать + вместо пробелов
                price_str = f"{min_val:,}".replace(",", " ").replace(" ", "+")
                params["sumPrice_start"] = price_str
        if max_price is not None:
            max_val = self._to_int_or_none(max_price)
            if max_val is not None:
                price_str = f"{max_val:,}".replace(",", " ").replace(" ", "+")
                params["sumPrice_end"] = price_str

        # Status - формат: "Приём заявок;Работа комиссии"
        # Важно: если purchase_stage указан, используем только его, иначе все статусы
        statuses = []
        if purchase_stage:
            stage_mapping = {
                "SUBMISSION": "Приём заявок",
                "EVALUATION": "Работа комиссии",
                "COMPLETED": "Архив",
                "CANCELLED": "Отменён",
            }
            if purchase_stage in stage_mapping:
                statuses.append(stage_mapping[purchase_stage])
        else:
            # По умолчанию все статусы (как в примере пользователя)
            statuses = ["Приём заявок", "Работа комиссии", "Архив", "Отменен", "Отменён"]

        if statuses:
            # Используем ключ "status[]" - requests правильно закодирует [] в %5B%5D
            # Значения будут автоматически URL-кодированы requests
            status_value = ";".join(statuses)
            params["status[]"] = status_value

        # Law - формат: "44fz;zakupki" (223-ФЗ это "zakupki")
        sections = []
        if law:
            law_mapping = {
                "44FZ": "44fz",
                "223FZ": "zakupki",
            }
            if law in law_mapping:
                sections.append(law_mapping[law])
        else:
            # По умолчанию оба (как в примере пользователя)
            sections = ["44fz", "zakupki"]

        if sections:
            # Используем ключ "sectionsCodes[]" - requests правильно закодирует [] в %5B%5D
            params["sectionsCodes[]"] = ";".join(sections)

        logger.debug(f"Tektorg URL params: {params}")
        return url, params

    def parse_results(self, html_content: str, progress_callback=None) -> list[dict]:
        # Убеждаемся, что контент правильно декодирован
        if isinstance(html_content, bytes):
            html_content = html_content.decode('utf-8', errors='ignore')
        
        soup = BeautifulSoup(html_content, "html.parser")
        results = []

        # Ищем ссылки на процедуры - более точный поиск
        # Ссылки вида: /44-fz/procedures/12345678 или /procedures/12345678
        all_links = soup.find_all("a", href=True)
        procedure_links = []
        seen_hrefs = set()
        
        for link in all_links:
            href = link.get("href", "").strip()
            if not href:
                continue
            
            # Проверяем, что это ссылка на процедуру (содержит ID)
            if "/procedures/" in href:
                # Извлекаем ID процедуры
                parts = href.split("/procedures/")
                if len(parts) > 1:
                    procedure_id = parts[1].split("/")[0].split("?")[0]
                    # ID должен быть числом
                    if procedure_id and procedure_id.isdigit() and href not in seen_hrefs:
                        seen_hrefs.add(href)
                        procedure_links.append(link)

        if not procedure_links:
            logger.warning("No procedure links found on tektorg.ru page")
            # Сохраняем HTML для отладки
            with open("debug_tektorg_page.html", "w", encoding="utf-8") as f:
                f.write(soup.prettify())
            return []

        logger.info(f"Found {len(procedure_links)} procedure links on tektorg.ru")

        for i, link in enumerate(procedure_links):
            if progress_callback:
                progress_callback(f"Обработка результата {i + 1} из {len(procedure_links)}...")

            try:
                href = link.get("href", "").strip()
                
                # Полный URL
                if href.startswith("/"):
                    full_url = urljoin(self.base_url, href)
                elif href.startswith("http"):
                    full_url = href
                else:
                    continue

                # Название - ищем более тщательно
                title = self._extract_title_tektorg(link)
                if not title or len(title.strip()) < 5:
                    # Если не нашли в ссылке, ищем в родительском контейнере
                    parent = link.find_parent()
                    if parent:
                        title = self._extract_title_from_container(parent)
                
                if not title or len(title.strip()) < 5:
                    logger.warning(f"Could not extract title for {full_url}, skipping")
                    continue
                
                # Очищаем название от лишних символов
                title = " ".join(title.split())

                # Цена - ищем в контейнере карточки
                price = self._extract_price_tektorg(link)

                results.append({
                    "Название": title.strip(),
                    "Цена": price,
                    "Ссылка": full_url,
                })
            except Exception as e:
                logger.warning(f"Failed to parse tektorg lot {i + 1}: {e}")
                import traceback
                logger.debug(traceback.format_exc())
                continue

        logger.info(f"Successfully parsed {len(results)} results from tektorg.ru")
        return results

    def _extract_title_tektorg(self, lot_element):
        """Извлекает название из элемента tektorg.ru."""
        # Пробуем найти текст в самом элементе
        text = lot_element.get_text(separator=" ", strip=True)
        if text and len(text) > 10:
            # Убираем лишние пробелы и берем первые 200 символов
            text = " ".join(text.split())
            return text[:200]

        # Ищем в атрибутах (title, aria-label и т.д.)
        for attr in ["title", "aria-label", "data-title"]:
            attr_text = lot_element.get(attr, "").strip()
            if attr_text and len(attr_text) > 10:
                return attr_text[:200]

        return None

    def _extract_title_from_container(self, container):
        """Извлекает название из контейнера карточки процедуры."""
        # Ищем заголовки и элементы с названием
        title_selectors = [
            "h1", "h2", "h3", "h4",
            "[class*='title']",
            "[class*='name']",
            "[class*='heading']",
        ]
        
        for selector in title_selectors:
            els = container.select(selector)
            for el in els:
                text = el.get_text(separator=" ", strip=True)
                if text and len(text) > 10:
                    return " ".join(text.split())[:200]
        
        # Если не нашли, берем весь текст контейнера, но только первые строки
        all_text = container.get_text(separator=" ", strip=True)
        if all_text:
            # Берем первые 200 символов и убираем лишние пробелы
            text = " ".join(all_text.split())[:200]
            if len(text) > 10:
                return text
        
        return None

    def _extract_price_tektorg(self, lot_element):
        """Извлекает цену из элемента tektorg.ru."""
        # Ищем цену в родительских элементах
        parent = lot_element.parent
        for _ in range(5):
            if not parent:
                break

            # Ищем элементы с ценой
            price_elements = parent.select("[class*='price'], [class*='sum'], [class*='cost']")
            for price_el in price_elements:
                price_text = price_el.get_text(strip=True)
                # Извлекаем числа
                numbers = re.findall(r"[\d\s]+", price_text.replace(",", "."))
                for num_str in numbers:
                    try:
                        # Убираем пробелы и преобразуем
                        num_clean = num_str.replace(" ", "").replace("\xa0", "")
                        if num_clean:
                            price = float(num_clean)
                            if price > 0:
                                return price
                    except ValueError:
                        continue

            parent = parent.parent

        return 0.0

    def get_documents_url(self, lot_url: str) -> tuple[str, dict] | None:
        """Для tektorg.ru документы обычно на той же странице лота."""
        # Документы находятся на странице лота, нужно будет парсить их оттуда
        # Пока возвращаем None - реализуем позже
        return None

