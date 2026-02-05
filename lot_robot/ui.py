import os
import sys
import subprocess
import tempfile
import re
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import threading
import pandas as pd
import logging
import webbrowser
import json


from config import PURCHASE_STAGES, LAWS, CONFIG
from ProcurmentSearcher import ProcurementSearcher
logger = logging.getLogger(__name__)

class ProcurementApp:
    """Main application class with improved UI and threading."""

    def __init__(self, root):
        self.root = root
        self.root.title("Tender Search & Analysis Robot")
        # Базовый размер под десктоп ≥1280 px
        self.root.geometry("1280x800")
        # состояние UI
        self.filters_collapsed = tk.BooleanVar(value=False)
        self.results = []
        self.searcher = ProcurementSearcher()
        self.search_thread = None
        self.analysis_results = []
        self._temp_doc_files = []
        self._lot_llm_cache = {}
        self._lot_llm_cache_path = None
        self._lot_titles_dump_path = None
        self._lot_titles_sent_path = None
        try:
            tmp = tempfile.NamedTemporaryFile(
                prefix="lot_title_cache_", suffix=".json", delete=False
            )
            self._lot_llm_cache_path = tmp.name
            tmp.close()
            self._lot_titles_dump_path = os.path.join(os.getcwd(), "lot_titles_all.txt")
            self._lot_titles_sent_path = os.path.join(os.getcwd(), "lot_titles_sent.txt")
        except Exception as e:
            logger.warning("Failed to create temp cache file: %s", e)
        # состояние правой панели анализа
        self.current_lot = None
        self.analysis_state_var = tk.StringVar(value="Idle")
        self.verdict_label_var = tk.StringVar(value="Анализ не выполнен")
        self.verdict_explanation_var = tk.StringVar(value="")
        
        self.llm_provider_options = sorted((CONFIG.get("LLM_PROVIDERS") or {}).keys())
        default_provider = (CONFIG.get("LLM_PROVIDER") or "").strip()
        if not default_provider and self.llm_provider_options:
            default_provider = self.llm_provider_options[0]
            CONFIG["LLM_PROVIDER"] = default_provider
        self.llm_provider_var = tk.StringVar(value=default_provider)
        self._setup_ui()
        self._load_lot_llm_cache()

        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)

    def _setup_ui(self):
        """Setup the user interface (split layout: search/results + analysis panel)."""
        # Общий фон, близкий к референсу (тёмная шапка, светлый контент)
        self.root.configure(bg="#111827")

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # Общий стиль ttk (таблица, заголовки и т.п.)
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure(
            "App.Treeview",
            font=("Segoe UI", 9),
            rowheight=22,
            borderwidth=0,
            background="#FFFFFF",
            fieldbackground="#FFFFFF",
        )
        style.configure(
            "App.Treeview.Heading",
            font=("Segoe UI", 9, "bold"),
            background="#E5E7EB",
            foreground="#111827",
            borderwidth=0,
        )
        style.map(
            "App.Treeview",
            background=[("selected", "#DBEAFE")],
            foreground=[("selected", "#111827")],
        )

        # Корневой контейнер
        container = tk.Frame(self.root, bg="#111827")
        container.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=0)  # шапка
        container.rowconfigure(1, weight=1)  # основное содержимое

        # Верхняя шапка приложения
        header = tk.Frame(container, bg="#374151", height=56)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_propagate(False)
        header.columnconfigure(1, weight=1)

        # Логотип (упрощённый прямоугольник с инициалами)
        logo_frame = tk.Frame(header, bg="#374151")
        logo_frame.grid(row=0, column=0, padx=(16, 12), pady=8, sticky="w")

        logo = tk.Label(
            logo_frame,
            text="ST",
            font=("Segoe UI Semibold", 14),
            bg="#0B5FFF",
            fg="#FFFFFF",
            padx=10,
            pady=4,
        )
        logo.pack()

        title_frame = tk.Frame(header, bg="#374151")
        title_frame.grid(row=0, column=1, sticky="w")

        title_label = tk.Label(
            title_frame,
            text="Tender Search & Analysis Robot",
            font=("Segoe UI Semibold", 14),
            bg="#374151",
            fg="#F9FAFB",
        )
        title_label.pack(anchor="w")

        subtitle_label = tk.Label(
            title_frame,
            text="Аналитика закупок и AI‑оценка лотов",
            font=("Segoe UI", 10),
            bg="#374151",
            fg="#D1D5DB",
        )
        subtitle_label.pack(anchor="w")

        # Основная область с панелями
        content = ttk.Frame(container, padding=10)
        content.grid(row=1, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        content.columnconfigure(0, weight=1)
        content.rowconfigure(0, weight=1)

        # Горизонтальный сплит
        paned = ttk.Panedwindow(content, orient=tk.HORIZONTAL)
        paned.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))

        # Левая панель (поиск/фильтры/результаты)
        left_frame = ttk.Frame(paned, padding=(8, 0, 8, 0))
        self.left_frame = left_frame
        left_frame.columnconfigure(1, weight=1)
        left_frame.rowconfigure(7, weight=0)  # блок документов
        left_frame.rowconfigure(8, weight=1)  # таблица результатов

        # Правая панель (анализ выбранного лота)
        right_frame = ttk.Frame(paned, padding=(0, 0, 0, 0))
        right_frame.columnconfigure(0, weight=1)
        right_frame.rowconfigure(2, weight=1)

        paned.add(left_frame, weight=3)
        paned.add(right_frame, weight=2)

        # Filters header (collapsible controls)
        filters_header = ttk.Frame(left_frame)
        filters_header.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(4, 0))
        filters_header.columnconfigure(0, weight=1)

        ttk.Label(
            filters_header,
            text="Фильтры",
            font=("Segoe UI", 10, "bold"),
        ).grid(row=0, column=0, sticky="w")

        self.filters_toggle_btn = ttk.Button(
            filters_header,
            text="Свернуть",
            command=self._toggle_filters,
            width=10,
        )
        self.filters_toggle_btn.grid(row=0, column=1, sticky="e")

        # Price filters frame
        price_frame = ttk.LabelFrame(
            left_frame, text="Фильтр по цене (руб)", padding="5"
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
            left_frame, text="Дополнительные фильтры", padding="5"
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


        # Sources selection frame
        sources_frame = ttk.LabelFrame(
            left_frame, text="Источники поиска", padding="5"
        )
        sources_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)

        self.source_vars = {}
        available_sources = ["zakupki.gov.ru", "tektorg.ru", "bidzaar.com"]
        for i, source_name in enumerate(available_sources):
            var = tk.BooleanVar(value=True)  # По умолчанию все включены
            self.source_vars[source_name] = var
            cb = ttk.Checkbutton(
                sources_frame,
                text=source_name,
                variable=var
            )
            cb.grid(row=0, column=i, sticky="w", padx=10)

        # Keywords selection frame
        keywords_frame = ttk.LabelFrame(
            left_frame, text="Ключевые слова", padding="5"
        )
        keywords_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        keywords_frame.columnconfigure(0, weight=1)

        self.search_keyword_vars = {}
        self.search_keyword_list = [k.strip() for k in (CONFIG.get("SEARCH_KEYWORDS") or []) if isinstance(k, str) and k.strip()]

        if not self.search_keyword_list:
            ttk.Label(
                keywords_frame,
                text="Список ключевых слов пуст. Заполните SEARCH_KEYWORDS в config.py",
                foreground="#B91C1C",
                wraplength=420,
            ).grid(row=0, column=0, sticky="w", pady=2)
        else:
            for i, kw in enumerate(self.search_keyword_list):
                var = tk.BooleanVar(value=True)
                self.search_keyword_vars[kw] = var
                cb = ttk.Checkbutton(keywords_frame, text=kw, variable=var)
                cb.grid(row=i, column=0, sticky="w")

        self.select_all_btn = ttk.Button(
            keywords_frame, text="Снять все", command=self._toggle_select_all_keywords
        )
        if not self.search_keyword_list:
            self.select_all_btn.config(state="disabled", text="Выбрать все")
        self.select_all_btn.grid(row=0, column=1, sticky="e", padx=(10, 0))

        self._filter_blocks = [price_frame, filters_frame, sources_frame, keywords_frame]

        # Buttons frame
        buttons_frame = ttk.Frame(left_frame)
        buttons_frame.grid(row=5, column=0, columnspan=2, pady=10)

        self.search_btn = ttk.Button(
            buttons_frame, text="Поиск", command=self.search
        )
        self.search_btn.pack(side=tk.LEFT, padx=5)

        self.export_btn = ttk.Button(
            buttons_frame,
            text="Сохранить в Excel",
            command=self.save_to_excel,
            state="disabled",
        )
        self.export_btn.pack(side=tk.LEFT, padx=5)

        self.clear_btn = ttk.Button(
            buttons_frame, text="Очистить", command=self.clear_results
        )
        self.clear_btn.pack(side=tk.LEFT, padx=5)

        self.reset_cache_btn = ttk.Button(
            buttons_frame, text="Сброс кэша", command=self._reset_llm_cache
        )
        self.reset_cache_btn.pack(side=tk.LEFT, padx=5)

        # Progress bar
        self.progress_var = tk.StringVar(value="Готов к поиску")
        progress_frame = ttk.Frame(left_frame)
        progress_frame.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        progress_frame.columnconfigure(0, weight=1)

        self.progress_bar = ttk.Progressbar(progress_frame, mode="indeterminate")
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E))

        self.status_label = ttk.Label(progress_frame, textvariable=self.progress_var)
        self.status_label.grid(row=1, column=0, sticky="w")

        # Документы frame
        doc_frame = ttk.LabelFrame(
            left_frame, text="Поиск в документах ТЗ", padding="5"
        )
        doc_frame.grid(row=7, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
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
            text="Анализ документов лота",
            command=self.analyze_documents,
        )
        self.analyze_btn.pack(side=tk.LEFT, padx=5)

        self.export_docs_btn = ttk.Button(
            doc_buttons_frame,
            text="Экспорт анализа документов",
            command=self.export_analysis,
            state="disabled",
        )
        self.export_docs_btn.pack(side=tk.LEFT, padx=5)

        # Results table
        self._setup_results_table(left_frame)

        # Правая панель анализа
        self._setup_analysis_panel(right_frame)

    def _setup_results_table(self, parent):
        """Setup the results table with scrollbars."""
        table_frame = ttk.Frame(parent)
        table_frame.grid(
            row=8, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10
        )
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)

        columns = ("Название", "Цена", "Оценка", "Источник", "Ссылка")
        self.tree_default_height = 15
        self.tree_expanded_height = 25
        self.tree = ttk.Treeview(
            table_frame,
            columns=columns,
            show="headings",
            height=self.tree_default_height,
            style="App.Treeview",
        )

        self.tree.heading("Название", text="Название", anchor="w")
        self.tree.heading("Цена", text="Цена (руб)", anchor="e")
        self.tree.heading("Оценка", text="Оценка", anchor="center", command=self._toggle_score_filter_popup)
        self.tree.heading("Источник", text="Источник", anchor="w")
        self.tree.heading("Ссылка", text="Ссылка", anchor="w")

        self.tree.column("Название", width=330, anchor="w")
        self.tree.column("Цена", width=120, anchor="e")
        self.tree.column("Оценка", width=80, anchor="center")
        self.tree.column("Источник", width=120, anchor="w")
        self.tree.column("Ссылка", width=280, anchor="w")

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
        self.tree.bind("<<TreeviewSelect>>", self._on_result_select)

        self._score_filter_set = None
        self._score_filter_popup = None

    def _toggle_filters(self):
        """Сворачивает/разворачивает блоки фильтров, не меняя их состояния."""
        new_state = not self.filters_collapsed.get()
        self.filters_collapsed.set(new_state)

        # Все блоки фильтров находятся в строках 2–4 левой панели:
        # price_frame, filters_frame, sources_frame.
        # Ищем их по grid_info и скрываем/показываем.
        blocks = getattr(self, "_filter_blocks", [])
        for child in blocks:
            if new_state:
                child.grid_remove()
            else:
                child.grid()

        self.filters_toggle_btn.config(text="Развернуть" if new_state else "Свернуть")
        if getattr(self, "tree", None):
            target_height = self.tree_expanded_height if new_state else self.tree_default_height
            self.tree.configure(height=target_height)

    def _toggle_score_filter_popup(self):
        if self._score_filter_popup and self._score_filter_popup.winfo_exists():
            self._score_filter_popup.destroy()
            self._score_filter_popup = None
            return

        popup = tk.Toplevel(self.root)
        popup.title("Фильтр оценки")
        popup.resizable(False, False)
        popup.transient(self.root)
        popup.grab_set()

        self._score_filter_popup = popup

        current = self._score_filter_set or set([1, 2, 3, 4, 5])
        vars_map = {}
        for i, score in enumerate([5, 4, 3, 2, 1]):
            var = tk.BooleanVar(value=score in current)
            vars_map[score] = var
            ttk.Checkbutton(popup, text=f"{score} ★", variable=var).grid(
                row=i, column=0, sticky="w", padx=10, pady=2
            )

        def set_all(state: bool):
            for var in vars_map.values():
                var.set(state)

        actions = ttk.Frame(popup)
        actions.grid(row=6, column=0, padx=10, pady=(6, 8), sticky="e")
        ttk.Button(actions, text="Все", command=lambda: set_all(True)).pack(side=tk.LEFT, padx=4)
        ttk.Button(actions, text="Ни одной", command=lambda: set_all(False)).pack(side=tk.LEFT, padx=4)

        def apply_filter():
            selected = {score for score, var in vars_map.items() if var.get()}
            if not selected:
                self._score_filter_set = None
            else:
                self._score_filter_set = selected
            self._render_results(self._get_filtered_results())
            popup.destroy()
            self._score_filter_popup = None

        ttk.Button(actions, text="Применить", command=apply_filter).pack(side=tk.LEFT, padx=4)

    def _get_filtered_results(self):
        if not self._score_filter_set:
            return self.results
        filtered = []
        for lot in self.results:
            score = lot.get("_score", 1)
            if score in self._score_filter_set:
                filtered.append(lot)
        return filtered

    def _render_results(self, results):
        self.tree.delete(*self.tree.get_children())
        for result in results:
            price_display = (
                f"{result['Цена']:,.2f}" if result["Цена"] > 0 else "Не указана"
            )
            score_display = self._format_score(result.get("_score", 1))
            source = result.get("Источник", "Неизвестно")
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
                    score_display,
                    source,
                    result["Ссылка"],
                ),
            )

    def _format_score(self, score: int) -> str:
        score = int(score) if isinstance(score, int) else 1
        if score < 1 or score > 5:
            score = 1
        return f"{score} ★"

    def _setup_analysis_panel(self, parent: ttk.Frame) -> None:
        """Создает правую панель анализа выбранного тендера."""
        # Заголовок панели
        header_frame = ttk.Frame(parent)
        header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        header_frame.columnconfigure(0, weight=1)

        self.detail_title_label = ttk.Label(
            header_frame,
            text="Лот не выбран",
            font=("Segoe UI", 12, "bold"),
            wraplength=380,
        )
        self.detail_title_label.grid(row=0, column=0, sticky="w")

        self.detail_meta_label = ttk.Label(
            header_frame,
            text="Выберите лот слева для просмотра деталей и запуска AI-анализа",
            font=("Segoe UI", 9),
            foreground="#64748B",
            wraplength=380,
        )
        self.detail_meta_label.grid(row=1, column=0, sticky="w", pady=(2, 0))

        # Блок статуса и кнопки запуска анализа
        cta_frame = ttk.LabelFrame(parent, text="AI-анализ документов", padding=8)
        cta_frame.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        cta_frame.columnconfigure(0, weight=1)
        cta_frame.columnconfigure(1, weight=0)

        self.analysis_status_label = ttk.Label(
            cta_frame,
            textvariable=self.analysis_state_var,
            foreground="#64748B",
        )
        self.analysis_status_label.grid(row=0, column=0, sticky="w")

        self.lot_analyze_btn = ttk.Button(
            cta_frame,
            text="Анализировать лот (LLM)",
            command=self.analyze_lot_report,
            state="disabled",
        )
        self.lot_analyze_btn.grid(row=0, column=1, sticky="e")

        
        self.llm_provider_label = ttk.Label(cta_frame, text="LLM provider")
        self.llm_provider_label.grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.llm_provider_combo = ttk.Combobox(
            cta_frame,
            textvariable=self.llm_provider_var,
            values=self.llm_provider_options,
            state="readonly",
            width=28,
        )
        self.llm_provider_combo.grid(row=1, column=1, sticky="e", pady=(6, 0))
        self.llm_provider_combo.bind("<<ComboboxSelected>>", self._on_llm_provider_change)
        # Блок итогового вердикта
        verdict_outer = ttk.Frame(parent)
        verdict_outer.grid(row=2, column=0, sticky="nsew")
        verdict_outer.columnconfigure(0, weight=1)
        verdict_outer.rowconfigure(1, weight=1)

        self.verdict_frame = tk.Frame(
            verdict_outer,
            bg="#E5E7EB",
            bd=2,
            relief="solid",
        )
        self.verdict_frame.grid(row=0, column=0, sticky="ew", pady=(0, 8), padx=(0, 0))
        self.verdict_frame.columnconfigure(0, weight=1)

        self.verdict_title_label = tk.Label(
            self.verdict_frame,
            textvariable=self.verdict_label_var,
            font=("Segoe UI", 11, "bold"),
            anchor="w",
            bg="#E5E7EB",
        )
        self.verdict_title_label.grid(row=0, column=0, sticky="ew", padx=10, pady=(6, 0))

        self.verdict_explanation_label = tk.Label(
            self.verdict_frame,
            textvariable=self.verdict_explanation_var,
            font=("Segoe UI", 9),
            anchor="w",
            justify="left",
            wraplength=380,
            bg="#E5E7EB",
        )
        self.verdict_explanation_label.grid(
            row=1, column=0, sticky="ew", padx=10, pady=(2, 8)
        )

        # Текстовый структурированный отчет
        report_frame = ttk.Frame(verdict_outer)
        report_frame.grid(row=1, column=0, sticky="nsew")
        report_frame.columnconfigure(0, weight=1)
        report_frame.rowconfigure(0, weight=1)

        self.report_text = scrolledtext.ScrolledText(
            report_frame,
            wrap=tk.WORD,
            state=tk.DISABLED,
            font=("Segoe UI", 9),
        )
        self.report_text.grid(row=0, column=0, sticky="nsew")

        # Нижняя панель действий
        actions_frame = ttk.Frame(parent)
        actions_frame.grid(row=3, column=0, sticky="ew", pady=(6, 0))
        actions_frame.columnconfigure(0, weight=1)

        self.export_pdf_btn = ttk.Button(
            actions_frame,
            text="Экспорт в PDF",
            command=self._export_report_to_pdf,
            state="disabled",
        )
        self.export_pdf_btn.grid(row=0, column=0, sticky="w")

        self.copy_report_btn = ttk.Button(
            actions_frame,
            text="Копировать отчет",
            command=self._copy_report_to_clipboard,
            state="disabled",
        )
        self.copy_report_btn.grid(row=0, column=1, sticky="w", padx=(6, 0))

        self.open_source_btn = ttk.Button(
            actions_frame,
            text="Открыть источник",
            command=self._open_current_lot_link,
            state="disabled",
        )
        self.open_source_btn.grid(row=0, column=2, sticky="w", padx=(6, 0))

    def _on_item_double_click(self, event):
        """Handle double-click on table item to open link."""
        item = self.tree.selection()[0] if self.tree.selection() else None
        if item:
            values = self.tree.item(item, "values")
            # Теперь ссылка на 4-й позиции (после Источника)
            if len(values) >= 4 and values[3] != "Ссылка не найдена":
                import webbrowser

                try:
                    webbrowser.open(values[3])
                except Exception as e:
                    messagebox.showerror("Ошибка", f"Не удалось открыть ссылку:\n{e}")

    def _on_result_select(self, event):
        """Обновляет правую панель при выборе строки в таблице."""
        selection = self.tree.selection()
        if not selection:
            self.current_lot = None
            self.detail_title_label.config(text="Лот не выбран")
            self.detail_meta_label.config(
                text="Выберите лот слева для просмотра деталей и запуска AI-анализа"
            )
            self.lot_analyze_btn.config(state="disabled")
            self.open_source_btn.config(state="disabled")
            self.analysis_state_var.set("Idle")
            self._reset_verdict_view()
            return

        item = selection[0]
        values = self.tree.item(item, "values")
        if len(values) < 4:
            return

        title, price, source, url = values[0], values[1], values[2], values[3]
        self.current_lot = {
            "title": title,
            "price": price,
            "source": source,
            "url": url,
        }
        meta_text = f"Цена: {price} · Источник: {source}"
        self.detail_title_label.config(text=title)
        self.detail_meta_label.config(text=meta_text)
        self.lot_analyze_btn.config(state="normal")
        self.open_source_btn.config(
            state="normal" if url and url != "Ссылка не найдена" else "disabled"
        )
        # сбрасываем прошлый результат анализа для нового выбора
        self.analysis_state_var.set("Idle")
        self._reset_verdict_view()

    def _reset_verdict_view(self):
        """Сбрасывает отображение вердикта и отчета."""
        self.verdict_label_var.set("Анализ не выполнен")
        self.verdict_explanation_var.set("")
        self._set_verdict_style("idle")
        self.report_text.config(state=tk.NORMAL)
        self.report_text.delete("1.0", tk.END)
        self.report_text.insert(
            tk.END,
            "Отчет по выбранному лоту будет показан здесь после выполнения AI-анализа.",
        )
        self.report_text.config(state=tk.DISABLED)
        self.export_pdf_btn.config(state="disabled")
        self.copy_report_btn.config(state="disabled")

    def _set_verdict_style(self, verdict_type: str) -> None:
        """Настраивает цвета блока вердикта в зависимости от типа."""
        # idle / suitable / uncertain / not_suitable
        if verdict_type == "suitable":
            bg = "#DCFCE7"
        elif verdict_type == "uncertain":
            bg = "#FEF3C7"
        elif verdict_type == "not_suitable":
            bg = "#FEE2E2"
        else:
            bg = "#E5E7EB"

        self.verdict_frame.configure(bg=bg)
        self.verdict_title_label.configure(bg=bg)
        self.verdict_explanation_label.configure(bg=bg)

    def _derive_verdict_from_summary(self, summary: str) -> tuple[str, str]:
        """
        Пытается вывести финальный вердикт (Suitable / Not suitable / Uncertain)
        из текстового резюме LLM.
        """
        s = (summary or "").lower()
        if not s or s.strip() == "—":
            return "Неопределенно", "uncertain"

        negative_markers = ["не наш профиль", "не подходит", "нецелесообразно", "не рекомендуется"]
        positive_markers = ["подходит", "целесообразно", "рекомендуется", "соответствует профилю"]

        if any(m in s for m in negative_markers):
            return "Не подходит", "not_suitable"
        if any(m in s for m in positive_markers):
            return "Подходит", "suitable"

        return "Неопределенно", "uncertain"

    def _fill_report_text(
        self,
        lot_title: str,
        lot_url: str,
        lot_source: str,
        lot_price: str,
        lot_deadline: str,
        llm_data: dict,
    ) -> None:
        """Формирует структурированный текст отчета в правой панели."""
        subject = llm_data.get("subject", "—")
        work_scope = llm_data.get("work_scope", "—")
        timelines = llm_data.get(
            "work_and_submission_timelines",
            "Сроки работ: —; Сроки подачи: —",
        )
        fit_summary = llm_data.get("fit_summary", "—")

        self.report_text.config(state=tk.NORMAL)
        self.report_text.delete("1.0", tk.END)

        # Финальный вердикт
        verdict_label, verdict_type = self._derive_verdict_from_summary(fit_summary)
        self.verdict_label_var.set(f"Итоговый вердикт: {verdict_label}")
        self.verdict_explanation_var.set(fit_summary or "—")
        self._set_verdict_style(verdict_type)

        # Структурированный отчет
        self.report_text.insert(tk.END, "ОТЧЕТ ПО ЛОТУ\n")
        self.report_text.insert(tk.END, "=" * 60 + "\n\n")

        self.report_text.insert(tk.END, "1) ЦЕЛЬ РАБОТ\n")
        self.report_text.insert(tk.END, "-" * 40 + "\n")
        self.report_text.insert(tk.END, f"{subject or '—'}\n\n")

        self.report_text.insert(tk.END, "2) ЗАДАЧИ / СОСТАВ РАБОТ\n")
        self.report_text.insert(tk.END, "-" * 40 + "\n")
        if work_scope and work_scope != "—":
            # разбиваем по ';' для удобства чтения
            parts = [p.strip() for p in work_scope.split(";") if p.strip()]
            for p in parts:
                self.report_text.insert(tk.END, f" • {p}\n")
            self.report_text.insert(tk.END, "\n")
        else:
            self.report_text.insert(tk.END, "—\n\n")

        self.report_text.insert(tk.END, "3) ОБЪЕМ И СРОКИ РАБОТ / ПОДАЧИ\n")
        self.report_text.insert(tk.END, "-" * 40 + "\n")
        self.report_text.insert(tk.END, f"{timelines}\n\n")

        self.report_text.insert(tk.END, "4) СРОКИ ЗАКУПКИ\n")
        self.report_text.insert(tk.END, "-" * 40 + "\n")
        self.report_text.insert(
            tk.END,
            f"Окончание подачи заявок (с сайта): {lot_deadline or '—'}\n\n",
        )

        self.report_text.insert(tk.END, "5) ОБЩИЕ СВЕДЕНИЯ О ЛОТЕ\n")
        self.report_text.insert(tk.END, "-" * 40 + "\n")
        self.report_text.insert(tk.END, f"Наименование: {lot_title}\n")
        self.report_text.insert(tk.END, f"Ссылка: {lot_url}\n")
        self.report_text.insert(tk.END, f"Площадка: {lot_source}\n")
        self.report_text.insert(tk.END, f"Стоимость: {lot_price}\n\n")

        self.report_text.config(state=tk.DISABLED)

        # включаем действия
        self.export_pdf_btn.config(state="normal")
        self.copy_report_btn.config(state="normal")

    def _export_report_to_pdf(self):
        """Простой экспорт отчета в текстовый PDF-файл (как plain text)."""
        content = self.report_text.get("1.0", tk.END).strip()
        if not content:
            messagebox.showwarning("Нет данных", "Отчет еще не сформирован")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF файлы", "*.pdf"), ("Все файлы", "*.*")],
            title="Сохранить отчет",
        )
        if not file_path:
            return

        try:
            # сохраняем как обычный текстовый контент — без сторонних библиотек PDF
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(content)
            messagebox.showinfo("Сохранено", f"Отчет сохранен в файл:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить отчет:\n{e}")

    def _copy_report_to_clipboard(self):
        """Копирует текст отчета в буфер обмена."""
        content = self.report_text.get("1.0", tk.END).strip()
        if not content:
            messagebox.showwarning("Нет данных", "Отчет еще не сформирован")
            return
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(content)
            messagebox.showinfo("Скопировано", "Отчет скопирован в буфер обмена")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось скопировать отчет:\n{e}")

    def _open_current_lot_link(self):
        """Открывает ссылку на текущий лот в браузере."""
        if not self.current_lot:
            return
        url = self.current_lot.get("url")
        if not url or url == "Ссылка не найдена":
            messagebox.showerror("Ошибка", "Неверная ссылка на лот")
            return
        try:
            webbrowser.open(url)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось открыть ссылку:\n{e}")

    def search(self):
        """Start search in a separate thread."""
        if self.search_thread and self.search_thread.is_alive():
            messagebox.showwarning("Поиск", "Поиск уже выполняется...")
            return

        keywords = []
        for kw, var in (self.search_keyword_vars or {}).items():
            if var.get():
                keywords.append(kw)
        if not keywords:
            messagebox.showerror(
                "Ошибка",
                "Выберите хотя бы одно ключевое слово для поиска.",
            )
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

        # Get selected sources
        selected_sources = [
            name for name, var in self.source_vars.items() if var.get()
        ]
        if not selected_sources:
            messagebox.showerror("Ошибка", "Выберите хотя бы один источник для поиска")
            return

        # Start search thread
        self.search_thread = threading.Thread(
            target=self._perform_search,
            args=(keywords, min_price, max_price, purchase_stage, law, selected_sources),
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

    def _perform_search(self, keywords, min_price, max_price, purchase_stage, law, source_names=None):
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
            self._render_results(self._get_filtered_results())

            # Update status
            if results:
                self.progress_var.set(f"Найдено результатов: {len(results)} (с оценкой)")
                self.export_btn.config(state="normal")
            else:
                self.progress_var.set("Результаты не найдены")
                messagebox.showinfo("Результат", "По вашему запросу ничего не найдено")

        # Start UI updates
        self.root.after(0, update_ui_start)

        try:
            search_keywords = [k.strip() for k in (keywords or []) if isinstance(k, str) and k.strip()]
            if not search_keywords:
                raise ValueError("SEARCH_KEYWORDS пустой. Добавьте слова в config.py")

            all_results = []
            total_keywords = len(search_keywords)
            for idx, kw in enumerate(search_keywords, start=1):
                if update_progress:
                    update_progress(f"Поиск по слову {idx}/{total_keywords}: {kw}")
                results = self.searcher.search_procurements(
                    kw, min_price, max_price, purchase_stage, law, update_progress, source_names
                )
                all_results.extend(results)

            # Дедупликация по названию + цене
            lot_map = {}
            for lot in all_results:
                key = self.searcher.make_lot_cache_key(
                    lot.get("Название"), lot.get("Цена")
                )
                if key not in lot_map:
                    lot_map[key] = lot

            logger.info("Lots total after dedupe: %s", len(lot_map))

            # Сохраняем полный список названий перед отправкой в LLM
            self._dump_all_lot_titles(list(lot_map.values()))

            # Используем кэш, чтобы не отправлять уже проверенные лоты в LLM
            cache = self._lot_llm_cache
            pending = []
            for key, lot in lot_map.items():
                cached = cache.get(key, {})
                score = cached.get("score") if isinstance(cached, dict) else cached
                if not isinstance(score, int) or score < 1 or score > 5:
                    pending.append((key, lot))

            batch_size = int(CONFIG.get("LLM_BATCH_SIZE", 20) or 20)
            if batch_size < 1:
                batch_size = 20

            if pending:
                total_batches = (len(pending) + batch_size - 1) // batch_size
                for batch_idx in range(total_batches):
                    batch = pending[batch_idx * batch_size:(batch_idx + 1) * batch_size]
                    self._append_sent_lot_titles([lot for _key, lot in batch])
                    items = []
                    for item_idx, (_key, lot) in enumerate(batch, start=1):
                        items.append(
                            {
                                "id": str(item_idx),
                                "title": lot.get("Название", ""),
                                "price": lot.get("Цена"),
                            }
                        )
                    if update_progress:
                        update_progress(
                            f"LLM: {batch_idx + 1}/{total_batches} — классификация по названию"
                        )
                    llm_result = self.searcher.call_llm_lot_title_filter(items)
                    for item_idx, (key, lot) in enumerate(batch, start=1):
                        score = llm_result.get(str(item_idx))
                        if not isinstance(score, int):
                            try:
                                score = int(score)
                            except Exception:
                                score = 1
                        if score < 1 or score > 5:
                            score = 1
                        cache[key] = {
                            "title": lot.get("Название", ""),
                            "price": lot.get("Цена"),
                            "url": lot.get("Ссылка"),
                            "source": lot.get("Источник"),
                            "score": score,
                        }
                    self._save_lot_llm_cache()
                logger.info("LLM checked lots: %s", len(pending))
            else:
                logger.info("LLM checked lots: 0 (all from cache)")

            # Возвращаем все лоты с оценкой
            rated_results = []
            for key, lot in lot_map.items():
                cached = cache.get(key, {})
                score = cached.get("score") if isinstance(cached, dict) else cached
                if not isinstance(score, int) or score < 1 or score > 5:
                    score = 1
                lot = dict(lot)
                lot["_score"] = score
                rated_results.append(lot)

            logger.info("Lots rated: %s", len(rated_results))

            self.root.after(0, lambda: update_ui_finish(rated_results))
        except Exception as e:
            logger.error(f"Search failed: {e}")
            err_text = str(e)
            self.root.after(0, lambda err=err_text: update_ui_finish([], err))

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
                        "Оценка": result.get("_score", 1),
                        "Источник": result.get("Источник", "Неизвестно"),
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
        self._score_filter_set = None
        self.export_btn.config(state="disabled")
        self.progress_var.set("Результаты очищены")

        # Reset filters
        self.min_price_entry.delete(0, tk.END)
        self.max_price_entry.delete(0, tk.END)
        self.stage_var.set("")
        self.law_var.set("")
    def _on_llm_provider_change(self, event=None):
        value = (self.llm_provider_var.get() or "").strip()
        if value:
            CONFIG["LLM_PROVIDER"] = value

    def _load_lot_llm_cache(self):
        path = self._lot_llm_cache_path
        if not path or not os.path.exists(path):
            return
        try:
            if os.path.getsize(path) == 0:
                return
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                self._lot_llm_cache = data
        except Exception as e:
            logger.warning("Failed to load lot cache: %s", e)

    def _save_lot_llm_cache(self):
        path = self._lot_llm_cache_path
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(self._lot_llm_cache, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logger.warning("Failed to save lot cache: %s", e)

    def _toggle_select_all_keywords(self):
        if not getattr(self, "search_keyword_vars", None):
            return
        all_selected = all(var.get() for var in self.search_keyword_vars.values())
        new_state = not all_selected
        for var in self.search_keyword_vars.values():
            var.set(new_state)
        if getattr(self, "select_all_btn", None):
            self.select_all_btn.config(text="Снять все" if new_state else "Выбрать все")

    def _reset_llm_cache(self):
        self._lot_llm_cache = {}
        self._save_lot_llm_cache()
        for path in (
            self._lot_titles_dump_path,
            self._lot_titles_sent_path,
        ):
            if not path:
                continue
            try:
                if os.path.exists(path):
                    os.remove(path)
            except Exception as e:
                logger.warning("Failed to remove file %s: %s", path, e)
        self.progress_var.set("Кэш LLM очищен")

    def _dump_all_lot_titles(self, lots: list[dict]) -> None:
        path = self._lot_titles_dump_path
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                for lot in lots:
                    title = (lot.get("Название") or "").strip()
                    price = lot.get("Цена")
                    f.write(f"{title}\t{price}\n")
        except Exception as e:
            logger.warning("Failed to write lot titles dump: %s", e)

    def _append_sent_lot_titles(self, lots: list[dict]) -> None:
        path = self._lot_titles_sent_path
        if not path:
            return
        try:
            with open(path, "a", encoding="utf-8") as f:
                for lot in lots:
                    title = (lot.get("Название") or "").strip()
                    price = lot.get("Цена")
                    f.write(f"{title}\t{price}\n")
        except Exception as e:
            logger.warning("Failed to append sent lot titles: %s", e)



    def _on_closing(self):
        """Handle application closing."""
        if self.search_thread and self.search_thread.is_alive():
            if not messagebox.askokcancel(
                    "Выход", "Поиск ещё выполняется. Закрыть приложение?"
            ):
                return

        # Удаляем все временные файлы документов
        for path in getattr(self, "_temp_doc_files", []):
            try:
                if os.path.exists(path):
                    os.remove(path)
            except Exception as e:
                logger.warning(f"Не удалось удалить временный файл {path}: {e}")

        cache_path = getattr(self, "_lot_llm_cache_path", None)
        if cache_path:
            try:
                if os.path.exists(cache_path):
                    os.remove(cache_path)
            except Exception as e:
                logger.warning(f"Не удалось удалить временный кэш {cache_path}: {e}")

        dump_path = getattr(self, "_lot_titles_dump_path", None)
        if dump_path:
            try:
                if os.path.exists(dump_path):
                    os.remove(dump_path)
            except Exception as e:
                logger.warning(f"Не удалось удалить файл {dump_path}: {e}")

        sent_path = getattr(self, "_lot_titles_sent_path", None)
        if sent_path:
            try:
                if os.path.exists(sent_path):
                    os.remove(sent_path)
            except Exception as e:
                logger.warning(f"Не удалось удалить файл {sent_path}: {e}")

        self.root.destroy()

    # МЕТОДЫ ДЛЯ РАБОТЫ С ДОКУМЕНТАМИ
    def analyze_lot_report(self):
        """Формирует отчет по выбранному лоту (общая инфо + аналитика документов через LLM)."""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Предупреждение", "Выберите лот для анализа")
            return

        item = selection[0]
        values = self.tree.item(item, "values")
        if len(values) < 4:
            messagebox.showerror("Ошибка", "Не удалось прочитать данные выбранного лота")
            return

        lot_title = values[0]
        lot_price_str = values[1]  # ✅ Цена из таблицы
        lot_source = values[2]
        lot_url = values[3]

        if not lot_url or lot_url == "Ссылка не найдена":
            messagebox.showerror("Ошибка", "Неверная ссылка на лот")
            return

        thread = threading.Thread(
            target=self._perform_lot_report,
            args=(lot_title, lot_url, lot_source, lot_price_str),
            daemon=True,
        )
        thread.start()

    def _perform_lot_report(self, lot_title: str, lot_url: str, lot_source: str, lot_price_str: str):
        def update_progress(message):
            self.root.after(0, lambda: self.progress_var.set(message))

        self.root.after(0, lambda: self.progress_bar.start())

        try:
            # ✅ Общая информация: дедлайн + цена из таблицы
            update_progress("Получение срока окончания подачи заявок...")
            lot_deadline = self.searcher.get_application_deadline(lot_url) or "—"

            lot_price = (lot_price_str or "").strip()
            if not lot_price:
                lot_price = "Не указана"

            # ✅ Документы + LLM
            update_progress("Скачивание документов лота...")
            documents = self.searcher.download_documents(lot_url, update_progress)

            update_progress("Объединение текста документов...")
            combined_text = self.searcher.build_lot_documents_text(documents)

            update_progress("LLM-анализ (цели/сроки/требования)...")
            try:
                llm_data = self.searcher.call_llm_lot_analysis(combined_text)
            except Exception as err:
                logger.warning("LLM analysis failed, fallback to dashes: %s", err)
                llm_data = {
                    "subject": "—",
                    "work_scope": "—",
                    "work_and_submission_timelines": "Сроки работ: —; Сроки подачи: —",
                    "fit_summary": "—",
                }

            self.root.after(
                0,
                lambda: self._fill_report_text(
                    lot_title=lot_title,
                    lot_url=lot_url,
                    lot_source=lot_source,
                    lot_price=lot_price,
                    lot_deadline=lot_deadline,
                    llm_data=llm_data,
                ),
            )

        except Exception as e:
            logger.error("Lot report failed: %s", e)
            err_text = str(e)
            self.root.after(
                0,
                lambda err=err_text: messagebox.showerror(
                    "Ошибка",
                    f"Не удалось сформировать отчет:\n{err}"
                )
            )
        finally:
            self.root.after(0, lambda: self.progress_bar.stop())

    def _show_lot_report_window(self, lot_title: str, lot_url: str, lot_source: str,
                                lot_price: str, lot_deadline: str,
                                documents: list[dict], llm_data: dict):
        win = tk.Toplevel(self.root)
        win.title("Отчет по лоту")
        win.geometry("900x650")

        text_frame = ttk.Frame(win)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        text_widget = scrolledtext.ScrolledText(text_frame, wrap=tk.WORD)
        text_widget.pack(fill=tk.BOTH, expand=True)

        # 1) Общая информация
        text_widget.insert(tk.END, "ОТЧЕТ ПО ЛОТУ\n", "title")
        text_widget.insert(tk.END, "=" * 60 + "\n\n")

        text_widget.insert(tk.END, "1) ОБЩАЯ ИНФОРМАЦИЯ\n", "subtitle")
        text_widget.insert(tk.END, f"Наименование лота: {lot_title}\n")
        text_widget.insert(tk.END, f"Ссылка: {lot_url}\n")
        text_widget.insert(tk.END, f"Площадка: {lot_source}\n")
        text_widget.insert(tk.END, f"Стоимость: {lot_price}\n")
        text_widget.insert(tk.END, f"Окончание подачи заявок: {lot_deadline}\n\n")

        # 2) Аналитика документов
        text_widget.insert(tk.END, "2) АНАЛИТИКА ДОКУМЕНТОВ\n", "subtitle")
        text_widget.insert(tk.END, f"Всего документов: {len(documents)}\n\n")

        text_widget.insert(tk.END, f"Предмет закупки: {llm_data.get('subject', '—')}\n\n")
        text_widget.insert(tk.END, f"Состав работ/услуг: {llm_data.get('work_scope', '—')}\n\n")
        text_widget.insert(tk.END,
                           f"Сроки работ и подачи: {llm_data.get('work_and_submission_timelines', 'Сроки работ: —; Сроки подачи: —')}\n\n")
        text_widget.insert(tk.END, f"Итоговое резюме (насколько подходит): {llm_data.get('fit_summary', '—')}\n\n")

        # список документов
        text_widget.insert(tk.END, "СПИСОК ДОКУМЕНТОВ\n", "subtitle")
        text_widget.insert(tk.END, "-" * 60 + "\n")
        if documents:
            for doc in documents:
                text_widget.insert(tk.END, f"• {doc.get('name')} ({doc.get('size')} bytes)\n")
        else:
            text_widget.insert(tk.END, "—\n")

        text_widget.tag_configure("title", font=("Arial", 12, "bold"))
        text_widget.tag_configure("subtitle", font=("Arial", 10, "bold"))
        text_widget.config(state=tk.DISABLED)


    def analyze_documents(self):
        """Анализирует документы выбранного лота."""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Предупреждение", "Выберите лот для анализа")
            return

        item = selection[0]
        values = self.tree.item(item, "values")
        # Ссылка теперь на 4-й позиции
        lot_url = values[3] if len(values) >= 4 else None

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
            analysis_results = self.searcher.search_in_documents(
                documents, keywords, update_progress
            )

            self.root.after(0, lambda: show_results(documents, analysis_results))

        except Exception as e:
            logger.error(f"Document analysis failed: {e}")
            self.root.after(
                0, lambda: messagebox.showerror("Ошибка", f"Ошибка анализа: {str(e)}")
            )
            self.progress_bar.stop()

    def _show_analysis_results(self, documents, analysis_results, keywords):
        """Показывает результаты анализа документов и даёт открыть их из памяти."""
        result_window = tk.Toplevel(self.root)
        result_window.title("Результаты анализа документов")
        result_window.geometry("900x600")

        # Верхняя часть — текстовый отчёт
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
                    text_widget.insert(tk.END, "   Контекст:\n")
                    for i, context in enumerate(result["sample_context"][:2]):
                        text_widget.insert(tk.END, f"     {i + 1}. {context}\n")

                text_widget.insert(tk.END, f"   Ссылка: {result['url']}\n\n")
        else:
            text_widget.insert(tk.END, "Ключевые слова не найдены в документах.\n\n")

        # --- ВСЕ ДОКУМЕНТЫ (всегда показываем, не только в else) ---
        text_widget.insert(tk.END, "ВСЕ НАЙДЕННЫЕ ДОКУМЕНТЫ:\n", "subtitle")
        text_widget.insert(tk.END, "=" * 50 + "\n\n")

        for doc in documents:
            text_widget.insert(tk.END, f"📄 {doc['name']} ({doc['size']} bytes)\n")

        # Настройка стилей текста
        text_widget.tag_configure("title", font=("Arial", 12, "bold"))
        text_widget.tag_configure("subtitle", font=("Arial", 10, "bold"))
        text_widget.tag_configure("document_name", font=("Arial", 9, "bold"))

        text_widget.config(state=tk.DISABLED)

        # --- НИЖНЯЯ ПАНЕЛЬ: выбор документа и открытие из памяти ---

        controls_frame = ttk.Frame(result_window)
        controls_frame.pack(fill=tk.X, padx=10, pady=(0, 10))

        ttk.Label(
            controls_frame,
            text="Открыть документ (временный файл, удалится при закрытии программы):",
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 5))

        doc_names = [doc["name"] for doc in documents]
        self._doc_select_var = tk.StringVar(value=doc_names[0] if doc_names else "")

        doc_combo = ttk.Combobox(
            controls_frame,
            textvariable=self._doc_select_var,
            values=doc_names,
            state="readonly",
            width=70,
        )
        doc_combo.grid(row=1, column=0, sticky="we", padx=(0, 5))

        def open_selected_doc():
            name = self._doc_select_var.get()
            if not name:
                messagebox.showwarning(
                    "Не выбран документ", "Выберите документ для открытия"
                )
                return

            # Находим документ по имени среди уже скачанных (они в памяти)
            doc = next((d for d in documents if d["name"] == name), None)
            if not doc:
                messagebox.showerror("Ошибка", "Не удалось найти выбранный документ")
                return

            self._open_document_from_memory(doc)

        open_btn = ttk.Button(
            controls_frame,
            text="Открыть документ",
            command=open_selected_doc,
        )
        open_btn.grid(row=1, column=1, sticky="e")

        controls_frame.columnconfigure(0, weight=1)

        # Сохраняем результаты анализа для экспорта
        self.analysis_results = analysis_results

    def _open_document_from_memory(self, doc):
        """
        Открывает документ, который хранится в памяти (doc['content']),
        через временный файл. Временный файл удаляется при закрытии приложения.
        """
        try:
            content = doc.get("content")
            if not content:
                messagebox.showerror(
                    "Ошибка", "У выбранного документа нет содержимого"
                )
                return

            name = doc.get("filename") or doc.get("name") or "document"
            # Определяем расширение
            suffix = ""
            if "." in name:
                suffix = name[name.rfind("."):]

            # создаём временный файл (delete=False, чтобы успеть его открыть)
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
            tmp.write(content)
            tmp_path = tmp.name
            tmp.close()

            # запоминаем путь, чтобы удалить при закрытии
            self._temp_doc_files.append(tmp_path)

            # Открываем системной программой
            if sys.platform.startswith("win"):
                os.startfile(tmp_path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", tmp_path])
            else:
                subprocess.Popen(["xdg-open", tmp_path])

        except Exception as e:
            logger.error(f"Failed to open document from memory: {e}")
            messagebox.showerror(
                "Ошибка",
                "Не удалось открыть документ.\n"
                "Возможно, на компьютере не установлена программа для этого типа файлов.",
            )


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
