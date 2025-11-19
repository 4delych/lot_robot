import os
import sys
import subprocess
import tempfile
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import threading
import pandas as pd
import logging
import webbrowser


from config import PURCHASE_STAGES, LAWS
from ProcurmentSearcher import ProcurementSearcher
logger = logging.getLogger(__name__)

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
        self._temp_doc_files = []
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

            name = doc.get("name") or "document"
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
