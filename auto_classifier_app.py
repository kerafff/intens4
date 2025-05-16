import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import pymorphy2
import threading
import re

# Лемматизация
morph = pymorphy2.MorphAnalyzer()

def lemmatize_text(text):
    """
    Лемматизирует текст, преобразуя слова в их нормальную форму
    """
    if pd.isna(text) or not isinstance(text, str):
        return []
    words = str(text).lower().split()
    lemmas = [morph.parse(word)[0].normal_form for word in words]
    return lemmas

# Ключевые слова
keyword_map = {
    "Нравится скорость отработки заявок": ["быстро", "оперативно", "скорость", "сразу", "в тот же день", "моментально"],
    "Нравится качество выполнения заявки": ["качественно", "отлично", "хорошо", "хорошо сделал", "хорошо сделали", "проблема решена", "все работает", "доволен работой", "качество"],
    "Нравится качество работы сотрудников": [
        "вежливо", "спасибо", "благодарю", "профессионально", 
        "вежливый персонал", "компетентные сотрудники", "помогли разобраться", "вежливые", "грамотные", "сотрудники молодцы", "хорошие сотрудники",
    
        "мастер", "мастера", "мастеру", "мастером", "мастеров", "мастерам", "мастерами", "мастерах",
        "специалист", "специалиста", "специалисту", "специалистом", "специалисте", "специалисты", "специалистов", "специалистам", "специалистами", "специалистах",
        "работник", "работника", "работнику", "работником", "работнике", "работники", "работников", "работникам", "работниками", "работниках",
        "сотрудник", "сотрудника", "сотруднику", "сотрудником", "сотруднике", "сотрудники", "сотрудников", "сотрудникам", "сотрудниками", "сотрудниках"
    ],
    "Понравилось выполнение заявки": ["выполнить", "сделать", "хорошо", "спасибо", "благодарю", "понравилось", "доволен", "отлично", "супер", "замечательно"],
    "Вопрос решен": ["решить", "устранить", "помочь", "закрыт", "выполнена", "вопрос закрыт", "проблема устранена", "все решено", "решили вопрос", "помогли решить"]
}

class CommentClassifierApp:
    def __init__(self, master):
        self.master = master
        master.title("Классификация комментариев v3")
        master.geometry("800x600")

        self.df = None
        self.input_csv_path = None
        self.selected_comment_column = tk.StringVar()
        self.no_category_column_name = "Без категории"

        # Настройка стилей
        style = ttk.Style()
        style.configure("TButton", padding=6, relief="flat", background="#ccc")
        style.configure("TLabel", padding=6)
        style.configure("Status.TLabel", font=("Helvetica", 10, "bold"))

        # --- Верхняя панель для загрузки --- 
        load_frame = ttk.Frame(master, padding="10 10 10 0")
        load_frame.pack(fill=tk.X)
        self.load_button = ttk.Button(load_frame, text="Загрузить CSV", command=self.load_csv)
        self.load_button.pack(side=tk.LEFT, padx=5)
        self.file_label = ttk.Label(load_frame, text="Файл не загружен", width=50)
        self.file_label.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        # --- Панель выбора колонки --- 
        column_select_frame = ttk.Frame(master, padding="10 5 10 5")
        column_select_frame.pack(fill=tk.X)
        ttk.Label(column_select_frame, text="Колонка с комментариями:").pack(side=tk.LEFT, padx=5)
        self.column_combobox = ttk.Combobox(column_select_frame, textvariable=self.selected_comment_column, state="readonly", width=45)
        self.column_combobox.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        self.column_combobox.bind("<<ComboboxSelected>>", self.on_column_selected)

        # --- Панель кнопок обработки и сохранения --- 
        process_save_frame = ttk.Frame(master, padding="10 5 10 10")
        process_save_frame.pack(fill=tk.X)
        self.process_button = ttk.Button(process_save_frame, text="Классифицировать", command=self.start_processing, state=tk.DISABLED)
        self.process_button.pack(side=tk.LEFT, padx=5)
        self.save_button = ttk.Button(process_save_frame, text="Сохранить CSV", command=self.save_csv, state=tk.DISABLED)
        self.save_button.pack(side=tk.LEFT, padx=5)

        # --- Статус и прогресс-бар --- 
        self.status_label = ttk.Label(master, text="", style="Status.TLabel", anchor=tk.CENTER)
        self.status_label.pack(pady=10, fill=tk.X)
        self.progress_bar = ttk.Progressbar(master, orient="horizontal", length=750, mode="determinate")
        self.progress_bar.pack(pady=10, padx=20, fill=tk.X)

        # --- Область для отображения данных ---
        data_frame = ttk.LabelFrame(master, text="Предпросмотр данных", padding="10")
        data_frame.pack(pady=10, padx=10, fill="both", expand=True)
        
        # Treeview для отображения данных
        self.tree = ttk.Treeview(data_frame)
        self.tree.pack(fill="both", expand=True)
        
        # скроллбары
        vsb = ttk.Scrollbar(self.tree, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(data_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        
        # --- Область для отображения категорий и ключевых слов ---
        info_frame = ttk.LabelFrame(master, text="Категории и ключевые слова (для справки)", padding="10")
        info_frame.pack(pady=10, padx=10, fill="both", expand=True)
        
        self.info_text = tk.Text(info_frame, wrap=tk.WORD, height=6, width=70, state=tk.DISABLED, relief=tk.FLAT, background=master.cget("background"))
        self.info_text.pack(fill="both", expand=True)
        self.populate_info_text()

    def populate_info_text(self):
       
        self.info_text.config(state=tk.NORMAL)
        self.info_text.delete("1.0", tk.END)
        for category, keywords in keyword_map.items():
            self.info_text.insert(tk.END, f"{category}:\n", "bold")
            # Ограничиваем количество отображаемых ключевых слов для удобства
            display_keywords = keywords[:10] if len(keywords) > 10 else keywords
            if len(keywords) > 10:
                display_text = f"  {', '.join(display_keywords)}... и еще {len(keywords) - 10} слов/форм"
            else:
                display_text = f"  {', '.join(display_keywords)}"
            self.info_text.insert(tk.END, f"{display_text}\n\n")
        self.info_text.tag_configure("bold", font=("Helvetica", 9, "bold"))
        self.info_text.config(state=tk.DISABLED)

    def load_csv(self):
       
        file_path = filedialog.askopenfilename(
            title="Выберите CSV файл",
            filetypes=(("CSV файлы", "*.csv"), ("Все файлы", "*.*"))
        )
        
        if not file_path:
            return
            
        self.input_csv_path = file_path
        self.df = None
        self.selected_comment_column.set("")
        self.column_combobox.config(values=[])
        self.column_combobox.config(state="disabled")
        self.process_button.config(state=tk.DISABLED)
        self.save_button.config(state=tk.DISABLED)
        self.progress_bar["value"] = 0
        
        try:
            self.df = pd.read_csv(file_path)
            
            # Удаляем старые колонки категорий, если они существуют
            categories_to_remove = list(keyword_map.keys()) + [self.no_category_column_name]
            for category in categories_to_remove:
                if category in self.df.columns:
                    self.df = self.df.drop(columns=[category])
            
            self.file_label.config(text=f"Загружен: {file_path.split('/')[-1]}")
            self.status_label.config(text="Файл успешно загружен. Выберите колонку с комментариями.", foreground="blue")
            
            available_columns = list(self.df.columns)
            self.column_combobox.config(values=available_columns, state="readonly")
            
            if available_columns:
                # Попытка найти колонку с комментариями
                common_text_cols = [col for col in available_columns if any(sub in col.lower() for sub in ['text', 'comment', 'review', 'отзыв', 'комментарий'])]
                if common_text_cols:
                    self.selected_comment_column.set(common_text_cols[0])
                else:
                    self.selected_comment_column.set(available_columns[0])
                self.on_column_selected()
                
                # Отображаем данные в таблице
                self.show_table(self.df.head(10))
            else:
                messagebox.showerror("Ошибка", "В CSV файле не найдено колонок.")
                self.file_label.config(text="Файл не загружен")
                return
                
        except Exception as e:
            messagebox.showerror("Ошибка загрузки", f"Не удалось загрузить файл: {e}")
            self.input_csv_path = None
            self.file_label.config(text="Файл не загружен")
            self.status_label.config(text="Ошибка загрузки файла.", foreground="red")

    def on_column_selected(self, event=None):
        
        if self.selected_comment_column.get() and self.df is not None:
            self.process_button.config(state=tk.NORMAL)
            self.status_label.config(text=f"Колонка '{self.selected_comment_column.get()}' выбрана. Готово к классификации.", foreground="green")
        else:
            self.process_button.config(state=tk.DISABLED)

    def classify_comment(self, comment_text):
       
        lemmas = lemmatize_text(comment_text)
        result = {}
        for category, keywords in keyword_map.items():
            result[category] = int(any(word in lemmas for word in keywords))
        # Добавляем "Без категории"
        result[self.no_category_column_name] = int(all(val == 0 for val in result.values()))
        return result

    def process_data(self):
        
        if self.df is None or not self.selected_comment_column.get():
            messagebox.showerror("Ошибка", "Нет данных для обработки или не выбрана колонка с комментариями.")
            self.status_label.config(text="Ошибка: данные или колонка не выбраны.", foreground="red")
            return

        current_comment_column = self.selected_comment_column.get()
        if current_comment_column not in self.df.columns:
            messagebox.showerror("Ошибка", f"Выбранная колонка '{current_comment_column}' больше не существует в DataFrame.")
            self.status_label.config(text="Ошибка: выбранная колонка не найдена.", foreground="red")
            return

        self.status_label.config(text="Идет классификация... Пожалуйста, подождите.", foreground="blue")
        self.load_button.config(state=tk.DISABLED)
        self.column_combobox.config(state=tk.DISABLED)
        self.process_button.config(state=tk.DISABLED)
        self.save_button.config(state=tk.DISABLED)
        self.progress_bar["value"] = 0
        self.progress_bar["maximum"] = len(self.df)

        # Удаляем существующие колонки категорий перед добавлением новых
        categories_to_remove = list(keyword_map.keys()) + [self.no_category_column_name]
        for category in categories_to_remove:
            if category in self.df.columns:
                self.df = self.df.drop(columns=[category])

        try:
            results = self.df[current_comment_column].apply(self.classify_comment)
            results_df = pd.DataFrame(list(results))
            
            # Обновляем прогресс-бар для визуальной обратной связи
            for i in range(len(self.df)):
                self.progress_bar["value"] = i + 1
                if (i + 1) % 50 == 0:  # Обновляем GUI не слишком часто
                    self.master.update_idletasks()
            
            # Добавляем результаты классификации в DataFrame
            self.df = pd.concat([self.df, results_df], axis=1)
            
            # Отображаем обновленные данные
            self.show_table(self.df.head(10))
            
            self.status_label.config(text="Классификация завершена! Файл готов к сохранению.", foreground="green")
            self.save_button.config(state=tk.NORMAL)
            messagebox.showinfo("Готово", "Комментарии успешно размечены.")
            
        except Exception as e:
            messagebox.showerror("Ошибка обработки", f"Произошла ошибка во время классификации: {e}")
            self.status_label.config(text="Ошибка во время обработки.", foreground="red")
        finally:
            self.load_button.config(state=tk.NORMAL)
            self.column_combobox.config(state="readonly")
            self.process_button.config(state=tk.NORMAL if self.selected_comment_column.get() else tk.DISABLED)
            self.master.update_idletasks()

    def start_processing(self):
       
        processing_thread = threading.Thread(target=self.process_data)
        processing_thread.daemon = True
        processing_thread.start()

    def save_csv(self):
       
        if self.df is None:
            messagebox.showerror("Ошибка", "Нет данных для сохранения.")
            return
            
        # Проверка, что хотя бы одна из категориальных колонок была добавлена
        if not any(cat_col in self.df.columns for cat_col in list(keyword_map.keys()) + [self.no_category_column_name]):
            messagebox.showerror("Ошибка", "Обработка не была выполнена или не добавила новые колонки.")
            return

        output_csv_path = filedialog.asksaveasfilename(
            title="Сохранить CSV файл как",
            defaultextension=".csv",
            filetypes=(("CSV файлы", "*.csv"), ("Все файлы", "*.*"))
        )

        if output_csv_path:
            try:
                self.df.to_csv(output_csv_path, index=False, encoding='utf-8-sig')
                self.status_label.config(text=f"Файл успешно сохранен: {output_csv_path.split('/')[-1]}", foreground="green")
                messagebox.showinfo("Сохранение успешно", f"Файл сохранен как {output_csv_path}")
            except Exception as e:
                messagebox.showerror("Ошибка сохранения", f"Не удалось сохранить файл: {e}")
                self.status_label.config(text="Ошибка сохранения файла.", foreground="red")
        else:
            self.status_label.config(text="Сохранение файла отменено.", foreground="orange")

    def show_table(self, df):
      
        # Очищаем существующие данные
        for col in self.tree.get_children():
            self.tree.delete(col)
            
        # Настраиваем колонки
        self.tree["columns"] = list(df.columns)
        self.tree["show"] = "headings"
        
        # Устанавливаем заголовки колонок
        for col in df.columns:
            self.tree.heading(col, text=col)
            # Устанавливаем ширину колонки в зависимости от типа данных
            if df[col].dtype == 'int64' or df[col].dtype == 'float64':
                self.tree.column(col, width=70, anchor="center")
            else:
                self.tree.column(col, width=150)
                
        # Добавляем данные
        for _, row in df.iterrows():
            self.tree.insert("", tk.END, values=list(row))

if __name__ == "__main__":
    root = tk.Tk()
    app = CommentClassifierApp(root)
    root.mainloop()
