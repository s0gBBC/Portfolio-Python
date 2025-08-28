#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import os
import pandas as pd
from tkinter import Tk, Listbox, Scrollbar, Button, Label, Frame, filedialog, messagebox
from tkinter.filedialog import askopenfilenames, askdirectory

class ExcelMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Объединение Excel файлов")

        # Инициализация атрибутов
        self.file_paths = []
        self.sheet_name = None
        self.key_columns = []
        self.merge_columns = []

        self.create_widgets()

    def create_widgets(self):
        """Создание всех виджетов на одном экране"""
        self.frame = Frame(self.root)
        self.frame.pack(padx=10, pady=10)

        # Кнопка для выбора файлов
        self.files_button = Button(self.frame, text="Выбрать файлы", command=self.select_files)
        self.files_button.grid(row=0, column=0, padx=5, pady=5)

        self.status_label = Label(self.frame, text="Выберите файлы для объединения.")
        self.status_label.grid(row=1, column=0, padx=5, pady=5)

    def select_files(self):
        """Функция для выбора файлов"""
        self.file_paths = askopenfilenames(
            title="Выберите файлы", 
            filetypes=[("Excel files", "*.xlsx;*.xls")]
        )
        
        if not self.file_paths:
            self.status_label.config(text="Файлы не выбраны.")
            return
            
        self.status_label.config(text=f"Выбрано {len(self.file_paths)} файлов.")
        self.files_button.config(state="disabled")
        
        # Проверяем, есть ли общий лист во всех файлах
        common_sheets = self.find_common_sheets()
        
        if not common_sheets:
            messagebox.showerror("Ошибка", "Нет общих листов во всех выбранных файлах.")
            return
            
        self.select_sheet(common_sheets)

    def find_common_sheets(self):
        """Находит общие листы во всех выбранных файлах"""
        sheets_list = []
        for file_path in self.file_paths:
            xls = pd.ExcelFile(file_path)
            sheets_list.append(set(xls.sheet_names))
            
        common_sheets = set.intersection(*sheets_list)
        return list(common_sheets)

    def select_sheet(self, sheet_names):
        """Выбор листа из файла"""
        self.clear_frame()

        Label(self.frame, text="Выберите лист для объединения:").grid(row=0, column=0, padx=5, pady=5)

        # Список листов
        self.sheet_listbox = Listbox(self.frame, height=5, width=50)
        for sheet in sheet_names:
            self.sheet_listbox.insert("end", sheet)
        self.sheet_listbox.grid(row=1, column=0, padx=5, pady=5)

        # Скроллбар
        scrollbar = Scrollbar(self.frame, orient="vertical", command=self.sheet_listbox.yview)
        scrollbar.grid(row=1, column=1, padx=5, pady=5, sticky="ns")
        self.sheet_listbox.config(yscrollcommand=scrollbar.set)

        # Кнопка для продолжения
        Button(self.frame, text="Продолжить", command=self.prepare_merge).grid(row=2, column=0, padx=5, pady=5)
        Button(self.frame, text="Назад", command=self.create_widgets).grid(row=2, column=1, padx=5, pady=5)

    def prepare_merge(self):
        """Подготовка к объединению - выбор ключевых столбцов"""
        selected_index = self.sheet_listbox.curselection()
        if not selected_index:
            self.status_label.config(text="Лист не выбран.")
            return

        self.sheet_name = self.sheet_listbox.get(selected_index)
        
        # Загружаем первый файл для получения столбцов
        df = pd.read_excel(self.file_paths[0], sheet_name=self.sheet_name)
        columns = df.columns.tolist()

        self.clear_frame()

        Label(self.frame, text="Выберите ключевые столбцы для объединения:").grid(row=0, column=0, padx=5, pady=5)
        Label(self.frame, text="(По этим столбцам будет происходить сопоставление строк)").grid(row=1, column=0, padx=5, pady=5)

        # Список столбцов (множественный выбор)
        self.columns_listbox = Listbox(self.frame, height=10, width=50, selectmode="multiple")
        for column in columns:
            self.columns_listbox.insert("end", column)
        self.columns_listbox.grid(row=2, column=0, padx=5, pady=5)

        # Скроллбар
        scrollbar = Scrollbar(self.frame, orient="vertical", command=self.columns_listbox.yview)
        scrollbar.grid(row=2, column=1, padx=5, pady=5, sticky="ns")
        self.columns_listbox.config(yscrollcommand=scrollbar.set)

        # Кнопки
        Button(self.frame, text="Продолжить", command=self.process_merge).grid(row=3, column=0, padx=5, pady=5)
        Button(self.frame, text="Назад", command=lambda: self.select_sheet(self.find_common_sheets())).grid(row=3, column=1, padx=5, pady=5)

    def process_merge(self):
        """Процесс объединения файлов"""
        selected_indices = self.columns_listbox.curselection()
        if not selected_indices:
            messagebox.showerror("Ошибка", "Не выбраны ключевые столбцы.")
            return
            
        self.key_columns = [self.columns_listbox.get(i) for i in selected_indices]
        
        # Загружаем все файлы
        dfs = []
        for file_path in self.file_paths:
            df = pd.read_excel(file_path, sheet_name=self.sheet_name)
            dfs.append(df)
            
        # Объединяем данные
        merged_df = self.merge_dataframes(dfs)
        
        # Сохраняем результат
        output_dir = askdirectory(title="Выберите папку для сохранения результата")
        if not output_dir:
            self.status_label.config(text="Папка не выбрана.")
            return
            
        output_path = os.path.join(output_dir, f"merged_{self.sheet_name}.xlsx")
        merged_df.to_excel(output_path, index=False)
        
        messagebox.showinfo("Успех", f"Файлы успешно объединены и сохранены как:\n{output_path}")
        self.root.after(2000, self.root.quit)

    def merge_dataframes(self, dfs):
        """
        Объединяет DataFrame, заполняя недостающие данные из других DataFrame
        """
        # Начинаем с первого DataFrame
        merged_df = dfs[0].copy()
        
        # Для каждого последующего DataFrame
        for df in dfs[1:]:
            # Объединяем по ключевым столбцам
            for _, row in df.iterrows():
                # Ищем совпадение в merged_df по ключевым столбцам
                mask = pd.Series(True, index=merged_df.index)
                for col in self.key_columns:
                    mask &= (merged_df[col] == row[col])
                
                if mask.any():  # Если нашли совпадение
                    idx = mask.idxmax()
                    # Обновляем только пустые или NaN значения
                    for col in df.columns:
                        if col not in self.key_columns:
                            if pd.isna(merged_df.at[idx, col]) or merged_df.at[idx, col] == '':
                                merged_df.at[idx, col] = row[col]
                else:  # Если не нашли - добавляем новую строку
                    merged_df = pd.concat([merged_df, row.to_frame().T], ignore_index=True)
        
        return merged_df

    def clear_frame(self):
        """Очистка фрейма перед добавлением новых элементов"""
        for widget in self.frame.winfo_children():
            widget.destroy()


# Запуск программы
if __name__ == "__main__":
    root = Tk()
    app = ExcelMergerApp(root)
    root.mainloop()

