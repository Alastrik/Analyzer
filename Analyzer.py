import tkinter as tk
import matplotlib
matplotlib.use('TkAgg')
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os


def get_unique_filename(base_name, extension=".xlsx"):
    filename = f"{base_name}_processed{extension}"
    counter = 1

    while os.path.exists(filename):
        filename = f"{base_name}_processed_{counter}{extension}"
        counter += 1

    return filename

class DataAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Анализатор данных")
        self.root.geometry("500x300")
        self.root.resizable(False, False)

        self.filename_var = tk.StringVar()
        self.format_var = tk.StringVar(value="txt")
        self.df = None

        main_frame = ttk.Frame(root, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="Название файла (без расширения):").pack(anchor=tk.W, pady=(0, 5))
        ttk.Entry(main_frame, textvariable=self.filename_var, width=40).pack(anchor=tk.W, pady=(0, 15))

        ttk.Label(main_frame, text="Формат файла:").pack(anchor=tk.W, pady=(0, 5))
        radio_frame = ttk.Frame(main_frame)
        radio_frame.pack(anchor=tk.W)

        ttk.Radiobutton(radio_frame, text=".txt (с разделителем)", variable=self.format_var, value="txt").pack(
            side=tk.LEFT, padx=(0, 15))
        ttk.Radiobutton(radio_frame, text=".csv", variable=self.format_var, value="csv").pack(side=tk.LEFT,
                                                                                              padx=(0, 15))
        ttk.Radiobutton(radio_frame, text=".xlsx", variable=self.format_var, value="xlsx").pack(side=tk.LEFT)

        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=20)

        ttk.Button(button_frame, text="Загрузить и проанализировать", command=self.load_and_analyze).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="Выбрать файл вручную", command=self.manual_file_select).pack(side=tk.LEFT)

    def manual_file_select(self):
        filepath = filedialog.askopenfilename(
            title="Выберите файл данных",
            filetypes=[
                ("Текстовые файлы", "*.txt"),
                ("CSV файлы", "*.csv"),
                ("Excel файлы", "*.xlsx"),
                ("Все файлы", "*.*")
            ]
        )
        if filepath:
            base_name = os.path.splitext(os.path.basename(filepath))[0]
            self.filename_var.set(base_name)

            ext = os.path.splitext(filepath)[1].lower()
            if ext == '.xlsx':
                self.format_var.set("xlsx")
            else:
                self.format_var.set("txt")

    def load_and_analyze(self):
        base_name = self.filename_var.get().strip()
        if not base_name:
            messagebox.showwarning("Ошибка", "Введите название файла!")
            return

        ext = self.format_var.get()
        filepath = base_name + '.' + ext

        if not os.path.exists(filepath):
            messagebox.showerror("Ошибка", f"Файл не найден:\n{filepath}")
            return

        try:
            if ext == "txt":
                with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
                    sample = f.read(1024)

                separators = [',', '\t', ';', ' ', '|']
                sep_counts = {sep: sample.count(sep) for sep in separators}
                best_sep = max(sep_counts, key=sep_counts.get)
                sep = best_sep if sep_counts[best_sep] > 0 else ','

                self.df = pd.read_csv(filepath, sep=sep, encoding='utf-8', engine='python')

            elif ext == "csv":
                self.df = pd.read_csv(
                    filepath,
                    encoding='utf-8-sig',
                    sep=',',
                    quotechar='"',
                    skipinitialspace=True,
                    engine='python'
                )

            else:
                xl = pd.ExcelFile(filepath)
                sheet_name = xl.sheet_names[0]

                df_temp = pd.read_excel(filepath, sheet_name=sheet_name, header=None)
                first_non_empty = df_temp.dropna(how='all').index[0]
                self.df = pd.read_excel(
                    filepath,
                    sheet_name=sheet_name,
                    header=first_non_empty,
                    dtype=str
                )
                for col in self.df.columns:
                    try:
                        self.df[col] = pd.to_numeric(self.df[col], errors='ignore')
                    except Exception:
                        pass

            if self.df.empty or len(self.df.columns) == 0:
                messagebox.showwarning("Предупреждение", "Файл не содержит данных или не удалось определить колонки.")
                return

            self.df = self.df.dropna(how='all', axis=1)

            self.show_report()
            self.plot_histogram()

            output_path = get_unique_filename(base_name, extension=".xlsx")
            self.save_full_report(output_path)

        except Exception as e:
            messagebox.showerror("Ошибка при загрузке", f"Не удалось загрузить файл:\n{str(e)}")

    def save_full_report(self, output_path):
        """Сохраняет полный аналитический отчёт в Excel"""
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            self.df.to_excel(writer, sheet_name='Data', index=False)

            desc = self.df.describe(include='all')
            desc.to_excel(writer, sheet_name='Statistics')

            info_df = pd.DataFrame({
                'Column': self.df.columns,
                'Dtype': self.df.dtypes.astype(str)
            })
            info_df.to_excel(writer, sheet_name='DataTypes', index=False)

            missing = pd.DataFrame({
                'Column': self.df.columns,
                'Missing_Count': self.df.isnull().sum().values,
                'Missing_Percent': (self.df.isnull().sum() / len(self.df) * 100).round(2).values
            })
            missing.to_excel(writer, sheet_name='MissingValues', index=False)

            cat_cols = self.df.select_dtypes(include=['object', 'category']).columns
            if len(cat_cols) > 0:
                cat_overview = []
                for col in cat_cols:
                    mode_val = self.df[col].mode()
                    top_value = mode_val.iloc[0] if not mode_val.empty else "—"
                    cat_overview.append({
                        'Column': col,
                        'Unique_Values': self.df[col].nunique(),
                        'Most_Frequent': top_value,
                        'Top_Freq_Count': (self.df[col] == top_value).sum()
                    })
                cat_df = pd.DataFrame(cat_overview)
                cat_df.to_excel(writer, sheet_name='CategoricalOverview', index=False)

    def show_report(self):
        with pd.option_context(
                'display.max_columns', None,
                'display.max_colwidth', None,
                'display.width', None,
                'display.max_rows', None
        ):
            buf = []
            buf.append("=== ИНФОРМАЦИЯ О ДАННЫХ ===\n")
            buf.append(str(self.df.dtypes))
            buf.append("\n\n=== ОПИСАТЕЛЬНАЯ СТАТИСТИКА ===\n")
            buf.append(str(self.df.describe(include='all')))

            report_window = tk.Toplevel(self.root)
            report_window.title("Отчёт о данных")
            report_window.geometry("900x600")

            text = tk.Text(report_window, wrap=tk.NONE)
            scroll_y = tk.Scrollbar(report_window, orient=tk.VERTICAL, command=text.yview)
            scroll_x = tk.Scrollbar(report_window, orient=tk.HORIZONTAL, command=text.xview)
            text.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

            text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
            scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

            text.insert(tk.END, "\n".join(buf))
            text.config(state=tk.DISABLED)

    def plot_histogram(self):
        numeric_cols = self.df.select_dtypes(include=[np.number]).columns
        if len(numeric_cols) == 0:
            messagebox.showinfo("Визуализация", "Нет числовых колонок для построения гистограмм.")
            return

        n = len(numeric_cols)

        if n <= 4:
            fig, axes = plt.subplots(n, 1, figsize=(6, 3 * n))
            if n == 1:
                axes = [axes]
            for ax, col in zip(axes, numeric_cols):
                self.df[col].dropna().hist(ax=ax, bins=20, color='lightgreen', edgecolor='black')
                ax.set_title(f'Гистограмма: {col}')
                ax.grid(True)
            plt.tight_layout()
            plt.show()
        else:
            for col in numeric_cols:
                plt.figure(figsize=(6, 4))
                plt.hist(self.df[col].dropna(), bins=20, color='lightgreen', edgecolor='black')
                plt.title(f'Гистограмма: {col}')
                plt.xlabel(col)
                plt.ylabel('Частота')
                plt.grid(True)
                plt.tight_layout()
                plt.show()

if __name__ == "__main__":
    root = tk.Tk()
    app = DataAnalyzerApp(root)
    root.mainloop()