import pandas as pd
import numpy as np
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from pathlib import Path
from datetime import datetime

def get_filenames_os_listdir(directory_path):
    """
    Retrieves all filenames in a given directory using os.listdir().
    Does not include subdirectories or files within subdirectories.
    """
    filenames = []
    for entry in os.listdir(directory_path):
        full_path = os.path.join(directory_path, entry)
        if os.path.isfile(full_path):
            filenames.append(entry)
    return filenames

def list_directories(path='.'):
    """
    Lists all directories within a given path.
    If no path is provided, it defaults to the current working directory.
    """
    directories = []
    for item in os.listdir(path):
        item_path = os.path.join(path, item)
        if os.path.isdir(item_path):
            directories.append(item)
    return directories


def link_count_cycle_numbers(df):
    dp_dict = pd.Series(data=list(df['Step Type']), index=list(df['DataPoint'])).to_dict()
    curr = None
    counter_dict = {'CC Chg': 0, 'CC DChg': 0}
    result_dict = dict()
    CCChg_counter = 0
    CCDChg_counter = 0

    for element in dp_dict.keys():
        if dp_dict[element] != curr:
            curr = dp_dict[element]
            if curr in ['CC Chg', 'CC DChg']:
                counter_dict[curr] += 1
                result_dict[element] = counter_dict[curr]
        else:
            if curr in ['CC Chg', 'CC DChg']:
                result_dict[element] = counter_dict[curr]

    result_df = pd.Series(result_dict, name="CycleNumber").reset_index()

    df = df.merge(result_df, how='left', left_on="DataPoint", right_on='index')

    return df

def get_average_current_per_cycle_from_file(filename):
    try:
        record_df = pd.read_excel(filename, sheet_name='record')
        record_df = link_count_cycle_numbers(record_df)
        if 'Current(mA)' in record_df.columns:
            avg_current_per_cycle = record_df.groupby(['CycleNumber','Step Type'])['Current(mA)'].mean().reset_index()
        elif 'Current(μA)' in record_df.columns:
            avg_current_per_cycle = record_df.groupby(['CycleNumber','Step Type'])['Current(μA)'].mean().reset_index()
            avg_current_per_cycle['Current(mA)'] = np.array(avg_current_per_cycle['Current(μA)'])/1000
        else:
            return "Error"
        avg_current_per_cycle['CycleNumber'] = avg_current_per_cycle['CycleNumber'].astype(int)
        avg_current_per_cycle = avg_current_per_cycle[avg_current_per_cycle["Step Type"]=='CC Chg'][["CycleNumber","Current(mA)"]].copy()
        return avg_current_per_cycle
    except:
        return "Error"

def get_en_eff_data_from_cycle_sheet(filename):
    cycle_df = pd.read_excel(filename, sheet_name='cycle')
    cycle_output_dict = {"min DChg. Cap.(mAh)": None, "max DChg. Cap.(mAh)": None, "avg Chg.-DChg. Eff(%)": None}
    try:
        cycle_output_dict["min DChg. Cap.(mAh)"] = min(cycle_df["DChg. Cap.(mAh)"])
        cycle_output_dict["max DChg. Cap.(mAh)"] = max(cycle_df["DChg. Cap.(mAh)"])
        cycle_output_dict["avg Chg.-DChg. Eff(%)"] = (cycle_df["Chg.-DChg. Eff(%)"]).mean()
    except:
        cycle_output_dict["min DChg. Cap.(mAh)"] = "Error"
        cycle_output_dict["max DChg. Cap.(mAh)"] = "Error"
        cycle_output_dict["avg Chg.-DChg. Eff(%)"] = "Error"
    return cycle_output_dict

def get_capacity_data_remarks_time_from_test_sheet(filename):
    test_df = pd.read_excel(filename, sheet_name='test', skiprows=10)
    try:
        theoretical_capacity = test_df[test_df["Step Name"]=="CC Chg"]["Capacity(mAh)"].iloc[0]
    except:
        theoretical_capacity = None
    test_df_2 = pd.read_excel(filename, sheet_name='test')
    remarks = test_df_2.iloc[2, 8]
    time = test_df_2.iloc[4,5]
    theoretical_capacity, remarks, time = theoretical_capacity if theoretical_capacity else "N/A", remarks if remarks else "N/A", time if time else "N/A"
    return theoretical_capacity, remarks, time

def get_device_info_from_unit_sheet(filename):
    # device присутствует в каждом файле из протестированных ~30, можно не обрабатывать исключения
    df = pd.read_excel(filename, sheet_name='unit')
    device = "{} {} {}".format(str(df.iloc[0, 1]),str(df.iloc[0, 2]),str(df.iloc[0, 3]))
    return device if device else "N/A"

def get_resistance_per_cycle_from_step_sheet(filename):
    try:
        step_df = pd.read_excel(filename, sheet_name='step')
        step_df = step_df[step_df["Step Type"].isin(["CC Chg","CC DChg"])].copy()
        step_df['Calculated Voltage, V'] = np.array(step_df["Energy(Wh)"])/np.array(step_df["Capacity(mAh)"])*1000
        step_df['prev Calculated Voltage, V'] = step_df.groupby('Cycle Index')['Calculated Voltage, V'].shift(1)
        avg_current_per_cycle = get_average_current_per_cycle_from_file(filename)
        step_df = step_df.merge(avg_current_per_cycle, how='inner', left_on=["Cycle Index"], right_on=["CycleNumber"])
        step_df["Resistance, Ohm"] = (np.array(step_df['prev Calculated Voltage, V']) - np.array(step_df['Calculated Voltage, V']))/(2*np.array(abs(step_df["Current(mA)"])))*1000
        resistance_per_cycle = step_df[step_df["Step Type"]=="CC DChg"][["Cycle Index","Resistance, Ohm"]].copy()
        return resistance_per_cycle
    except ZeroDivisionError:
        return "ZeroDivisionError"
    except:
        return "Error"


def make_output_statistics_df(filename):
    # Сбор данных из файла excel
    resistance_per_cycle = get_resistance_per_cycle_from_step_sheet(filename)
    en_eff_data_from_cycle_sheet = get_en_eff_data_from_cycle_sheet(filename)
    theoretical_capacity, remarks, time = get_capacity_data_remarks_time_from_test_sheet(filename)
    device = get_device_info_from_unit_sheet(filename)

    # Создание словаря
    output_statistics_dict = dict()
    output_statistics_dict["Файл"] = os.path.basename(filename)
    output_statistics_dict["Описание"] = remarks
    output_statistics_dict["Номер на стенде"] = device
    output_statistics_dict["Дата и время"] = time

    output_statistics_dict["Среднее сопротивление поляризации, Ом"] = resistance_per_cycle if type(
        resistance_per_cycle) == str else resistance_per_cycle["Resistance, Ohm"].mean()
    output_statistics_dict["Выборочное стандартное отклонение, Ом"] = resistance_per_cycle if type(
        resistance_per_cycle) == str else resistance_per_cycle["Resistance, Ohm"].std()
    output_statistics_dict["Минимальная емкость разряда, мАч"] = en_eff_data_from_cycle_sheet["min DChg. Cap.(mAh)"]
    output_statistics_dict["Максимальная емкость разряда, мАч"] = en_eff_data_from_cycle_sheet["max DChg. Cap.(mAh)"]
    output_statistics_dict["Теоретическая емкость, мАч"] = theoretical_capacity
    output_statistics_dict["Среднее значение эффективности заряда-разряда, %"] = en_eff_data_from_cycle_sheet[
        "avg Chg.-DChg. Eff(%)"]

    # Создание dataframe
    output_statistics_df = pd.DataFrame([output_statistics_dict])

    return output_statistics_df

def create_folder_name():
        return str(datetime.now().strftime("%Y-%m-%d-%H-%M-%S"))


class FileSelectorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("File Selector")
        self.root.geometry("700x500")
        self.root.resizable(True, True)

        # Переменная для хранения путей к файлам
        self.input_full_paths_list = []

        self.create_widgets()

    def create_widgets(self):
        # Основной фрейм
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Конфигурация весов для растягивания
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

        # Заголовок
        title_label = ttk.Label(main_frame, text="Выбор файлов для обработки",
                                font=("Arial", 14, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))

        # Фрейм для кнопок выбора и сброса
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=1, column=0, columnspan=3, pady=(0, 10), sticky=(tk.W, tk.E))

        # Кнопка "Выбрать файлы"
        self.select_button = ttk.Button(button_frame, text="Выбрать файлы",
                                        command=self.select_files)
        self.select_button.pack(side=tk.LEFT, padx=(0, 10))

        # Кнопка "Сбросить"
        self.reset_button = ttk.Button(button_frame, text="Сбросить",
                                       command=self.reset_files,
                                       state=tk.DISABLED)
        self.reset_button.pack(side=tk.LEFT)

        # Метка с количеством выбранных файлов
        self.file_count_label = ttk.Label(main_frame, text="Выбрано файлов: 0",
                                          font=("Arial", 10))
        self.file_count_label.grid(row=2, column=0, columnspan=3, pady=(0, 20), sticky=tk.W)

        # Фрейм для списка файлов с прокруткой
        list_frame = ttk.Frame(main_frame)
        list_frame.grid(row=3, column=0, columnspan=3, pady=(0, 20), sticky=(tk.W, tk.E, tk.N, tk.S))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        # Создаем Treeview для отображения выбранных файлов
        columns = ("filename", "path")
        self.file_tree = ttk.Treeview(list_frame, columns=columns, show="tree headings", height=10)
        self.file_tree.heading("filename", text="Имя файла")
        self.file_tree.heading("path", text="Путь")
        self.file_tree.column("filename", width=200)
        self.file_tree.column("path", width=400)

        # Scrollbar для Treeview
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.file_tree.yview)
        self.file_tree.configure(yscrollcommand=scrollbar.set)

        self.file_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))

        # Кнопка "Запустить"
        self.run_button = ttk.Button(main_frame, text="Запустить",
                                     command=self.run_processing,
                                     state=tk.DISABLED)
        self.run_button.grid(row=4, column=0, columnspan=3, pady=(20, 0))

    def select_files(self):
        """Открывает диалог выбора файлов и обновляет список"""
        filetypes = [
            ("Все файлы", "*.*"),
            ("Текстовые файлы", "*.txt"),
            ("Изображения", "*.png *.jpg *.jpeg *.bmp *.gif"),
            ("Документы", "*.doc *.docx *.pdf"),
        ]

        files = filedialog.askopenfilenames(
            title="Выберите файлы для обработки",
            filetypes=filetypes
        )

        if files:
            self.input_full_paths_list = list(files)
            self.update_file_list()
            self.update_buttons_state()

    def reset_files(self):
        """Сбрасывает выбранные файлы"""
        self.input_full_paths_list = []
        self.update_file_list()
        self.update_buttons_state()

    def update_file_list(self):
        """Обновляет отображение списка файлов"""
        # Очищаем текущий список
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)

        # Добавляем новые файлы
        for file_path in self.input_full_paths_list:
            path = Path(file_path)
            self.file_tree.insert("", "end", values=(path.name, str(path.parent)))

        # Обновляем счетчик файлов
        self.file_count_label.config(text=f"Выбрано файлов: {len(self.input_full_paths_list)}")

    def update_buttons_state(self):
        """Обновляет состояние кнопок в зависимости от выбранных файлов"""
        if self.input_full_paths_list:
            self.reset_button.config(state=tk.NORMAL)
            self.run_button.config(state=tk.NORMAL)
        else:
            self.reset_button.config(state=tk.DISABLED)
            self.run_button.config(state=tk.DISABLED)

    def run_processing(self):
        """Запускает обработку файлов (заглушка для вашего кода)"""
        if not self.input_full_paths_list:
            messagebox.showwarning("Предупреждение", "Не выбрано ни одного файла!")
            return

        # Здесь будет ваш основной код обработки
        print("Запуск обработки файлов:")
        df_array = []
        for i, file_path in enumerate(self.input_full_paths_list, 1):
            try:
                df_array.append(make_output_statistics_df(file_path))
            except:
                pass
        result_df = pd.concat(df_array)
        result_df.to_excel(
            "PR_plots/" + str(datetime.now().strftime("%Y-%m-%d-%H-%M-%S")) + '_Cycling_statistics_total.xlsx',
            sheet_name='Статистика', index=False)

        # Пример сообщения об успешном завершении
        messagebox.showinfo("Успех", f"Обработка {len(self.input_full_paths_list)} файлов завершена!")


def main():
    root = tk.Tk()
    app = FileSelectorApp(root)
    root.mainloop()


if __name__ == "__main__":
    folder_name = "PR_plots"
    os.makedirs(folder_name, exist_ok=True)
    print(f"Folder '{folder_name}' ensured to exist.")
    main()