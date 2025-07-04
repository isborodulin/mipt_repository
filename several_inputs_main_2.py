import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
from colorsys import hls_to_rgb
from matplotlib.ticker import MaxNLocator
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from datetime import datetime


def get_cycle_number_list_from_string(string):
    return_array = []

    if string == '':
        return return_array
    curr = ''
    last_curr = ''
    for i in range(len(string)):
        if string[i] == ' ' and curr:
            try:
                return_array.append(int(curr))
                curr = ''
            except ValueError:
                this_array = [i for i in range(int(curr[:curr.find('-')]), int(curr[curr.find('-') + 1:]) + 1)]
                return_array = return_array + this_array
                curr = ''
        elif string[i] == ' ' and not curr:
            pass
        else:
            try:
                if string[i] == '-':
                    pass
                else:
                    error_tryer = int(string[i])
                curr += string[i]
                last_curr = string[i]
            except ValueError:
                print("Недопустимые символы во входной строке, попробуйте снова")
                return 0
    if curr:
        try:
            return_array.append(int(curr))
        except ValueError:
            this_array = [i for i in range(int(curr[:curr.find('-')]), int(curr[curr.find('-') + 1:]) + 1)]
            return_array = return_array + this_array
    return return_array


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


def generate_color_gradient(numbers):
    """
    Generate a dictionary mapping integers to pastel colors in a green-to-red gradient.

    Args:
        numbers: List of positive integers > 1

    Returns:
        Dictionary with integers as keys and hex color strings as values
    """
    if not numbers:
        return {}

    numbers = np.array(numbers)
    min_val, max_val = np.min(numbers), np.max(numbers)

    # Normalize the numbers to [0, 1] range
    if min_val == max_val:
        normalized = np.zeros_like(numbers, dtype=float)
    else:
        normalized = (numbers - min_val) / (max_val - min_val)

    color_dict = {}
    for num, norm in zip(numbers, normalized):
        # Calculate hue (0=red, 1/3=green in HSL space)
        hue = (1 - norm) * (120 / 360)  # 0° is red, 120° is green

        # Fixed high lightness and medium saturation for pastel effect
        lightness = 0.5  # High lightness for pastel
        saturation = 0.7  # Medium saturation

        # Convert HSL to RGB
        r, g, b = hls_to_rgb(hue, lightness, saturation)

        # Convert to 8-bit RGB and then to hex
        hex_color = "#{:02X}{:02X}{:02X}".format(
            int(r * 255),
            int(g * 255),
            int(b * 255)
        )
        color_dict[int(num)] = hex_color

    return color_dict


def plot_charge_discharge_curves(df_counted_dict, columns_to_plot, input_file_name, folder_name, density_plot=False,
                                 active_material_mass=1):
    max_cycle = 0
    min_cycle = 1000
    max_list_of_cycles = []

    for col in list(df_counted_dict.keys()):
        curr_min = min(list(df_counted_dict[col]['CycleNumber'].unique()))
        curr_max = max(list(df_counted_dict[col]['CycleNumber'].unique()))
        if curr_min < min_cycle:
            min_cycle = curr_min
        if curr_max > max_cycle:
            max_cycle = curr_max
    max_list_of_cycles = range(min_cycle, max_cycle+1)

    color_gradient = generate_color_gradient(max_list_of_cycles)
    styling_dict = {
        "A": '-',
        "B": '--',
        "C": ':',
        "D": '-.',
        "E": '-'
    }

    font = {'family': 'times', 'size': 14}
    plt.rc('font', **font)

    fig, ax = plt.subplots(figsize=(11, 8), dpi=250)
    plt.minorticks_on()
    fig.suptitle("Кривые заряда и разряда при постоянных токах", y=0.98)

    # Create a dictionary to store all handles and labels
    all_handles = []
    all_labels = []

    # Create a dictionary to store legend titles and their handles
    legend_info = {}

    for col in list(df_counted_dict.keys()):
        if input_file_name[col]:
            dict_preprocessing = dict()
            df = df_counted_dict[col].copy()

            if density_plot and columns_to_plot['x'] == "Capacity(mAh)":
                df["Capacity(mAh)"] = np.array(df["Capacity(mAh)"]) / active_material_mass[col]

            # Preprocess data
            for step_type in ['CC Chg', 'CC DChg']:
                dict_preprocessing[step_type] = dict()
                for cycle_number in list(df['CycleNumber'].unique()):
                    dict_preprocessing[step_type][cycle_number] = dict()
                    for axis in columns_to_plot.keys():
                        dict_preprocessing[step_type][cycle_number][axis] = list(
                            df[(df["CycleNumber"] == cycle_number) & (df["Step Type"] == step_type)][
                                columns_to_plot[axis]])

            # Plot data and collect handles/labels
            for step_type in ['CC Chg', 'CC DChg']:
                for cycle_number in list(df['CycleNumber'].unique()):
                    if step_type == 'CC DChg':
                        line, = plt.plot(dict_preprocessing[step_type][cycle_number]['x'],
                                         dict_preprocessing[step_type][cycle_number]['y'],
                                         label=f"{input_file_name[col]}: {cycle_number}",
                                         color=color_gradient[cycle_number],
                                         linestyle=styling_dict[col])
                        all_handles.append(line)
                        all_labels.append(f"{input_file_name[col]}: {cycle_number}")
                    else:
                        plt.plot(dict_preprocessing[step_type][cycle_number]['x'],
                                 dict_preprocessing[step_type][cycle_number]['y'],
                                 color=color_gradient[cycle_number],
                                 linestyle=styling_dict[col])

    # Create a single legend with all entries
    ax.legend(
        handles=all_handles,
        labels=all_labels,
        loc='upper left',
        bbox_to_anchor=(1.02, 1),  # Places legend outside right
        borderaxespad=0.0  # Padding between plot and legend
    )

    plt.tight_layout(rect=[0, 0, 0.85, 0.95])

    # Axis labels
    if density_plot and columns_to_plot['x'] == "Capacity(mAh)":
        plt.xlabel("Удельная емкость, мАч/г")
    elif columns_to_plot['x'] == "Capacity(mAh)":
        plt.xlabel("Емкость, мАч")

    if columns_to_plot['y'] == "Voltage(V)":
        plt.ylabel("Напряжение, В")

    # Create filename
    filename = "_".join([input_file_name[col] for col in df_counted_dict.keys() if input_file_name[col]])
    filename += "_discharge_charge_plot_single_current"
    if density_plot:
        filename += "_density_plot"

    plt.savefig(fname=f"{folder_name}/{filename}.png", format='png')
    plt.close()
    return 0


def plot_discharge_capacity_per_cycle(df_counted_dict, active_material_mass, theoretical_capacity, input_file_name, folder_name):
    fig, ax = plt.subplots(figsize=(13, 8), dpi=250)
    fig.suptitle("Максимальная емкость каждого цикла разрядки, мАч", y=0.98)
    ax.xaxis.set_major_locator(MaxNLocator(integer=True))
    font = {'family': 'times', 'size': 14}
    plt.rc('font', **font)

    marker_dict = {
        "A":'.',
        "B":'v',
        "C":'s',
        "D":'*',
        "E":'o'
    }
    with pd.ExcelWriter(folder_name+"/"+"_max_capacities"+".xlsx") as writer:
        for col in list(df_counted_dict.keys()):
            if input_file_name[col]:
                max_capacities_per_cycle = df_counted_dict[col][df_counted_dict[col]['Step Type'] == 'CC DChg'].groupby('CycleNumber')[
                    'Capacity(mAh)'].max().reset_index()
                max_capacities_per_cycle.to_excel(writer, sheet_name=input_file_name[col])
                plt.plot(list(max_capacities_per_cycle['CycleNumber']), list(max_capacities_per_cycle['Capacity(mAh)']),
                         label='Экспериментальные значения ' + input_file_name[col], marker=marker_dict[col], color='#000000')

                plt.axhline(y=theoretical_capacity[col] * active_material_mass[col], linestyle='--', color='gray',
                            label='Теоретические значения '+ input_file_name[col])
                ax.text(ax.get_xlim()[1] + 0.1, theoretical_capacity[col] * active_material_mass[col],
                        f'{theoretical_capacity[col] * active_material_mass[col]:.2f}')
    plt.ylabel("Емкость, мАч")

    max_capacity = 0
    for col in list(df_counted_dict.keys()):
        if theoretical_capacity[col] * active_material_mass[col] > max_capacity:
            max_capacity = theoretical_capacity[col] * active_material_mass[col]

    plt.ylim(0, max_capacity * 1.5)
    plt.xlabel("Номер цикла")

    plt.legend(
        loc='upper left',
        bbox_to_anchor=(1.02, 1),  # Outside right
        borderaxespad=0.0
    )
    plt.tight_layout(rect=[0, 0, 0.85, 0.95])

    filename = ""
    for col in list(df_counted_dict.keys()):
        if input_file_name[col]:
            filename+=input_file_name[col]

    filename = filename + "_capacity_per_cycle"
    plt.savefig(fname=f"{folder_name}/{filename}.png", format='png')

    return 0

def cut_N_letters_from_a_string(string, N):
    if len(string)<=N:
        return string
    else:
        return string[:N]

def create_folder_name_based_on_input_filenames(input_filename_dict):
    output_name = ""
    for col in list(input_filename_dict.keys()):
        if input_filename_dict[col]:
            output_name += (input_filename_dict[col] + "_")
        return output_name + str(datetime.now().strftime("%Y-%m-%d-%H-%M-%S"))







class LabReportApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Отрисовка графиков CC циклирования Neware")
        self.root.geometry("1200x500")

        # Variables to store user input - now as dictionaries
        self.cycles_inputs = {col: tk.StringVar() for col in ['A', 'B', 'C', 'D', 'E']}
        self.active_material_masses = {col: tk.DoubleVar() for col in ['A', 'B', 'C', 'D', 'E']}
        self.theoretical_capacities = {col: tk.DoubleVar() for col in ['A', 'B', 'C', 'D', 'E']}
        self.input_file_names = {col: tk.StringVar() for col in ['A', 'B', 'C', 'D', 'E']}
        self.display_file_names = {col: tk.StringVar() for col in ['A', 'B', 'C', 'D', 'E']}
        self.is_running = False

        # Main frame
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # Close button (red X)
        self.close_button = ttk.Button(
            self.main_frame,
            text="X",
            style="Red.TButton",
            command=self.on_closing
        )
        self.close_button.pack(side=tk.RIGHT, anchor=tk.NE, padx=5, pady=5)
        self.create_user_guide()

        # Create the input table
        self.create_input_table()

        # Run button
        self.create_run_button()

        # Configure styles
        self.configure_styles()

    def create_user_guide(self):
        # Sample random text (replace with your actual guide)
        guide_text = """Одновременно можно визуализировать от 1 до 5 файлов. Чтобы вывести все циклы, окно "Выбранные циклы" можно оставить пустым. Конкретные циклы можно выбрать перечислением через пробел и/или диапазоном через дефис (например, 1 2 4-5). Масса активного материала нужна для вывода удельных значений, теоретическая удельная емкость - для удобства сравнения эксперимента с теорией. Результат отработки скрипта находится в одной папке с исполняемым скриптом, папка называется plots и генерируется автоматически. """

        guide_label = ttk.Label(
            self.main_frame,
            text=guide_text,
            wraplength=1000,  # Adjust as needed
            justify=tk.LEFT
        )
        guide_label.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=10)




    def configure_styles(self):
        style = ttk.Style()
        style.configure("Red.TButton", foreground="red", font=('Helvetica', 14, 'bold'))
        style.configure("Run.TButton", font=('Helvetica', 14, 'bold'))
        style.configure("Green.TButton", foreground="green", font=('Helvetica', 14, 'bold'))
        style.configure("Drop.TFrame", background="#f0f0f0", relief=tk.RAISED, borderwidth=1)
        style.configure("ColumnHeader.TLabel", font=('Helvetica', 14, 'bold'))
        style.configure("RowHeader.TLabel", font=('Helvetica', 14))

    def create_input_table(self):
        # Create a frame for the table
        table_frame = ttk.Frame(self.main_frame)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        # Create column headers (A-E)
        for col_idx, col in enumerate(['A', 'B', 'C', 'D', 'E']):
            header = ttk.Label(table_frame, text=col, style="ColumnHeader.TLabel", anchor='center')
            header.grid(row=0, column=col_idx+1, padx=5, pady=5, sticky='nsew')

        # Create rows with input fields
        self.create_table_row(table_frame, 1, "Выбранные циклы:", self.cycles_inputs)
        self.create_table_row(table_frame, 2, "Масса активного материала, г:", self.active_material_masses)
        self.create_table_row(table_frame, 3, "Теоретическая удельная емкость, мАч/г:", self.theoretical_capacities)
        self.create_file_row(table_frame, 4, "Выбрать файл (.xlsx):")

        # Configure grid weights
        for i in range(5):
            table_frame.grid_columnconfigure(i+1, weight=1)
        for i in range(4):
            table_frame.grid_rowconfigure(i+1, weight=1)

    def create_table_row(self, parent, row_num, label_text, variables_dict):
        # Row header
        header = ttk.Label(parent, text=label_text, style="RowHeader.TLabel")
        header.grid(row=row_num, column=0, padx=5, pady=5, sticky='w')

        # Input fields
        for col_idx, col in enumerate(['A', 'B', 'C', 'D', 'E']):
            entry = ttk.Entry(parent, textvariable=variables_dict[col])
            entry.grid(row=row_num, column=col_idx+1, padx=5, pady=5, sticky='nsew')

    def create_file_row(self, parent, row_num, label_text):
        # Row header
        header = ttk.Label(parent, text=label_text, style="RowHeader.TLabel")
        header.grid(row=row_num, column=0, padx=5, pady=5, sticky='w')

        # File input fields (drag and drop)
        for col_idx, col in enumerate(['A', 'B', 'C', 'D', 'E']):
            drop_frame = ttk.Frame(parent, style="Drop.TFrame", height=50)
            drop_frame.grid(row=row_num, column=col_idx+1, padx=5, pady=5, sticky='nsew')

            drop_label = ttk.Label(drop_frame, text=f"Файл {col}")
            drop_label.pack(expand=True)

            drop_frame.bind("<Button-1>", lambda e, c=col: self.browse_file(c))
            drop_frame.bind("<Enter>", lambda e: drop_frame.config(style="Drop.TFrame"))
            drop_frame.bind("<Leave>", lambda e: drop_frame.config(style="Drop.TFrame"))

            file_label = ttk.Label(drop_frame, textvariable=self.display_file_names[col])
            file_label.pack(fill=tk.X)

    def browse_file(self, column):
        file_path = filedialog.askopenfilename(
            title=f"Select Input File for column {column}",
            filetypes=[("All files", "*.*")]
        )
        if file_path:
            #filename = os.path.basename(file_path)
            filename = file_path
            self.input_file_names[column].set(filename)
            self.display_file_names[column].set(cut_N_letters_from_a_string(os.path.basename(file_path), 15))
            setattr(self, f'full_file_path_{column}', os.path.basename(file_path))

    def create_run_button(self):
        self.run_button = ttk.Button(
            self.main_frame,
            text="Run",
            style="Run.TButton",
            command=self.run_script
        )
        self.run_button.pack(fill=tk.X, pady=20, ipady=10)

    def run_script(self):
        if self.is_running:
            return

        # Prepare the input dictionaries
        cycles_input_string = {col: var.get() or None for col, var in self.cycles_inputs.items()}
        active_material_mass = {col: var.get() if var.get() != 0 else None for col, var in self.active_material_masses.items()}
        theoretical_capacity = {col: var.get() if var.get() != 0 else None for col, var in self.theoretical_capacities.items()}
        input_file_name = {col: var.get() or None for col, var in self.input_file_names.items()}

        # Change button to yellow (running)
        self.run_button.config(style="Yellow.TButton")
        self.is_running = True
        self.root.update()  # Force UI update

        try:
            # Call the processing directly (no threading)
            self.execute_script(cycles_input_string, active_material_mass, theoretical_capacity, input_file_name)

            # Change button to green (success)
            self.run_button.config(style="Green.TButton")

        except Exception as e:
            # Change button to red (error)
            self.run_button.config(style="RedError.TButton")
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

        finally:
            self.is_running = False
            self.root.update()  # Force UI update

    def execute_script(self, cycles_input_string, active_material_mass, theoretical_capacity, input_file_name):
        # This function remains untouched as requested
        try:
            # Here you would call your actual script
            columns_to_plot = {"x": "Capacity(mAh)", "y": "Voltage(V)"}
            print("Running script with parameters:")
            print(f"Cycles: {cycles_input_string}")
            print(f"Active Material Mass: {active_material_mass}")
            print(f"Theoretical Capacity: {theoretical_capacity}")
            print(f"Input file: {input_file_name}")

            df_counted_dict = dict()
            # Process each column separately
            for col in list(input_file_name.keys()):

                if input_file_name[col]:
                    df = pd.read_excel(input_file_name[col], sheet_name="record")
                    df['DataPoint'] = df['DataPoint'].apply(lambda x: int(x))

                    input_file_name[col] = col + "_" +cut_N_letters_from_a_string(os.path.basename(input_file_name[col]), 5)

                    # Определение циклов
                    df_counted = link_count_cycle_numbers(df)
                    df_counted = df_counted[df_counted['CycleNumber'] > 0].copy()
                    df_counted['CycleNumber'] = df_counted['CycleNumber'].astype(int)

                    # Выделение циклов (из тех, что введены в юзер-интерфейсе)
                    if cycles_input_string[col]:
                        trimming_array = get_cycle_number_list_from_string(cycles_input_string[col])
                        if len(trimming_array) != 0:
                            df_counted = df_counted[df_counted['CycleNumber'].isin(trimming_array)].copy()
                    df_counted_dict[col] = df_counted.copy()




            # Создание папки с output-файлами внутри plots

            folder_name = "plots/"+create_folder_name_based_on_input_filenames(input_file_name)
            os.makedirs(folder_name, exist_ok=True)
            print(f"Folder '{folder_name}' ensured to exist.")

            # Сохранение сырых табличных данных в таблицу

            with pd.ExcelWriter(folder_name + "/" + "raw_cycling_data" + ".xlsx") as writer:
                for col in list(df_counted_dict.keys()):
                    if 'Current(mA)' in df_counted_dict[col].columns:
                        df_counted_dict[col][['Step Type', 'CycleNumber', 'Voltage(V)', 'Capacity(mAh)', 'Current(mA)']].to_excel(writer, sheet_name=input_file_name[col])
                    elif 'Current(μA)' in df_counted_dict[col].columns:
                        df_counted_dict[col][
                            ['Step Type', 'CycleNumber', 'Voltage(V)', 'Capacity(mAh)', 'Current(μA)']].to_excel(writer,
                                                                                                                 sheet_name=
                                                                                                                 input_file_name[
                                                                                                                     col])
                    else:
                        pass

            # Отрисовка графиков
            plot_charge_discharge_curves(df_counted_dict, columns_to_plot, input_file_name,folder_name, density_plot=False)
            plot_charge_discharge_curves(df_counted_dict, columns_to_plot, input_file_name,folder_name, density_plot=True,
                                        active_material_mass=active_material_mass)
            plot_discharge_capacity_per_cycle(df_counted_dict, active_material_mass, theoretical_capacity, input_file_name, folder_name)

            # Change button to green (success)
            self.run_button.config(style="Green.TButton")

        except Exception as e:
            error_msg = str(e)
            # Change button to red (error)
            self.root.after(0, lambda: self.run_button.config(style="Red.TButton"))
            self.root.after(0, lambda: messagebox.showerror("Error_", error_msg))
            print(error_msg)

        finally:
            self.root.after(0, lambda: setattr(self, 'is_running', False))

    def on_closing(self):
        if messagebox.askokcancel("Quit", "Do you want to quit?"):
            self.root.destroy()

if __name__ == "__main__":

    # Создаем папку с выходными файлами
    folder_name = "plots"
    os.makedirs(folder_name, exist_ok=True)
    print(f"Folder '{folder_name}' ensured to exist.")

    root = tk.Tk()

    # Configure additional styles
    style = ttk.Style()
    style.configure("Yellow.TButton", background="yellow", foreground="black")
    style.configure("Green.TButton", background="green", foreground="white")
    style.configure("Red.TButton", background="red", foreground="white")
    style.configure("RedError.TButton", background="red", foreground="white")

    app = LabReportApp(root)
    root.mainloop()