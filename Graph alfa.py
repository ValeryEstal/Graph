import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import numpy as np
from scipy.optimize import curve_fit
import os

# Функция для создания файла Excel
def create_and_open_excel(n):
    data = {f'X{i+1}': [] for i in range(n)}
    data.update({f'Y{i+1}': [] for i in range(n)})
    data.update({f'P{i+1}': [] for i in range(n)})
    df = pd.DataFrame(data)
    df.to_excel('data.xlsx', index=False)
    os.startfile('data.xlsx')
    return df

def input_values_in_excel(df):
    for col in df.columns:
        for idx in range(len(df)):
            val = float(entry_values.get())
            df.at[idx, col] = val
    df.to_excel('data.xlsx', index=False)

def visualize_and_approximate(df, approx_choice, graph_title, x_label, y_label):
    x_cols = [col for col in df.columns if col.startswith('X')]
    y_cols = [col for col in df.columns if col.startswith('Y')]
    p_cols = [col for col in df.columns if col.startswith('P')]
    fig, ax = plt.subplots()
    markers = ['o', 's', '^', 'v', 'D', '*']
    
    # Словарь для хранения названий зависимостей
    dependency_labels = {}
    approx_formulas = [] # Список для хранения формул аппроксимации

    for i, (x_col, y_col, p_col) in enumerate(zip(x_cols, y_cols, p_cols)):
        x = df[x_col]
        y = df[y_col]
        p = df[p_col]
        marker = markers[i % len(markers)]
        
        # Получение названия зависимости от пользователя
        dependency_label = tk.simpledialog.askstring("Название зависимости", f"Введите название зависимости {y_col} от {x_col}:")
        if dependency_label is None:
            dependency_label = f'{y_col} от {x_col}' # Если пользователь отменил ввод, используем стандартное название
        dependency_labels[y_col] = dependency_label # Сохраняем название

        ax.errorbar(x, y, yerr=p, fmt=marker, label=dependency_labels[y_col]) # Используем полученное название

        if approx_choice != 0:
            if approx_choice == 1:
                popt, pcov = np.polyfit(x, y, 1, w=1/p, cov=True)
                ax.plot(x, np.polyval(popt, x), label=f'y = {popt[0]:.2f}x + {popt[1]:.2f}')
                approx_formulas.append(f'y = {popt[0]:.2f}x + {popt[1]:.2f}, Погрешности: {np.sqrt(np.diag(pcov))}')
            elif approx_choice == 2:
                popt, pcov = curve_fit(lambda x, a, b: a * np.log(x) + b, x, y, sigma=p, absolute_sigma=True)
                ax.plot(x, popt[0] * np.log(x) + popt[1], label=f'y = {popt[0]:.2f}ln(x) + {popt[1]:.2f}')
                approx_formulas.append(f'y = {popt[0]:.2f}ln(x) + {popt[1]:.2f}')
            elif approx_choice == 3:
                popt, pcov = curve_fit(lambda x, a, b: a * np.power(x, b), x, y, sigma=p, absolute_sigma=True)
                ax.plot(x, popt[0] * np.power(x, popt[1]), label=f'y = {popt[0]:.2f}x^{popt[1]:.2f}')
                approx_formulas.append(f'y = {popt[0]:.2f}x^{popt[1]:.2f}')

    ax.set_title(graph_title)
    ax.set_xlabel(x_label)
    ax.set_ylabel(y_label)
    plt.grid(True)
    
    # Добавляем легенду под графиком
    plt.legend(bbox_to_anchor=(0.1, -0.09), loc='center right', ncol=3)
    
    # Выводим тип аппроксимации и погрешности
    if approx_choice != 0:
        ax.text(0.3, -0.1, f"Тип аппроксимации: {['Без аппроксимации', 'Линейная', 'Логарифмическая', 'Степенная'][approx_choice]}", ha='center', va='top', transform=ax.transAxes)
        
        # Формируем текст с формулами аппроксимации и погрешностями
        approx_text = "\n".join(approx_formulas)
        if approx_text:
            ax.text(0.7, -0.1, f"{approx_text}", ha='center', va='top', transform=ax.transAxes)

    # Создание окна для отображения графика
    fig_canvas = FigureCanvasTkAgg(fig, master=root)
    canvas_widget = fig_canvas.get_tk_widget()
    canvas_widget.pack(expand=True, fill="both")
    
    # Отображение графика
    fig_canvas.draw()

    # Кнопка сохранения графика
    def save_graph():
        file_path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG files", "*.png")])
        if file_path:
            plt.savefig(file_path)
            messagebox.showinfo("Успех", "График успешно сохранен")

    save_button = tk.Button(root, text="Сохранить график", command=save_graph)
    save_button.pack()

    # Кнопка печати графика
   # def print_graph():
    #    fig.canvas.print_figure(dpi=fig.dpi)
     #   messagebox.showinfo("Успех", "График успешно отправлен на печать")

    #print_button = tk.Button(root, text="Печать графика", command=print_graph)
    #print_button.pack()

def run_program():
    n = int(entry_n.get())
    df = create_and_open_excel(n)
    print("Пожалуйста, введите значения в таблицу в открытом Excel-документе. Сохраните файл. Нажмите Enter после завершения.")
    input_values_in_excel(df)
    build_graph_button.config(state="normal")

def build_graph():
    approx_choice = int(entry_approx_choice.get())
    graph_title = entry_graph_title.get()
    x_label = entry_x_label.get()
    y_label = entry_y_label.get()

    df = pd.read_excel('data.xlsx')
    visualize_and_approximate(df, approx_choice, graph_title, x_label, y_label)

# Создание окна Tkinter
root = tk.Tk()
root.title("Построение графика")

# Элементы интерфейса
label_n = tk.Label(root, text="Введите количество осей X, Y и P:")
entry_n = tk.Entry(root)
label_values = tk.Label(root, text="Введите значения:")
entry_values = tk.Entry(root)
button_run = tk.Button(root, text="Запустить программу", command=run_program)
build_graph_button = tk.Button(root, text="Построить график", command=build_graph, state="disabled")

label_approx_choice = tk.Label(root, text="Выберите тип аппроксимации (1 - прямая, 2 - логарифм, 3 - степенная функция, 0 - без аппроксимации):")
entry_approx_choice = tk.Entry(root)
label_graph_title = tk.Label(root, text="Введите название графика:")
entry_graph_title = tk.Entry(root)
label_x_label = tk.Label(root, text="Введите название оси X:")
entry_x_label = tk.Entry(root)
label_y_label = tk.Label(root, text="Введите название оси Y:")
entry_y_label = tk.Entry(root)

# Размещение элементов в окне
label_n.pack()
entry_n.pack()
label_values.pack()
entry_values.pack()
button_run.pack()
build_graph_button.pack()

label_approx_choice.pack()
entry_approx_choice.pack()
label_graph_title.pack()
entry_graph_title.pack()
label_x_label.pack()
entry_x_label.pack()
label_y_label.pack()
entry_y_label.pack()

# Увеличиваем размер окна
 # Ширина x Высота root.geometry("800x800")

root.mainloop()
