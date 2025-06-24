import tkinter
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

class AnalizadorExcel:
    def __init__(self, window):
        self.window = window
        self.window.title("Analizador de Excel")
        self.window.geometry("1100x700")

        # Frame superior que tiene el botón para cargar el archivo y el menú para seleccionar la columna
        self.top_frame = tkinter.Frame(self.window, padx=20, pady=15)
        self.top_frame.pack()

        # Botón para cargar el archivo
        self.load_button = tkinter.Button(self.top_frame, text="Cargar Archivo Excel", command=self.load_excel)
        self.load_button.grid(row=0, column=0, padx=25, pady=5)

        # Menú desplegable para seleccionar columna
        self.var_column = tkinter.StringVar()
        self.columns_menu = ttk.Combobox(self.top_frame, textvariable=self.var_column, state="disabled")
        self.columns_menu.grid(row=0, column=1, padx=25, pady=5)
        self.columns_menu.bind("<<ComboboxSelected>>", self.generate_table)

        # Frame del medio para el analisis
        self.middle_frame = tkinter.Frame(self.window, padx=10, pady=10)
        self.middle_frame.pack(fill=tkinter.BOTH, expand=True)

        #Tabla de frecuencia
        self.tree = ttk.Treeview(self.middle_frame)
        self.tree.pack(pady=10, fill=tkinter.BOTH, expand=True)

        #Grafico
        self.graph_frame = tkinter.Frame(self.middle_frame)
        self.graph_frame.pack(fill=tkinter.BOTH, expand=True)

        # Frame inferior para el botón de guardar análisis
        self.bottom_frame = tkinter.Frame(self.window, pady=15)
        self.bottom_frame.pack()

        # Botón para guardar análisis en Excel
        self.save_button = tkinter.Button(self.bottom_frame, text="Guardar Análisis", command=self.save_analysis, state=tkinter.DISABLED)
        self.save_button.pack()

        # Variables para almacenar datos
        self.df = None
        self.df_frequency = None

    def load_excel(self):
        file = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx *.xls")])
        if not file:
            return  # Si no selecciona archivo, no hace nada

        try:
            self.df = pd.read_excel(file)  # Lee el archivo Excel con pandas
            columns = list(self.df.columns)

            # Activa y llena el menú desplegable con las columnas del Excel
            self.columns_menu['values'] = columns
            self.columns_menu.set("Seleccione una columna")
            self.columns_menu.config(state="readonly")

            # Limpia tabla previa y desactiva botón guardar
            self.save_button.config(state=tkinter.DISABLED)
            self.tree.delete(*self.tree.get_children())

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer el archivo:\n{e}")

    def generate_table(self, event=None):
        column = self.var_column.get()
        if not column or column == "Seleccione una columna":
            return

        # Cálculo de frecuencias
        absolute_frequency = self.df[column].value_counts()
        relative_frequency = self.df[column].value_counts(normalize=True).round(4)

        # DataFrame que junta ambas frecuencias
        df_frec = pd.DataFrame({
            column: absolute_frequency.index,
            'Frecuencia Absoluta': absolute_frequency.values,
            'Frecuencia Relativa': relative_frequency.values
        })

        self.df_frequency = df_frec

        # Mostrar la tabla y activar botón guardar
        self.show_table(self.df_frequency)
        self.plot_graph(self.df_frequency, column)
        self.save_button.config(state=tkinter.NORMAL)

    def show_table(self, df):
        self.tree.delete(*self.tree.get_children())  # Limpia tabla

        self.tree["columns"] = list(df.columns)  # Define columnas
        self.tree["show"] = "headings"           # Oculta columna vacía inicial

        # Configura encabezados y anchos centrados
        for col in df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="center")

        # Inserta filas
        for _, row in df.iterrows():
            self.tree.insert("", "end", values=list(row))

    def plot_graph(self, df, column_name):
        # Limpia el gráfico anterior si existe
        for widget in self.graph_frame.winfo_children():
            widget.destroy()

        fig, ax = plt.subplots(figsize=(8, 5))
        bars = ax.bar(df[column_name], df['Frecuencia Absoluta'], color='skyblue')

        # Añadir porcentajes sobre las barras
        for bar, porcentaje in zip(bars, df['Frecuencia Relativa']):
            altura = bar.get_height()
            ax.annotate(f'{porcentaje*100:.1f}%', 
                        xy=(bar.get_x() + bar.get_width() / 2, altura),
                        xytext=(0, 5),
                        textcoords="offset points",
                        ha='center', va='bottom')

        ax.set_title("Frecuencia de " + column_name)
        ax.set_xlabel(column_name)
        ax.set_ylabel("Frecuencia Absoluta")
        plt.xticks(rotation=45, ha='right')

        fig.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=self.graph_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tkinter.BOTH, expand=True)

    def save_analysis(self):
        save_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Archivo Excel", "*.xlsx")])
        if save_file:
            try:
                self.df_frequency.to_excel(save_file, index=False)
                messagebox.showinfo("Éxito", "El análisis fue guardado exitosamente.")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo guardar el archivo:\n{e}")


window = tkinter.Tk()
app = AnalizadorExcel(window)
window.mainloop()