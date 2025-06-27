import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from scipy import stats

def main():
    # Ventana principal
    root = tk.Tk()
    root.title("Análisis de Monitoreo Ambiental")
    root.geometry("1000x800")

    # Scroll vertical
    canvas = tk.Canvas(root)
    scrollbar = ttk.Scrollbar(root, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=scrollbar.set)

    scroll_frame = ttk.Frame(canvas)

    #Usar rueda del mouse
    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
    
    canvas.bind_all("<MouseWheel>", _on_mousewheel)

    # Scroll se ajusta al tamaño del canvas
    def resize_canvas(event):
        canvas.itemconfig(canvas_window, width=event.width)

    # Window para redimensionar
    canvas_window = canvas.create_window((0, 0), window=scroll_frame, anchor="n")
    canvas.bind("<Configure>", resize_canvas)

    scroll_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # Botón para cargar el archivo Excel
    ttk.Button(scroll_frame, text="Cargar Excel", command=lambda: cargar_excel(scroll_frame)).pack(pady=10, anchor="center")

    # Ejecutar aplicación
    root.mainloop()

# Cargar datos
def cargar_excel(parent):
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        return

    # Leer datos
    df = pd.read_excel(file_path)

    # Cálculo de frecuencia por localización
    freq_abs = df['Localización'].value_counts()
    freq_rel = df['Localización'].value_counts(normalize=True) * 100

    freq = pd.DataFrame({
        'Localización': freq_abs.index,
        'Frecuencia Absoluta': freq_abs.values,
        'Frecuencia Relativa': freq_rel.values.round(2)
    })

    freq_frame = ttk.Frame(parent)
    freq_frame.pack(fill="x", pady=10, anchor="center")

    # Tabla de frecuencua
    freq_tree = ttk.Treeview(freq_frame)
    freq_tree.pack(side="left", fill="x", expand=True)

    freq_tree["columns"] = list(freq.columns)
    freq_tree["show"] = "headings"

    for col in freq.columns:
        freq_tree.heading(col, text=col)

    for _, row in freq.iterrows():
        freq_tree.insert("", "end", values=list(row))

    # Cálculo de estadísticas por localización
    stats_df = df.groupby('Localización')['Profundidad (m)'].agg([
        ('Media', 'mean'),
        ('Mediana', 'median'),
        ('Moda', lambda x: stats.mode(x, keepdims=True)[0][0]),
        ('Desv Std', 'std'),
        ('Varianza', 'var')
    ]).reset_index()

    stats_frame = ttk.Frame(parent)
    stats_frame.pack(fill="x", pady=10, anchor="center")

    stats_tree = ttk.Treeview(stats_frame)
    stats_tree.pack(side="left", fill="x", expand=True)

    stats_tree["columns"] = list(stats_df.columns)
    stats_tree["show"] = "headings"
    for col in stats_df.columns:
        stats_tree.heading(col, text=col)
    for _, row in stats_df.iterrows():
        stats_tree.insert("", "end", values=list(row))

    # Gráfico de barras de frecuencia por localización 
    fig1, ax1 = plt.subplots()
    bars = ax1.bar(freq['Localización'], freq['Frecuencia Absoluta'])
    total = freq['Frecuencia Absoluta'].sum()
    for bar in bars:
        height = bar.get_height()
        ax1.text(bar.get_x() + bar.get_width()/2, height, f"{height / total:.1%}", ha='center', va='bottom')
    ax1.set_title("Frecuencia Absoluta por Localización")
    ax1.set_ylabel("Frecuencia Absoluta")
    ax1.set_xlabel("Localización")
    plt.xticks(rotation=45)

    canvas_plot1 = FigureCanvasTkAgg(fig1, master=parent)
    canvas_plot1.draw()
    canvas_plot1.get_tk_widget().pack(pady=10, anchor="center")

    # Gráfico de línea de tendencia por año
    if 'Año de Monitoreo' in df.columns:
        df_year = df.groupby(['Año de Monitoreo', 'Localización'])['Profundidad (m)'].mean().reset_index()
        pivot = df_year.pivot(index='Año de Monitoreo', columns='Localización', values='Profundidad (m)').fillna(0)
        fig2, ax2 = plt.subplots()
        pivot.plot(ax=ax2, marker='o')
        ax2.set_title("Tendencia de Profundidad Promedio por Año")
        ax2.set_ylabel("Profundidad (m)")
        ax2.set_xlabel("Año")
        plt.xticks(rotation=45)

        canvas_plot2 = FigureCanvasTkAgg(fig2, master=parent)
        canvas_plot2.draw()
        canvas_plot2.get_tk_widget().pack(pady=10, anchor="center")

    # Guardar resultados a Excel
    def guardar_resultado():
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if save_path:
            with pd.ExcelWriter(save_path) as writer:
                freq.to_excel(writer, sheet_name="Frecuencia Absoluta", index=False)
                stats_df.to_excel(writer, sheet_name="Estadisticas", index=False)
                if 'Año de Monitoreo' in df.columns:
                    pivot.to_excel(writer, sheet_name="Tendencia")
            messagebox.showinfo("Éxito", "Datos guardados exitosamente")

    ttk.Button(parent, text="Guardar Resultado", command=guardar_resultado).pack(pady=20, anchor="center")

if __name__ == "__main__":
    main()