import tkinter as tk
from tkinter import ttk, messagebox
from pymongo.mongo_client import MongoClient
from pymongo.server_api import ServerApi
from openpyxl import Workbook
from datetime import datetime
from dotenv import load_dotenv
import os
from tkcalendar import DateEntry  # Importa DateEntry desde tkcalendar
from openpyxl.chart import PieChart, Reference

# Cargamos el ENV
load_dotenv()

# Conexión a MongoDB
uri = os.getenv("MONGO_URI")
client = MongoClient(uri, server_api=ServerApi('1'))
db = client[os.getenv("CLIENT")]
collection = db[os.getenv("COLLECTION")]

# Obtener categorías existentes
categorias = [
    "Comida",
    "Transporte",
    "Ocio",
    "Compras",
    "Otros",
    "Universidad",
    "Salud",
    "Servicios",
    "Hogar",
    "Entretenimiento",
    "Viajes",
    "Imprevistos",
    "Educación",
    "Deudas",
    "Tecnología",
    "Ropa",
    "Regalos",
    "Ahorros",
    "Mascotas",
    "Familia",
]


def seleccionar_fecha():
    def seleccionar():
        fecha = cal.get_date()
        fecha_entry.delete(0, tk.END)
        fecha_entry.insert(0, fecha)
        top.destroy()

    top = tk.Toplevel(root)
    top.title("Seleccionar Fecha")
    cal = DateEntry(top, width=12, background='darkblue', foreground='white', borderwidth=2)
    cal.pack(padx=10, pady=10)
    btn_seleccionar = ttk.Button(top, text="Seleccionar", command=seleccionar)
    btn_seleccionar.pack(pady=10)

def agregar_gasto():
    descripcion = entry_descripcion.get()
    monto = entry_monto.get()
    categoria = combo_categorias.get()
    fecha = fecha_entry.get()
    gastoNecesario = entry_necesario.get()

    if descripcion and monto and categoria and fecha and gastoNecesario:
        try:
            monto = float(monto)
        except ValueError:
            messagebox.showerror("Error", "Por favor, introduzca un monto válido (número).")
            return

        timestamp = datetime.now()

        gasto = {
            "descripcion": descripcion,
            "monto": monto,
            "categoria": categoria,
            "fecha": fecha,
            "timestamp": timestamp,
            "necesidad": gastoNecesario
        }

        collection.insert_one(gasto)
        messagebox.showinfo("Éxito", "Gasto agregado correctamente.")
        entry_descripcion.delete(0, tk.END)
        entry_monto.delete(0, tk.END)
    else:
        messagebox.showerror("Error", "Por favor, complete todos los campos.")

def generar_reporte():
    gastos = list(collection.find({}))

    if not gastos:
        messagebox.showwarning("Advertencia", "No hay gastos registrados.")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Gastos"

    ws.append(["Fecha del gasto", "Detalles del gasto", "Monto", "Categoría", "Timestamp", "Necesario", "Mes del gasto"])

    for gasto in gastos:
        fecha = gasto["fecha"]
        descripcion = gasto["descripcion"]
        monto = gasto["monto"]
        categoria = gasto["categoria"]
        gastoNecesario = gasto["necesidad"]
        timestamp = gasto["timestamp"].strftime("%Y-%m-%d %H:%M:%S")
        mes = datetime.strptime(fecha, "%Y-%m-%d").strftime("%B")
        ws.append([fecha, descripcion, monto, categoria, timestamp, gastoNecesario, mes])

    ws.append([])
    ws.append(["Total de gastos", "=SUM(C2:C{})".format(len(gastos) + 1)])

    ws = wb.create_sheet("Resumen")
    ws.append(["Categoría", "Monto"])
    categorias = {}
    for gasto in gastos:
        categoria = gasto["categoria"]
        monto = gasto["monto"]
        if categoria in categorias:
            categorias[categoria] += monto
        else:
            categorias[categoria] = monto

    for categoria, monto in categorias.items():

        ws.append([categoria, monto])

    pie = PieChart()
    labels = Reference(ws, min_col=1, min_row=2, max_row=len(categorias) + 1)
    data = Reference(ws, min_col=2, min_row=1, max_row=len(categorias) + 1)
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(labels)
    pie.title = "Gastos por categoría"
    ws.add_chart(pie, "A10")

    nombre_archivo = f"reporte_gastos_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
    wb.save(nombre_archivo)
    messagebox.showinfo("Éxito", f"Archivo Excel '{nombre_archivo} puede ser encontrado en la carpeta del proyecto")
    os.system(f"start {nombre_archivo}")

def center_window(window, width, height):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width / 2) - (width / 2)
    y = (screen_height / 2) - (height / 2) - 50
    window.geometry('%dx%d+%d+%d' % (width, height, x, y))

root = tk.Tk()
root.title("Gestor de Presupuesto Mensual")

width = 600
height = 400

center_window(root, width, height)

frame_agregar = tk.Frame(root)
frame_agregar.pack(padx=10, pady=10)

label_descripcion = tk.Label(frame_agregar, text="Descripción:")
label_descripcion.grid(row=0, column=0, padx=5, pady=5)
entry_descripcion = tk.Entry(frame_agregar)
entry_descripcion.grid(row=0, column=1, padx=5, pady=5)

label_monto = tk.Label(frame_agregar, text="Monto:")
label_monto.grid(row=1, column=0, padx=5, pady=5)
entry_monto = tk.Entry(frame_agregar)
entry_monto.grid(row=1, column=1, padx=5, pady=5)

label_categoria = tk.Label(frame_agregar, text="Categoría:")
label_categoria.grid(row=2, column=0, padx=5, pady=5)
combo_categorias = ttk.Combobox(frame_agregar, values=categorias)
combo_categorias.grid(row=2, column=1, padx=5, pady=5)
combo_categorias.current(0)

label_fecha = tk.Label(frame_agregar, text="Fecha:")
label_fecha.grid(row=4, column=0, padx=5, pady=5)

fecha_entry = tk.Entry(frame_agregar)  # Cambié a Entry para la fecha
fecha_entry.grid(row=4, column=1, padx=5, pady=5)

btn_seleccionar_fecha = ttk.Button(frame_agregar, text="Seleccionar Fecha", command=seleccionar_fecha)
btn_seleccionar_fecha.grid(row=4, column=2, padx=5, pady=5)

label_necesario = tk.Label(frame_agregar, text="Fue un gasto necesario o innecesario?")
label_necesario.grid(row=5, column=0, padx=5, pady=5)
entry_necesario = ttk.Entry(frame_agregar)
entry_necesario.grid(row=5, column=1, padx=5, pady=5)

btn_agregar = ttk.Button(frame_agregar, text="Agregar Gasto", command=agregar_gasto)
btn_agregar.grid(row=6, columnspan=3, padx=5, pady=5)

frame_reporte = tk.Frame(root)
frame_reporte.pack(padx=10, pady=10)

btn_reporte = ttk.Button(frame_reporte, text="Generar Reporte", command=generar_reporte)
btn_reporte.pack(padx=5, pady=5)

try:
    client.admin.command('ping')
    print("Starting app...")
    root.mainloop()
    print("Closing app...")
except Exception as e:
    print(e)
