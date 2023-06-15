import tkinter as tk
from PIL import Image, ImageTk
import serial
import threading
import requests
import requests_cache
import pandas as pd
import time
from datetime import datetime
from tkinter import simpledialog
from tkinter import ttk

requests_cache.install_cache("weather_cache", expire_after=3600)

window = tk.Tk()
window.title("Weather App")
window.configure(bg="#1682BE")
window.attributes('-fullscreen', True)

container = tk.Frame(window, bg="#1682BE", width=400, height=590, padx=32, pady=28)
container.pack()

default_image = Image.new("RGB", (240, 240), color="gray")
default_photo = ImageTk.PhotoImage(default_image)

weather_image_label = tk.Label(container, image=default_photo, bg="#1682BE")
weather_image_label.pack(pady=30)

temperature_label = tk.Label(container, text="22°C", font=("Poppins", 64, "bold"), fg="#fff", bg="#1682BE")
temperature_label.pack()

description_label = tk.Label(container, text="Lobos, Buenos Aires", font=("Poppins", 22), fg="#fff", bg="#1682BE")
description_label.pack()

weather_details = tk.Frame(container, bg="#1682BE")
weather_details.pack(pady=30)

humidity_frame = tk.Frame(weather_details, bg="#1682BE")
humidity_frame.pack(side="left", padx=(10, 0))  # Mover el marco de humedad hacia la izquierda

humidity_icon_label = tk.Label(humidity_frame, text="Humedad (%)", font=("FontAwesome", 24), fg="#fff", bg="#1682BE")
humidity_icon_label.pack()

humidity_text_label = tk.Label(humidity_frame, text="", font=("Roboto", 18), fg="#fff", bg="#1682BE")
humidity_text_label.pack()

tk.Label(weather_details, text="   ", bg="#1682BE").pack(side="left")

pressure_frame = tk.Frame(weather_details, bg="#1682BE")
pressure_frame.pack(side="left", padx=(0, 0))

pressure_icon_label = tk.Label(pressure_frame, text="Presion (hPa)", font=("FontAwesome", 24), fg="#fff", bg="#1682BE")
pressure_icon_label.pack()

pressure_text_label = tk.Label(pressure_frame, text="", font=("Roboto", 18), fg="#fff", bg="#1682BE")
pressure_text_label.pack()

city = "Lobos"
api_key = "c64b5bf7622508aa844014db88936841"

def get_weather_description(city):
    url = f"https://api.openweathermap.org/data/2.5/weather?q={city}&appid={api_key}&units=metric"
    response = requests.get(url)
    data = response.json()
    description = data["weather"][0]["description"]
    return description

def get_weather_data():
    with requests_cache.disabled():
        description = get_weather_description(city)
    return description


def update_weather_data(data):
    temperature, humidity, pressure = data.split(',')
    temperature_label.config(text="{}°C".format(temperature))
    humidity_text_label.config(text="{}%".format(humidity))
    pressure = pressure.rstrip(';')  # Eliminar el punto y coma al final
    if pressure:  # Verificar si el valor de presión es válido
        pressure_text_label.config(text="{}".format(pressure.replace(';', '')))

def update_weather_image(description):
    try:
        image_path = "images/"
        if "clear sky" in description.lower():
            image_path += "clear.png"
        elif "overcast clouds" in description.lower():
            image_path += "clouds.png"
        elif "scattered clouds" in description.lower():
            image_path += "clouds.png"
        elif "light rain" in description.lower():
            image_path += "drizzle.png"
        elif "mist" in description.lower():
            image_path += "mist.png"
        elif "snow" in description.lower():
            image_path += "snow.png"
        else:
            image_path += "default.png"

        image = Image.open(image_path)
        image = image.resize((240, 240))
        photo = ImageTk.PhotoImage(image)
        weather_image_label.configure(image=photo)
        weather_image_label.image = photo
    except (FileNotFoundError, OSError):
        weather_image_label.configure(image=default_photo)
        weather_image_label.image = default_photo

def read_serial_data():
    port = "COM5"  # Cambiar según el puerto serie que estés utilizando
    baudrate = 115200  # Velocidad de transmisión, ajustar según tus requerimientos

    ser = serial.Serial(port, baudrate)

    while True:
        if ser.in_waiting > 0:
            data = ser.readline().decode().strip()
            window.after(0, update_weather_data, data)

    ser.close()

def update_weather():
    description = get_weather_data()
    window.after(0, update_weather_image, description)


is_saving = False  # Variable para controlar el guardado en Excel

def save_to_excel():
    global is_saving
    if is_saving:
        return  # Si ya se está guardando, no hacer nada
    is_saving = True

    df = pd.DataFrame(columns=["Hora", "Temperatura (°C)", "Humedad", "Presion (hPa)"])
    interval = simpledialog.askinteger("Intervalo de espera", "Ingrese la cantidad de segundos de espera:")

    def save_data_to_excel():
        nonlocal df, interval

        # Crear la tabla en la interfaz gráfica
        table_frame = tk.Frame(container, bg="#1682BE")
        table_frame.pack(pady=10)

        tree = ttk.Treeview(table_frame, columns=("Hora", "Temperatura", "Humedad", "Presion"), show="headings", selectmode="browse")
        tree.column("Hora", width=100, anchor="center")
        tree.column("Temperatura", width=100, anchor="center")
        tree.column("Humedad", width=100, anchor="center")
        tree.column("Presion", width=100, anchor="center")
        tree.heading("Hora", text="Hora")
        tree.heading("Temperatura", text="Temperatura (°C)")
        tree.heading("Humedad", text="Humedad")
        tree.heading("Presion", text="Presion (hPa)")
        tree.pack()

        while is_saving:
            temperature = float(temperature_label.cget("text").replace("°C", ""))
            humidity = float(humidity_text_label.cget("text").replace("%", ""))
            pressure = float(pressure_text_label.cget("text").rstrip(';'))
            current_time = datetime.now().strftime("%H:%M:%S")
            df.loc[len(df)] = [current_time, temperature, humidity, pressure]
            df.to_excel("datos.xlsx", index=False, float_format="%.1f")

            # Limpiar la tabla antes de mostrar los datos actualizados
            for item in tree.get_children():
                tree.delete(item)

            # Insertar los datos en la tabla
            for index, row in df.iterrows():
                tree.insert("", "end", values=(row["Hora"], row["Temperatura (°C)"], row["Humedad"], row["Presion (hPa)"]))

            time.sleep(interval)

    threading.Thread(target=save_data_to_excel, daemon=True).start()

def stop_saving():
    global is_saving
    is_saving = False

serial_thread = threading.Thread(target=read_serial_data)
serial_thread.daemon = True
serial_thread.start()

weather_thread = threading.Thread(target=update_weather)
weather_thread.daemon = True
weather_thread.start()

button_frame = tk.Frame(container, bg="#1682BE")
button_frame.pack(pady=10)

save_button = tk.Button(button_frame, text="Guardar en Excel", command=save_to_excel, font=("Poppins", 14, "bold"), bg="#00abf0", fg="#fff", width=17, borderwidth=0)
save_button.pack(side="left", padx=(0, 5))

stop_button = tk.Button(button_frame, text="Detener Guardado", command=stop_saving, font=("Poppins", 14, "bold"), bg="#00abf0", fg="#fff", width=17, borderwidth=0)
stop_button.pack(side="left", padx=(5, 0))

def toggle_fullscreen():
    # Verificar si la ventana está actualmente en modo de pantalla completa
    if window.attributes('-fullscreen'):
        window.attributes('-fullscreen', False)  # Salir del modo de pantalla completa
    else:
        window.attributes('-fullscreen', True)  # Entrar en el modo de pantalla completa

# Crear el botón para alternar entre el modo de pantalla completa y la ventana normal
fullscreen_button = tk.Button(window, text="FS", command=toggle_fullscreen, font=("Poppins", 14, "bold"), bg="#00abf0", fg="#fff", width=4, borderwidth=0)
fullscreen_button.place(x=window.winfo_screenwidth() - 50, y=window.winfo_screenheight() - 65, anchor="se")


def close_app():
    window.destroy()

window.protocol("WM_DELETE_WINDOW", close_app)

window.mainloop()
