"""
Importación de las librerías que vamos a utilizar para la mayoría de la aplicación
"""

from tkinter import Tk, Label, PhotoImage
import json

"""
Leer información que esta dentro del json
"""

with open("app/data.json", "r") as data:
    jsonDATA = json.load(data)


"""
Funciones globales que vamos a estar utilizando para toda la aplicación de instalación
"""


class FUNCTIONS:
    def centerWindows(window: any, width: int, height: int):
        """
        window => recibe el parámetro que esta manejando la ventana
        width => es el ancho de la ventana
        height => es el alto de la ventana
        """

        # * configuramos el valor de la posición [x] y de la posición [y]
        x = (window.winfo_screenwidth() - width) // 2
        y = (window.winfo_screenheight() - height) // 2

        # * insertar la información dentro del geometry de la ventana

        window.geometry(f"{width}x{height}+{x}+{y}")


# TODO: Aquí estará toda la aplicación que usaremos para la instalación de office ltsc


# ! inicio de la aplicación
class Main:
    def __init__(self) -> None:
        # * configuración de la venta
        windows = Tk()
        windows.withdraw()
        windows.title("Instalador de Office LTSC")
        windows.resizable(False, False)
        windows.iconbitmap(jsonDATA["icon"])
        FUNCTIONS.centerWindows(windows, 600, 400)

        # * configuración del panel lateral
        panel_left = Label(windows, background="#0079ff")
        panel_left.place(width=250, height=500)

        # * imagen del panel lateral
        image_panel_left_resource = PhotoImage(file=jsonDATA["logo"])
        image_panel_left = Label(
            windows, image=image_panel_left_resource, background="#0079ff"
        )
        image_panel_left.place(x=50, y=50)

        # * mostrar la ventana al usuario
        windows.deiconify()
        windows.mainloop()


# ? iniciar la aplicación
if __name__ == "__main__":
    Main()
