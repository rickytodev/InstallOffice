"""
Aplicación que se encargara de la instalación de la paquetería de office
"""

# ? Importar librerías que utilizaremos dentro de la aplicación
from tkinter import Tk, PhotoImage, Label, messagebox, Text, ttk, Checkbutton, StringVar
import requests
import json


# ? Imágenes que utilizaremos
class Images:
    def __init__(self) -> None:
        self.icon = requests.get(_json["url"] + "oAc4AcHiRo2r9AL3Q" + ".png")
        self.app = requests.get(_json["url"] + _json["id_img"] + ".png")


# ? Funciones globales que utilizaremos
class Functions:
    def centerWND(window: any, _width: int, _height: int):
        """
        window => es la ventana que vamos a centrar
        _width => es el ancho de la ventana
        _height => es el alto de la ventana
        """

        # * definimos el apartado de [x]
        x = (window.winfo_screenwidth() - _width) // 2

        # * definimos el apartado de [y]
        y = (window.winfo_screenheight() - _height) // 2

        # * insertar la información dentro de la ventana
        window.geometry(f"{_width}x{_height}+{x}+{y}")


# ? Apartado inicial de la aplicación


class Install:
    def __init__(self) -> None:
        # * configurar la ventana que mostraremos al usuario
        windows = Tk()
        windows.withdraw()
        windows.title(
            f"https://rickytodev.pages.dev/applications&websites/{_json['application']}"
        )
        icon = PhotoImage(data=IMG.icon.content)
        windows.iconphoto(False, icon)
        windows.config(background="white")
        Functions.centerWND(windows, 600, 400)
        windows.resizable(False, False)

        # * barra de progreso
        pgr = ttk.Progressbar(windows, orient="horizontal")
        pgr.place(width=550, height=30, x=25, y=154)

        # * line de la parte inferior
        line_bottom = Label(windows, background="#cccccc")
        line_bottom.place(width=600, height=1, y=338)

        # * panel inferior
        panel_bottom = Label(windows, background="#f0f0f0")
        panel_bottom.place(width=600, height=60, y=340)

        # * botón para cancelar
        cancel_button = ttk.Button(
            windows,
            text="Cancelar",
            padding=(6, 3),
            takefocus=False,
            command=lambda: [windows.destroy()],
            cursor="hand2",
        )
        cancel_button.place(x=26, y=355)

        # * botón para regresar
        return_button = ttk.Button(
            windows,
            text="Regresar",
            padding=(6, 3),
            takefocus=False,
            cursor="hand2",
            command=lambda: [windows.destroy(), Main()],
        )
        return_button.place(x=390, y=355)

        # * botón para continuar
        installer_button = ttk.Button(
            windows,
            text="Instalar",
            padding=(6, 3),
            takefocus=False,
            cursor="hand2",
            command=lambda: [
                self.initINSTALLER(pgr, cancel_button, return_button, installer_button)
            ],
        )
        installer_button.place(x=490, y=355)

        # * mostrar la ventana al usuario
        windows.deiconify()
        windows.mainloop()

    def initINSTALLER(self, pgr: any, bt1: any, bt2: any, bt3: any):
        pgr.start()
        bt1.config(state="disabled", cursor="arrow")
        bt2.config(state="disabled", cursor="arrow")
        bt3.config(state="disabled", cursor="arrow")


class Main:
    def __init__(self) -> None:
        # * configurar la ventana que mostraremos al usuario
        windows = Tk()
        windows.withdraw()
        windows.title(
            f"https://rickytodev.pages.dev/applications&websites/{_json['application']}"
        )
        icon = PhotoImage(data=IMG.icon.content)
        windows.iconphoto(False, icon)
        windows.config(background="white")
        Functions.centerWND(windows, 600, 400)
        windows.resizable(False, False)

        # * términos y condiciones
        bck_terms = Label(windows, background="#cccccc")
        bck_terms.place(width=550, height=250, x=25, y=25)

        terms = Text(windows, background="white", borderwidth=0)
        terms.place(width=548, height=248, x=26, y=26)
        terms.insert("1.0", _json["terms_conditions"])
        terms.config(state="disabled")

        # * botón para aceptar términos y condiciones
        chk_bt_action = StringVar(windows)
        chk_bt_action.set(0)
        chk_bt = Checkbutton(
            windows,
            variable=chk_bt_action,
            text=" Aceptar Términos & Condiciones",
            takefocus=False,
            background="white",
            activebackground="white",
            cursor="hand2",
            command=lambda: [self.nextPAGE(chk_bt_action, continue_button)],
        )
        chk_bt.place(x=26, y=296)

        # * line de la parte inferior
        line_bottom = Label(windows, background="#cccccc")
        line_bottom.place(width=600, height=1, y=338)

        # * panel inferior
        panel_bottom = Label(windows, background="#f0f0f0")
        panel_bottom.place(width=600, height=60, y=340)

        # * botón para cancelar
        cancel_button = ttk.Button(
            windows,
            text="Cancelar",
            padding=(6, 3),
            takefocus=False,
            command=lambda: [windows.destroy()],
            cursor="hand2",
        )
        cancel_button.place(x=26, y=355)

        # * botón para continuar
        continue_button = ttk.Button(
            windows,
            text="Continuar",
            padding=(6, 3),
            takefocus=False,
            state="disabled",
            command=lambda: [windows.destroy(), Install()],
        )
        continue_button.place(x=490, y=355)

        # * mostrar la ventana al usuario
        windows.deiconify()
        windows.mainloop()

    def nextPAGE(self, check: any, button: any):
        if check.get() == "1":
            button.config(state="normal", cursor="hand2")
        else:
            button.config(state="disabled", cursor="arrow")


# ? Iniciar aplicación
if __name__ == "__main__":
    try:
        GET = requests.get("https://www.google.com")
        if GET.status_code == 200:
            # * leer la información del [settings.json] de la url
            response = requests.get(
                "https://rickytodev.pages.dev/installer/officeltsc/settings.json"
            )
            json_data = response.content.decode("utf-8")
            _json = json.loads(json_data)

            # * guarda la clase dentro de un valor global para poder ser utilizada por toda la aplicación
            IMG = Images()

            # * iniciar la aplicación con toda la información necesaria
            Main()

    except requests.exceptions.ConnectionError:
        messagebox.showerror(
            "@ERROR", "OCUPAS UNA CONEXIÓN A INTERNET PARA INICIAR EL INSTALADOR"
        )
