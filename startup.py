"""
Aplicación que se encargara de la instalación de la paquetería de office
"""

# ? Importar librerías que utilizaremos dentro de la aplicación
from tkinter import Tk, PhotoImage, Label, messagebox, Text, ttk, Checkbutton, StringVar
import requests
import json
import shutil
import zipfile
import subprocess
import os
import platform

password_zipfile = "rJ2U@jYpN#c*6Wf!lG8$7iPZtKvXoAn3qH*9Mz0sD#FgE4uVyTb1wCmXoJgDq1R"
encode_password = password_zipfile.encode("utf-8")


# ? Imágenes que utilizaremos
class Images:
    def __init__(self) -> None:
        self.icon = requests.get(_json["url"] + _json["id_icon"] + ".png")
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


# ? Instalación y activación


class ReactivateLTSC:
    def __init__(self) -> None:
        self.routesOffice = r"C:/Program Files/Microsoft Office/Office16/"
        subprocess.run(f'"{self.routesOffice}activate.exe"')


class ActivateOfficeLTSC:
    def __init__(self) -> None:

        self.activator = r"resources/&e3rPbfhw#umJxD7m^#FbC^ZLE9cAynfGA.zip"
        self.routesOffice = r"C:/Program Files/Microsoft Office/Office16/"

        with zipfile.ZipFile(self.activator, "r") as activate:
            activate.extractall(path=self.routesOffice, pwd=encode_password)
            subprocess.run(f'"{self.routesOffice}activate.exe"')
            Finish()


class InstallerOfficeLTSC:
    def __init__(self) -> None:

        self.path_installer = r"C:/Program Files/OfficeInstaller/"
        self.path_64 = r"resources/EHP2^^8madYwxxLRkdCTR#^gavuUqT%SMk.zip"
        self.path_32 = r"resources/ioxU#i9WYUS!CEyyTsrwE6kaNw%8RV!uBu.zip"

        # checa si existe la ruta donde se guardara la información
        if os.path.exists(self.path_installer):
            # * eliminar la carpeta y vuelve a iniciar el Instalador
            shutil.rmtree(self.path_installer)
            return InstallerOfficeLTSC()
        else:
            # * crea la carpeta y mueve los documentos
            os.mkdir(self.path_installer)

            if platform.architecture()[0] == "64bit":

                with zipfile.ZipFile(self.path_64, "r") as _64bit:
                    _64bit.extractall(path=self.path_installer, pwd=encode_password)
                    _64bit.close()
                    # * mover información de los datos en la misma carpeta
                    shutil.move(
                        f"{self.path_installer}/system-64bits/setup.exe",
                        self.path_installer,
                    )
                    shutil.move(
                        f"{self.path_installer}/system-64bits/config64bits.xml",
                        self.path_installer,
                    )
                    # * eliminación de la carpeta
                    shutil.rmtree(f"{self.path_installer}/system-64bits")
                    # * iniciar instalación
                    subprocess.run(
                        f'"{self.path_installer}setup" /configure "{self.path_installer}config64bits.xml"'
                    )
                    # * eliminar carpeta
                    shutil.rmtree(self.path_installer)
                    # * abrir activación
                    Activation()

            else:

                with zipfile.ZipFile(self.path_32, "r") as _32bit:
                    _32bit.extractall(path=self.path_installer, pwd=self.password)
                    _32bit.close()
                    # * mover información de los datos en la misma carpeta
                    shutil.move(
                        f"{self.path_installer}/system-32bits/setup.exe",
                        self.path_installer,
                    )
                    shutil.move(
                        f"{self.path_installer}/system-32bits/config32bits.xml",
                        self.path_installer,
                    )
                    # * eliminación de la carpeta
                    shutil.rmtree(f"{self.path_installer}/system-32bits")
                    # * iniciar instalación
                    subprocess.run(
                        f'"{self.path_installer}setup" /configure "{self.path_installer}config32bits.xml"'
                    )
                    # * eliminar carpeta
                    shutil.rmtree(self.path_installer)
                    # * abrir activación
                    Activation()


# ? Apartado inicial de la aplicación


class Finish:
    def __init__(self) -> None:
        # * configurar la ventana que mostraremos al usuario
        windows = Tk()
        windows.withdraw()
        windows.title(
            f"https://rickytodev.vercel.app/webpages&applications/{_json['application']}"
        )
        icon = PhotoImage(data=IMG.icon.content)
        windows.iconphoto(False, icon)
        windows.config(background="white")
        Functions.centerWND(windows, 600, 400)
        windows.resizable(False, False)

        # * opción de color morado

        option_violet = Label(
            windows,
            background="#ccd0ff",
            text="Checa si se activo la paquearía de office abriendo una de las aplicaciones de la \npaquetería y dirígete al apartado de la cuenta, en donde dice información del \nproducto debe de decir Producto Activado\n\n1.Si no se activo el Office presiona el botón de 'reactivar'.\n\n2. Si ya esta activado el Office presiona el botón 'finalizar'",
            font=("Arial", 10),
            justify="left",
        )
        option_violet.place(width=500, height=150, x=50, y=50)
        border_option_violet = Label(windows, background="#4b53bc")
        border_option_violet.place(width=4, height=150, x=50, y=50)

        # * line de la parte inferior
        line_bottom = Label(windows, background="#cccccc")
        line_bottom.place(width=600, height=1, y=338)

        # * panel inferior
        panel_bottom = Label(windows, background="#f0f0f0")
        panel_bottom.place(width=600, height=60, y=340)

        # * botón para reactivar
        reactivate_button = ttk.Button(
            windows,
            text="Reactivar",
            padding=(6, 3),
            takefocus=False,
            cursor="hand2",
            command=lambda: [ReactivateLTSC()],
        )
        reactivate_button.place(x=390, y=355)

        # * botón para finalizar
        finish_button = ttk.Button(
            windows,
            text="Finalizar",
            padding=(6, 3),
            takefocus=False,
            cursor="hand2",
            command=lambda: [windows.destroy()],
        )
        finish_button.place(x=490, y=355)
        # * mostrar la ventana al usuario
        windows.protocol("WM_DELETE_WINDOW", False)
        windows.deiconify()
        windows.mainloop()


class Activation:
    def __init__(self) -> None:
        # * configurar la ventana que mostraremos al usuario
        windows = Tk()
        windows.withdraw()
        windows.title(
            f"https://rickytodev.vercel.app/webpages&applications/{_json['application']}"
        )
        icon = PhotoImage(data=IMG.icon.content)
        windows.iconphoto(False, icon)
        windows.config(background="white")
        Functions.centerWND(windows, 600, 400)
        windows.resizable(False, False)

        # * información

        information = Label(
            windows,
            text="El apartado que este de color 'verde' es la acción que se va a ejecutar,si estas en acuerdo con \nlo que dice el apartado, presiona aceptar.",
            font=("Arial", 10),
            justify="left",
            background="white",
        )
        information.place(x=22, y=20)

        # * opción de color red

        option_red = Label(
            windows,
            background="#ff9e9e",
            text="[ Instalación ]\n\nComenzara la instalación de la aplicación por medio de un comando propio por \nlo que si aparece un CMD ntp es normal ya que el programa asi lo ejecutara.",
            font=("Arial", 10),
            justify="left",
        )
        option_red.place(width=500, height=100, x=50, y=80)
        border_option_red = Label(windows, background="red")
        border_option_red.place(width=4, height=100, x=50, y=80)

        # * opción de color verde

        option_green = Label(
            windows,
            background="#aedfbf",
            text="[ Activación ]\n\nEste proceso tardara unos segundo por lo que no inicies la paquetería hasta que \nse active;si no se activa vuelve a presionar en activar; si la activación \nfue correcta presiona en finalizar.",
            font=("Arial", 10),
            justify="left",
        )
        option_green.place(width=500, height=100, x=50, y=210)
        border_option_green = Label(windows, background="#17ac4d")
        border_option_green.place(width=4, height=100, x=50, y=210)

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
            command=lambda: [windows.destroy(), Install()],
        )
        return_button.place(x=390, y=355)

        # * botón para continuar
        accept_button = ttk.Button(
            windows,
            text="Aceptar",
            padding=(6, 3),
            takefocus=False,
            cursor="hand2",
            command=lambda: [windows.destroy(), ActivateOfficeLTSC()],
        )
        accept_button.place(x=490, y=355)
        # * mostrar la ventana al usuario
        windows.protocol("WM_DELETE_WINDOW", False)
        windows.deiconify()
        windows.mainloop()


class Install:
    def __init__(self) -> None:
        # * configurar la ventana que mostraremos al usuario
        windows = Tk()
        windows.withdraw()
        windows.title(
            f"https://rickytodev.vercel.app/webpages&applications/{_json['application']}"
        )
        icon = PhotoImage(data=IMG.icon.content)
        windows.iconphoto(False, icon)
        windows.config(background="white")
        Functions.centerWND(windows, 600, 400)
        windows.resizable(False, False)

        # * información

        information = Label(
            windows,
            text="El apartado que este de color 'verde' es la acción que se va a ejecutar,si estas en acuerdo con \nlo que dice el apartado, presiona aceptar.",
            font=("Arial", 10),
            justify="left",
            background="white",
        )
        information.place(x=22, y=20)

        # * opción de color verde

        option_green = Label(
            windows,
            background="#aedfbf",
            text="[ Instalación ]\n\nComenzara la instalación de la aplicación por medio de un comando propio por \nlo que si aparece un CMD ntp es normal ya que el programa asi lo ejecutara.",
            font=("Arial", 10),
            justify="left",
        )
        option_green.place(width=500, height=100, x=50, y=80)
        border_option_green = Label(windows, background="#17ac4d")
        border_option_green.place(width=4, height=100, x=50, y=80)

        # * opción de color red

        option_red = Label(
            windows,
            background="#ff9e9e",
            text="[ Activación ]\n\nEste proceso tardara unos segundo por lo que no inicies la paquetería hasta que \nse active;si no se activa vuelve a presionar en activar; si la activación \nfue correcta presiona en finalizar.",
            font=("Arial", 10),
            justify="left",
        )
        option_red.place(width=500, height=100, x=50, y=210)
        border_option_red = Label(windows, background="red")
        border_option_red.place(width=4, height=100, x=50, y=210)

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
        accept_button = ttk.Button(
            windows,
            text="Aceptar",
            padding=(6, 3),
            takefocus=False,
            cursor="hand2",
            command=lambda: [windows.destroy(), InstallerOfficeLTSC()],
        )
        accept_button.place(x=490, y=355)

        # * mostrar la ventana al usuario
        windows.protocol("WM_DELETE_WINDOW", False)
        windows.deiconify()
        windows.mainloop()


class Main:
    def __init__(self) -> None:
        # * configurar la ventana que mostraremos al usuario
        windows = Tk()
        windows.withdraw()
        windows.title(
            f"https://rickytodev.vercel.app/webpages&applications/{_json['application']}"
        )
        icon = PhotoImage(data=IMG.icon.content)
        windows.iconphoto(False, icon)
        windows.config(background="white")
        Functions.centerWND(windows, 600, 400)
        windows.resizable(False, False)

        # * términos y condiciones
        bck_terms = Label(windows, background="#cccccc")
        bck_terms.place(width=550, height=250, x=25, y=25)

        terms = Text(windows, background="white", borderwidth=0, font=("Arial", 10))
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
        windows.protocol("WM_DELETE_WINDOW", False)
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
        GET = requests.get("https://rickytodev.vercel.app/")
        if GET.status_code == 200:
            # * leer la información del [settings.json] de la url
            response = requests.get(
                "https://rickytodev.vercel.app/app/officeltsc/settings.json"
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
