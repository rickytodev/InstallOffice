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
class ActivateOfficeLTSC:
    def __init__(self) -> None:

        self.activator = r"resources/&e3rPbfhw#umJxD7m^#FbC^ZLE9cAynfGA.zip"
        self.password = (
            "rJ2U@jYpN#c*6Wf!lG8$7iPZtKvXoAn3qH*9Mz0sD#FgE4uVyTb1wCmXoJgDq1R"
        )
        self.routesOffice = r"C:/Program Files/Microsoft Office/Office16/"

        with zipfile.ZipFile(self.activator, "r") as activate:
            activate.extractall(path=self.routesOffice, pwd=self.password)
            subprocess.run(f'"{self.routesOffice}activate.exe"')


class InstallerOfficeLTSC:
    def __init__(self, progress: any) -> None:

        self.path_installer = r"C:/Program Files/OfficeInstaller"
        self.password = (
            "rJ2U@jYpN#c*6Wf!lG8$7iPZtKvXoAn3qH*9Mz0sD#FgE4uVyTb1wCmXoJgDq1R"
        )
        self.path_64 = r"resources/EHP2^^8madYwxxLRkdCTR#^gavuUqT%SMk.zip"
        self.path_32 = r"resources/ioxU#i9WYUS!CEyyTsrwE6kaNw%8RV!uBu.zip"

        # checa si existe la ruta donde se guardara la información
        if os.path.exists(self.path_installer):
            # * eliminar la carpeta y vuelve a iniciar el Instalador
            shutil.rmtree(self.path_installer)
            return InstallerOfficeLTSC(progress)
        else:
            # * crea la carpeta y mueve los documentos
            os.mkdir(self.path_installer)

            if platform.architecture[0] == "64bit":

                with zipfile.ZipFile(self.path_64, "r") as _64bit:
                    _64bit.extractall(path=self.path_installer, pwd=self.password)
                    _64bit.close()
                    # * mover información de los datos en la misma carpeta
                    shutil.move(
                        f"{self.path_installer}/system-64bits/setup.exe",
                        self.path_installer,
                    )
                    shutil.move(
                        f"{self.path_installer}/system-64bits/config.xml",
                        self.path_installer,
                    )
                    # * eliminación de la carpeta
                    shutil.rmtree(f"{self.path_installer}/system-64bits")
                    # * iniciar instalación
                    subprocess.run(
                        f'"{self.path_installer}/setup" /configure "{self.path_installer}/config/config64bits.xml"'
                    )
                    # * eliminar carpeta
                    shutil.rmtree(self.path_installer)

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
                        f"{self.path_installer}/system-32bits/config.xml",
                        self.path_installer,
                    )
                    # * eliminación de la carpeta
                    shutil.rmtree(f"{self.path_installer}/system-32bits")
                    # * iniciar instalación
                    subprocess.run(
                        f'"{self.path_installer}/setup" /configure "{self.path_installer}/config/config32bits.xml"'
                    )
                    # * eliminar carpeta
                    shutil.rmtree(self.path_installer)


# ? Apartado inicial de la aplicación
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
        InstallerOfficeLTSC(pgr)


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
