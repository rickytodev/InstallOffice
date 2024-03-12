from tkinter import Tk, Label, ttk, PhotoImage, Text, Button, Entry
import platform
import os
import shutil
import zipfile
import subprocess


path_file = 'C:/Program Files/OfficeInstaller'
zip_password = 'rJ2U@jYpN#c*6Wf!lG8$7iPZtKvXoAn3qH*9Mz0sD#FgE4uVyTb1wCmXoJgDq1R'
zip_file = f'tools/system-{platform.architecture()[0]}s.zip'


def center_windows(window: any, width: int, height: int):
    position_x = (window.winfo_screenwidth() - width) // 2
    position_y = (window.winfo_screenheight() - height) // 2
    window.geometry(f'{width}x{height}+{position_x}+{position_y}')


def install():
    if os.path.exists(path_file):

        shutil.rmtree(path_file)

        install()

    else:
        os.mkdir(path_file)

        with zipfile.ZipFile(zip_file, 'r') as data:
            data.extractall(path=path_file, pwd=zip_password.encode())
            data.close()
            shutil.move(f'{path_file}/system-{platform.architecture()[0]}s/setup.exe', path_file)
            shutil.move(f'{path_file}/system-{platform.architecture()[0]}s/config{platform.architecture()[0]}s.xml',
                        path_file)
            shutil.rmtree(f'{path_file}/system-{platform.architecture()[0]}s')
            subprocess.run(f'"{path_file}/setup" /configure "{path_file}/config{platform.architecture()[0]}s.xml"')
            shutil.rmtree(path_file)
            shutil.copy(r'tools/settings-for-activate-packages.zip',
                        'C:/Program Files/Microsoft Office/Office16/activate-office.zip')

            with zipfile.ZipFile('C:/Program Files/Microsoft Office/Office16/activate-office.zip', 'r') as activate_data:
                activate_data.extractall(path='C:/Program Files/Microsoft Office/Office16', pwd=zip_password.encode())
                activate_data.close()
                os.remove('C:/Program Files/Microsoft Office/Office16/activate-office.zip')
                subprocess.run('"C:/Program Files/Microsoft Office/Office16/activate.exe"')
                Finish()


class Finish:
    def __init__(self):
        windows = Tk()
        windows.withdraw()
        windows.overrideredirect(True)
        windows.attributes('-topmost', True)
        center_windows(windows, 400, 190)

        border_bg_color = Label(windows, background='#e7380c')
        border_bg_color.place(width=390, height=180, x=5, y=5)

        border_bg_window = Label(windows)
        border_bg_window.place(width=388, height=178, x=6, y=6)

        logo_img = PhotoImage(file=r'img/logo.png')
        logo = Label(windows, image=logo_img)
        logo.pack(pady=26)

        border_bit = Label(windows, text='Installation Completed !!', font=('Arial', 12), foreground='#e7380c')
        border_bit.place(width=350, height=30, x=25, y=95)

        button_finish = Button(windows, text='Finish', font=('Arial', 11), takefocus=False, background='#e7380c',
                                borderwidth=0, foreground='white', cursor='hand2', justify='center',
                                activeforeground='white', activebackground='#b32b09', command=lambda: [
                windows.destroy(), exit()])
        button_finish.place(width=125, height=35, x=270, y=150)

        windows.deiconify()
        windows.mainloop()

class SelectBits:
    def __init__(self):
        windows = Tk()
        windows.withdraw()
        windows.overrideredirect(True)
        windows.attributes('-topmost', True)
        center_windows(windows, 400, 190)

        border_bg_color = Label(windows, background='#e7380c')
        border_bg_color.place(width=390, height=180, x=5, y=5)

        border_bg_window = Label(windows)
        border_bg_window.place(width=388, height=178, x=6, y=6)

        logo_img = PhotoImage(file=r'img/logo.png')
        logo = Label(windows, image=logo_img)
        logo.pack(pady=26)

        border_bit = Label(windows, background='#e7380c')
        border_bit.place(width=350, height=30, x=25, y=95)

        bit_system = Entry(windows, borderwidth=0, font=('Arial', 9), justify='center', disabledforeground='black')
        bit_system.place(width=348, height=28, x=26, y=96)
        bit_system.insert('end', f'{platform.architecture()[0]}s')
        bit_system.configure(state='disabled')

        button_install = Button(windows, text='Install', font=('Arial', 11), takefocus=False, background='#e7380c',
                                borderwidth=0, foreground='white', cursor='hand2', justify='center',
                                activeforeground='white', activebackground='#b32b09', command=lambda: [
                windows.destroy(), install()])
        button_install.place(width=125, height=35, x=270, y=150)

        windows.deiconify()
        windows.mainloop()


class TerminusAndConditioned:
    def __init__(self):
        windows = Tk()
        windows.withdraw()
        windows.overrideredirect(True)
        windows.attributes('-topmost', True)
        center_windows(windows, 400, 370)

        border_bg_color = Label(windows, background='#e7380c')
        border_bg_color.place(width=390, height=360, x=5, y=5)

        border_bg_window = Label(windows)
        border_bg_window.place(width=388, height=358, x=6, y=6)

        logo_img = PhotoImage(file=r'img/logo.png')
        logo = Label(windows, image=logo_img)
        logo.pack(pady=26)

        bgk_license_color = Label(windows, background='#e7380c')
        bgk_license_color.place(width=350, height=200, x=25, y=95)

        bgk_license = Label(windows)
        bgk_license.place(width=348, height=198, x=26, y=96)

        with open('LICENSE.txt', 'r', encoding='utf-8') as file:
            data = file.read()

        license_text = Text(windows, font=('Arial', 9), foreground='black', takefocus=False,
                            borderwidth=0, background='#f0f0f0')
        license_text.place(width=344, height=194, x=28, y=98)
        license_text.insert('1.0', data)
        license_text.configure(state='disabled')

        button_continue = Button(windows, text='Continue', font=('Arial', 11), takefocus=False, background='#e7380c',
                                 borderwidth=0, foreground='white', cursor='hand2', justify='center',
                                 activeforeground='white', activebackground='#b32b09',
                                 command=lambda: [windows.destroy(), SelectBits()])
        button_continue.place(width=125, height=35, x=270, y=330)

        windows.deiconify()
        windows.mainloop()


class Main:
    def __init__(self):
        windows = Tk()
        windows.withdraw()
        windows.overrideredirect(True)
        windows.attributes('-topmost', True)
        center_windows(windows, 400, 150)

        border_bg_color = Label(windows, background='#e7380c')
        border_bg_color.place(width=390, height=140, x=5, y=5)

        border_bg_window = Label(windows)
        border_bg_window.place(width=388, height=138, x=6, y=6)

        logo_img = PhotoImage(file=r'img/logo.png')
        logo = Label(windows, image=logo_img)
        logo.pack(pady=26)

        style_progress = ttk.Style()
        style_progress.theme_use('clam')
        style_progress.configure('Horizontal.TProgressbar', troughcolor='#f0f0f0', background='#e7380c',
                                 darkcolor="#e7380c", lightcolor="#e7380c", bordercolor="#e7380c")
        progress = ttk.Progressbar(windows, style='Horizontal.TProgressbar', orient='horizontal')
        progress.place(width=350, height=30, x=25, y=95)
        progress.start()

        windows.deiconify()
        windows.after(6000, lambda: [progress.stop(), windows.destroy(), TerminusAndConditioned()])
        windows.mainloop()


if __name__ == "__main__":
    Main()
