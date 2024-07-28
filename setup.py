import os
import shutil
import subprocess
import zipfile
import time
import platform


# import passkey
with open("tools/パスジップ.key", "r") as data:
    k = data.read()
    KEY = k.encode("utf-8")

# file installer
f = "C:/OfficeLTSC/"


# remove before files the updates
def remove_files():
    if os.path.exists(f):
        shutil.rmtree(f)
        return True

    return True


if remove_files() == True:
    # configure message
    i = 1

    # ? message
    print(f"[{i}] Ready for install new update")

    # create new file
    os.mkdir(f)

    # ? message
    i += 1
    print(f"[{i}] New file is create")

    # export .zip in the file
    with zipfile.ZipFile("tools/リソース.zip", "r") as zip_extract:
        zip_extract.extractall(path=f, pwd=KEY)

    # configure break
    time.sleep(1)

    # ? message
    i += 1
    print(f"[{i}] Extract correctly the リソース.zip")

    # ? message
    i += 1
    print(f"[{i}] Init installation for architecture the {platform.architecture()[0]}")

    # init install
    r = f + platform.architecture()[0]
    command = f'"{r}/setup" /configure "{r}/config{platform.architecture()[0]}s.xml"'
    subprocess.run(command, shell=False)

    # ? message
    i += 1
    print(f"[{i}] Installation the Office LTSC is complete")

    # activation
    subprocess.run(
        f"{r}/activate{platform.architecture()[0]}.cmd",
        check=True,
        shell=True,
        capture_output=True,
        text=True,
    )

    # ? message
    i += 1
    print(f"[{i}] Activation the OfficeLTSC is complete")

    # remove installation
    remove_files()

    # ? message
    i += 1
    print(f"[{i}] Remove files the installation is complete")
