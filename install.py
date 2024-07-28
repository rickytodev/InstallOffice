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
def check_file():
    if os.path.exists(f):
        shutil.rmtree(f)
        return True

    return True


if check_file() == True:
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
