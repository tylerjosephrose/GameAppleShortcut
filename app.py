import json
import os
import glob
import subprocess
import difflib
import ctypes
import cv2

from PIL import Image
import pytesseract
import pandas as pd
from io import StringIO
import pyautogui
import time
import numpy as np
import os

from flask import Flask, request, Response, jsonify


TEXT_TO_CLICK = 'Play'


app = Flask(__name__)


SCAN_DIRECTORIES = ['C:\\Games\\', 'C:\\Steam\\steamapps\\common', 'D:\\Games', 'D:\\Program Files (x86)\\Steam\\steamapps\common',
                    'E:\\Battle.net', 'E:\\Program Files (x86)\\Ubisoft', 'E:\\SteamLibrary\\steamapps\common']


@app.route('/apps')
def apps():
    apps = _get_list_of_apps()
    return jsonify(list(apps.keys()))


@app.route('/setup')
def setup():
    query_params = request.args
    if 'app' in query_params:
        app_to_run = query_params['app']
        return _setup(app_to_run)
    else:
        return Response('No app provided', status=400)


def _minimize_windows():
    pyautogui.hotkey('winleft', 'd')


def _get_list_of_apps():
    results = {}
    for dir in SCAN_DIRECTORIES:
        #print(f'Searching: {dir}')
        os.chdir(dir)
        for subdir in os.listdir():
            if os.path.isdir(subdir):
                os.chdir(subdir)
                executables = glob.glob('*.exe')
                #print(executables)
                if len(executables) > 0:
                    sorted_executables = difflib.get_close_matches(subdir, executables, cutoff=0.1)
                    for exe in sorted_executables:
                        if 'Launcher' in exe:
                            sorted_executables = [exe]
                            break
                    #print(f'Closest match for {subdir} is {sorted_executables}')
                    results[subdir] = os.path.join(dir, subdir, sorted_executables[0])
                os.chdir('..')
            else:
                continue

    return results


def _setup(app):
    _minimize_windows()
    apps = _get_list_of_apps()
    if app in apps:
        _kill_process('BeatSyncConsole.exe')
        # The following line is what I want to use, but it requires python3.7
        executable = os.path.basename(apps[app])
        path = os.path.dirname(apps[app])
        if 'Launcher' in executable:
            subprocess.Popen(executable, cwd=path, creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP, close_fds=True, shell=True)
            time.sleep(5)
            pyautogui.moveTo(0, 0)
            pyautogui.FAILSAFE = False

            im = pyautogui.screenshot()
            base_path = os.path.dirname(os.path.realpath(__file__))
            image_file = os.path.join(base_path, 'Screenshot.png')
            im.save(image_file)

            pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
            tessdata_dir_config = '--tessdata-dir "C:\\Program Files\\Tesseract-OCR\\tessdata"'
            # Example config: '--tessdata-dir "C:\\Program Files (x86)\\Tesseract-OCR\\tessdata"'
            # It's important to include double quotes around the dir path.

            # Since the blue and white for the play button are too close in color, convert the image to black and white to make the lines distinct
            im = cv2.imread(image_file, cv2.IMREAD_UNCHANGED)
            grayImage = cv2.cvtColor(im, cv2.COLOR_BGR2GRAY)
  
            (thresh, im) = cv2.threshold(grayImage, 127, 255, cv2.THRESH_BINARY)
            cv2.imwrite(image_file, im)
            im = Image.fromarray(im)

            box_boundaries = pytesseract.image_to_data(im, lang='eng', config=tessdata_dir_config)
            #print(box_boundaries)
            box_boundaries_string = StringIO(box_boundaries)
            # engine='python' is there to prevent reading EOF in the results
            df = pd.read_csv(box_boundaries_string, sep='\t', engine='python')
            #with pd.option_context('display.max_rows', None, 'display.max_columns', None):  # more options can be specified also
                #print(df)

            entry = df[df['text'].str.contains(TEXT_TO_CLICK, na=False)]
            #print(entry)
            center = (float((entry.left+entry.width + entry.left)/2), float((entry.top+entry.height + entry.top)/2))
            #print(im.size)
            center_pct = (center[0]/im.size[0], center[1]/im.size[1])

            screensize = pyautogui.size()
            location = (center_pct[0]*screensize[0], center_pct[1]*screensize[1])

            pyautogui.moveTo(location)
            pyautogui.click()
            return Response(f'Success: Running {app}', status=200)
        else:
            subprocess.Popen(executable, cwd=path, creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP, close_fds=True, shell=True)
            #subprocess.Popen(apps[app], close_fds=True)
            return Response(f'Success: Running {app}', status=200)
    else:
        return Response(f'Failed: {app} does not exist', status=400)


def _kill_process(name):
    os.system(f'taskkill /IM "{name}" /F')


if __name__ == '__main__':
    ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)
    app.run(debug=True, host='0.0.0.0')
