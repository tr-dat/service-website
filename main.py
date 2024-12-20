import glob
import hashlib
import json
import os
import random
import os.path
import shutil
import logging
import ssl

from os import path
import subprocess
import threading
import time
from turtle import distance

import psutil
from pathlib import Path

import pyautogui
import requests
import win32gui
import win32con
import win32api
import win32process
import re
from Crypto.PublicKey import RSA
from flask import Flask, request

from flask_cors import cross_origin, CORS


logger = logging.getLogger()
logger.setLevel(logging.INFO)

LOCALAPPDATA = os.getenv('LOCALAPPDATA')
APPDATA = os.getenv('APPDATA')
# EXT_PATH = os.path.join(Path().absolute(), "dist")
PROGRAM_FILE_PATH = os.getenv('ProgramFiles')
PROGRAM_FILE_X86_PATH = os.getenv('ProgramFiles(x86)')

CHROME_APP_PATH = os.path.join(LOCALAPPDATA, "Google\\Chrome")
CHROME_APP = os.path.join(PROGRAM_FILE_PATH, "Google\\Chrome\\Application\\chrome.exe")

EDGE_APP_PATH = os.path.join(LOCALAPPDATA, "Microsoft\\Edge")
EDGE_APP = os.path.join(PROGRAM_FILE_X86_PATH, "Microsoft\\Edge\\Application\\msedge.exe")


# @staticmethod
# def prepareExt():
#     EXT_PATH = r"C:\Users\Admin\PycharmProjects\Tool\extension\violentmonkey\dist"
#     manifestFile = os.path.join(EXT_PATH, "manifest.json")
#     print("Path ", manifestFile)
#     if path.exists(manifestFile):
#         f = open(manifestFile)
#         print("f: ", f)
#         try:
#             data = json.load(f)
#             key = makePublicKeyExt()
#             print("key: ", key)
#             if key:
#                 data['key'] = key
#             json_object = json.dumps(data, indent=4)
#             with open(manifestFile, "w", encoding='utf-8') as outfile:
#                 outfile.write(json_object)
#         except:
#             pass

def all_ok(hwnd, param=[]):
    param.append(hwnd)
    return True


def callbackChrome(hwnd, ps):
    t = win32gui.GetClassName(hwnd)
    if t == "Chrome_WidgetWin_1":
        ps.append(hwnd)
    return True


def getOpeningBrowserChromium():
    tmpS = []
    win32gui.EnumWindows(callbackChrome, tmpS)
    allHwnds = []
    for item in tmpS:
        win32gui.EnumChildWindows(item, all_ok, allHwnds)
        allHwnds.append(item)
    return allHwnds


def mouseMoveWheelz(br, x, y, d):
    global listChromeOpening
    l_param = win32api.MAKELONG(int(float(x)), int(float(y)))
    for hwndId in listChromeOpening:
        if br == get_process_name_by_hwnd(hwndId):
            try:
                p = win32gui.GetParent(hwndId)
                if p == 0:
                    # print("P: ", p)
                    win32gui.SendMessage(hwndId, win32con.WM_ACTIVATE, win32con.WA_ACTIVE, 0)
                    if d == 1:
                        win32api.PostMessage(hwndId, win32con.WM_VSCROLL, win32con.SB_LINEUP, l_param)
                    else:
                        win32api.PostMessage(hwndId, win32con.WM_VSCROLL, win32con.SB_LINEDOWN, l_param)
            except:
                pass

def mouseMoveWheel(x, y, d):
    global listChromeOpening
    l_param = win32api.MAKELONG(int(float(x)), int(float(y)))
    for hwndId in listChromeOpening:
        win32gui.SendMessage(hwndId, win32con.WM_ACTIVATE, win32con.WA_ACTIVE, 0)
        if d == 1:
            win32api.PostMessage(hwndId, win32con.WM_VSCROLL, win32con.SB_LINEUP, l_param)
        else:
            win32api.PostMessage(hwndId, win32con.WM_VSCROLL, win32con.SB_LINEDOWN, l_param)

def send_ctrl_w(window_handle):
    # Gửi thông điệp WM_KEYDOWN với mã phím Ctrl
    win32gui.PostMessage(window_handle, win32con.WM_KEYDOWN, win32con.VK_CONTROL, 0)

    # Gửi thông điệp WM_KEYDOWN với mã phím T
    win32gui.PostMessage(window_handle, win32con.WM_KEYDOWN, ord('W'), 0)

    # Gửi thông điệp WM_KEYUP để giải phóng các phím đã nhấn
    win32gui.PostMessage(window_handle, win32con.WM_KEYUP, ord('W'), 0)
    win32gui.PostMessage(window_handle, win32con.WM_KEYUP, win32con.VK_CONTROL, 0)
    print("Ctrl W")
# def clickToAllHwnd(br, x, y):
#     print("Clicking")
#     global listChromeOpening
#     l_param = win32api.MAKELONG(int(float(x)), int(float(y)))
#     # print(f"The browser is: {active_browser}")
#     # if br == active_browser:
#     for hwndId in listChromeOpening:
#         if br == get_process_name_by_hwnd(hwndId):
#             print("Handle: ", hwndId)
#             try:
#                 p = win32gui.GetParent(hwndId)
#                 if p == 0:
#                     print("Click HWND: ", hwndId)
#                     win32gui.PostMessage(hwndId, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, l_param)
#                     win32gui.PostMessage(hwndId, win32con.WM_LBUTTONUP, win32con.MK_LBUTTON, l_param)
#             except:
#                 pass

def clickToAllHwnd(br, x, y):
    print("Clicking")
    global listChromeOpening
    l_param = win32api.MAKELONG(int(float(x)), int(float(y)))
    for hwndId in listChromeOpening:
        if br == get_process_name_by_hwnd(hwndId):
            # print("zzz", get_process_name_by_hwnd(hwndId))
            # print("Handle: ", hwndId)
            try:
                p = win32gui.GetParent(hwndId)
                if p == 0:

                    # send_ctrl_w(hwndId)
                    print("Click HWND: ", hwndId)
                    win32gui.PostMessage(hwndId, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, l_param)
                    win32gui.PostMessage(hwndId, win32con.WM_LBUTTONUP, win32con.MK_LBUTTON, l_param)
            except:
                pass

def mouseMoveToAllHwnd(x, y):
    global listChromeOpening
    l_param = win32api.MAKELONG(int(float(x)), int(float(y)))
    for hwndId in listChromeOpening:
        try:
            # p = win32gui.GetParent(hwndId)
            # if p == 0:
            win32gui.PostMessage(hwndId, win32con.WM_MOUSEMOVE, win32con.WM_SETCURSOR, win32api.MAKELONG(0, 0))
            time.sleep(0.2)
            win32gui.PostMessage(hwndId, win32con.WM_MOUSEMOVE, win32con.WM_SETCURSOR, l_param)
            time.sleep(0.2)
        except:
            pass


def putText(br):
    try:
        global listChromeOpening
        response = requests.get("https://yt.dencontrungage.com/Api/test")
        if response.status_code == 200:
            data = response.json()
            objects = data.get('data', [])
            all_urls = []
            for obj in objects:
                url = obj.get('url')
                if url:
                    all_urls.append(url)
            for hwndId in listChromeOpening:
                if br == get_process_name_by_hwnd(hwndId):
                    time.sleep(2)
                    p = win32gui.GetParent(hwndId)
                    if p == 0:
                        text = all_urls[random.randint(0, len(all_urls))]
                        print("text: ", text)
                        pyautogui.typewrite(text, interval=0.1)
        else:
            print("status code:", response.status_code)
            return None
    except Exception as e:
        print("Error!: ", e)
        return None

                # x = 200
                # y = 400
                # l_param = win32api.MAKELONG(int(float(x)), int(float(y)))
                # print("Click HWND: ", hwndId)
                # win32gui.PostMessage(hwndId, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, l_param)
                # win32gui.PostMessage(hwndId, win32con.WM_LBUTTONUP, win32con.MK_LBUTTON, l_param)
                # print("write HWND: ", hwndId)
                # pyautogui.hotkey('ctrl', '1')
                # win32gui.SendMessage(hwndId, win32con.WM_SETTEXT, 0, text)
                # pyautogui.typewrite(cmtText, interval=0.0000001)
    # print("leng",leng)
    # pyautogui.press("enter")
    # time.sleep(leng)
def sendCtrlW(br):
    global listChromeOpening
    for hwndId in listChromeOpening:
        if br == get_process_name_by_hwnd(hwndId):
            p = win32gui.GetParent(hwndId)
            if p == 0:
                time.sleep(random.randint(3, 5))
                pyautogui.hotkey('ctrl', '1')
                time.sleep(1)
                pyautogui.hotkey('ctrl', 'w')
                #-------
                # win32gui.PostMessage(hwndId, win32con.WM_KEYDOWN, win32con.VK_CONTROL, 0)
                # # Gửi tin nhắn WM_KEYDOWN cho phím W
                # win32gui.PostMessage(hwndId, win32con.WM_KEYDOWN, ord('W'), 0)
                # # Gửi tin nhắn WM_CHAR cho phím W
                # win32gui.PostMessage(hwndId, win32con.WM_CHAR, ord('W'), 0)
                # # Gửi tin nhắn WM_KEYUP cho phím W
                # win32gui.PostMessage(hwndId, win32con.WM_KEYUP, ord('W'), 0)
                # # Gửi tin nhắn WM_KEYUP cho phím Ctrl
                # win32gui.PostMessage(hwndId, win32con.WM_KEYUP, win32con.VK_CONTROL, 0)

                # win32gui.PostMessage(hwndId, win32con.WM_KEYDOWN, win32con.VK_CONTROL, 0)
                # win32gui.PostMessage(hwndId, win32con.WM_KEYDOWN, ord('W'), 0)
                # time.sleep(1)
                # win32gui.PostMessage(hwndId, win32con.WM_KEYUP, ord('W'), 0)
                # win32gui.PostMessage(hwndId, win32con.WM_KEYUP, win32con.VK_CONTROL, 0)
def mWheel(br, x, y):
    global listChromeOpening
    l_param = win32api.MAKELONG(int(float(x)), int(float(y)))
    for hwndId in listChromeOpening:
        if br == get_process_name_by_hwnd(hwndId):
            p = win32gui.GetParent(hwndId)
            print("hwnd: ", p)
            if p == 0:
                time.sleep(0.2)
                win32gui.SendMessage(hwndId, win32con.WM_ACTIVATE, win32con.WA_ACTIVE, 0)
                # print("p = 0")
                # win32api.PostMessage(hwndId, win32con.WM_VSCROLL, win32con.SB_LINEDOWN, l_param)
                    # win32api.PostMessage(hwndId, win32con.WM_VSCROLL, win32con.SB_LINEUP, l_param)
                    # win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, 0, 0, -3 * 60, 0)
                count = random.randint(1, 3)
                # print("count: ", count)
                for i in range(8):
                    print("i: ", i)
                    # time.sleep(0.2)
                    # win32gui.SendMessage(hwndId, win32con.WM_ACTIVATE, win32con.WA_ACTIVE, 0)
                    time.sleep(random.randint(3, 5))
                    win32api.PostMessage(hwndId, win32con.WM_VSCROLL, win32con.SB_LINEDOWN, l_param)
                if(count == 1):
                    print("count: ", count)
                    time.sleep(random.randint(3, 4))
                    win32api.PostMessage(hwndId, win32con.WM_VSCROLL, win32con.SB_LINEUP, l_param)

                    # # countz = random.randint(3, 6)
                    # # print("countz: ", countz)
                    # time.sleep(3)
                    # for i in range(3):
                    #     time.sleep(random.randint(4, 5))
                    #     win32api.PostMessage(hwndId, win32con.WM_VSCROLL, win32con.SB_LINEUP, l_param)

# def get_browser_process_name():
#     active_window_handle = win32gui.GetForegroundWindow()
#     active_window_pid = win32process.GetWindowThreadProcessId(active_window_handle)[1]
#     # Get the process handle
#     try:
#         desired_access = win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ
#         process_handle = win32api.OpenProcess(desired_access, False, active_window_pid)
#         # process_handle = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, active_window_pid)
#         executable_path = win32process.GetModuleFileNameEx(process_handle, 0)
#         process_name = os.path.basename(executable_path)
#         browsers = ["chrome.exe", "firefox.exe", "msedge.exe", "iexplore.exe", "opera.exe"]
#         for browser in browsers:
#             if browser.lower() in process_name.lower():
#                 return process_name
#     except Exception as e:
#         return None
def get_browser_process_name():
    browser_names = ['chrome.exe', 'firefox.exe', 'msedge.exe', 'opera.exe']
    tmpS = []
    win32gui.EnumWindows(callbackChrome, tmpS)
    allHwnds = []
    for item in tmpS:
        win32gui.EnumChildWindows(item, all_ok, allHwnds)
        allHwnds.append(item)
    for hwnd in allHwnds:
        process_name = get_process_name_by_hwnd(hwnd)
        if process_name in browser_names:
            print(f"Hwnd: {hwnd}, Process Name: {process_name}")

def get_process_name_by_hwnd(hwnd):
    process_id = win32process.GetWindowThreadProcessId(hwnd)[1]
    if process_id > 0:
        try:
            process = psutil.Process(process_id)
            return process.name()
        except psutil.NoSuchProcess:
            return None
    else:
        return None
# CERT_FILE = r'C:\Users\Admin\PycharmProjects\Tool\src\TSL\certificate.pem'
# KEY_FILE = r'C:\Users\Admin\PycharmProjects\Tool\src\TSL\private_key.pem'

HOST_NAME = '127.0.0.1'
PORT = 16242
webApp = Flask(__name__)
cors = CORS(webApp)
webApp.config['CORS_HEADERS'] = 'Content-Type'
# @webApp.route("/localapi", methods=['GET', "POST"])
@webApp.route("/localapi", methods=['GET', "OPTIONS", "POST"])
@cross_origin()

def defaultUrl():
    args = request.args
    if args:
        if args.get("act") == "cl":
            time.sleep(random.randint(2, 3))
            # print("Click here")
            x = args.get("x")
            y = args.get("y")
            # br = args.get("br")
            br = "chrome.exe"
            clickToAllHwnd(br, x, y)
        if args.get("act") == "ctrlT":
            time.sleep(random.randint(2, 4))
            pyautogui.hotkey('ctrl', 't')
            time.sleep(random.randint(3, 5))
            putText("chrome.exe")
            time.sleep(random.randint(2, 4))
            pyautogui.press("enter")
        if args.get("act") == "mWheel":
            x = args.get("x")
            y = args.get("y")
            br = "chrome.exe"
            mWheel(br, x, y)
            # time.sleep(10)

        if args.get("act") == "ctrlW":
            br = "chrome.exe"
            sendCtrlW(br)
            # pyautogui.hotkey('ctrl', '1')
            # time.sleep(3)
            # pyautogui.hotkey('ctrl', 'w')
        if args.get("act") == "ptext":
            putText("chrome.exe")
        if args.get("act") == "mm":
            x = args.get("x")
            y = args.get("y")
            mouseMoveToAllHwnd(x, y)
        if args.get("act") == "wheel":
            x = args.get("x")
            y = args.get("y")
            dz = args.get("direction")
            d = int(dz)
            br = "chrome.exe"
            mouseMoveWheelz(br, x, y, d)

    response = webApp.response_class(
        status=200,
        mimetype='application/json'
    )
    return response

listChromeOpening = []
def createBrowser():
    global CHROME_APP, listChromeOpening
    # # cmd = "\"{}\" -no-default-browser-check --new-window https://google.com".format(CHROME_APP)
    # cmd = "\"{}\" -no-default-browser-check --new-window https://bing.com".format(EDGE_APP)
    # subprocess.Popen(cmd)
    # time.sleep(3)
    listChromeOpening = getOpeningBrowserChromium()
    # makePublicKeyExt()
    # time.sleep(2)


if __name__ == "__main__":
    # prepareExt()
    get_browser_process_name()
    tCheck = threading.Thread(target=createBrowser)
    tCheck.start()
    webApp.run(host=HOST_NAME, port=PORT)

    # Tạo context SSL
    # context = ssl.SSLContext(ssl.PROTOCOL_TLSv1_2)
    # context.load_cert_chain(CERT_FILE, KEY_FILE)
    # webApp.run(host=HOST_NAME, port=PORT, ssl_context=context)


