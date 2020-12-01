import threading
import mouse as mouse
import time
import pyscreenshot as ImageGrab
import win32gui
import win32ui
import win32api
import keyboard
import pytesseract
import ctypes
from googletrans import Translator
from mtranslate import translate
import tkinter
import pyttsx3
import win32con
from ctypes import *


# Change Wrong Resolution (DPI 125%) (to 1920x1080)
awareness = ctypes.c_int()
ctypes.windll.shcore.SetProcessDpiAwareness(2)

# CHANGE 'tesseract.exe' PATH
pytesseract.pytesseract.tesseract_cmd = r'C:/Users/USER/AppData/Local/Tesseract-OCR/tesseract.exe'


def draw_rectangle(x: int, y: int):
    dc = win32gui.GetDC(0)
    dcObj = win32ui.CreateDCFromHandle(dc)
    red = win32api.RGB(255, 0, 0)  # Red
    # dcObj.SetBkColor(red)
    hwnd = win32gui.WindowFromPoint((0, 0))
    monitor = (0, 0, win32api.GetSystemMetrics(0), win32api.GetSystemMetrics(1))
    while True:
        if not check:
            break
        n = win32gui.GetCursorPos()
        dcObj.DrawFocusRect((x, y, n[0], n[1]))
        win32gui.InvalidateRect(hwnd, monitor, True)  # Refresh the entire monitor


def draw_text(original_text: str, translation_text: str, coor: list):
    dc = win32gui.GetDC(0)
    dcObj = win32ui.CreateDCFromHandle(dc)
    red = win32api.RGB(255, 0, 0)  # Red
    # dcObj.SetBkColor(red)
    dcObj.DrawFocusRect((coor[0], coor[1], coor[2], coor[3]))

    label = tkinter.Label(text=original_text + '\n' + translation_text, font=('Times', '12'), fg='red', bg='white')
    label.master.overrideredirect(True)  # without window
    label.master.geometry('+' + str(coor[0]) + '+' + str(coor[3]))
    label.master.lift()
    label.master.wm_attributes("-topmost", True)
    label.master.wm_attributes("-disabled", True)
    # label.master.wm_attributes("-transparentcolor", "white")
    label.pack()
    appear_time = int(0.4 * len(translation_text.split()) * 1000)  # appear time text (0.4 sec per word)
    label.after(3000 + appear_time, label.master.destroy)
    label.mainloop()


def say_text(speech: str):
    engine = pyttsx3.init()
    engine.setProperty('rate', 125)
    voices = engine.getProperty('voices')
    engine.setProperty('voice', voices[1].id)
    engine.say(speech)
    engine.runAndWait()


if __name__ == '__main__':

    # Change Cursor
    # SetSystemCursor = windll.user32.SetSystemCursor  # reference to function
    # SetSystemCursor.restype = c_int  # return
    # SetSystemCursor.argtype = [c_int, c_int]  # arguments
    # hCursor = win32api.LoadCursor(0, win32con.OCR_APPSTARTING)

    check = True

    while True:
        # Check if alt+t was pressed
        if keyboard.is_pressed('alt+t'):
            print('alt+t Key was pressed')

            # if SetSystemCursor(hCursor, win32con.OCR_CROSS) == 0:
            #     print('Error in setting the cursor')

            mouse.wait()
            x_start = mouse.get_position()[0]
            y_start = mouse.get_position()[1]

            worker = threading.Thread(target=draw_rectangle, args=(x_start, y_start), )
            worker.start()

            time.sleep(1)
            mouse.wait()
            x_end = mouse.get_position()[0]
            y_end = mouse.get_position()[1]

            # if SetSystemCursor(hCursor, win32con.OCR_NORMAL) == 0:
            #     print('Error in revert the cursor')

            check = False
            time.sleep(1)
            check = True
            # TODO: add checked x<x1 && y<y1
            try:
                img = ImageGrab.grab(bbox=(x_start, y_start, x_end, y_end))
                img.save('text.jpg')
            except Exception as e:
                print('Wrong End Coordinates.')
                continue
            time.sleep(1)
            text = pytesseract.image_to_string('text.jpg', lang='eng')
            text = str.join(" ", text.splitlines())
            print(text)
            translator = Translator()
            translate_counter = 0
            while True:
                try:
                    translate_counter += 1
                    translation = translator.translate(text=text, src='en', dest='iw').text
                    print(translation)
                    break
                except Exception as e:
                    # googletrans problem (sometimes)
                    if translate_counter == 3:
                        translation = translate(text, 'iw')
                        print(translation)
                        break
                    print(e)
                    translator = Translator()

            coordinates = [x_start, y_start, x_end, y_end]
            draw_text(text, translation, coordinates)
            say_text(text)
