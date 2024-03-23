import pyautogui
import pygetwindow
import pydirectinput
import time
import random
import mss
import numpy as np
from PIL import Image
import gc
import win32api
import win32con
import bettercam
import cv2

def main():
    """
    Main function for the program
    """

    # Set to True to see the live feed
    visuals = False

    # Time to wait between each note
    deplayBetweenNotes = 0.4

    quitKey = "Q"

    # Finds all Windows with the title "New World"
    newWorldWindows = pygetwindow.getWindowsWithTitle("New World")

    # TODO - Fix this, cause it could choose the wrong window
    # Find the Window titled exactly "New World" (typically the actual game)
    for window in newWorldWindows:
        if window.title == "New World":
            newWorldWindow = window
            break

    # Select that Window
    newWorldWindow.activate()

    # Move your mouse to the center of the game window
    centerW = newWorldWindow.left + (newWorldWindow.width/2)
    centerH = newWorldWindow.top + (newWorldWindow.height/2)

    print(centerH, centerW)
    # pyautogui.moveTo(centerW, centerH)
    # win32api.mouse_event(win32con.MOUSEEVENTF_ABSOLUTE, int(centerW), int(centerH), 0, 0)

    win32api.SetCursorPos((int(centerW), int(centerH)))

    # Clicky Clicky
    time.sleep(0.1)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0)
    time.sleep(0.01)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,0,0)
    time.sleep(0.5)

    # TODO - newWorldWindow.top keeps giving a negative number, so I'm just going to use 0 for now
    region = (
        newWorldWindow.left + round(newWorldWindow.width/3) - 50,
        newWorldWindow.top + 800,
        newWorldWindow.left + (round(newWorldWindow.width/3)) + 50,
        newWorldWindow.top + newWorldWindow.height - 100
    )

    camera = bettercam.create(region=region, output_color="BGRA", max_buffer_len=512)
    camera.start(target_fps=120, video_mode=True)
    # This should resolve issues with the first cast being short
    time.sleep(2)

    while win32api.GetAsyncKeyState(ord(quitKey)) == 0:
        time.sleep(deplayBetweenNotes)
        npImg = np.array(camera.get_latest_frame())

        sctImg = Image.fromarray(npImg)

        if visuals:
            cv2.imshow('Live Feed', npImg)
            if (cv2.waitKey(1) & 0xFF) == ord(quitKey):
                exit()

        try:
            if pyautogui.locate("imgs/click.PNG", sctImg, grayscale=True, confidence=.60) is not None:
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0)
                time.sleep(0.01)
                win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN,0,0)
                time.sleep(0.01)
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,0,0)
                time.sleep(0.01)
                win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP,0,0)
                print("Clicked")
                continue
        except Exception as e:
            pass

        try:
            if pyautogui.locate("imgs/a.PNG", sctImg, grayscale=True, confidence=.60) is not None:
                win32api.keybd_event('A', 0,0,0)  # Press the 'a' key (key down)
                time.sleep(0.02)
                win32api.keybd_event('A',0 ,win32con.KEYEVENTF_KEYUP ,0)  # Release the 'a' key (key up)
                print("a")
                continue
        except Exception as e:
            pass    

        try:
            if pyautogui.locate("imgs/s.PNG", sctImg, grayscale=True, confidence=.60) is not None:
                win32api.keybd_event(ord('S'), 0,0,0)  # Press the 'a' key (key down)
                time.sleep(0.02)
                win32api.keybd_event(ord('S'),0 ,win32con.KEYEVENTF_KEYUP ,0)  # Release the 'a' key (key up)
                print("s")
                continue
        except Exception as e:
            pass   

        try:
            if pyautogui.locate("imgs/d.PNG", sctImg, grayscale=True, confidence=.60) is not None:
                win32api.keybd_event(ord('D'), 0,0,0)  # Press the 'a' key (key down)
                time.sleep(0.01)
                win32api.keybd_event(ord('D'),0 ,win32con.KEYEVENTF_KEYUP ,0)  # Release the 'a' key (key up)
                print("d")
                continue
        except Exception as e:
            pass   
        try:
            if pyautogui.locate("imgs/w.PNG", sctImg, grayscale=True, confidence=.60) is not None:
                win32api.keybd_event(ord('W'), 0,0,0)  # Press the 'a' key (key down)
                time.sleep(0.01)
                win32api.keybd_event(ord('W'),0 ,win32con.KEYEVENTF_KEYUP ,0)  # Release the 'a' key (key up)
                print("w")
                continue
        except Exception as e:
            pass  

        try:
            if pyautogui.locate("imgs/space.PNG", sctImg, grayscale=True, confidence=.60) is not None:
                win32api.keybd_event(ord(' '), 0,0,0)  # Press the 'a' key (key down)
                time.sleep(0.01)
                win32api.keybd_event(ord(' '),0 ,win32con.KEYEVENTF_KEYUP ,0)  # Release the 'a' key (key up)
                print("space")
                continue
        except Exception as e:
            pass  


# Runs the main function
if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        import traceback
        traceback.print_exc()
