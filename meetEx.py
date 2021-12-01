from PIL import Image
from PIL import ImageGrab

import win32gui
import win32ui
import win32api
import win32con, win32com.client

import cv2
import keyboard
import numpy as np
from ctypes import windll

class MeetExit:
    def __init__(self, RectSize=100):
        self.run = True
        self.hwnd = None
        self.toplist, self.winlist = [], []
        self.bbox = ("minX", "minY", "maxX", "maxY") # or (left, bottom, right, top)
        self.color = [0, 0] # [the current color, how many times the color changes] 

        self.changePos = True
        self.temp = True
        self.pos = (357, 445)
        self.RECTSIZE = RectSize

        def enum_cb(hwnd, results):
            # get all of the active windows
            self.winlist.append((hwnd, win32gui.GetWindowText(hwnd)))
        
        win32gui.EnumWindows(enum_cb, self.toplist)

    def getHWND(self):
        # filter the names of active windows, get the one that has "Meet" in it
        googleMeet = ([(hwnd, title) for hwnd, title in self.winlist if "Meet" in title])
        self.hwnd, meet = googleMeet[0]
        win32gui.ShowWindow(self.hwnd, 4)

        self.generateBbox()

    def generateBbox(self):
        # make bounding box of the chosen window
        # this function will be called on the main loop so
        # that it changes according to current window
        self.bbox = win32gui.GetWindowRect(self.hwnd)
        self.bbox = tuple(int(xy*1.5) for xy in self.bbox)

        self.width, self.height = round(self.bbox[2] - self.bbox[0]), \
                                  round(self.bbox[3] - self.bbox[1])

    def getWindowCapture(self):
        img = ImageGrab.grab(self.bbox, all_screens=True)
        return img
        
    def getActiveWindow(self): # credit to: https://stackoverflow.com/a/24352388
        hwndDC = win32gui.GetWindowDC(self.hwnd)
        mfcDC  = win32ui.CreateDCFromHandle(hwndDC)
        saveDC = mfcDC.CreateCompatibleDC()

        saveBitMap = win32ui.CreateBitmap()
        saveBitMap.CreateCompatibleBitmap(mfcDC, self.width, self.height)
        saveBitMap.CreateCompatibleBitmap(mfcDC, self.width, self.height)
        saveDC.SelectObject(saveBitMap)

        saveDC.BitBlt((0, 0), (self.width, self.height),  mfcDC,  (0, 0),  win32con.SRCCOPY)
        
        # Change the line below depending on whether you want the whole window
        # or just the client area. 
        windll.user32.PrintWindow(self.hwnd, saveDC.GetSafeHdc(), 3)

        bmpinfo = saveBitMap.GetInfo()
        bmpstr = saveBitMap.GetBitmapBits(True)

        im = Image.frombuffer(
            'RGB',
            (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
            bmpstr, 'raw', 'BGRX', 0, 1)

        win32gui.DeleteObject(saveBitMap.GetHandle())
        saveDC.DeleteDC()
        mfcDC.DeleteDC()
        win32gui.ReleaseDC(self.hwnd, hwndDC)

        return im

    def interestRect(self, rawImg):
        # show the user what's in the rectangle on separate window
        self.rectPos()
        x, y = self.pos

        x, y = x*2, y*2
        RECTSIZE = self.RECTSIZE*2
        
        # crop the image of what user wants
        try:
            rectWindow = np.array(rawImg)[y:y+RECTSIZE, x:x+RECTSIZE]
            cv2.imshow("Participant", rectWindow)

            # get average pixel
            avgPixel = np.average(rectWindow)
            return avgPixel
        except Exception as e: print(e)

    def rectPos(self):
        # turn lock/release rectangle position when user presses "esc"
        if keyboard.is_pressed("esc"): 
            self.temp = True
            self.changePos = not self.changePos

        # rectangle of what user wants to see, in this case its the participants
        if self.changePos: 
            if self.temp: win32api.SetCursorPos(self.pos); self.temp = False
            x, y = win32gui.GetCursorPos()
            
            # call 2 if statements just to safe some unnecessary process, get position only when the
            # user wants to
            if (0 < x and x < self.width//2 - self.RECTSIZE) and \
               (0 < y and y < self.height//2 - self.RECTSIZE):
                self.pos = (x, y)
                self.color = [0, 0]

    def checkForChanges(self, current):
        # average pixel changes only when the amount of people changes
        # for example the amount of black color in 1 is less than the amount of black color
        # in 8, with this information we can use it to detect changes of participants inside a meeting
        # (participant entering AND leaving the meeting will be counted as changes, so keep that in mind)
        MAXCHANGES = 20
        if self.color[0] != current and not self.changePos: 
            self.color = [current, self.color[1]+1]

        # for loading
        print("[" + "="*self.color[1] + "-"*(MAXCHANGES-(self.color[1]+1)) + "]" + \
              "%i/%i"%(self.color[1], MAXCHANGES), end="\r")

        if self.color[1] >= MAXCHANGES:
            # shut down the main loop, a bit of a naive solution but at least it works
            self.refreshPage()
            self.run = False

    def refreshPage(self):
        # set the focus onto the window, and refresh the page
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.SendKeys('%')
        win32gui.ShowWindow(self.hwnd, 4)
        win32gui.SetForegroundWindow(self.hwnd)
        
        keyboard.press_and_release("F5")

    def overlay(self, frame):
        gray = np.zeros((int(self.height//2), int(self.width//2), 3), np.uint8)
        gray[:] = (180, 180, 180)

        _x, _y = self.width//8, self.height//4
        gray = cv2.putText(gray, 'You can close this window now', (_x-60, _y), 
                            cv2.FONT_HERSHEY_SIMPLEX, .8, (255, 255, 255), 1, cv2.LINE_AA)
                                    
        gray = cv2.putText(gray, 'Or press "Esc" to set the rectangle position', (_x, _y+20), 
                            cv2.FONT_HERSHEY_SIMPLEX, .4, (255, 255, 255), 1, cv2.LINE_AA)
        
        overlayFrame = cv2.addWeighted(frame, 0.1, gray, 0.8, 5)
        return overlayFrame

    def drawWindow(self, img=None):
        # redraw all of the windows
        if img==None: img = self.getInactiveWindow()

        self.generateBbox()
        avgPixel = self.interestRect(img)
        self.checkForChanges(avgPixel)

        if self.changePos:
            img_np = np.array(img)

            global frame
            frame = cv2.cvtColor(img_np, cv2.COLOR_BGR2RGB)
            frame = cv2.resize(frame, (int(self.width//2), int(self.height//2)))
        
            x, y = self.pos
            cv2.rectangle(frame, (x, y), (x+self.RECTSIZE, y+self.RECTSIZE), (213, 214, 216), 5)
            cv2.imshow("Window", frame)
        
        if not self.changePos and self.temp:
            overlay = self.overlay(frame)
            cv2.imshow("Window", overlay)
            self.temp = False

def main():
    window = MeetExit(50)
    window.getHWND()

    while window.run:
        """
        You can choose which method you want to use for capturing the window, getWindowCapture()
        will capture the window as-is, meaning it will capture whatever is on your screen
        (the input is x, y, width, height so if anything overlaps the window, it will block the view).

        getActiveWindow(), however, wouldn't capture any blocking window; but it takes a LOT of memory.
        it can take up to 50% of memory usage on top of all processes (it can reach memory error). Recommended
        for 1 monitor users.
        """
        img = window.getWindowCapture()
        window.drawWindow(img)

        if cv2.waitKey(200) & 0xFF == ord('q'):
            break

    cv2.destroyAllWindows()

if __name__ == "__main__":
    main()