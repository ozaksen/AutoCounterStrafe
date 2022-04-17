"""mousemacro.py defines the following functions:

click() -- calls left mouse click
hold() -- presses and holds left mouse button
release() -- releases left mouse button

rightclick() -- calls right mouse click
righthold() -- calls right mouse hold
rightrelease() -- calls right mouse release

middleclick() -- calls middle mouse click
middlehold() -- calls middle mouse hold
middlerelease() -- calls middle mouse release

move(x,y) -- moves mouse to x/y coordinates (in pixels)
getpos() -- returns mouse x/y coordinates (in pixels)
slide(x,y) -- slides mouse to x/y coodinates (in pixels)
              also supports optional speed='slow', speed='fast'
"""

from ctypes import*
from ctypes.wintypes import *
from time import sleep
import win32ui
import win32api
import win32gui
import win32con
import PIL.ImageGrab
import numpy as np


__all__ = ['click', 'hold', 'release', 'rightclick', 'righthold', 'rightrelease', 'middleclick', 'middlehold', 'middlerelease', 'move', 'slide', 'getpos']

# START SENDINPUT TYPE DECLARATIONS
PUL = POINTER(c_ulong)

class KeyBdInput(Structure):
    _fields_ = [("wVk", c_ushort),
             ("wScan", c_ushort),
             ("dwFlags", c_ulong),
             ("time", c_ulong),
             ("dwExtraInfo", PUL)]

class HardwareInput(Structure):
    _fields_ = [("uMsg", c_ulong),
             ("wParamL", c_short),
             ("wParamH", c_ushort)]

class MouseInput(Structure):
    _fields_ = [("dx", c_long),
             ("dy", c_long),
             ("mouseData", c_ulong),
             ("dwFlags", c_ulong),
             ("time",c_ulong),
             ("dwExtraInfo", PUL)]

class Input_I(Union):
    _fields_ = [("ki", KeyBdInput),
              ("mi", MouseInput),
              ("hi", HardwareInput)]

class Input(Structure):
    _fields_ = [("type", c_ulong),
             ("ii", Input_I)]

class POINT(Structure):
    _fields_ = [("x", c_ulong),
             ("y", c_ulong)]
# END SENDINPUT TYPE DECLARATIONS

  #  LEFTDOWN   = 0x00000002,
  #  LEFTUP     = 0x00000004,
  #  MIDDLEDOWN = 0x00000020,
  #  MIDDLEUP   = 0x00000040,
  #  MOVE       = 0x00000001,
  #  ABSOLUTE   = 0x00008000,
  #  RIGHTDOWN  = 0x00000008,
  #  RIGHTUP    = 0x00000010

MIDDLEDOWN = 0x00000020
MIDDLEUP   = 0x00000040
MOVE       = 0x00000001
ABSOLUTE   = 0x00008000
RIGHTDOWN  = 0x00000008
RIGHTUP    = 0x00000010


FInputs = Input * 2
extra = c_ulong(0)

click = Input_I()
click.mi = MouseInput(0, 0, 0, 2, 0, pointer(extra))
release = Input_I()
release.mi = MouseInput(0, 0, 0, 4, 0, pointer(extra))

x = FInputs( (0, click), (0, release) )
#user32.SendInput(2, pointer(x), sizeof(x[0])) CLICK & RELEASE

x2 = FInputs( (0, click) )
#user32.SendInput(2, pointer(x2), sizeof(x2[0])) CLICK & HOLD

x3 = FInputs( (0, release) )
#user32.SendInput(2, pointer(x3), sizeof(x3[0])) RELEASE HOLD


def move(x,y):
    #windll.user32.SetCursorPos(x,y)
    win32api.mouse_event(win32con.MOUSEEVENTF_MOVE, x, y)

def getpos():
    global pt
    pt = POINT()
    windll.user32.GetCursorPos(byref(pt))
    return pt.x, pt.y

def slide(a,b,speed=0):
    while True:
        if speed == 'slow':
            sleep(0.005)
            Tspeed = 2
        if speed == 'fast':
            sleep(0.001)
            Tspeed = 5
        if speed == 0:
            sleep(0.001)
            Tspeed = 3

        x = getpos()[0]
        y = getpos()[1]
        if abs(x-a) < 5:
            if abs(y-b) < 5:
                break

        if a < x:
            x -= Tspeed
        if a > x:
            x += Tspeed
        if b < y:
            y -= Tspeed
        if b > y:
            y += Tspeed
        move(x,y)


def click():
    windll.user32.SendInput(2,pointer(x),sizeof(x[0]))

def hold():
    windll.user32.SendInput(2, pointer(x2), sizeof(x2[0]))

def release():
    windll.user32.SendInput(2, pointer(x3), sizeof(x3[0]))


def rightclick():
    windll.user32.mouse_event(RIGHTDOWN,0,0,0,0)
    windll.user32.mouse_event(RIGHTUP,0,0,0,0)

def righthold():
    windll.user32.mouse_event(RIGHTDOWN,0,0,0,0)

def rightrelease():
    windll.user32.mouse_event(RIGHTUP,0,0,0,0)


def middleclick():
    windll.user32.mouse_event(MIDDLEDOWN,0,0,0,0)
    windll.user32.mouse_event(MIDDLEUP,0,0,0,0)

def middlehold():
    windll.user32.mouse_event(MIDDLEDOWN,0,0,0,0)

def middlerelease():
    windll.user32.mouse_event(MIDDLEUP,0,0,0,0)

def PressKey(ascii):
    win32api.keybd_event(ascii,0,0,0) 

def ReleaseKey(ascii):
    win32api.keybd_event(ascii,0,win32con.KEYEVENTF_KEYUP,0) #Realize the Ctrl button


def KeyPress():
    sleep(3)
    PressKey(65) # press Q
    sleep(.05)
    #ReleaseKey(0x10) #release Q

def get_pixel_colour(i_x, i_y):
    import win32gui
    i_desktop_window_id = win32gui.GetDesktopWindow()
    i_desktop_window_dc = win32gui.GetWindowDC(i_desktop_window_id)
    long_colour = win32gui.GetPixel(i_desktop_window_dc, i_x, i_y)
    i_colour = int(long_colour)
    return (i_colour & 0xff), ((i_colour >> 8) & 0xff), ((i_colour >> 16) & 0xff)

def get_pixel_colourImg(i_x, i_y):
	
    img = PIL.ImageGrab.grab(bbox =(1920, 6, 3839, 1085), all_screens=True)
    r=0
    g=1
    b=2

    r_query = 254
    g_query = 254
    b_query = 84
    array = np.array(img)
    res = np.where((array[:,:,r] >= r_query) & (array[:,:,g] >= g_query) & (array[:,:,b] <= b_query))
    coordinates = zip(res[0], res[1]) 
    unique_coordinates = list(set(list(coordinates)))
    print(unique_coordinates)
    unique_coordinates.sort(key=lambda y: y[1], reverse= True)
    #img.show()
    print(1920+unique_coordinates[0][0],6+unique_coordinates[0][1])
    curr = getpos()
    slide((curr[0]-(1920+unique_coordinates[0][0])) * 2.55,(curr[1]-(1920+unique_coordinates[0][1])) * 2.55)
    #return img.load()[i_x, i_y]
 

def scan_screen():
    m = getpos()
    col = get_pixel_colour(m[0],m[1])
    print (get_pixel_colourImg(m[0],m[1]))
    #print(m)
    #print(col)
    if  85 > col[0] > 70 and 50 > col[1] > 60 and 25 > col[0] > 35:
        print('found')
        exit()
        return m
    return False


if __name__ == "__main__":

    state_left = win32api.GetKeyState(17)  # Left button down = 0 or 1. Button up = -127 or -128
    left_key = False
    back_key = False
    right_key = False
    forward_key = False
    #scan = scan_screen()
    # sleep(3)
    # print(getpos())
    # c = getpos()
    # move(int((1920- c[0] )*2.55), int((6 - c[0]) * 2.55))
    #exit()
    while True:
        a = win32api.GetKeyState(17)
        # sleep(1)
        #exit()
        #if scan != False:
        #    move(scan[0],scan[1])
        #print(get_pixel_colour(mouse_pos[0], mouse_pos[1]))
        # print(getpos())
        
        if a != state_left:  # Button state changed
            state_left = a
            #print(a)
            if a < 0:
                print('Left Button Pressed')
                if win32api.GetAsyncKeyState(ord('A')) < 0:
                    print('pressing D')
                    PressKey(68)
                    left_key = True
                if win32api.GetAsyncKeyState(ord('W')) < 0:
                    print('pressing S')
                    PressKey(83)
                    back_key = True
                if win32api.GetAsyncKeyState(ord('S')) < 0:
                    print('pressing W')
                    PressKey(87)
                    forward_key = True
                if win32api.GetAsyncKeyState(ord('D')) < 0:
                    print('pressing A')
                    PressKey(65)
                    right_key = True
            else:
                print('Left Button Released')
                if left_key:
                    print('pressing D')
                    ReleaseKey(68)
                if back_key:
                    print('pressing S')
                    ReleaseKey(83)
                if forward_key:
                    print('pressing W')
                    ReleaseKey(87)
                if right_key:
                    print('pressing A')
                    ReleaseKey(65)
                left_key = False
                back_key = False
                right_key = False
                forward_key = False
        sleep(0.001)