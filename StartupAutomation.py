'''Functions for common automation and window repositioning tasks.

Specialized for Windows only.
'''

import win32gui, win32con, win32process, win32api, win32console
#python -m pip install pypiwin32
import time

def run_program(commandLine):
    '''Creates a new process with a new console window.
    
    Returns the handle to the process.
    '''
    startupInfo = win32process.STARTUPINFO()
    (hProcess, hThread, processId, threatId) = win32process.CreateProcess(
        None, commandLine, None, None, False, 
        win32process.CREATE_NEW_CONSOLE, None, None, startupInfo)
    #Not using os.startfile() because I'd like to return the process handle
    #return os.startfile(path)
    return hProcess

def get_all_window_handles():
    '''Returns a list of all window handles for all processes.'''
    windowList = []
    def EnumWindowsProc(hwnd, lParam):
        windowList.append(hwnd)
        return True
    success = win32gui.EnumWindows(EnumWindowsProc, None)
    return windowList

def get_process_window_handles(hProcess):
    '''Returns a list of all window handles for a single process.
    
    hProcess is a process handle, which can be saved from run_program()
    '''
    windowList = []
    pid = win32process.GetProcessId(hProcess)
    def EnumWindowsProc(hwnd, lParam):
        (threadId, windowPid) = win32process.GetWindowThreadProcessId(hwnd)
        if windowPid == pid:
            windowList.append(hwnd)
        return True
    success = win32gui.EnumWindows(EnumWindowsProc, None)
    return windowList

def get_window_handle(windowName):
    '''Returns the window handle for the window with the given title.'''
    return win32gui.FindWindow(None, windowName)

def get_current_window_handle():
    '''Returns the window handle for the console window associated with this 
    process.
    '''
    return win32console.GetConsoleWindow()

def get_window_name(hwnd):
    '''Returns the title of the window with the given window handle.'''
    return win32gui.GetWindowText(hwnd)

def get_window_coords(hwnd):
    '''Retrieves the coordinates of the window with the given handle.
    
    Coordinates returned in the 4-tuple: (left, top, right, bottom)
    '''
    (left, top, right, bottom) = win32gui.GetWindowRect(hwnd)
    return (left, top, right, bottom)

def set_window_coords(hwnd, left, top, right=0, bottom=0):
    '''Repositions a window to new coordinates.
    
    If the right and bottom coordinates are ignored, the window size is 
    preserved.
    '''
    width = right - left
    height = bottom - top
    if (right == 0) or (bottom == 0):
        (curLeft, curTop, curRight, curBottom) = get_window_coords(hwnd)
        width = curRight - curLeft
        height = curBottom - curTop
        right = left + width
        bottom = top + height
    return win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, 
                                 left, top, width, height, 0)

def set_window_size(hwnd, width, height):
    '''Resizes a window without repositioning.'''
    (curLeft, curTop, curRight, curBottom) = get_window_coords(hwnd)
    newRight = curLeft + width
    newBottom = curTop + height
    return setWindowCoords(hwnd, curLeft, curTop, newRight, newBottom)

def bring_window_to_foreground(hwnd):
    '''Sets the foreground window to the window of the given handle.'''
    win32gui.SetForegroundWindow(hwnd)

def press_key(virtualKeyCode):
    '''Presses the given virtual key down and up with no delay.'''
    win32api.keybd_event(virtualKeyCode, 0, 0, 0)
    win32api.keybd_event(virtualKeyCode, 0, 2, 0)

def press_shift_something(virtualKeyCode):
    '''Holds shift while pressing a virtual key, for typing capital letters.'''
    win32api.keybd_event(win32con.VK_SHIFT, 0, 0, 0)
    time.sleep(0.1)
    press_key(virtualKeyCode)
    time.sleep(0.1)
    win32api.keybd_event(win32con.VK_SHIFT, 0, 2, 0)

def press_return():
    '''Virtually hits the enter key.'''
    press_key(win32con.VK_RETURN)

def press_printscreen():
    '''Presses the print screen key.'''
    press_key(win32con.VK_SNAPSHOT)
    time.sleep(0.1)

def type_string(text):
    '''Virtually types the text string.
    
    Attempts typing capital letters and special character keys.
    '''
    for char in text:
        if 65 <= ord(char) <= 90:
            #Capital letters
            press_shift_something(ord(char))
        elif 97 <= ord(char) <= 122:
            #Lowercase letters
            press_key(ord(char.upper()))
        elif 48 <= ord(char) <= 57:
            #Numbers
            press_key(ord(char))
        elif char == '\t':
            press_key(win32con.VK_TAB)
        elif char == '\n':
            press_key(win32con.VK_RETURN)
        elif char == ' ':
            press_key(win32con.VK_SPACE)
        elif char == '+':
            press_key(0xBB)
        elif char == '=':
            press_shift_something(0xBB)
        elif char == ',':
            press_key(0xBC)
        elif char == '<':
            press_shift_something(0xBC)
        elif char == '-':
            press_key(0xBD)
        elif char == '_':
            press_shift_something(0xBD)
        elif char == '.':
            press_key(0xBE)
        elif char == '>':
            press_shift_something(0xBE)
        elif char == '/':
            press_key(0xBF)
        elif char == '?':
            press_shift_something(0xBF)
        elif char == '\'':
            press_key(0xDE)
        elif char == '"':
            press_shift_something(0xDE)
        elif char == '!':
            press_shift_something(ord('1'))
        elif char == '@':
            press_shift_something(ord('2'))
        elif char == '#':
            press_shift_something(ord('3'))
        elif char == '$':
            press_shift_something(ord('4'))
        elif char == '%':
            press_shift_something(ord('5'))
        elif char == '^':
            press_shift_something(ord('6'))
        elif char == '&':
            press_shift_something(ord('7'))
        elif char == '*':
            press_shift_something(ord('8'))
        elif char == '(':
            press_shift_something(ord('9'))
        elif char == ')':
            press_shift_something(ord('0'))
        else:
            print("Can't type '"+char+"' key.", ord(char))
            return False
        time.sleep(0.1)
    return True

def set_cursor(x, y):
    '''Moves the cursor to screen coordinates (x,y).'''
    return win32api.SetCursorPos((x, y))

def left_click():
    '''Simulates a mouse left-click with no delay.'''
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0)

def capslock_on():
    '''Determines if caps lock is pressed, returning True or False.'''
    capsLockOn = win32api.GetKeyState(win32con.VK_CAPITAL)
    if capsLockOn:
        return True
    else:
        return False
