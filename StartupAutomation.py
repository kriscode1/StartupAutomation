'''Functions for common automation and window repositioning tasks.

Author: Kristofer Christakos
Last modified date: March 22, 2017

Specialized for Windows only.
'''

import win32gui, win32con, win32process, win32api, win32console
#python -m pip install pypiwin32
import time

MULI_KEYPRESS_SLEEP_DELAY = 0.1

#################### Process Related Functions ####################

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

#################### Window Related Functions ####################

def get_all_window_handles():
    '''Returns a list of all window handles for all processes.'''
    windowList = []
    def EnumWindowsProc(hwnd, lParam):
        windowList.append(hwnd)
        return True
    success = win32gui.EnumWindows(EnumWindowsProc, None)
    return windowList

def get_process_window_handles(process_handle):
    '''Returns a list of all window handles for a single process.
    
    process_handle can be retrieved and saved from run_program()
    '''
    windowList = []
    pid = win32process.GetProcessId(process_handle)
    def EnumWindowsProc(hwnd, lParam):
        (threadId, windowPid) = win32process.GetWindowThreadProcessId(hwnd)
        if windowPid == pid:
            windowList.append(hwnd)
        return True
    success = win32gui.EnumWindows(EnumWindowsProc, None)
    return windowList

def get_current_window_handle():
    '''Returns the window handle for the console window associated with this 
    process.
    '''
    return win32console.GetConsoleWindow()

def get_window_handle(window_name):
    '''Returns the window handle for the window with the given title.'''
    return win32gui.FindWindow(None, window_name)

def get_window_name(window_handle):
    '''Returns the title of the window with the given window handle.'''
    return win32gui.GetWindowText(window_handle)

def get_window_coords(window_handle):
    '''Retrieves the coordinates of the window with the given handle.
    
    Coordinates returned in the 4-tuple: (left, top, right, bottom)
    '''
    (left, top, right, bottom) = win32gui.GetWindowRect(window_handle)
    return (left, top, right, bottom)

def set_window_coords(window_handle, left, top, right=0, bottom=0):
    '''Repositions a window to new coordinates.
    
    If the right and bottom coordinates are ignored, the window size is 
    preserved.
    '''
    width = right - left
    height = bottom - top
    if (right == 0) or (bottom == 0):
        (curLeft, curTop, curRight, curBottom) = get_window_coords(window_handle)
        width = curRight - curLeft
        height = curBottom - curTop
        right = left + width
        bottom = top + height
    return win32gui.SetWindowPos(window_handle, win32con.HWND_TOP, 
                                 left, top, width, height, 0)

def set_window_size(window_handle, width, height):
    '''Resizes a window without repositioning.'''
    (curLeft, curTop, curRight, curBottom) = get_window_coords(window_handle)
    newRight = curLeft + width
    newBottom = curTop + height
    return setWindowCoords(window_handle, curLeft, curTop, newRight, newBottom)

def bring_window_to_foreground(window_handle):
    '''Sets the foreground window to the window of the given handle.'''
    win32gui.SetForegroundWindow(window_handle)

def get_foreground_window():
    '''Retrieves the window handle of the foreground window.'''
    return win32gui.GetForegroundWindow()

def window_is_foreground(window_handle):
    '''Confirms the given window is the foreground window.'''
    if (window_handle != 0) and (window_handle == get_foreground_window()): return True
    return False

def get_parent_window_at_coordinates(x, y):
    '''Returns the handle of the parent window at (x,y) screen coordinates.'''
    window_handle = win32gui.WindowFromPoint((x,y))
    if window_handle == 0: window_handle = None
    return window_handle

def get_child_window_at_coordinates(x, y):
    '''Returns the handle of the first child window at (x,y) screen coordinates.
    
    Returns None if no child window contains the coordinates.
    '''
    parent_handle = win32gui.WindowFromPoint((x,y))
    if parent_handle == 0: parent_handle = None
    child_handle = None
    if (parent_handle != None):
        child_handle = win32gui.ChildWindowFromPoint(parent_handle, (x,y))
        if child_handle == 0: child_handle = None
        if child_handle == parent_handle: child_handle = None
    return child_handle

def get_window_tree_at_coordinates(x, y):
    '''Returns a list of window handles at (x,y) screen coordinates.
    
    The first handle in the list is the parent window. The other handles 
    are the first child window of the previous handle that contains the 
    coordinates. These are iteratively found with 
    get_child_window_at_coordinates(x, y). 
    '''
    window_handles = []
    parent_handle = get_parent_window_at_coordinates(x, y)
    while parent_handle != None:
        window_handles.append(parent_handle)
        child_handle = win32gui.ChildWindowFromPoint(parent_handle, (x,y))
        if child_handle == 0: child_handle = None
        if child_handle == parent_handle: parent_handle = None
        else: parent_handle = child_handle
    return window_handles

#################### Keyboard Related Functions ####################

def press_key(virtual_key_code):
    '''Presses the given virtual key down and up with no delay.
    
    Use type_string() instead if possible.
    '''
    win32api.keybd_event(virtual_key_code, 0, 0, 0)
    win32api.keybd_event(virtual_key_code, 0, 2, 0)

def press_return():
    '''Presses the return key. Same as press_enter().'''
    press_key(win32con.VK_RETURN)

def press_enter():
    '''Presses the enter key. Same as press_return().'''
    press_return()

def press_enter_safely(foreground_window_handle):
    '''Verifies the foreground window before pressing the enter key.'''
    if (foreground_window_handle != 0) and window_is_foreground(foreground_window_handle):
        press_return()
        return True
    return False

def press_printscreen():
    '''Presses the print screen key.'''
    press_key(win32con.VK_SNAPSHOT)

def type_string(text):
    '''Virtually types the text string.
    
    Attempts typing capital letters and special character keys.
    '''
    def press_shift_something(virtual_key_code):
        '''Holds shift while pressing a virtual key, for typing capital letters.'''
        win32api.keybd_event(win32con.VK_SHIFT, 0, 0, 0)
        time.sleep(MULI_KEYPRESS_SLEEP_DELAY)
        press_key(virtual_key_code)
        time.sleep(MULI_KEYPRESS_SLEEP_DELAY)
        win32api.keybd_event(win32con.VK_SHIFT, 0, 2, 0)
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
        elif char == '\n':
            press_return()
        else:
            print("Can't type '"+char+"' key.", ord(char))
            return False
        time.sleep(MULI_KEYPRESS_SLEEP_DELAY)
    return True

def type_string_safely(text, foreground_window_handle):
    '''Verifies the foreground window before every virtual keypress.'''
    if (foreground_window_handle == 0): return False
    for char in text:
        if window_is_foreground(foreground_window_handle): type_string(char)
        else: return False
    return True

def capslock_on():
    '''Determines if caps lock is pressed, returning True or False.'''
    capsLockOn = win32api.GetKeyState(win32con.VK_CAPITAL)
    if capsLockOn:
        return True
    else:
        return False

#################### Cursor Related Functions ####################

def get_cursor():
    '''Retrieves the cursor screen coordinates (x,y).'''
    return win32api.GetCursorPos()

def set_cursor(x, y):
    '''Moves the cursor to screen coordinates (x,y).'''
    return win32api.SetCursorPos((x, y))

def left_click():
    '''Simulates a mouse left-click with no delay.'''
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0)
