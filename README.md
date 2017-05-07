# Purpose

Startup Automation is a library of Python 3 functions used for common automation and window repositioning tasks. 

For example, I like to keep a [bandwidth monitor][bm] running in the top corner of my screen whenever I turn on my computer. I wrote a Python script to run on startup that uses StartupAutomation.py to automatically open and move the window. 

[bm]: http://kriscoder.com/projects/spotbottle "Spot Bottle - command line resource monitor for spotting bottlenecks"

Also included with StartupAutomation:

* CursorPositionTool.py is a tool to give the screen coordinates at wherever the mouse cursor is placed. 
* WindowInfoTool.py is a tool to give the name, position, and size of whichever window the mouse cursor is moved over. 

# Example Startup Script

    import StartupAutomation as sa
    import time
    
    # start the desired program
    process_handle = sa.run_program("C:\\SomeProgram.exe")
    time.sleep(1)
    
    # get the window handles for the program
    window_list = sa.get_process_window_handles(process_handle)
    
    # reposition the window where I want it, at (1590, 0)
    # (I know the program only has only one window, and one window handle)
    sa.set_window_coords(window_list[0], 1590, 0)

# Build Notes

* For Microsoft Windows only
* Written in Python 3
* Must install the pypiwin32 module first:

        python -m pip install pypiwin32

# API

## FUNCTIONS

### Process Related Functions

    run_program(commandLine)
        Creates a new process with a new console window.

        Returns the handle to the process.

    get_process_image_name(process_handle)
        Attempts to retrieve the executable path of the given process.
        
        Parameter process_handle must have been created with the PROCESS_QUERY_INFORMATION or at least PROCESS_QUERY_LIMITED_INFORMATION access right. This is automatically true for handles returned by run_program(). Returns an empty string "" on failure. 
    
    get_process_id_from_window(window_handle)
        Gets the PID of the process owning the given window.
    
    get_process_image_name_from_window(window_handle)
        Attempts to retrieve the executable path of the process owning the given window.
    
### Window Related Functions
    
    get_all_window_handles()
        Returns a list of all window handles for all processes.
    
    get_process_window_handles(process_handle)
        Returns a list of all window handles for a single process.
        
        process_handle can be retrieved and saved from run_program()
    
    get_current_window_handle()
        Returns the window handle for the console window associated with this 
        process.
    
    get_window_handle(window_name)
        Returns the window handle for the window with the given title.
    
    get_window_handle_multiple_tries(window_name, seconds_to_wait, number_of_attempts)
        Tries get_window_handle() multiple times, with a time delay between attempts.
    
    get_window_name(window_handle)
        Returns the title of the window with the given window handle.
    
    get_window_coords(window_handle)
        Retrieves the coordinates of the window with the given handle.
        
        Coordinates returned in the 4-tuple: (left, top, right, bottom)
    
    set_window_coords(window_handle, left, top, right=0, bottom=0)
        Repositions a window to new coordinates.
        
        If the right and bottom coordinates are ignored, the window size is 
        preserved.
    
    set_window_size(window_handle, width, height)
        Resizes a window without repositioning.
    
    bring_window_to_foreground(window_handle)
        Sets the foreground window to the window of the given handle.
    
    get_foreground_window()
        Retrieves the window handle of the foreground window.
    
    window_is_foreground(window_handle)
        Confirms the given window is the foreground window.
    
    get_parent_window_at_coordinates(x, y)
        Returns the handle of the parent window at (x,y) screen coordinates.
    
    get_child_window_at_coordinates(x, y)
        Returns the handle of the first child window at (x,y) screen coordinates.
        
        Returns None if no child window contains the coordinates.
    
    get_window_tree_at_coordinates(x, y)
        Returns a list of window handles at (x,y) screen coordinates.
        
        The first handle in the list is the parent window. The other handles 
        are the first child window of the previous handle that contains the 
        coordinates. These are iteratively found with 
        get_child_window_at_coordinates(x, y).
    
### Keyboard Related Functions
    type_string(text)
        Virtually types the text string.
        
        Attempts typing capital letters and special character keys.
    
    type_string_safely(text, foreground_window_handle)
        Verifies the foreground window before every virtual keypress.
    
    press_printscreen()
        Presses the print screen key.
    
    press_return()
        Presses the return key. Same as press_enter().
    
    press_enter()
        Presses the enter key. Same as press_return().
    
    press_enter_safely(foreground_window_handle)
        Verifies the foreground window before pressing the enter key.
    
    press_key(virtual_key_code)
        Presses the given virtual key down and up with no delay.
        
        Use type_string() instead if possible.
    
    capslock_on()
        Determines if caps lock is pressed, returning True or False.
    
### Cursor Related Functions
    
    get_cursor()
        Retrieves the cursor screen coordinates (x,y).
    
    set_cursor(x, y)
        Moves the cursor to screen coordinates (x,y).
    
    left_click()
        Simulates a mouse left-click with no delay.

## DATA
    MULI_KEYPRESS_SLEEP_DELAY = 0.1

# Improvements

Please feel free to contribute to this library on Github: <https://github.com/kriscode1/StartupAutomation>

If you find this library useful, I'd like to hear about it too.
