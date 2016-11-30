# StartupAutomation
Functions for common automation and window repositioning tasks.

Specialized for Windows only. Made in python 3.5.2.

FUNCTIONS

    bring_window_to_foreground(hwnd)
        Sets the foreground window to the window of the given handle.

    capslock_on()
        Determines if caps lock is pressed, returning True or False.

    get_all_window_handles()
        Returns a list of all window handles for all processes.

    get_current_window_handle()
        Returns the window handle for the console window associated with this
        process.

    get_process_window_handles(hProcess)
        Returns a list of all window handles for a single process.

        hProcess is a process handle, which can be saved from run_program()

    get_window_coords(hwnd)
        Retrieves the coordinates of the window with the given handle.

        Coordinates returned in the 4-tuple: (left, top, right, bottom)

    get_window_handle(windowName)
        Returns the window handle for the window with the given title.

    get_window_name(hwnd)
        Returns the title of the window with the given window handle.

    left_click()
        Simulates a mouse left-click with no delay.

    press_key(virtualKeyCode)
        Presses the given virtual key down and up with no delay.

    press_printscreen()
        Presses the print screen key.

    press_return()
        Virtually hits the enter key.

    press_shift_something(virtualKeyCode)
        Holds shift while pressing a virtual key, for typing capital letters.

    run_program(commandLine)
        Creates a new process with a new console window.

        Returns the handle to the process.

    set_cursor(x, y)
        Moves the cursor to screen coordinates (x,y).

    set_window_coords(hwnd, left, top, right=0, bottom=0)
        Repositions a window to new coordinates.

        If the right and bottom coordinates are ignored, the window size is
        preserved.

    set_window_size(hwnd, width, height)
        Resizes a window without repositioning.

    type_string(text)
        Virtually types the text string.

        Attempts typing capital letters and special character keys.
