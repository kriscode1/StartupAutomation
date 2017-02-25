# StartupAutomation
Functions for common automation and window repositioning tasks.

Windows only, made in python 3. 

## FUNCTIONS

### Process Related Functions

    run_program(commandLine)
        Creates a new process with a new console window.

        Returns the handle to the process.

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
    
    press_printscreen()
        Presses the print screen key.
    
    press_return()
        Presses the enter key.
    
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
