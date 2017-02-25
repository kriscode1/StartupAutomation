'''Displays information of the window underneath the mouse cursor.

Author: Kristofer Christakos

Parent window displayed first, then the first child window if possible, 
and further grandchildren. Does not display all child windows, only the first
returned by get_child_window_at_coordinates(x, y), written using Microsoft's 
function ChildWindowFromPoint.
'''

import StartupAutomation as sa
import time

while True:
    (cursor_x, cursor_y) = sa.get_cursor()
    window_handles = sa.get_window_tree_at_coordinates(cursor_x, cursor_y)
    print("Window Handles:", window_handles)
    for window_handle in window_handles:
        print(window_handle, ' "', sa.get_window_name(window_handle), '"', sep='')
        (left, top, right, bottom) = sa.get_window_coords(window_handle)
        print("Coords: ", (left, top, right, bottom), "  ", right-left, "x", bottom-top)
    print("")
    time.sleep(1)
