'''Continuously displays the cursor screen coordinates.

Tool to determine coordinates when writing automation scripts.

Author: Kristofer Christakos
'''
import win32api
#python -m pip install pypiwin32
import time

while True:
    pos = win32api.GetCursorPos()
    print(pos)
    time.sleep(0.5)
