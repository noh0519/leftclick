import pyautogui
import keyboard
import time

stime = 60
try:
    while 1:
        for cnt in range(stime, 0, -1):
            print('Count : {}'.format(cnt))
            time.sleep(1)
        print('Left Click!')
        pyautogui.leftClick()
except KeyboardInterrupt:
    print('Exit Macro')


'''
mpos = [[1200, 800], [3600, 800]]
stime = 60
try:
    while 1:
        for mx, my in mpos:
            pyautogui.leftClick(x=mx,y=my)
            time.sleep(stime)
except KeyboardInterrupt:
    print('Exit MyMacro')
'''


'''
while 1:
    position = pyautogui.position()
    print(position)
    time.sleep(3)
'''