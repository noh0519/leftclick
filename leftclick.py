import pyautogui
import keyboard
import time
import win32gui
import win32con
import json
import os
import atexit
import sys
import ctypes

# 설정 파일 경로 (exe 파일과 같은 디렉토리에 저장)
if getattr(sys, 'frozen', False):
    # exe 파일로 실행될 때
    CONFIG_FILE = os.path.join(os.path.dirname(sys.executable), "leftclick.json")
else:
    # 개발 환경에서 실행될 때
    CONFIG_FILE = "leftclick.json"

def get_console_window():
    """콘솔 창 핸들을 안전하게 가져옵니다."""
    try:
        # 먼저 kernel32.dll에서 GetConsoleWindow 함수 직접 호출 시도
        kernel32 = ctypes.windll.kernel32
        hwnd = kernel32.GetConsoleWindow()
        # 유효한 핸들인지 확인
        if hwnd != 0:
            try:
                # 창이 실제로 존재하는지 확인
                rect = win32gui.GetWindowRect(hwnd)
                # 유효한 크기인지 확인
                if rect[2] > rect[0] and rect[3] > rect[1]:
                    return hwnd
            except:
                pass
        
        # 콘솔 창을 찾지 못하면 현재 활성 창 사용
        hwnd = win32gui.GetForegroundWindow()
        if hwnd != 0:
            try:
                rect = win32gui.GetWindowRect(hwnd)
                if rect[2] > rect[0] and rect[3] > rect[1]:
                    return hwnd
            except:
                pass
        
        # 마지막 대안으로 창 제목으로 찾기
        def enum_windows_callback(hwnd, windows):
            if win32gui.IsWindowVisible(hwnd):
                window_title = win32gui.GetWindowText(hwnd)
                if any(keyword in window_title.lower() for keyword in ['python', 'leftclick', 'cmd', 'powershell', 'terminal']):
                    windows.append(hwnd)
            return True
        
        windows = []
        win32gui.EnumWindows(enum_windows_callback, windows)
        if windows:
            return windows[0]
        
        return 0
    except Exception as e:
        print(f"Error getting console window: {e}")
        return 0

def save_window_position():
    """현재 콘솔 창의 위치와 크기를 저장합니다."""
    try:
        # 콘솔 창 핸들 가져오기
        hwnd = get_console_window()
        
        if hwnd != 0:
            # 창 위치와 크기 가져오기
            rect = win32gui.GetWindowRect(hwnd)
            
            # 유효한 창 크기인지 확인
            if rect[2] > rect[0] and rect[3] > rect[1]:
                config = {
                    "x": rect[0],
                    "y": rect[1],
                    "width": rect[2] - rect[0],
                    "height": rect[3] - rect[1]
                }
                
                # JSON 파일로 저장
                with open(CONFIG_FILE, 'w') as f:
                    json.dump(config, f)
                print(f"Window position saved: {config}")
            else:
                print(f"Invalid window size detected: {rect}")
        else:
            print("Could not get valid window handle")
    except Exception as e:
        print(f"Failed to save window position: {e}")

def restore_window_position():
    """저장된 위치로 콘솔 창을 복원합니다."""
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r') as f:
                config = json.load(f)
            
            # 잠시 대기 후 창 위치 복원 (exe 파일 실행 시 안정성 향상)
            time.sleep(0.5)
            
            # 현재 콘솔 창 핸들 가져오기
            hwnd = get_console_window()
            
            if hwnd != 0:
                # 창 위치와 크기 복원
                win32gui.SetWindowPos(
                    hwnd,
                    win32con.HWND_TOP,
                    config["x"],
                    config["y"],
                    config["width"],
                    config["height"],
                    win32con.SWP_SHOWWINDOW
                )
                print(f"Window position restored: {config}")
            else:
                print("Could not get valid window handle for restoration")
        else:
            print("No saved window configuration found")
    except Exception as e:
        print(f"Failed to restore window position: {e}")

# 프로그램 종료 시 창 위치 저장 (정상 종료용)
# Ctrl + C 입력시에도 KeyboardInterrupt 예외 처리 후 정상 종료 하므로 해당 코드 실행됨
atexit.register(save_window_position)

# 프로그램 시작 시 창 위치 복원
restore_window_position()

stime = 60
try:
    while 1:
        for cnt in range(stime, 0, -1):
            print('Count : {}'.format(cnt))
            time.sleep(1)
        print('Left Click!')
        pyautogui.leftClick()
except KeyboardInterrupt:
    print('Interrupted by user')

# input('Press Enter to exit...')
