import tkinter as tk
from tkinter import messagebox
import threading
import time

# --- 核心功能库 (需要安装: pip install pywin32 pynput pyautogui) ---
try:
    import win32gui
    import win32api
    import win32con
    from pynput import keyboard
    import pyautogui
except ImportError as e:
    print(f"错误：缺少必要的库: {e}")
    print("请运行: pip install pywin32 pynput pyautogui")
    exit()

# --- 配置 ---
# 布局定义: {布局名称: [(x, y, width, height), ...]}
# 坐标和尺寸相对于屏幕
SCREEN_WIDTH, SCREEN_HEIGHT = pyautogui.size() # 获取屏幕分辨率

LAYOUTS = {
    "Side By Side": [
        (0, 0, SCREEN_WIDTH // 2, SCREEN_HEIGHT),                    # 左半屏
        (SCREEN_WIDTH // 2, 0, SCREEN_WIDTH // 2, SCREEN_HEIGHT)     # 右半屏
    ],
    "Three Columns": [
        (0, 0, SCREEN_WIDTH // 3, SCREEN_HEIGHT),
        (SCREEN_WIDTH // 3, 0, SCREEN_WIDTH // 3, SCREEN_HEIGHT),
        (SCREEN_WIDTH // 3 * 2, 0, SCREEN_WIDTH // 3, SCREEN_HEIGHT)
    ],
    "Top Left Corner": [ # 示例：只移动一个窗口到左上角1/4区域
        (0, 0, SCREEN_WIDTH // 2, SCREEN_HEIGHT // 2)
    ]
    # 可以添加更多自定义布局...
}

# 鼠标位置热键: {'<modifier>+<key>': (x, y)}
# 使用 pynput 的格式: https://pynput.readthedocs.io/en/latest/keyboard.html#keyboard-keys
# 注意：确保坐标 (x, y) 在你的屏幕范围内
MOUSE_HOTKEYS = {
    '<ctrl>+<alt>+1': (100, 100),      # 移动到 (100, 100) 并激活窗口
    '<ctrl>+<alt>+2': (SCREEN_WIDTH - 100, 100), # 移动到右上角附近并激活
    '<ctrl>+<alt>+q': (SCREEN_WIDTH // 2, SCREEN_HEIGHT // 2) # 移动到屏幕中心
}

# 布局应用热键: {'<modifier>+<key>': '布局名称'}
LAYOUT_HOTKEYS = {
    '<ctrl>+<alt>+l': "Side By Side",
    '<ctrl>+<alt>+k': "Three Columns"
}

# --- Windows API 辅助函数 ---

def get_visible_windows():
    """获取所有可见的顶层窗口句柄和标题"""
    windows = []
    def callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
            # 排除一些系统窗口或空标题窗口 (可以根据需要调整)
            style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
            if style & win32con.WS_VISIBLE and not (style & win32con.WS_DISABLED):
                 # 进一步过滤，例如排除程序管理器等
                class_name = win32gui.GetClassName(hwnd)
                if class_name != "Progman" and class_name != "WorkerW":
                    windows.append((hwnd, win32gui.GetWindowText(hwnd)))
        return True
    win32gui.EnumWindows(callback, None)
    # 尝试按 Z 序（用户看到的堆叠顺序）排序，最近活动的在前
    # 注意：EnumWindows 不保证顺序，这只是一个尝试
    active_hwnd = win32gui.GetForegroundWindow()
    try:
        # 将活动窗口移到列表前面（如果存在）
        active_index = [w[0] for w in windows].index(active_hwnd)
        active_window = windows.pop(active_index)
        windows.insert(0, active_window)
    except ValueError:
        pass # 活动窗口不在列表中
    return windows


def activate_window(hwnd):
    """尝试激活指定句柄的窗口"""
    try:
        # 尝试不同的方法来确保窗口被激活
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE) # 先确保窗口不是最小化
        # win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
        # win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
        win32gui.SetForegroundWindow(hwnd)
        print(f"激活窗口句柄: {hwnd}")
        return True
    except Exception as e:
        print(f"激活窗口 {hwnd} 失败: {e}")
        # 可能需要管理员权限，或者目标窗口不允许被其他程序激活
        return False

def move_and_resize_window(hwnd, x, y, width, height):
    """移动并调整窗口大小"""
    try:
        # 参数: hwnd, x, y, width, height, repaint(bool)
        win32gui.MoveWindow(hwnd, x, y, width, height, True)
        print(f"移动窗口 {hwnd} 到 ({x},{y}) 尺寸 ({width}x{height})")
    except Exception as e:
        print(f"移动窗口 {hwnd} 失败: {e}")

def get_window_at_pos(x, y):
    """获取指定屏幕坐标下的顶层窗口句柄"""
    try:
        return win32gui.WindowFromPoint((x, y))
    except Exception as e:
        print(f"获取坐标 ({x},{y}) 处的窗口失败: {e}")
        return None

# --- 热键回调函数 ---

def apply_layout_action(layout_name):
    """应用指定的窗口布局"""
    if layout_name not in LAYOUTS:
        print(f"错误：布局 '{layout_name}' 未定义")
        return

    layout_zones = LAYOUTS[layout_name]
    visible_windows = get_visible_windows()

    print(f"应用布局 '{layout_name}' 到 {min(len(visible_windows), len(layout_zones))} 个窗口...")

    # 简单策略：将前 N 个可见窗口应用到前 N 个区域
    # 更复杂的策略可能需要用户选择窗口或基于窗口标题等
    for i, zone in enumerate(layout_zones):
        if i < len(visible_windows):
            hwnd, title = visible_windows[i]
            print(f"  -> 窗口 '{title}' ({hwnd}) 应用到区域 {i+1}")
            x, y, w, h = zone
            move_and_resize_window(hwnd, x, y, w, h)
        else:
            break # 没有更多窗口可应用

def mouse_hotkey_action(coords):
    """移动鼠标到指定坐标并激活窗口"""
    x, y = coords
    print(f"触发鼠标热键，移动到 ({x}, {y})")
    try:
        pyautogui.moveTo(x, y, duration=0.1) # 移动鼠标
        time.sleep(0.05) # 短暂等待确保鼠标已到位
        hwnd = get_window_at_pos(x, y)
        if hwnd:
            print(f"找到窗口句柄: {hwnd} ({win32gui.GetWindowText(hwnd)})")
            activate_window(hwnd)
        else:
            print(f"坐标 ({x},{y}) 处没有找到窗口")
    except Exception as e:
        print(f"执行鼠标热键动作失败: {e}")


# --- 热键监听器 ---
listener_thread = None
hotkey_listener = None

def start_hotkey_listener():
    global hotkey_listener
    bindings = {}

    # 绑定鼠标位置热键
    for hotkey, coords in MOUSE_HOTKEYS.items():
        # 使用 lambda 捕获当前的 coords 值
        bindings[hotkey] = lambda c=coords: mouse_hotkey_action(c)
        print(f"绑定鼠标热键: {hotkey} -> {coords}")

    # 绑定布局应用热键
    for hotkey, layout_name in LAYOUT_HOTKEYS.items():
        bindings[hotkey] = lambda ln=layout_name: apply_layout_action(ln)
        print(f"绑定布局热键: {hotkey} -> {layout_name}")

    if not bindings:
        print("没有定义任何热键。")
        return

    try:
        # pynput.keyboard.GlobalHotKeys 在其自己的线程中运行回调
        hotkey_listener = keyboard.GlobalHotKeys(bindings)
        hotkey_listener.start()
        print("热键监听器已启动。")
        # hotkey_listener.join() # 不在这里 join，让它在后台运行
    except Exception as e:
        print(f"启动热键监听器失败: {e}")
        messagebox.showerror("错误", f"启动热键监听器失败:\n{e}\n\n可能是权限问题或热键已被占用。")


def stop_hotkey_listener():
    global hotkey_listener
    if hotkey_listener and hotkey_listener.is_alive():
        print("正在停止热键监听器...")
        hotkey_listener.stop()
        hotkey_listener = None
        print("热键监听器已停止。")

# --- GUI (非常基础的 Tkinter 示例) ---
def create_gui():
    root = tk.Tk()
    root.title("窗口布局与鼠标助手")
    root.geometry("350x250") # 设置窗口大小

    label = tk.Label(root, text="点击按钮应用布局或使用热键：")
    label.pack(pady=10)

    # 添加按钮手动应用布局
    for name in LAYOUTS:
        # 使用 lambda 捕获当前的 name 值
        btn = tk.Button(root, text=f"应用布局: {name}", command=lambda n=name: apply_layout_action(n))
        btn.pack(pady=5)

    status_label = tk.Label(root, text="热键监听器正在运行...", fg="green")
    status_label.pack(pady=10)

    # 退出时停止监听器
    def on_closing():
        if messagebox.askokcancel("退出", "确定要退出并停止热键监听吗?"):
            stop_hotkey_listener()
            root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_closing)

    # 启动热键监听线程
    global listener_thread
    listener_thread = threading.Thread(target=start_hotkey_listener, daemon=True)
    listener_thread.start()

    root.mainloop() # 阻塞直到窗口关闭

# --- 主程序入口 ---
if __name__ == "__main__":
    print(f"当前屏幕分辨率: {SCREEN_WIDTH}x{SCREEN_HEIGHT}")
    print("启动应用程序...")
    # 注意：在某些系统上，操作其他窗口可能需要管理员权限运行此脚本
    create_gui()
    print("应用程序已退出。")