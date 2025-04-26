import tkinter as tk
from tkinter import messagebox, scrolledtext
import threading
import time
import sys # 用于退出程序

# --- 核心功能库 (需要安装: pip install pywin32 pynput pyautogui) ---
try:
    import win32api
    import win32gui
    import win32con
    import win32process
    from pynput import keyboard
    import pyautogui
    import pprint
except ImportError as e:
    print(f"错误：缺少必要的库: {e}")
    print("请先运行: pip install pywin32 pynput pyautogui")
    sys.exit(1) # 退出程序

# --- 配置 ---
# 获取虚拟屏幕的尺寸 (包含所有显示器)
VIRTUAL_SCREEN_LEFT = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)
VIRTUAL_SCREEN_TOP = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN)
VIRTUAL_SCREEN_WIDTH = win32api.GetSystemMetrics(win32con.SM_CXVIRTUALSCREEN)
VIRTUAL_SCREEN_HEIGHT = win32api.GetSystemMetrics(win32con.SM_CYVIRTUALSCREEN)

# 布局定义: {布局名称: [(x, y, width, height), ...]}
# 使用虚拟屏幕绝对坐标
LAYOUTS = {
    # 示例：假设主屏在左，副屏在右，都是 1920x1080
    # 需要根据你的实际屏幕布局调整坐标！
    "主屏左半+副屏右半": [
        (VIRTUAL_SCREEN_LEFT, VIRTUAL_SCREEN_TOP, 1920 // 2, 1080),               # 主屏左半
        (VIRTUAL_SCREEN_LEFT + 1920 + (1920 // 2), VIRTUAL_SCREEN_TOP, 1920 // 2, 1080) # 假设副屏在主屏右侧
    ],
     "虚拟屏幕 2x2 网格": [
        (VIRTUAL_SCREEN_LEFT, VIRTUAL_SCREEN_TOP, VIRTUAL_SCREEN_WIDTH // 2, VIRTUAL_SCREEN_HEIGHT // 2), # Top-Left
        (VIRTUAL_SCREEN_LEFT + VIRTUAL_SCREEN_WIDTH // 2, VIRTUAL_SCREEN_TOP, VIRTUAL_SCREEN_WIDTH // 2, VIRTUAL_SCREEN_HEIGHT // 2), # Top-Right
        (VIRTUAL_SCREEN_LEFT, VIRTUAL_SCREEN_TOP + VIRTUAL_SCREEN_HEIGHT // 2, VIRTUAL_SCREEN_WIDTH // 2, VIRTUAL_SCREEN_HEIGHT // 2), # Bottom-Left
        (VIRTUAL_SCREEN_LEFT + VIRTUAL_SCREEN_WIDTH // 2, VIRTUAL_SCREEN_TOP + VIRTUAL_SCREEN_HEIGHT // 2, VIRTUAL_SCREEN_WIDTH // 2, VIRTUAL_SCREEN_HEIGHT // 2) # Bottom-Right
    ],
    "仅主屏三列": [ # 假设主屏分辨率 1920x1080 且位于 (0,0)
        (0, 0, 1920 // 3, 1080),
        (1920 // 3, 0, 1920 // 3, 1080),
        (1920 // 3 * 2, 0, 1920 // 3, 1080)
    ]
    # 添加更多自定义布局...
}

# 鼠标位置热键: {'<modifier>+<key>': (x, y)}
# 使用虚拟屏幕绝对坐标
MOUSE_HOTKEYS = {
    '<ctrl>+<alt>+1': (VIRTUAL_SCREEN_LEFT + 100, VIRTUAL_SCREEN_TOP + 100), # 主屏左上角附近
    '<ctrl>+<alt>+2': (VIRTUAL_SCREEN_LEFT + VIRTUAL_SCREEN_WIDTH - 100, VIRTUAL_SCREEN_TOP + 100), # 虚拟屏幕右上角附近
    '<ctrl>+<alt>+q': (VIRTUAL_SCREEN_LEFT + VIRTUAL_SCREEN_WIDTH // 2, VIRTUAL_SCREEN_TOP + VIRTUAL_SCREEN_HEIGHT // 2) # 虚拟屏幕中心
}

# 布局应用热键: {'<modifier>+<key>': '布局名称'}
LAYOUT_HOTKEYS = {
    '<ctrl>+<alt>+l': "仅主屏三列",
    '<ctrl>+<alt>+g': "虚拟屏幕 2x2 网格",
    '<ctrl>+<alt>+p': "主屏左半+副屏右半" # 示例多屏热键
}

# --- 窗口过滤配置 ---
MIN_WINDOW_WIDTH = 50
MIN_WINDOW_HEIGHT = 50
BLACKLIST_CLASSES = {"progman", "workerw", "shell_traywnd", "dv2controlhost", "msctls_statusbar32", "systemtray_notifywnd"}
BLACKLIST_TITLES = {"", "program manager", "settings"} # 排除空标题和某些系统窗口
OWN_PROCESS_ID = win32api.GetCurrentProcessId() # 获取当前脚本进程ID


# --- Windows API 辅助函数 (复用之前的函数) ---
def get_screen_info():
    """获取所有显示器的信息 (用于参考，布局使用虚拟坐标)"""
    monitors = []
    monitor_count = 0
    def callback(hmonitor, hdc, rect, lparam):
        nonlocal monitor_count
        monitor_count += 1
        try:
            info = win32api.GetMonitorInfo(hmonitor)
            r = info.get('Monitor')
            w = r[2] - r[0]
            h = r[3] - r[1]
            wa = info.get('Work')
            monitors.append({
                'id': monitor_count, 'handle': hmonitor, 'device': info.get('Device', 'N/A'),
                'is_primary': info.get('Flags') == win32con.MONITORINFOF_PRIMARY,
                'resolution': {'width': w, 'height': h},
                'position': {'left': r[0], 'top': r[1], 'right': r[2], 'bottom': r[3]},
                'work_area': {'left': wa[0], 'top': wa[1], 'width': wa[2] - wa[0], 'height': wa[3] - wa[1]}
            })
        except Exception as e: pass
        return True
    try: win32api.EnumDisplayMonitors(None, None, callback)
    except Exception as e: print(f"EnumDisplayMonitors 调用失败: {e}")
    return monitors

def get_suitable_windows():
    """获取所有可见、有标题、非工具窗口、非最小化且尺寸合适的顶层窗口"""
    windows = []
    def callback(hwnd, _):
        # 检查进程ID，避免操作自己 (如果需要GUI的话)
        # try:
        #     _, pid = win32process.GetWindowThreadProcessId(hwnd)
        #     if pid == OWN_PROCESS_ID:
        #         return True # 跳过自己
        # except Exception:
        #     pass # 获取进程ID失败，继续检查其他

        if not win32gui.IsWindowVisible(hwnd): return True
        if win32gui.IsIconic(hwnd): return True

        ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        if ex_style & win32con.WS_EX_TOOLWINDOW: return True

        class_name = win32gui.GetClassName(hwnd).lower()
        if class_name in BLACKLIST_CLASSES: return True

        title = win32gui.GetWindowText(hwnd)
        if not title or title.lower() in BLACKLIST_TITLES: return True

        try:
            rect = win32gui.GetWindowRect(hwnd)
            width = rect[2] - rect[0]
            height = rect[3] - rect[1]
            if width < MIN_WINDOW_WIDTH or height < MIN_WINDOW_HEIGHT: return True
        except Exception:
             return True # 获取Rect失败则跳过

        placement = win32gui.GetWindowPlacement(hwnd)
        if placement[1] == win32con.SW_SHOWMINIMIZED: return True

        windows.append((hwnd, title))
        return True

    try: win32gui.EnumWindows(callback, None)
    except Exception as e: print(f"EnumWindows 失败: {e}")

    # 尝试按 Z 序（活动窗口优先）
    active_hwnd = win32gui.GetForegroundWindow()
    try:
        active_index = [w[0] for w in windows].index(active_hwnd)
        active_window = windows.pop(active_index)
        windows.insert(0, active_window)
    except ValueError: pass

    # print(f"识别到 {len(windows)} 个适合排布的窗口。") # 减少打印
    return windows

def activate_window(hwnd):
    """尝试激活指定句柄的窗口"""
    try:
        # 先取消最小化
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        # 尝试设置前景
        win32gui.SetForegroundWindow(hwnd)
        # print(f"激活窗口: {win32gui.GetWindowText(hwnd)} ({hwnd})")
        return True
    except Exception as e:
        print(f"激活窗口 {hwnd} ('{win32gui.GetWindowText(hwnd)}') 失败: {e}")
        return False

def move_and_resize_window(hwnd, x, y, width, height):
    """移动并调整窗口大小"""
    try:
        # 使用 SetWindowPos 通常更可靠，尤其是在改变 Z 顺序或激活状态时
        # flags = win32con.SWP_NOACTIVATE | win32con.SWP_NOZORDER # 可选 flags
        win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, x, y, width, height, 0)
        # win32gui.MoveWindow(hwnd, x, y, width, height, True)
        # print(f"移动窗口 {hwnd} 到 ({x},{y}) 尺寸 ({width}x{height})")
    except Exception as e:
        print(f"移动窗口 {hwnd} ('{win32gui.GetWindowText(hwnd)}') 失败: {e}")

def get_window_at_pos(x, y):
    """获取指定屏幕坐标下的顶层窗口句柄"""
    try:
        return win32gui.WindowFromPoint((x, y))
    except Exception as e:
        print(f"获取坐标 ({x},{y}) 处的窗口失败: {e}")
        return None

# --- 热键回调函数 (基本不变) ---
def apply_layout_action(layout_name):
    print(f"\n[{time.strftime('%H:%M:%S')}] 触发布局: {layout_name}")
    if layout_name not in LAYOUTS:
        print(f"  错误：布局 '{layout_name}' 未定义")
        return

    layout_zones = LAYOUTS[layout_name]
    suitable_windows = get_suitable_windows()

    num_zones = len(layout_zones)
    num_windows = len(suitable_windows)

    if num_windows == 0:
        print("  没有找到适合排布的窗口。")
        return

    print(f"  识别到 {num_windows} 个窗口，准备应用到 {num_zones} 个区域...")

    arranged_count = 0
    for i, zone in enumerate(layout_zones):
        if i < num_windows:
            hwnd, title = suitable_windows[i]
            print(f"    -> 应用区域 {i+1} 到窗口 '{title}' ({hwnd})")
            x, y, w, h = zone
            # 确保窗口可见且非最小化
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            time.sleep(0.05) # 短暂等待
            move_and_resize_window(hwnd, x, y, w, h)
            arranged_count += 1
        else:
            break

    print(f"  布局 '{layout_name}' 应用完成，共排列了 {arranged_count} 个窗口。")
    if num_windows > arranged_count:
        print(f"  注意: 共有 {num_windows} 个窗口被识别，但布局只有 {num_zones} 个区域。")

def mouse_hotkey_action(coords):
    x, y = coords
    print(f"\n[{time.strftime('%H:%M:%S')}] 触发鼠标定位: 移动到 ({x}, {y})")
    try:
        pyautogui.moveTo(x, y, duration=0.1)
        time.sleep(0.05) # 等待鼠标稳定
        hwnd = get_window_at_pos(x, y)
        root_hwnd = win32gui.GetAncestor(hwnd, win32con.GA_ROOT) # 获取根窗口句柄
        if root_hwnd and root_hwnd != 0:
             title = win32gui.GetWindowText(root_hwnd)
             if title: # 只激活有标题的根窗口
                 print(f"  找到窗口: '{title}' ({root_hwnd})，尝试激活...")
                 activate_window(root_hwnd)
             else:
                 print(f"  坐标 ({x},{y}) 处的窗口无标题或非预期类型。")
        else:
            print(f"  坐标 ({x},{y}) 处没有找到合适的窗口。")
    except Exception as e:
        print(f"  执行鼠标热键动作失败: {e}")

# --- 热键监听器管理 ---
listener_thread = None
hotkey_listener = None
listener_running = threading.Event() # 用于控制监听器线程状态

def start_hotkey_listener(status_callback=None):
    global hotkey_listener
    if listener_running.is_set():
        print("监听器已在运行。")
        return

    bindings = {}
    # 绑定鼠标位置热键
    for hotkey, coords in MOUSE_HOTKEYS.items():
        bindings[hotkey] = lambda c=coords: mouse_hotkey_action(c)
    # 绑定布局应用热键
    for hotkey, layout_name in LAYOUT_HOTKEYS.items():
        bindings[hotkey] = lambda ln=layout_name: apply_layout_action(ln)

    if not bindings:
        print("错误：没有定义任何热键。")
        if status_callback: status_callback("错误: 未定义热键", "red")
        return

    def run_listener():
        global hotkey_listener
        try:
            print("启动热键监听器线程...")
            hotkey_listener = keyboard.GlobalHotKeys(bindings)
            listener_running.set() # 标记为正在运行
            if status_callback: status_callback("监听器运行中...", "green")
            print("热键监听器已启动。按定义的快捷键触发动作。")
            hotkey_listener.run() # .run() 会阻塞，直到 .stop() 被调用
        except (RuntimeError, Exception) as e: # RuntimeError 可能在线程已停止时再次 stop 引发
            print(f"热键监听器出错或停止: {e}")
            listener_running.clear() # 标记为已停止
            if status_callback: status_callback(f"监听器停止: {e}", "red")
        finally:
            listener_running.clear() # 确保停止时标记清除
            print("热键监听器线程已结束。")
            if status_callback: status_callback("监听器已停止", "orange")

    global listener_thread
    # 使用 daemon=True, 但更推荐显式 join 或管理
    listener_thread = threading.Thread(target=run_listener) #, daemon=True)
    listener_thread.start()

def stop_hotkey_listener(status_callback=None):
    global hotkey_listener, listener_thread
    if not listener_running.is_set():
        print("监听器未在运行。")
        return

    print("正在尝试停止热键监听器...")
    if hotkey_listener:
        try:
            hotkey_listener.stop()
        except Exception as e:
             print(f"停止监听器时出错: {e}") # 可能已经停止
    listener_running.clear() # 标记停止

    # 等待线程结束 (可选，但推荐)
    if listener_thread and listener_thread.is_alive():
        print("等待监听线程退出...")
        listener_thread.join(timeout=2.0) # 等待最多2秒
        if listener_thread.is_alive():
             print("警告：监听线程未能按预期停止。")

    hotkey_listener = None
    listener_thread = None
    print("热键监听器已请求停止。")
    if status_callback: status_callback("监听器已停止", "orange")

# --- GUI ---
class AppGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("窗口助手 (布局 & 鼠标热键)")
        self.root.geometry("450x400")

        # 显示基本信息
        info_text = f"虚拟屏幕: {VIRTUAL_SCREEN_WIDTH}x{VIRTUAL_SCREEN_HEIGHT} at ({VIRTUAL_SCREEN_LEFT},{VIRTUAL_SCREEN_TOP})\n"
        info_text += f"当前时间: {time.strftime('%Y-%m-%d %H:%M:%S %Z')}\n"
        tk.Label(root, text=info_text, justify=tk.LEFT).pack(pady=5, anchor='w', padx=10)

        # 显示已定义的热键
        tk.Label(root, text="已定义热键:", font=('Arial', 10, 'bold')).pack(pady=5, anchor='w', padx=10)
        hotkey_frame = tk.Frame(root)
        hotkey_frame.pack(fill=tk.X, padx=10)

        st = scrolledtext.ScrolledText(hotkey_frame, wrap=tk.WORD, height=10, relief=tk.SOLID, bd=1)
        st.pack(fill=tk.BOTH, expand=True)
        st.insert(tk.INSERT, "--- 鼠标定位热键 ---\n")
        for key, val in MOUSE_HOTKEYS.items():
            st.insert(tk.INSERT, f"{key} -> 移至 {val}\n")
        st.insert(tk.INSERT, "\n--- 窗口布局热键 ---\n")
        for key, val in LAYOUT_HOTKEYS.items():
            st.insert(tk.INSERT, f"{key} -> 应用布局 '{val}'\n")
        st.configure(state='disabled') # 使其只读

        # 状态标签
        self.status_label = tk.Label(root, text="监听器未运行", fg="gray", font=('Arial', 10))
        self.status_label.pack(pady=10)

        # 控制按钮
        button_frame = tk.Frame(root)
        button_frame.pack(pady=10)
        self.start_button = tk.Button(button_frame, text="启动监听器", command=self.start_listener)
        self.start_button.pack(side=tk.LEFT, padx=10)
        self.stop_button = tk.Button(button_frame, text="停止监听器", command=self.stop_listener, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=10)

        # 退出时停止监听器
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # 尝试自动启动监听器
        self.start_listener()

    def update_status(self, message, color):
        self.status_label.config(text=message, fg=color)
        if listener_running.is_set():
            self.start_button.config(state=tk.DISABLED)
            self.stop_button.config(state=tk.NORMAL)
        else:
            self.start_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED)

    def start_listener(self):
        start_hotkey_listener(self.update_status)

    def stop_listener(self):
        stop_hotkey_listener(self.update_status)

    def on_closing(self):
        print("GUI 关闭请求...")
        self.stop_listener() # 确保监听器停止
        self.root.destroy()   # 关闭 Tkinter 窗口

# --- 主程序入口 ---
if __name__ == "__main__":
    print("--- 窗口助手启动 ---")
    print(f"虚拟屏幕尺寸: {VIRTUAL_SCREEN_WIDTH}x{VIRTUAL_SCREEN_HEIGHT} "
          f"位于 ({VIRTUAL_SCREEN_LEFT},{VIRTUAL_SCREEN_TOP})")
    print(f"当前进程 ID: {OWN_PROCESS_ID}")
    # print("\n可用屏幕信息:")
    # pprint.pprint(get_screen_info(), indent=2) # 打印屏幕信息供参考

    # 注意：操作其他窗口可能需要管理员权限
    if not win32api.GetModuleHandle("pywin32_system_tray"): # 简单检查是否以admin运行 (不完全可靠)
         is_admin = False
         try:
             # 尝试获取管理员权限相关的令牌信息
             import ctypes
             is_admin = ctypes.windll.shell32.IsUserAnAdmin() != 0
         except Exception:
             pass
         if not is_admin:
              print("\n警告：脚本似乎没有以管理员权限运行。")
              print("如果热键或窗口操作无效，请尝试以管理员身份运行。\n")


    # 创建并运行 GUI
    main_root = tk.Tk()
    app = AppGUI(main_root)
    main_root.mainloop()

    print("\n--- 窗口助手退出 ---")
    # 确保线程完全结束 (如果 on_closing 没完全停止的话)
    stop_hotkey_listener()