import tkinter as tk
from tkinter import messagebox
import pygetwindow as gw
import time
import threading
import ctypes
import pyautogui

class InputMethodManager:
    """输入法管理类"""
    def __init__(self):
        self.lang_id_is_english = False

    def get_current_language_tag(self):
        """获取当前输入法的语言标签"""
        language_tag = ctypes.create_unicode_buffer(8)
        ctypes.windll.user32.GetKeyboardLayoutNameW(language_tag)
        return language_tag.value.strip()

    def switch_to_english(self):
        """通过模拟键盘快捷键切换到英文输入法"""
        try:
            pyautogui.hotkey('alt', 'shift')  # 模拟 Alt + Shift 切换
            print("输入法切换已请求")
            time.sleep(1)  # 等待输入法切换生效
        except Exception as e:
            print(f"输入法切换失败: {e}")

class InputSwitcherApp:
    """输入法切换器应用程序类"""
    def __init__(self, master):
        self.master = master
        self.master.title("强制保持英文输入法")

        self.input_method_manager = InputMethodManager()
        self.target_pid = None
        self.target_window = None
        self.is_running = False

        self._init_ui()
        self.populate_window_list()

    def _init_ui(self):
        """初始化用户界面控件"""
        tk.Label(self.master, text="选择要监控的窗口:").pack(pady=10)

        self.selected_window_label = tk.Label(self.master, text="未选择窗口")
        self.selected_window_label.pack(pady=5)

        self.window_listbox = tk.Listbox(self.master, width=60)
        self.window_listbox.pack(pady=5)

        tk.Button(self.master, text="选择窗口", command=self.select_window).pack(pady=5)
        tk.Button(self.master, text="开始监控", command=self.start_monitoring).pack(pady=5)
        tk.Button(self.master, text="取消监控", command=self.stop_monitoring).pack(pady=5)
        tk.Button(self.master, text="刷新窗口列表", command=self.populate_window_list).pack(pady=5)

        self.status_label = tk.Label(self.master, text="", fg="green")
        self.status_label.pack(pady=10)

    def populate_window_list(self):
        """获取并显示所有窗口"""
        try:
            windows = gw.getAllTitles()
            self.window_listbox.delete(0, tk.END)  # 清空列表框
            self.window_listbox.insert(tk.END, *windows)  # 直接插入列表
        except Exception as e:
            messagebox.showerror("错误", f"获取窗口失败: {e}")

    def select_window(self):
        """选择要监控的窗口"""
        selected_index = self.window_listbox.curselection()
        if selected_index:
            window_title = self.window_listbox.get(selected_index)
            self.target_window = gw.getWindowsWithTitle(window_title)[0]  # 获取窗口对象
            self.target_pid = self.get_pid_by_handle(self.target_window._hWnd)
            self.selected_window_label.config(text=f"已选择窗口: {self.target_window.title}")
        else:
            messagebox.showwarning("警告", "未选择有效窗口")

    def monitor_window(self):
        """监控选定窗口的输入法状态"""
        self.status_label.config(text=f"监控中: PID {self.target_pid}")
        while self.is_running:
            self._check_input_method()
            time.sleep(1)  # 每秒检查一次

    def _check_input_method(self):
        """检查当前输入法，并进行切换"""
        if self.target_window:
            active_window = gw.getActiveWindow()
            if active_window and self.target_pid == self.get_pid_by_handle(active_window._hWnd):
                current_lang_tag = self.input_method_manager.get_current_language_tag()
                print(f"当前输入法语言标签: {current_lang_tag}")  # 输出当前语言标签

                # 检查当前输入法是否为英文
                if current_lang_tag.lower() == "0409":  # 英文输入法
                    if not self.input_method_manager.lang_id_is_english:
                        self.input_method_manager.lang_id_is_english = True
                        print("检测到输入法切换为英文")
                    else:
                        print("当前已为英文输入法，无需切换。")
                else:  # 非英文输入法
                    if self.input_method_manager.lang_id_is_english:
                        self.input_method_manager.lang_id_is_english = False
                        print("检测到输入法切换为非英文，准备切换为英文")

                    # 切换为英文输入法
                    self.input_method_manager.switch_to_english()

                    # 等待输入法切换生效
                    time.sleep(1)  # 等待输入法切换生效

                    # 重新获取输入法状态
                    current_lang_tag = self.input_method_manager.get_current_language_tag()
                    print(f"切换后的输入法语言标签: {current_lang_tag}")  # 查看切换后的语言标签
                    # 更新状态
                    self.input_method_manager.lang_id_is_english = (current_lang_tag.lower() == "0409")

            else:
                print("当前活跃窗口不是选中的目标窗口，跳过检测。")

    def start_monitoring(self):
        """开始监控选中的窗口"""
        if self.target_pid is not None and not self.is_running:
            self.is_running = True
            self.status_label.config(text=f"正在启动监控... (PID: {self.target_pid})")
            threading.Thread(target=self.monitor_window, daemon=True).start()
        else:
            messagebox.showwarning("警告", "请先选择一个窗口")

    def stop_monitoring(self):
        """停止监控"""
        if self.is_running:
            self.is_running = False
            self.status_label.config(text="监控已停止")

    @staticmethod
    def get_pid_by_handle(hwnd):
        """获取窗口句柄对应的进程ID"""
        pid = ctypes.c_ulong()
        ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        return pid.value

if __name__ == "__main__":
    """主程序入口"""
    root = tk.Tk()
    app = InputSwitcherApp(root)
    root.mainloop()
