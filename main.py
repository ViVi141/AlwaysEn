import tkinter as tk
from tkinter import messagebox
import pygetwindow as gw
import time
import threading
import ctypes
import ctypes.wintypes

# 兼容定义：在某些 Python 发行版中 ctypes.wintypes 可能未定义 HKL
try:
    HKLType = ctypes.wintypes.HKL  # type: ignore[attr-defined]
except Exception:
    HKLType = ctypes.c_void_p

# 版本信息
APP_VERSION = "2.0"
APP_BUILD_DATE = "2025-08-19"

class InputMethodManager:
    """输入法管理类"""
    # 语言常量：主语言为英语（0x09），具体布局使用美国英语 HKL 字符串
    PRIMARY_LANG_EN = 0x0009
    LAYOUT_EN_US = "00000409"

    def __init__(self):
        self.lang_id_is_english = False

    def get_current_keyboard_layout(self, hwnd):
        """获取指定窗口句柄的当前输入法布局（返回16进制字符串，如"0409"）"""
        try:
            pid = ctypes.c_ulong()  # 用于存储进程ID
            tid = ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            layout_id = ctypes.windll.user32.GetKeyboardLayout(tid)
            return f"{layout_id & 0xFFFF:04X}"  # 返回低16位，确保是0409格式
        except Exception as e:
            print(f"获取输入法布局失败: {e}")
            return None

    @staticmethod
    def is_english_langid(langid_value):
        """根据 LANGID 判断是否为英语（检测主语言位是否为 0x09）"""
        primary_lang = langid_value & 0x03FF  # 低10位为主语言
        return primary_lang == InputMethodManager.PRIMARY_LANG_EN


    def force_english_for_hwnd(self, hwnd):
        """更底层：对指定窗口强制请求切换到英文 HKL（优先 WM_INPUTLANGCHANGEREQUEST，失败再 AttachThreadInput+ActivateKeyboardLayout）"""
        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32

        WM_INPUTLANGCHANGEREQUEST = 0x0050
        SMTO_ABORTIFHUNG = 0x0002
        KLF_ACTIVATE = 0x00000001
        KLF_SETFORPROCESS = 0x00000100

        # 1) 确保已加载英文布局，获得 HKL
        hkl = user32.LoadKeyboardLayoutW(self.LAYOUT_EN_US, 0)
        if hkl == 0:
            print("LoadKeyboardLayoutW 失败")
            return

        # 2) 首选：向窗口发送 WM_INPUTLANGCHANGEREQUEST
        try:
            result = ctypes.c_ulong()
            send_ok = user32.SendMessageTimeoutW(
                ctypes.wintypes.HWND(hwnd),
                ctypes.wintypes.UINT(WM_INPUTLANGCHANGEREQUEST),
                ctypes.wintypes.WPARAM(0),
                ctypes.wintypes.LPARAM(hkl),
                ctypes.wintypes.UINT(SMTO_ABORTIFHUNG),
                ctypes.wintypes.UINT(200),
                ctypes.byref(result)
            )
        except Exception as e:
            send_ok = 0
            print(f"SendMessageTimeoutW 失败: {e}")

        if send_ok:
            print("已通过 WM_INPUTLANGCHANGEREQUEST 请求切换到英文")
            time.sleep(0.1)
            return

        # 3) 备选：Attach 到目标线程后激活 HKL
        try:
            pid = ctypes.c_ulong()
            target_tid = user32.GetWindowThreadProcessId(ctypes.wintypes.HWND(hwnd), ctypes.byref(pid))
            current_tid = kernel32.GetCurrentThreadId()
            attached = user32.AttachThreadInput(current_tid, target_tid, True)
            try:
                act = user32.ActivateKeyboardLayout(HKLType(hkl), KLF_ACTIVATE | KLF_SETFORPROCESS)
                if act == 0:
                    raise OSError("ActivateKeyboardLayout 返回 0")
                print("已通过 AttachThreadInput+ActivateKeyboardLayout 强制切换英文")
            finally:
                if attached:
                    user32.AttachThreadInput(current_tid, target_tid, False)
        except Exception as e:
            print(f"AttachThreadInput/ActivateKeyboardLayout 失败: {e}")

class InputSwitcherApp:
    """输入法切换器应用程序类"""
    def __init__(self, master):
        self.master = master
        self.master.title(f"强制保持英文输入法 v{APP_VERSION}")
        try:
            self.master.minsize(560, 420)
        except Exception:
            pass

        self.input_method_manager = InputMethodManager()
        self.target_pid = None
        self.target_window = None
        self.target_root_hwnd = None
        self.is_running = False
        self.lock = threading.Lock()  # 线程锁
        self.last_lang_tag = ""
        self.window_items = []
        self.var_topmost = tk.BooleanVar(value=False)
        self.var_debug = tk.BooleanVar(value=False)
        self._last_debug_log_times = {}
        self._debug_min_interval_seconds = 2.0

        self._init_ui()
        self.populate_window_list()

    def _init_ui(self):
        """初始化用户界面控件"""
        selection_group = tk.LabelFrame(self.master, text="窗口选择", padx=8, pady=8)
        selection_group.pack(fill="both", expand=True, padx=10, pady=10)

        self.selected_window_label = tk.Label(selection_group, text="未选择窗口")
        self.selected_window_label.pack(anchor="w", pady=(0, 6))

        list_frame = tk.Frame(selection_group)
        list_frame.pack(fill="both", expand=True)

        self.window_listbox = tk.Listbox(list_frame, width=60, height=12)
        self.window_listbox.pack(side="left", fill="both", expand=True)
        scrollbar = tk.Scrollbar(list_frame, orient="vertical", command=self.window_listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self.window_listbox.config(yscrollcommand=scrollbar.set)
        self.window_listbox.bind("<Double-1>", lambda e: self.select_window())
        self.window_listbox.bind("<Return>", lambda e: self.select_window())

        btn_frame = tk.Frame(self.master)
        btn_frame.pack(fill="x", padx=10, pady=(0, 8))
        tk.Button(btn_frame, text="选择窗口", command=self.select_window, width=12).pack(side="left", padx=4)
        self.btn_start = tk.Button(btn_frame, text="开始监控", command=self.start_monitoring, width=12, state="disabled")
        self.btn_start.pack(side="left", padx=4)
        self.btn_stop = tk.Button(btn_frame, text="取消监控", command=self.stop_monitoring, width=12, state="disabled")
        self.btn_stop.pack(side="left", padx=4)
        tk.Button(btn_frame, text="刷新窗口列表", command=self.populate_window_list, width=14).pack(side="left", padx=4)
        tk.Checkbutton(btn_frame, text="窗口置顶", variable=self.var_topmost, command=self.toggle_topmost).pack(side="right", padx=4)
        tk.Checkbutton(btn_frame, text="调试日志", variable=self.var_debug).pack(side="right", padx=4)

        self.status_label = tk.Label(self.master, text="就绪", fg="green", anchor="w", bd=1, relief="sunken")
        self.status_label.pack(fill="x", padx=10, pady=(0, 10))

    @staticmethod
    def _get_root_hwnd(hwnd):
        """获取窗口根句柄（GA_ROOT）"""
        try:
            user32 = ctypes.windll.user32
            GA_ROOT = 2
            root = user32.GetAncestor(ctypes.wintypes.HWND(hwnd), ctypes.wintypes.UINT(GA_ROOT))
            return int(root)
        except Exception:
            return int(hwnd)

    def _post_ui(self, func, *args, **kwargs):
        """将UI更新调度到主线程执行"""
        try:
            self.master.after(0, lambda: func(*args, **kwargs))
        except Exception as e:
            print(f"UI 更新调度失败: {e}")

    def update_status(self, text):
        self._post_ui(self.status_label.config, text=text)

    def show_error_threadsafe(self, title, message):
        self._post_ui(messagebox.showerror, title, message)

    def set_status_style_by_lang(self, is_english):
        color = "green" if is_english else "red"
        self._post_ui(self.status_label.config, fg=color)

    def toggle_topmost(self):
        try:
            self.master.wm_attributes('-topmost', bool(self.var_topmost.get()))
        except Exception as e:
            print(f"设置窗口置顶失败: {e}")

    def _debug_log(self, key: str, message: str):
        """受控调试日志：需开启开关，且同 key 日志至少间隔 N 秒输出一次"""
        try:
            if not self.var_debug.get():
                return
            now = time.time()
            last = self._last_debug_log_times.get(key, 0.0)
            if now - last >= self._debug_min_interval_seconds:
                print(message)
                self._last_debug_log_times[key] = now
        except Exception:
            pass

    def populate_window_list(self):
        """获取并显示所有窗口"""
        try:
            try:
                all_windows = gw.getAllWindows()
            except Exception:
                all_windows = []

            items = []
            for w in all_windows:
                try:
                    title = getattr(w, 'title', '')
                    if not title or not title.strip():
                        continue
                    pid = self.get_pid_by_handle(w._hWnd)
                    display = f"{title} (PID {pid})"
                    items.append((display, w))
                except Exception:
                    continue

            seen = set()
            unique_items = []
            for display, w in items:
                if display in seen:
                    continue
                seen.add(display)
                unique_items.append((display, w))

            self.window_items = [w for (display, w) in unique_items]
            self.window_listbox.delete(0, tk.END)
            for display, _ in unique_items:
                self.window_listbox.insert(tk.END, display)
        except Exception as e:
            messagebox.showerror("错误", f"获取窗口失败: {e}")

    def select_window(self):
        """选择要监控的窗口"""
        selected_index = self.window_listbox.curselection()
        if not selected_index:
            messagebox.showwarning("警告", "未选择有效窗口")
            return
        idx = selected_index[0]
        try:
            self.target_window = self.window_items[idx]
            self.target_pid = self.get_pid_by_handle(self.target_window._hWnd)
            self.target_root_hwnd = self._get_root_hwnd(self.target_window._hWnd)
            self.selected_window_label.config(text=f"已选择窗口: {self.target_window.title} (PID {self.target_pid})")
            if hasattr(self, 'btn_start'):
                self.btn_start.config(state="normal")
        except Exception as e:
            messagebox.showerror("错误", f"选择窗口失败: {e}")

    def monitor_window(self):
        """监控选定窗口的输入法状态"""
        self.update_status(f"监控中: PID {self.target_pid}")
        self.is_running = True
        while self.is_running:
            try:
                self._check_input_method()
                # 实时状态显示
                try:
                    user32 = ctypes.windll.user32
                    fg_hwnd = int(user32.GetForegroundWindow() or 0)
                    fg_root = self._get_root_hwnd(fg_hwnd) if fg_hwnd else None
                    is_active_target = (fg_root is not None and self.target_root_hwnd is not None and fg_root == self.target_root_hwnd)
                except Exception:
                    is_active_target = False

                if is_active_target:
                    status = f"监控中: PID {self.target_pid} | HKL: {self.last_lang_tag or '----'} | 窗口: {self.target_window.title}"
                else:
                    status = f"等待目标窗口激活 | 目标: {self.target_window.title}"
                self.update_status(status)
            except Exception as e:
                print(f"监控过程中发生错误: {e}")
                self.show_error_threadsafe("错误", f"监控过程中发生错误: {e}")
            finally:
                time.sleep(0.1)  # 每0.1秒检查一次，更频繁地响应改变

    def _check_input_method(self):
        """检查当前输入法，并进行切换"""
        if self.target_window:
            user32 = ctypes.windll.user32
            fg_hwnd = int(user32.GetForegroundWindow() or 0)
            if fg_hwnd:
                fg_root = self._get_root_hwnd(fg_hwnd)
            else:
                fg_root = None
            if fg_root is not None and self.target_root_hwnd is not None and fg_root == self.target_root_hwnd:
                # 针对当前前台窗口线程读取 HKL（更符合焦点线程的输入法状态）
                current_lang_tag = self.input_method_manager.get_current_keyboard_layout(fg_hwnd)
                if current_lang_tag is None:
                    return
                self._debug_log("hkl", f"当前输入法语言标签: {current_lang_tag}")
                self.last_lang_tag = current_lang_tag

                with self.lock:  # 防止竞争条件
                    # 将字符串（如"0409"）解析为整数 LANGID 并判断是否英语（支持英美/英英等）
                    try:
                        langid_value = int(current_lang_tag, 16)
                    except ValueError:
                        langid_value = 0

                    is_english = self.input_method_manager.is_english_langid(langid_value)
                    if is_english:
                        if not self.input_method_manager.lang_id_is_english:
                            self.input_method_manager.lang_id_is_english = True
                            print("检测到输入法切换为英文")
                        try:
                            self.set_status_style_by_lang(True)
                        except Exception:
                            pass
                    else:
                        if self.input_method_manager.lang_id_is_english:
                            self.input_method_manager.lang_id_is_english = False
                            print("检测到输入法切换为非英文，准备切换为英文")
                        # 切换为英文输入法（更底层针对当前前台窗口）
                        try:
                            self.input_method_manager.force_english_for_hwnd(fg_hwnd)
                        except Exception as e:
                            print(f"force_english_for_hwnd 失败: {e}")
                        try:
                            self.set_status_style_by_lang(False)
                        except Exception:
                            pass
            else:
                self._debug_log("skip_non_target", "当前前台窗口不属于所选顶层窗口，跳过检测。")

    def start_monitoring(self):
        """开始监控选中的窗口"""
        if self.target_pid is not None and not self.is_running:
            self.is_running = True
            self.status_label.config(text=f"正在启动监控... (PID: {self.target_pid})")
            threading.Thread(target=self.monitor_window, daemon=True).start()
            if hasattr(self, 'btn_start'):
                self.btn_start.config(state="disabled")
            if hasattr(self, 'btn_stop'):
                self.btn_stop.config(state="normal")
        else:
            messagebox.showwarning("警告", "请先选择一个窗口并开始监控。")

    def stop_monitoring(self):
        """停止监控"""
        if self.is_running:
            self.is_running = False
            self.status_label.config(text="监控已停止")
            if hasattr(self, 'btn_start'):
                self.btn_start.config(state="normal")
            if hasattr(self, 'btn_stop'):
                self.btn_stop.config(state="disabled")
        else:
            messagebox.showwarning("警告", "当前没有正在监控的窗口。")

    @staticmethod
    def get_pid_by_handle(hwnd):
        """获取窗口句柄对应的进程ID"""
        try:
            pid = ctypes.c_ulong()
            ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            return pid.value
        except Exception as e:
            print(f"获取进程ID失败: {e}")
            messagebox.showerror("错误", f"获取进程ID失败: {e}")
            return None

if __name__ == "__main__":
    """主程序入口"""
    root = tk.Tk()
    app = InputSwitcherApp(root)
    root.mainloop()
