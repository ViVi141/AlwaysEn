import tkinter as tk
from tkinter import messagebox
import psutil
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
APP_VERSION = "2.1"
APP_BUILD_DATE = "2025-08-22"

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
            self.master.minsize(600, 480)
        except Exception:
            pass

        self.input_method_manager = InputMethodManager()
        self.target_pid = None
        self.target_window = None
        self.target_root_hwnd = None
        self.target_process_name = None
        self.target_process_path = None
        self.monitor_mode = "window"  # "window"、"process" 或 "path"
        self.is_running = False
        self.lock = threading.Lock()  # 线程锁
        self.last_lang_tag = ""
        self.window_items = []
        self.process_items = []
        self.var_topmost = tk.BooleanVar(value=False)
        self.var_debug = tk.BooleanVar(value=False)
        self._last_debug_log_times = {}
        self._debug_min_interval_seconds = 2.0

        self._init_ui()
        self.populate_window_list()
        self.populate_process_list()
        self.switch_monitor_mode()  # 初始化UI状态

    def _init_ui(self):
        """初始化用户界面控件"""
        # 监控模式选择
        mode_frame = tk.Frame(self.master)
        mode_frame.pack(fill="x", padx=10, pady=(10, 5))
        tk.Label(mode_frame, text="监控模式:").pack(side="left")
        self.var_monitor_mode = tk.StringVar(value="window")
        tk.Radiobutton(mode_frame, text="窗口监控", variable=self.var_monitor_mode, value="window", 
                      command=self.switch_monitor_mode).pack(side="left", padx=(10, 5))
        tk.Radiobutton(mode_frame, text="程序监控", variable=self.var_monitor_mode, value="process", 
                      command=self.switch_monitor_mode).pack(side="left", padx=5)
        tk.Radiobutton(mode_frame, text="路径监控", variable=self.var_monitor_mode, value="path", 
                      command=self.switch_monitor_mode).pack(side="left", padx=5)

        # 窗口选择组
        self.window_group = tk.LabelFrame(self.master, text="窗口选择", padx=8, pady=8)
        self.window_group.pack(fill="both", expand=True, padx=10, pady=5)

        self.selected_window_label = tk.Label(self.window_group, text="未选择窗口")
        self.selected_window_label.pack(anchor="w", pady=(0, 6))

        list_frame = tk.Frame(self.window_group)
        list_frame.pack(fill="both", expand=True)

        self.window_listbox = tk.Listbox(list_frame, width=60, height=8)
        self.window_listbox.pack(side="left", fill="both", expand=True)
        scrollbar = tk.Scrollbar(list_frame, orient="vertical", command=self.window_listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self.window_listbox.config(yscrollcommand=scrollbar.set)
        self.window_listbox.bind("<Double-1>", lambda e: self.select_window())
        self.window_listbox.bind("<Return>", lambda e: self.select_window())

        # 程序选择组
        self.process_group = tk.LabelFrame(self.master, text="程序选择", padx=8, pady=8)
        self.process_group.pack(fill="both", expand=True, padx=10, pady=5)

        self.selected_process_label = tk.Label(self.process_group, text="未选择程序")
        self.selected_process_label.pack(anchor="w", pady=(0, 6))

        process_list_frame = tk.Frame(self.process_group)
        process_list_frame.pack(fill="both", expand=True)

        self.process_listbox = tk.Listbox(process_list_frame, width=60, height=8)
        self.process_listbox.pack(side="left", fill="both", expand=True)
        process_scrollbar = tk.Scrollbar(process_list_frame, orient="vertical", command=self.process_listbox.yview)
        process_scrollbar.pack(side="right", fill="y")
        self.process_listbox.config(yscrollcommand=process_scrollbar.set)
        self.process_listbox.bind("<Double-1>", lambda e: self.select_process())
        self.process_listbox.bind("<Return>", lambda e: self.select_process())

        # 路径监控组
        self.path_group = tk.LabelFrame(self.master, text="路径监控", padx=8, pady=8)
        self.path_group.pack(fill="both", expand=True, padx=10, pady=5)

        self.selected_path_label = tk.Label(self.path_group, text="未选择程序路径")
        self.selected_path_label.pack(anchor="w", pady=(0, 6))

        path_input_frame = tk.Frame(self.path_group)
        path_input_frame.pack(fill="x", pady=(0, 6))
        
        tk.Label(path_input_frame, text="程序路径:").pack(side="left")
        self.path_entry = tk.Entry(path_input_frame, width=50)
        self.path_entry.pack(side="left", fill="x", expand=True, padx=(5, 5))
        tk.Button(path_input_frame, text="浏览", command=self.browse_path, width=8).pack(side="right")

        # 路径验证和状态显示
        self.path_status_label = tk.Label(self.path_group, text="", fg="gray")
        self.path_status_label.pack(anchor="w", pady=(0, 6))

        # 按钮框架
        btn_frame = tk.Frame(self.master)
        btn_frame.pack(fill="x", padx=10, pady=(0, 8))
        tk.Button(btn_frame, text="选择窗口", command=self.select_window, width=12).pack(side="left", padx=4)
        tk.Button(btn_frame, text="选择程序", command=self.select_process, width=12).pack(side="left", padx=4)
        tk.Button(btn_frame, text="验证路径", command=self.validate_path, width=12).pack(side="left", padx=4)
        self.btn_start = tk.Button(btn_frame, text="开始监控", command=self.start_monitoring, width=12, state="disabled")
        self.btn_start.config(state="disabled")
        self.btn_start.pack(side="left", padx=4)
        self.btn_stop = tk.Button(btn_frame, text="取消监控", command=self.stop_monitoring, width=12, state="disabled")
        self.btn_stop.pack(side="left", padx=4)
        tk.Button(btn_frame, text="刷新列表", command=self.refresh_lists, width=12).pack(side="left", padx=4)
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

    def switch_monitor_mode(self):
        """切换监控模式"""
        self.monitor_mode = self.var_monitor_mode.get()
        
        # 隐藏所有组
        self.window_group.pack_forget()
        self.process_group.pack_forget()
        self.path_group.pack_forget()
        
        if self.monitor_mode == "window":
            self.window_group.pack(fill="both", expand=True, padx=10, pady=5)
            self.selected_window_label.config(text="未选择窗口")
            self.selected_process_label.config(text="未选择程序")
            self.selected_path_label.config(text="未选择程序路径")
        elif self.monitor_mode == "process":
            self.process_group.pack(fill="both", expand=True, padx=10, pady=5)
            self.selected_window_label.config(text="未选择窗口")
            self.selected_process_label.config(text="未选择程序")
            self.selected_path_label.config(text="未选择程序路径")
        else:  # path mode
            self.path_group.pack(fill="both", expand=True, padx=10, pady=5)
            self.selected_window_label.config(text="未选择窗口")
            self.selected_process_label.config(text="未选择程序")
            self.selected_path_label.config(text="未选择程序路径")
        
        # 重置选择状态
        self.target_pid = None
        self.target_window = None
        self.target_root_hwnd = None
        self.target_process_name = None
        self.target_process_path = None
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="disabled")
        self.is_running = False
        self.update_status("就绪")

    def populate_process_list(self):
        """获取并显示所有进程"""
        try:
            import psutil
            processes = []
            for proc in psutil.process_iter(['pid', 'name', 'exe']):
                try:
                    info = proc.info
                    if info['name'] and info['exe']:
                        display = f"{info['name']} (PID {info['pid']})"
                        processes.append((display, info['pid'], info['name']))
                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                    continue
            
            # 按名称排序
            processes.sort(key=lambda x: x[2].lower())
            
            self.process_items = processes
            self.process_listbox.delete(0, tk.END)
            for display, _, _ in processes:
                self.process_listbox.insert(tk.END, display)
        except ImportError:
            # 如果没有psutil，使用Windows API获取进程列表
            self.populate_process_list_winapi()
        except Exception as e:
            messagebox.showerror("错误", f"获取进程列表失败: {e}")

    def populate_process_list_winapi(self):
        """使用Windows API获取进程列表（备用方案）"""
        try:
            import win32process
            import win32gui
            import win32con
            
            processes = []
            def enum_windows_callback(hwnd, processes):
                try:
                    if win32gui.IsWindowVisible(hwnd):
                        title = win32gui.GetWindowText(hwnd)
                        if title and title.strip():
                            _, pid = win32process.GetWindowThreadProcessId(hwnd)
                            try:
                                proc = psutil.Process(pid)
                                name = proc.name()
                                display = f"{name} (PID {pid})"
                                processes.append((display, pid, name))
                            except:
                                display = f"Unknown (PID {pid})"
                                processes.append((display, pid, "Unknown"))
                    return True
                except:
                    return True
            
            win32gui.EnumWindows(enum_windows_callback, processes)
            
            # 去重并按名称排序
            seen = set()
            unique_processes = []
            for display, pid, name in processes:
                if (pid, name) not in seen:
                    seen.add((pid, name))
                    unique_processes.append((display, pid, name))
            
            unique_processes.sort(key=lambda x: x[2].lower())
            self.process_items = unique_processes
            self.process_listbox.delete(0, tk.END)
            for display, _, _ in unique_processes:
                self.process_listbox.insert(tk.END, display)
                
        except Exception as e:
            messagebox.showerror("错误", f"获取进程列表失败: {e}")

    def select_process(self):
        """选择要监控的程序"""
        selected_index = self.process_listbox.curselection()
        if not selected_index:
            messagebox.showwarning("警告", "未选择有效程序")
            return
        idx = selected_index[0]
        try:
            display, pid, name = self.process_items[idx]
            self.target_pid = pid
            self.target_process_name = name
            self.target_window = None
            self.target_root_hwnd = None
            self.selected_process_label.config(text=f"已选择程序: {name} (PID {pid})")
            self.btn_start.config(state="normal")
        except Exception as e:
            messagebox.showerror("错误", f"选择程序失败: {e}")

    def refresh_lists(self):
        """刷新窗口和进程列表"""
        if self.monitor_mode == "window":
            self.populate_window_list()
        elif self.monitor_mode == "process":
            self.populate_process_list()
        # 路径监控模式不需要刷新列表

    def browse_path(self):
        """浏览选择程序路径"""
        from tkinter import filedialog
        file_path = filedialog.askopenfilename(
            title="选择要监控的程序",
            filetypes=[
                ("可执行文件", "*.exe"),
                ("所有文件", "*.*")
            ]
        )
        if file_path:
            self.path_entry.delete(0, tk.END)
            self.path_entry.insert(0, file_path)
            self.validate_path()

    def validate_path(self):
        """验证程序路径"""
        path = self.path_entry.get().strip()
        if not path:
            self.path_status_label.config(text="请输入程序路径", fg="red")
            self.btn_start.config(state="disabled")
            return False
        
        import os
        if not os.path.exists(path):
            self.path_status_label.config(text="文件不存在", fg="red")
            self.btn_start.config(state="disabled")
            return False
        
        if not os.path.isfile(path):
            self.path_status_label.config(text="不是有效的文件", fg="red")
            self.btn_start.config(state="disabled")
            return False
        
        # 获取文件名（不含扩展名）
        filename = os.path.splitext(os.path.basename(path))[0]
        self.target_process_path = path
        self.target_process_name = filename
        
        self.path_status_label.config(text=f"✓ 路径有效: {filename}", fg="green")
        self.selected_path_label.config(text=f"已选择程序: {filename}")
        self.btn_start.config(state="normal")
        return True

    def _is_window_belongs_to_path(self, hwnd):
        """检查窗口是否属于指定路径的程序"""
        try:
            if not hwnd or not self.target_process_path:
                return False
            
            # 获取窗口对应的进程ID
            pid = ctypes.c_ulong()
            ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            
            if pid.value == 0:
                return False
            
            # 通过进程ID获取进程路径
            import psutil
            try:
                proc = psutil.Process(pid.value)
                proc_exe = proc.exe()
                
                # 比较路径（忽略大小写，处理路径分隔符）
                import os
                target_path = os.path.normpath(self.target_process_path.lower())
                proc_path = os.path.normpath(proc_exe.lower())
                
                return target_path == proc_path
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                return False
                
        except Exception as e:
            self._debug_log("path_check_error", f"路径检查失败: {e}")
            return False

    def _is_window_belongs_to_process(self, hwnd, target_pid):
        """检查窗口是否属于指定进程"""
        try:
            if not hwnd:
                return False
            pid = ctypes.c_ulong()
            ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            return pid.value == target_pid
        except Exception:
            return False

    def _get_window_title(self, hwnd):
        """获取窗口标题"""
        try:
            if not hwnd:
                return "Unknown"
            length = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
            if length == 0:
                return "Untitled"
            buffer = ctypes.create_unicode_buffer(length + 1)
            ctypes.windll.user32.GetWindowTextW(hwnd, buffer, length + 1)
            return buffer.value or "Untitled"
        except Exception:
            return "Unknown"

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
        """监控选定窗口、程序或路径的输入法状态"""
        if self.monitor_mode == "window":
            self.update_status(f"监控中: 窗口 PID {self.target_pid}")
        elif self.monitor_mode == "process":
            self.update_status(f"监控中: 程序 {self.target_process_name} (PID {self.target_pid})")
        else:  # path mode
            self.update_status(f"监控中: 程序路径 {self.target_process_name}")
        
        self.is_running = True
        while self.is_running:
            try:
                self._check_input_method()
                # 实时状态显示
                try:
                    user32 = ctypes.windll.user32
                    fg_hwnd = int(user32.GetForegroundWindow() or 0)
                    fg_root = self._get_root_hwnd(fg_hwnd) if fg_hwnd else None
                    
                    if self.monitor_mode == "window":
                        is_active_target = (fg_root is not None and self.target_root_hwnd is not None and fg_root == self.target_root_hwnd)
                    elif self.monitor_mode == "process":
                        # 程序监控模式：检查前台窗口是否属于目标程序
                        is_active_target = self._is_window_belongs_to_process(fg_hwnd, self.target_pid)
                    else:  # path mode
                        # 路径监控模式：检查前台窗口是否属于指定路径的程序
                        is_active_target = self._is_window_belongs_to_path(fg_hwnd)
                        
                except Exception:
                    is_active_target = False

                if is_active_target:
                    if self.monitor_mode == "window":
                        status = f"监控中: 窗口 PID {self.target_pid} | HKL: {self.last_lang_tag or '----'} | 窗口: {self.target_window.title}"
                    elif self.monitor_mode == "process":
                        status = f"监控中: 程序 {self.target_process_name} | HKL: {self.last_lang_tag or '----'} | 当前窗口: {self._get_window_title(fg_hwnd)}"
                    else:  # path mode
                        status = f"监控中: 程序 {self.target_process_name} | HKL: {self.last_lang_tag or '----'} | 当前窗口: {self._get_window_title(fg_hwnd)}"
                else:
                    if self.monitor_mode == "window":
                        status = f"等待目标窗口激活 | 目标: {self.target_window.title}"
                    elif self.monitor_mode == "process":
                        status = f"等待程序 {self.target_process_name} 的窗口激活"
                    else:  # path mode
                        status = f"等待程序 {self.target_process_name} 启动并激活"
                self.update_status(status)
            except Exception as e:
                print(f"监控过程中发生错误: {e}")
                self.show_error_threadsafe("错误", f"监控过程中发生错误: {e}")
            finally:
                time.sleep(0.1)  # 每0.1秒检查一次，更频繁地响应改变

    def _check_input_method(self):
        """检查当前输入法，并进行切换"""
        user32 = ctypes.windll.user32
        fg_hwnd = int(user32.GetForegroundWindow() or 0)
        
        if not fg_hwnd:
            return
            
        # 根据监控模式判断是否需要处理
        should_process = False
        
        if self.monitor_mode == "window":
            # 窗口监控模式：检查是否为选定的窗口
            if self.target_window and self.target_root_hwnd:
                fg_root = self._get_root_hwnd(fg_hwnd)
                should_process = (fg_root is not None and fg_root == self.target_root_hwnd)
        elif self.monitor_mode == "process":
            # 程序监控模式：检查前台窗口是否属于目标程序
            should_process = self._is_window_belongs_to_process(fg_hwnd, self.target_pid)
        else:  # path mode
            # 路径监控模式：检查前台窗口是否属于指定路径的程序
            should_process = self._is_window_belongs_to_path(fg_hwnd)
        
        if should_process:
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
            if self.monitor_mode == "window":
                self._debug_log("skip_non_target", "当前前台窗口不属于所选顶层窗口，跳过检测。")
            elif self.monitor_mode == "process":
                self._debug_log("skip_non_target", "当前前台窗口不属于目标程序，跳过检测。")
            else:  # path mode
                self._debug_log("skip_non_target", "当前前台窗口不属于指定路径的程序，跳过检测。")

    def start_monitoring(self):
        """开始监控选中的窗口、程序或路径"""
        if not self.is_running:
            # 检查是否有有效的监控目标
            if self.monitor_mode == "window" and self.target_pid is not None:
                self.is_running = True
                self.status_label.config(text=f"正在启动窗口监控... (PID: {self.target_pid})")
            elif self.monitor_mode == "process" and self.target_pid is not None:
                self.is_running = True
                self.status_label.config(text=f"正在启动程序监控... {self.target_process_name} (PID: {self.target_pid})")
            elif self.monitor_mode == "path" and self.target_process_path is not None:
                self.is_running = True
                self.status_label.config(text=f"正在启动路径监控... {self.target_process_name}")
            else:
                if self.monitor_mode == "window":
                    messagebox.showwarning("警告", "请先选择一个窗口并开始监控。")
                elif self.monitor_mode == "process":
                    messagebox.showwarning("警告", "请先选择一个程序并开始监控。")
                else:  # path mode
                    messagebox.showwarning("警告", "请先选择并验证程序路径。")
                return
            
            threading.Thread(target=self.monitor_window, daemon=True).start()
            if hasattr(self, 'btn_start'):
                self.btn_start.config(state="disabled")
            if hasattr(self, 'btn_stop'):
                self.btn_stop.config(state="normal")
        else:
            messagebox.showwarning("警告", "监控已在运行中。")

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
