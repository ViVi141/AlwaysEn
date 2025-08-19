# AlwaysEn

> 版本：v2.0（2025-08-19）

## 简介

该应用程序是一个用Python编写的图形用户界面（GUI）工具，旨在强制保持在特定窗口中的英文输入法状态。它使用 `tkinter` 库进行界面构建，并利用 `pygetwindow` 获取窗口；通过底层 WinAPI（`WM_INPUTLANGCHANGEREQUEST`、`AttachThreadInput`、`ActivateKeyboardLayout`）对目标/前台窗口线程强制激活英文输入法布局（优先美国英语），提高稳定性。

## 功能

- **选择窗口**: 用户可以从可用窗口列表中选择需要监控的目标窗口（已过滤空标题并去重）。
- **监控输入法状态**: 实时监控选定窗口的输入法状态；若为非英文则自动切回英文（支持多种英文布局识别，例如美式/英式）。
- **UI 操作**: 通过简单的按钮和标签，用户可以容易地启动、停止监控和刷新窗口列表。

## 系统要求

- Python 3.x
- `tkinter`
- `pygetwindow`

## 安装

1. 请确保您已安装 Python 3.x。
2. 使用 `pip` 安装依赖库：
   ```bash
   pip install pygetwindow pywin32 pillow mouseinfo
   ```

## 使用说明

1. 运行程序：
   ```bash
   python main.py
   ```
2. 在弹出的窗口中，选择要监控的窗口。
3. 点击“开始监控”按钮开始监控该窗口的输入法状态。
4. 如需停止，可随时点击“取消监控”。
5. 若目标窗口需要更高权限，请以管理员身份运行本程序。

## 更新日志

- v2.0（2025-08-19）
  - 仅保留底层 WinAPI 切换方式：`WM_INPUTLANGCHANGEREQUEST`、`AttachThreadInput`、`ActivateKeyboardLayout`，移除热键循环切换方案。
  - 适配多进程/多窗口应用：通过根窗口句柄（`GetAncestor(GA_ROOT)`）判断是否属于同一顶层窗口，并针对前台窗口线程读取/切换输入法。
  - UI 优化：窗口分组、滚动列表、置顶开关、按钮状态、实时状态栏（显示 HKL/PID/窗口）、调试日志开关与节流。
  - 清理依赖：移除 `pyautogui`，保留 `pygetwindow`、`pywin32`、`pillow`、`mouseinfo`。

## 贡献

欢迎任何人提出建议或进行贡献！请提交问题、功能请求或直接提交拉取请求（PR）。

## 许可证
本项目采用 MIT 许可证，详情请查看 LICENSE 文件。