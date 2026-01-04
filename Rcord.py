import tkinter as tk
from tkinter import messagebox
import sqlite3
import calendar
import re
import math
import os
import sys
import json
import winreg
from datetime import datetime, timedelta, date
import warnings

# 忽略警告
warnings.filterwarnings("ignore")

import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.toast import ToastNotification
from PIL import Image, ImageTk 

# ================= 配置常量 =================
THEME_NAME = "litera" 

# 字体配置
FONT_CLOCK = ("Segoe UI", 56, "bold")      
FONT_DATE = ("Microsoft YaHei UI", 11)     
FONT_DATA_NUM = ("Segoe UI", 20, "bold")   
FONT_NORMAL = ("Microsoft YaHei UI", 10)   
FONT_BOLD = ("Microsoft YaHei UI", 10, "bold") 
FONT_CAL_SMALL = ("Microsoft YaHei UI", 8) 

DAILY_NET_HOURS = 7.0
LUNCH_START = "12:00"
LUNCH_END = "13:00"

SIZE_SMALL = (20, 20)
SIZE_BTN   = (24, 24)
SIZE_LARGE = (100, 100)

STATUS_IDLE = 0     
STATUS_WORKING = 1  
STATUS_ABNORMAL = 9 

# ================= 资源路径工具 =================
def get_resource_path(relative_path):
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

if getattr(sys, 'frozen', False):
    EXEC_DIR = os.path.dirname(sys.executable)
else:
    EXEC_DIR = os.path.dirname(os.path.abspath(__file__))

DB_NAME = os.path.join(EXEC_DIR, "work_log_v2.db")
CONFIG_FILE = os.path.join(EXEC_DIR, "config.json")
ASSETS_DIR = get_resource_path("assets") 
ICON_FILENAME = "Record.png" 

class WorkAppPro(ttk.Window):
    def __init__(self):
        try:
            from ctypes import windll
            myappid = 'mycompany.worktimepro.gui.v1' 
            windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        except: pass
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1) 
        except:
            try: windll.user32.SetProcessDPIAware() 
            except: pass

        super().__init__(themename=THEME_NAME)
        self.withdraw()
        self.title("WorkTime Pro")
        
        self.config = self.load_config()
        self.imgs = {}
        self.load_assets()

        self.today_date = datetime.now().strftime("%Y-%m-%d")
        
        self.current_start_dt = None

        self.is_working = False 
        self.var_time = tk.StringVar()
        self.var_date = tk.StringVar()
        self.var_worked = tk.StringVar(value="0.00h")
        self.var_target = tk.StringVar(value=f"{DAILY_NET_HOURS}h")
        self.var_btn_text = tk.StringVar(value="上班打卡")
        self.var_status_text = tk.StringVar(value="早安")
        self.var_autostart = tk.BooleanVar(value=self.config.get("auto_start", False))
        
        self.init_db()
        self.setup_ui()
        self.refresh_main_data()
        self.start_clock_loop()
        
        self.after(500, self.check_first_run)
        self.center_and_show(400, 660)

    def load_assets(self):
        assets_config = {
            "sun":      ("Sun.png", SIZE_LARGE),
            "start":    ("Start.png", SIZE_LARGE),
            "working":  ("Working.png", SIZE_LARGE),
            "coffee_L": ("Coffee.png", SIZE_LARGE), 
            "beach":    ("Beach.png", SIZE_LARGE),
            "vacation_L": ("Vacation.png", SIZE_LARGE),
            "party":    ("Party.png", SIZE_LARGE),
            "sleep":    ("Sleep.png", SIZE_LARGE),
            "flash":    ("Flash.png", SIZE_BTN),
            "clock":    ("Clock.png", SIZE_BTN),
            "coffee_S": ("Coffee.png", SIZE_BTN), 
            "vacation_S": ("Vacation.png", SIZE_BTN), 
            "settings": ("Settings.png", SIZE_BTN),
            "calendar": ("Calendar.png", SIZE_BTN),
            "banana":   ("Banana.png", SIZE_BTN),
            "save":     ("Save.png", SIZE_BTN),
            "target":    ("Target.png", SIZE_SMALL),
            "stopwatch": ("Stopwatch.png", SIZE_SMALL),
            "idea":      ("Idea.png", SIZE_SMALL)
        }
        for key, (filename, target_size) in assets_config.items():
            path = os.path.join(ASSETS_DIR, filename)
            if os.path.exists(path):
                try:
                    pil_img = Image.open(path)
                    pil_img = pil_img.resize(target_size, Image.Resampling.LANCZOS)
                    self.imgs[key] = ImageTk.PhotoImage(pil_img)
                except: pass

    def format_time_str(self, time_str):
        if not time_str: return None
        t = time_str.strip().replace("：", ":")
        match = re.match(r"^(\d{1,2})[:](\d{1,2})$", t)
        if match:
            h = int(match.group(1))
            m = int(match.group(2))
            if 0 <= h <= 23 and 0 <= m <= 59:
                return f"{h:02d}:{m:02d}"
        return None

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f: return json.load(f)
            except: return {}
        return {}

    def save_config(self):
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f: json.dump(self.config, f)

    def check_first_run(self):
        if not self.config.get("has_run_before", False):
            if messagebox.askyesno("✨ 欢迎使用", "这是您第一次运行。\n是否需要设置为开机自动启动？"):
                self.var_autostart.set(True)
                self.toggle_autostart(silent=True)
            self.config["has_run_before"] = True
            self.save_config()

    def toggle_autostart(self, silent=False):
        enable = self.var_autostart.get()
        app_name = "WorkTime Pro"
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        try:
            if getattr(sys, 'frozen', False):
                run_path = f'"{sys.executable}"'
            else:
                python_exe = sys.executable.replace("python.exe", "pythonw.exe")
                script_path = os.path.abspath(__file__)
                run_path = f'"{python_exe}" "{script_path}"'
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_ALL_ACCESS)
            if enable:
                winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ, run_path)
                if not silent: ToastNotification("设置成功", "已开启开机自启", bootstyle="success").show_toast()
            else:
                try:
                    winreg.DeleteValue(key, app_name)
                    if not silent: ToastNotification("设置成功", "已关闭开机自启", bootstyle="info").show_toast()
                except: pass
            winreg.CloseKey(key)
            self.config["auto_start"] = enable
            self.save_config()
        except Exception as e:
            if not silent: messagebox.showerror("权限错误", str(e))
            self.var_autostart.set(not enable)

    def reset_database(self):
        if messagebox.askyesno("危险操作", "⚠️ 确定要清空所有数据吗？"):
            try:
                conn = sqlite3.connect(DB_NAME)
                conn.execute("DROP TABLE IF EXISTS attendance")
                conn.commit(); conn.close()
                self.init_db()
                self.refresh_main_data()
                ToastNotification("重置完成", "数据库已重建", bootstyle="success").show_toast()
            except Exception as e: messagebox.showerror("错误", str(e))

    def center_and_show(self, w, h, win=None):
        target = win if win else self
        target.update_idletasks()
        ws, hs = self.winfo_screenwidth(), self.winfo_screenheight()
        x = (ws - w) // 2
        y = (hs - h) // 2
        target.geometry(f"{w}x{h}+{x}+{y}")
        target.deiconify()

    def init_db(self):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute('''CREATE TABLE IF NOT EXISTS attendance (
                date TEXT PRIMARY KEY, punches TEXT, duration REAL, day_type INTEGER DEFAULT 0, status INTEGER DEFAULT 0 
        )''')
        try: cursor.execute("ALTER TABLE attendance ADD COLUMN punches TEXT")
        except: pass
        try: cursor.execute("ALTER TABLE attendance ADD COLUMN status INTEGER DEFAULT 0")
        except: pass
        conn.commit(); conn.close()

    def setup_ui(self):
        header = ttk.Frame(self, padding=(15, 10))
        header.pack(fill="x")
        
        # 🟢 1. 将按钮保存为 self 属性，以便后续定位菜单
        self.btn_setting = ttk.Button(header, image=self.imgs.get("settings"), bootstyle="link-dark", width=3)
        self.btn_setting.pack(side="left")
        
        # 🟢 2. 绑定新的自定义菜单函数
        self.btn_setting.configure(command=self.open_setting_menu)

        ttk.Button(header, text=" 月度记录", image=self.imgs.get("calendar"), compound="left", 
                   bootstyle="outline-primary", command=self.open_calendar_window, cursor="hand2").pack(side="right")

        card_frame = ttk.Frame(self, padding=0)
        card_frame.pack(fill="x", expand=False, padx=15, pady=0)

        time_box = ttk.Frame(card_frame)
        time_box.pack(fill="x", pady=(5, 0)) 
        ttk.Label(time_box, textvariable=self.var_time, font=FONT_CLOCK, bootstyle="dark", anchor="center").pack(fill="x")
        ttk.Label(time_box, textvariable=self.var_date, font=FONT_DATE, bootstyle="secondary", anchor="center").pack(fill="x")

        stat_box = ttk.Frame(card_frame, padding=(5, 15))
        stat_box.pack(fill="x")
        
        f_left = ttk.Frame(stat_box)
        f_left.pack(side="left", expand=True)
        ttk.Label(f_left, text=" 目标时长", image=self.imgs.get("target"), compound="left", font=FONT_NORMAL, bootstyle="dark").pack()
        ttk.Label(f_left, textvariable=self.var_target, font=FONT_DATA_NUM, bootstyle="info").pack()

        ttk.Separator(stat_box, orient="vertical").pack(side="left", fill="y", padx=10)

        f_right = ttk.Frame(stat_box)
        f_right.pack(side="left", expand=True)
        ttk.Label(f_right, text=" 当前时长", image=self.imgs.get("stopwatch"), compound="left", font=FONT_NORMAL, bootstyle="dark").pack()
        self.lbl_worked = ttk.Label(f_right, textvariable=self.var_worked, font=FONT_DATA_NUM, bootstyle="success")
        self.lbl_worked.pack()

        msg_container = ttk.Frame(self, padding=(25, 10)) 
        msg_container.pack(fill="x", pady=(0, 5))
        
        self.msg_lbl_title = ttk.Label(msg_container, text=" 当前状态 ", image=self.imgs.get("idea"), compound="left", bootstyle="primary", font=("微软雅黑", 9, "bold"))
        self.msg_box = ttk.Labelframe(msg_container, labelwidget=self.msg_lbl_title, padding=(10, 10), bootstyle="primary")
        self.msg_box.pack(fill="x")

        status_inner = ttk.Frame(self.msg_box)
        status_inner.pack(anchor="center") 
        
        self.lbl_icon = ttk.Label(status_inner, image=self.imgs.get("sun"), anchor="center")
        self.lbl_icon.grid(row=0, column=0, padx=(0, 15), sticky="e")
        
        self.lbl_text = ttk.Label(status_inner, textvariable=self.var_status_text, 
                                  font=FONT_NORMAL, width=14, 
                                  anchor="w", justify="left", bootstyle="primary")
        self.lbl_text.grid(row=0, column=1, sticky="w")

        ttk.Frame(self).pack(fill="both", expand=True)

        btn_area = ttk.Frame(self, padding=(25, 20))
        btn_area.pack(side="bottom", fill="x", pady=(0, 5)) 
        
        self.btn_mid = ttk.Button(btn_area, text=" 中途记录", image=self.imgs.get("coffee_S"), compound="left",
                                  bootstyle="outline-dark", command=self.handle_mid_punch, width=14, state="disabled")
        self.btn_mid.pack(anchor="center", pady=(0, 10))

        self.btn_main = ttk.Button(btn_area, textvariable=self.var_btn_text, 
                                   image=self.imgs.get("flash"), compound="left",
                                   command=self.handle_main_action, bootstyle="success")
        self.btn_main.pack(fill="x", ipady=12)


    # 🟢【最终版】深色悬浮菜单：灰色边框 + 紧凑布局 + 无悬停变色
    def open_setting_menu(self):
        # 1. 如果菜单已存在，先关闭
        if hasattr(self, 'menu_win') and self.menu_win.winfo_exists():
            self.menu_win.destroy()
            return

        # 2. 定义深色皮肤配色
        BG_COLOR = "#2c2c2e"       # 菜单主背景 (深灰)
        FG_COLOR = "#ffffff"       # 纯白文字
        DIVIDER_COLOR = "#48484a"  # 分割线颜色
        # 🟢 修改点：将边框改为明显的灰色
        BORDER_COLOR = "#8e8e93"   
        TOGGLE_ON_COLOR = "#34c759" # 开关开启 (绿)
        TOGGLE_OFF_COLOR = "#636366"# 开关关闭 (灰)

        # 3. 创建无边框窗口
        self.menu_win = tk.Toplevel(self)
        self.menu_win.overrideredirect(True)       # 去除标题栏
        self.menu_win.attributes('-topmost', True) # 始终置顶
        self.menu_win.configure(bg=BG_COLOR)

        # 🟢 边框容器：背景设为灰色(BORDER_COLOR)，padding=1 形成 1px 边框
        main_container = tk.Frame(self.menu_win, bg=BORDER_COLOR, padx=1, pady=1)
        main_container.pack(fill="both", expand=True)
        
        # 内容容器：背景设为深色(BG_COLOR)
        content_frame = tk.Frame(main_container, bg=BG_COLOR)
        content_frame.pack(fill="both", expand=True)

        # 内部类：自定义手绘开关
        class CanvasToggle(tk.Canvas):
            def __init__(self, parent, variable, command=None, bg=BG_COLOR):
                super().__init__(parent, width=44, height=24, bg=bg, highlightthickness=0, bd=0, cursor="hand2")
                self.var = variable
                self.cmd = command
                self.bind("<Button-1>", self.toggle)
                self.render()

            def render(self):
                self.delete("all")
                is_on = self.var.get()
                fill_color = TOGGLE_ON_COLOR if is_on else TOGGLE_OFF_COLOR
                # 绘制轨道
                self.create_oval(1, 1, 23, 23, fill=fill_color, outline=fill_color) 
                self.create_rectangle(12, 1, 32, 23, fill=fill_color, outline=fill_color)
                self.create_oval(21, 1, 43, 23, fill=fill_color, outline=fill_color)
                # 绘制滑块
                cx = 32 if is_on else 12
                self.create_oval(cx-10, 2, cx+10, 22, fill="#ffffff", outline="")

            def toggle(self, event=None):
                self.var.set(not self.var.get())
                self.render()
                if self.cmd: self.cmd()

        # --- 通用菜单行创建函数 ---
        def create_row(icon_key, text, is_toggle=False, toggle_var=None, command=None, text_color=FG_COLOR):
            # 紧凑行高
            row = tk.Frame(content_frame, bg=BG_COLOR, height=35)
            row.pack(fill="x")
            
            # 紧凑内边距
            inner = tk.Frame(row, bg=BG_COLOR, padx=10, pady=5)
            inner.pack(fill="both", expand=True)

            # 1. 图标
            if icon_key and self.imgs.get(icon_key):
                lbl_icon = tk.Label(inner, image=self.imgs.get(icon_key), bg=BG_COLOR, bd=0)
                lbl_icon.pack(side="left", padx=(0, 8))
            
            # 2. 文字
            lbl_text = tk.Label(inner, text=text, font=("Microsoft YaHei UI", 9), 
                                fg=text_color, bg=BG_COLOR, bd=0)
            lbl_text.pack(side="left")

            # 3. 右侧控件
            toggle_btn = None
            if is_toggle and toggle_var:
                toggle_btn = CanvasToggle(inner, variable=toggle_var, command=command, bg=BG_COLOR)
                toggle_btn.pack(side="right")
            
            # 点击交互
            def on_click(e):
                if is_toggle and toggle_btn:
                    toggle_btn.toggle()
                elif command:
                    command()

            lbl_text.bind("<Button-1>", on_click)
            inner.bind("<Button-1>", on_click)
            if not is_toggle:
                row.configure(cursor="hand2")

            return row

        # --- 菜单内容 ---
        create_row("flash", "开机自启", is_toggle=True, toggle_var=self.var_autostart, command=self.toggle_autostart)
        
        # 分割线
        tk.Frame(content_frame, bg=DIVIDER_COLOR, height=1).pack(fill="x", padx=10)

        # 清空数据
        def clean_action():
            self.menu_win.destroy()
            self.reset_database()
        create_row("banana", "清空数据", is_toggle=False, command=clean_action, text_color="#ff6b6b")

        # 4. 窗口定位
        self.menu_win.update_idletasks()
        width = 160
        height = main_container.winfo_reqheight()
        
        root_x = self.btn_setting.winfo_rootx()
        root_y = self.btn_setting.winfo_rooty() + self.btn_setting.winfo_height() + 5
        
        if root_x + width > self.winfo_screenwidth():
            root_x = self.winfo_screenwidth() - width - 5
            
        self.menu_win.geometry(f"{width}x{height}+{root_x}+{root_y}")

        # 5. 焦点丢失关闭
        def on_focus_out(event):
            if self.menu_win:
                self.menu_win.destroy()

        self.menu_win.bind("<FocusOut>", on_focus_out)
        self.menu_win.focus_force()

    # 继承 calculate_logic 的逻辑 (含午休扣除 + 0.5h 取整)
    def update_realtime_duration(self):
        if self.is_working and self.current_start_dt:
            now = datetime.now()
            # 1. 计算原始总秒数
            total_sec = (now - self.current_start_dt).total_seconds()
            
            # 2. 计算午休扣除时长
            l_s = datetime.strptime(f"{self.today_date} {LUNCH_START}", "%Y-%m-%d %H:%M")
            l_e = datetime.strptime(f"{self.today_date} {LUNCH_END}", "%Y-%m-%d %H:%M")
            
            overlap_start = max(self.current_start_dt, l_s)
            overlap_end = min(now, l_e)
            
            deduction_sec = 0.0
            if overlap_end > overlap_start:
                deduction_sec = (overlap_end - overlap_start).total_seconds()
            
            # 3. 计算净工时 (小时)
            raw_net_hours = max(0, (total_sec - deduction_sec) / 3600.0)
            
            # 4. 执行取整逻辑：向下取整到 0.5 (例如 1.6 -> 1.5, 1.9 -> 1.5)
            # 公式：floor(hours * 2) / 2
            display_hours = math.floor(raw_net_hours * 2) / 2.0
            
            self.var_worked.set(f"{display_hours:.1f}h")

    def start_clock_loop(self):
        if not self.winfo_exists(): return
        now = datetime.now()
        weeks = ["周一","周二","周三","周四","周五","周六","周日"]
        self.var_time.set(now.strftime("%H:%M"))
        self.var_date.set(f"{now.strftime('%Y-%m-%d')}  {weeks[now.weekday()]}")
        
        # 实时更新计算
        self.update_realtime_duration()
        
        self.after(1000, self.start_clock_loop)

    def refresh_main_data(self):
        rec = self.get_record(self.today_date)
        if not rec:
            self.set_state_idle()
            return
        status = rec['status']
        day_type = rec['type']
        punches = rec['punches'].split(',') if rec['punches'] else []

        if day_type != 0:
            self.set_state_finished(rec['duration'], day_type)
            return

        if status == STATUS_WORKING:
            if punches:
                try:
                    self.current_start_dt = datetime.strptime(f"{self.today_date} {punches[0]}", "%Y-%m-%d %H:%M")
                except:
                    self.current_start_dt = None
            
            self.set_state_working(len(punches))
            self.update_realtime_duration()
        else:
            self.current_start_dt = None
            if len(punches) > 0:
                self.set_state_finished(rec['duration'], 0)
            else:
                self.set_state_idle()

    def set_state_idle(self):
        self.is_working = False
        self.current_start_dt = None
        self.var_worked.set("0.00h")
        self.var_btn_text.set(" 上班记录")
        self.btn_main.configure(bootstyle="success", state="normal", image=self.imgs.get("flash"))
        self.btn_mid.configure(state="disabled") 
        self.lbl_worked.configure(bootstyle="secondary")
        self.lbl_icon.configure(image=self.imgs.get("start")) 
        self.var_status_text.set("新的一天\n准备出发！")
        self.msg_box.configure(bootstyle="primary")
        self.msg_lbl_title.configure(bootstyle="primary")
        self.lbl_text.configure(bootstyle="primary") 

    def set_state_working(self, count):
        self.is_working = True
        self.var_btn_text.set(f" 下班记录") 
        self.btn_main.configure(bootstyle="warning", state="normal", image=self.imgs.get("clock"))
        self.lbl_worked.configure(bootstyle="primary")
        if count >= 6: self.btn_mid.configure(state="disabled")
        else: self.btn_mid.configure(state="normal")
        self.lbl_icon.configure(image=self.imgs.get("working")) 
        self.var_status_text.set("工作中\n静待干饭")
        self.msg_box.configure(bootstyle="warning") 
        self.msg_lbl_title.configure(bootstyle="warning")
        self.lbl_text.configure(bootstyle="warning")

    def set_state_finished(self, dur, type_code):
        self.is_working = False
        self.current_start_dt = None
        self.var_worked.set(f"{dur}h")
        self.btn_mid.configure(state="disabled") 
        
        if type_code == 1: 
            self.var_btn_text.set(" 非工作日")
            self.btn_main.configure(bootstyle="info", state="normal", image=self.imgs.get("coffee_S"))
            self.lbl_icon.configure(image=self.imgs.get("coffee_L"))
            self.var_status_text.set("好好休息")
            self.msg_box.configure(bootstyle="info")
            self.msg_lbl_title.configure(bootstyle="info")
            self.lbl_text.configure(bootstyle="info")
            
        elif type_code in [2, 3]: 
            self.var_btn_text.set(" 今日休假")
            self.btn_main.configure(bootstyle="info", state="normal", image=self.imgs.get("vacation_S"))
            self.lbl_icon.configure(image=self.imgs.get("beach"))
            self.var_status_text.set("假期愉快！")
            self.msg_box.configure(bootstyle="info")
            self.msg_lbl_title.configure(bootstyle="info")
            self.lbl_text.configure(bootstyle="info")
            
        else: 
            self.var_btn_text.set(" 管理记录")
            self.btn_main.configure(bootstyle="primary", state="normal", image=self.imgs.get("settings"))
            self.lbl_icon.configure(image=self.imgs.get("party"))
            self.var_status_text.set("已下班\n享受生活吧")
            self.msg_box.configure(bootstyle="success")
            self.msg_lbl_title.configure(bootstyle="success")
            self.lbl_text.configure(bootstyle="success")

    def ask_punch_time(self, title="记录确认"):
        dialog = tk.Toplevel(self)
        dialog.withdraw()
        dialog.title(title)
        
        w, h = 280, 200
        ws, hs = self.winfo_screenwidth(), self.winfo_screenheight()
        x, y = (ws - w) // 2, (hs - h) // 2
        dialog.geometry(f"{w}x{h}+{x}+{y}")
        
        ttk.Label(dialog, text="请确认记录时间", font=FONT_BOLD).pack(pady=(20, 10))
        
        v_time = tk.StringVar(value=datetime.now().strftime("%H:%M"))
        e = ttk.Entry(dialog, textvariable=v_time, font=("Segoe UI", 20, "bold"), justify="center", width=6)
        e.pack(pady=5)
        e.focus_set()
        
        result_container = {"time": None}
        
        def on_confirm(event=None):
            raw_t = v_time.get()
            formatted_time = self.format_time_str(raw_t)
            
            if not formatted_time:
                messagebox.showerror("格式错误", "请输入正确的时间格式\n例如: 09:30 或 9:30\n支持中文冒号", parent=dialog)
                return
            
            result_container["time"] = formatted_time
            dialog.destroy()
            
        dialog.bind('<Return>', on_confirm)

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(fill="x", pady=20, padx=25)
        ttk.Button(btn_frame, text="取消", bootstyle="secondary", command=dialog.destroy).pack(side="left", expand=True)
        ttk.Button(btn_frame, text="确认", bootstyle="primary", command=on_confirm).pack(side="left", expand=True)

        dialog.transient(self) 
        dialog.grab_set() 
        dialog.deiconify()
        self.wait_window(dialog)
        return result_container["time"]

    def handle_main_action(self):
        rec = self.get_record(self.today_date)
        if rec and rec['type'] != 0:
            self.open_edit_dialog(self.today_date)
            return
        if rec and rec['status'] == STATUS_IDLE and rec['punches']:
            self.open_edit_dialog(self.today_date)
            return
        if self.is_working:
            self.perform_clock_out(rec)
        else:
            self.perform_clock_in(rec)

    def handle_mid_punch(self):
        if not self.is_working: return
        rec = self.get_record(self.today_date)
        punches = rec['punches'].split(',') if (rec and rec['punches']) else []
        user_time = self.ask_punch_time("中途记录")
        if not user_time: return
        punches.append(user_time)
        punches.sort()
        self.update_db(punches, 0.0, STATUS_WORKING)
        self.refresh_main_data()
        if len(punches) >= 5: self.btn_mid.configure(state="disabled")
        else: self.btn_mid.configure(state="normal")
        ToastNotification("记录成功", f"已添加: {user_time}", bootstyle="info").show_toast()

    def perform_clock_in(self, rec):
        user_time = self.ask_punch_time("上班记录")
        if not user_time: return
        punches = rec['punches'].split(',') if (rec and rec['punches']) else []
        punches.append(user_time)
        punches.sort()
        self.update_db(punches, 0.0, STATUS_WORKING)
        self.refresh_main_data()
        ToastNotification("上班啦", f"时间: {user_time}", bootstyle="success").show_toast()

    def perform_clock_out(self, rec):
        user_time = self.ask_punch_time("下班记录")
        if not user_time: return
        punches = rec['punches'].split(',') if (rec and rec['punches']) else []
        punches.append(user_time)
        punches.sort()
        duration = self.calculate_logic(punches)
        self.update_db(punches, duration, STATUS_IDLE)
        self.refresh_main_data()
        ToastNotification("下班啦", f"今日工时: {duration}h", bootstyle="success").show_toast()

    def update_db(self, punches, duration, status, day_type=0):
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("INSERT OR REPLACE INTO attendance (date, punches, duration, day_type, status) VALUES (?, ?, ?, ?, ?)", 
                 (self.today_date, ",".join(punches), duration, day_type, status))
        conn.commit(); conn.close()

    def calculate_logic(self, punches):
        if not punches or len(punches) < 2: return 0.0
        start_str, end_str = punches[0], punches[-1]
        try:
            fmt = "%H:%M"
            t1 = datetime.strptime(start_str, fmt)
            t2 = datetime.strptime(end_str, fmt)
            if t2 < t1: t2 += timedelta(days=1)
            raw_hours = (t2 - t1).total_seconds() / 3600.0
            l_s = datetime.strptime(LUNCH_START, fmt)
            l_e = datetime.strptime(LUNCH_END, fmt)
            overlap_start = max(t1, l_s)
            overlap_end = min(t2, l_e)
            deduction = 0.0
            if overlap_end > overlap_start: deduction = (overlap_end - overlap_start).total_seconds() / 3600.0
            net_hours = max(0, raw_hours - deduction)
            return math.floor(net_hours * 2) / 2.0 
        except: return 0.0

    def get_record(self, d):
        conn = sqlite3.connect(DB_NAME); c = conn.cursor()
        try: 
            c.execute("SELECT punches, duration, day_type, status FROM attendance WHERE date=?",(d,))
            r = c.fetchone()
            st = r[3] if (r and len(r)>3) else 0
            return {'punches':r[0], 'duration':r[1], 'type':r[2], 'status':st} if r else None
        except: return None
        finally: conn.close()

    def open_calendar_window(self):
        cal_win = ttk.Toplevel(self)
        cal_win.withdraw()
        cal_win.title("月度记录")

        # 🟢 移除了之前的 style.configure 代码，使用内置 danger 样式

        nav = ttk.Frame(cal_win, padding=10)
        nav.pack(fill="x")
        ttk.Button(nav, text="◀", command=lambda: chg(-1), bootstyle="outline-dark", width=4).pack(side="left")
        lbl_title = ttk.Label(nav, text="...", font=("Segoe UI", 12, "bold"), bootstyle="dark")
        lbl_title.pack(side="left", expand=True)
        ttk.Button(nav, text="▶", command=lambda: chg(1), bootstyle="outline-dark", width=4).pack(side="right")
        
        head = ttk.Frame(cal_win, padding=5)
        head.pack(fill="x")
        for i, t in enumerate("一二三四五六日"):
            c = "danger" if i==6 else "dark"
            ttk.Label(head, text=t, bootstyle=c, anchor="center", font=FONT_BOLD).pack(side="left", expand=True, fill="x")
            
        grid = ttk.Frame(cal_win, padding=(5,0,5,5))
        grid.pack(fill="both", expand=True)
        
        stats_frame = ttk.Labelframe(cal_win, text=" 当月统计 ", padding=10, bootstyle="info")
        stats_frame.pack(fill="x", padx=10, pady=10)
        
        f_req = ttk.Frame(stats_frame); f_req.pack(side="left", expand=True)
        ttk.Label(f_req, text="应出勤", font=("微软雅黑", 9), bootstyle="secondary").pack()
        lbl_stat_req = ttk.Label(f_req, text="0h", font=FONT_BOLD, bootstyle="dark"); lbl_stat_req.pack()
        ttk.Separator(stats_frame, orient="vertical").pack(side="left", fill="y", padx=5)
        
        f_act = ttk.Frame(stats_frame); f_act.pack(side="left", expand=True)
        ttk.Label(f_act, text="已出勤", font=("微软雅黑", 9), bootstyle="secondary").pack()
        lbl_stat_act = ttk.Label(f_act, text="0h", font=FONT_BOLD, bootstyle="success"); lbl_stat_act.pack()
        ttk.Separator(stats_frame, orient="vertical").pack(side="left", fill="y", padx=5)
        
        f_abs = ttk.Frame(stats_frame); f_abs.pack(side="left", expand=True)
        ttk.Label(f_abs, text="缺  勤", font=("微软雅黑", 9), bootstyle="secondary").pack()
        lbl_stat_abs = ttk.Label(f_abs, text="0h", font=FONT_BOLD, bootstyle="danger"); lbl_stat_abs.pack()
        
        self.cal_year, self.cal_month = datetime.now().year, datetime.now().month
        def render():
            for w in grid.winfo_children(): w.destroy()
            lbl_title.config(text=f"{self.cal_year}年 {self.cal_month}月")
            conn = sqlite3.connect(DB_NAME); c = conn.cursor()
            query = f"{self.cal_year}-{self.cal_month:02d}-%"
            c.execute("SELECT date, punches, duration, day_type, status FROM attendance WHERE date LIKE ?", (query,))
            rows = c.fetchall()
            conn.close()
            recs = {r[0]: {'punches':r[1], 'duration':r[2], 'type':r[3], 'status':r[4]} for r in rows}
            cal_data = calendar.monthcalendar(self.cal_year, self.cal_month)
            today_str = date.today().strftime("%Y-%m-%d")
            total_req, total_act, total_absent = 0.0, 0.0, 0.0
            
            for r, week in enumerate(cal_data):
                grid.rowconfigure(r, weight=1)
                for c, d in enumerate(week):
                    grid.columnconfigure(c, weight=1)
                    if d==0: continue
                    d_str = f"{self.cal_year}-{self.cal_month:02d}-{d:02d}"
                    rec = recs.get(d_str)
                    is_sunday = (c == 6)
                    is_work_day_default = not is_sunday
                    should_count_req = is_work_day_default
                    if rec:
                        if rec['type'] == 1: should_count_req = False
                        elif rec['type'] in [2, 3]: should_count_req = True
                        elif rec['type'] == 0: should_count_req = True
                    if should_count_req: total_req += DAILY_NET_HOURS
                    day_dur = rec['duration'] if rec else 0.0
                    total_act += day_dur
                    if d_str < today_str and should_count_req:
                        total_absent += max(0, DAILY_NET_HOURS - day_dur)
                        
                    bg, txt = "light", str(d)
                    if rec:
                        if rec['type']==1: bg, txt = "secondary", f"{d}\n非"
                        elif rec['type']==2: bg, txt = "warning", f"{d}\n假"
                        elif rec['type']==3: bg, txt = "info", f"{d}\n调"
                        else:
                            if rec['status'] == STATUS_WORKING: 
                                if d_str != today_str:
                                    bg, txt = "warning-outline", f"{d}\n异" 
                            else:
                                diff = rec['duration'] - DAILY_NET_HOURS
                                diff_str = f"{diff:+.1f}"
                                txt = f"{d}\n\n{rec['duration']:.1f}h\n{diff_str}h"
                                if rec['duration'] >= DAILY_NET_HOURS: 
                                    bg = "success"
                                else: 
                                    bg = "primary"
                    else:
                        if d_str < today_str and is_work_day_default: 
                            # 🟢 核心修改：直接使用内置 danger 样式 (实心红)
                            bg, txt = "danger", f"{d}\n缺"
                        elif is_sunday: bg = "secondary-outline"
                    
                    if d_str == today_str: 
                        target_color = "warning"
                        if "outline" in bg: bg = target_color
                        elif bg == "light": 
                            bg = target_color
                            txt = f"{d}\n今"

                    btn = ttk.Button(grid, text=txt, bootstyle=bg, command=lambda x=d_str: self.open_edit_dialog(x, cal_win, render))
                    btn.grid(row=r, column=c, sticky="nsew", padx=1, pady=1)
            lbl_stat_req.config(text=f"{total_req:.1f}h")
            lbl_stat_act.config(text=f"{total_act:.1f}h")
            lbl_stat_abs.config(text=f"{total_absent:.1f}h")
        def chg(x):
            self.cal_month += x
            if self.cal_month>12: self.cal_month, self.cal_year = 1, self.cal_year+1
            elif self.cal_month<1: self.cal_month, self.cal_year = 12, self.cal_year-1
            render()
        render()
        self.center_and_show(480, 650, cal_win)

    def open_edit_dialog(self, d_str, parent=None, callback=None):
        win = ttk.Toplevel(parent if parent else self)
        win.withdraw()
        win.title("记录管理")
            
        win.resizable(True, True) 
        
        rec = self.get_record(d_str)
        def_punches = rec['punches'].split(',') if (rec and rec['punches']) else []
        def_type = rec['type'] if rec else 0
        
        top = ttk.Frame(win, bootstyle="primary", padding=15)
        top.pack(fill="x")
        ttk.Label(top, text=f"📅  {d_str}", font=("Segoe UI", 16, "bold"), bootstyle="inverse-primary").pack()
        
        bot = ttk.Frame(win, padding=20)
        bot.pack(side="bottom", fill="x")
        content = ttk.Frame(win, padding=20)
        content.pack(fill="both", expand=True)
        
        v_type = tk.IntVar(value=def_type)
        entry_list = [] 
        f_type = ttk.Labelframe(content, text=" 类型 ", padding=10)
        f_type.pack(fill="x", pady=(0, 15))
        
        frame_input = ttk.Frame(content)
        frame_note = ttk.Frame(content)
        f_punches = ttk.Labelframe(frame_input, padding=10, bootstyle="default") 
        f_punches.pack(fill="both", expand=True)
        
        tool_frame = ttk.Frame(f_punches)
        tool_frame.pack(fill="x", pady=(0, 5))
        ttk.Label(tool_frame, text="记录时间 (HH:MM)", font=FONT_BOLD, bootstyle="dark").pack(side="left")
        
        rows_frame = ttk.Frame(f_punches)
        rows_frame.pack(fill="both", expand=True)
        
        note_icon = ttk.Label(frame_note, image=self.imgs.get("coffee_L"), anchor="center")
        note_icon.pack(pady=(20, 10))
        note_title = ttk.Label(frame_note, text="非工作日", font=("微软雅黑", 20, "bold"), anchor="center")
        note_title.pack(pady=(0, 10))
        note_desc = ttk.Label(frame_note, text="...", font=("微软雅黑", 11), justify="center", anchor="center", bootstyle="secondary")
        note_desc.pack()
        
        def switch_view():
            ty = v_type.get()
            frame_input.pack_forget()
            frame_note.pack_forget()
            if ty == 0:
                frame_input.pack(fill="both", expand=True)
                if len(entry_list) == 0: add_entry_row(); add_entry_row()
            else:
                frame_note.pack(fill="both", expand=True)
                if ty == 1:
                    note_icon.config(image=self.imgs.get("coffee_L")); note_title.config(text="非工作日", foreground="#E68585")
                    note_desc.config(text="好好休息\n不计入应出勤时长")
                elif ty == 2:
                    note_icon.config(image=self.imgs.get("beach")); note_title.config(text="法定节假日", foreground="#FF9800")
                    note_desc.config(text="假期愉快\n默认计入7小时出勤")
                elif ty == 3:
                    note_icon.config(image=self.imgs.get("sleep")); note_title.config(text="调休", foreground="#20BC99")
                    note_desc.config(text="补休/调休\n默认计入7小时出勤")
                    
        ttk.Radiobutton(f_type, text="工作日", variable=v_type, value=0, command=switch_view).grid(row=0, column=0, sticky="w", padx=10, pady=5)
        ttk.Radiobutton(f_type, text="非工作日", variable=v_type, value=1, command=switch_view).grid(row=0, column=1, sticky="w", padx=10, pady=5)
        ttk.Radiobutton(f_type, text="法定假", variable=v_type, value=2, command=switch_view).grid(row=0, column=2, sticky="w", padx=10, pady=5)
        ttk.Radiobutton(f_type, text="调休", variable=v_type, value=3, command=switch_view).grid(row=0, column=3, sticky="w", padx=10, pady=5)
        
        def add_entry_row(val=None):
            if len(entry_list) >= 6: return
            if val is None: val = datetime.now().strftime("%H:%M")
            row = ttk.Frame(rows_frame)
            row.pack(fill="x", pady=4) 
            ttk.Label(row, text=f"{len(entry_list)+1}.", width=3, bootstyle="dark", font=FONT_BOLD).pack(side="left")
            e = ttk.Entry(row, font=("Segoe UI", 12), justify="center")
            e.insert(0, val)
            e.pack(side="left", fill="x", expand=True)
            entry_list.append((row, e))
            
        def remove_last_row():
            if len(entry_list) > 2:
                row, e = entry_list.pop()
                row.destroy()
                
        ttk.Button(tool_frame, text="+", width=4, bootstyle="success-outline", command=lambda: add_entry_row(None)).pack(side="right")
        ttk.Button(tool_frame, text="-", width=4, bootstyle="secondary-outline", command=remove_last_row).pack(side="right", padx=5)
        
        if def_punches:
            for p in def_punches: add_entry_row(p)
        switch_view()
        
        def run_del():
            if messagebox.askyesno("确认删除", "确定清空当日记录?", parent=win):
                conn = sqlite3.connect(DB_NAME)
                conn.execute("DELETE FROM attendance WHERE date=?", (d_str,))
                conn.commit(); conn.close()
                if d_str == self.today_date: self.refresh_main_data()
                win.destroy()
                if callback: callback()
                
        def run_save():
            ty = v_type.get()
            new_punches = []
            dur = 0.0
            status = STATUS_IDLE 
            if ty == 0:
                for row, e in entry_list:
                    val = e.get().strip()
                    if val:
                        formatted = self.format_time_str(val)
                        if not formatted:
                            messagebox.showerror("格式错误", f"时间格式不正确: {val}\n请使用 HH:MM", parent=win)
                            return
                        new_punches.append(formatted)
                new_punches.sort()
                dur = self.calculate_logic(new_punches)
            elif ty == 1: dur = 0.0
            else: dur = DAILY_NET_HOURS
            conn = sqlite3.connect(DB_NAME)
            conn.execute("INSERT OR REPLACE INTO attendance (date, punches, duration, day_type, status) VALUES (?,?,?,?,?)",
                         (d_str, ",".join(new_punches), dur, ty, status))
            conn.commit(); conn.close()
            if d_str == self.today_date: self.refresh_main_data()
            win.destroy()
            if callback: callback()
            
        ttk.Button(bot, text=" 清空", image=self.imgs.get("banana"), compound="left", bootstyle="danger-outline", width=10, command=run_del).pack(side="left")
        ttk.Button(bot, text=" 保存", image=self.imgs.get("save"), compound="left", bootstyle="primary", width=12, command=run_save).pack(side="right")
        self.center_and_show(450, 620, win)

if __name__ == "__main__":
    app = WorkAppPro()
    app.mainloop()