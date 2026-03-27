#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MES 终端扫描系统
功能:
1. 设备登录/登出，自动延期 token(24h)
2. 获取绑定工单列表，支持下拉选择和刷新
3. 连续扫码，调用出站接口
4. 离线模式：扫码数据存 SQLite，上线后上传
5. 接口可配置，管理员密码验证后修改
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import requests
import json
import sqlite3
import threading
import os
import winsound
from datetime import datetime, timedelta
from typing import Optional, List, Dict
import csv
import configparser

# ==================== 配置文件管理 ====================
CONFIG_FILE = "config.ini"
DEFAULT_CONFIG = {
    "api": {
        "base_url": "http://172.16.0.10:8080/Mes/",
        "login_path": "/Device/Login",
        "prolong_path": "/Device/reLogin",
        "crossing_path": "/Device/Cross_station",
        "mmo_list_path": "/Device/get_MmoList"
    },
    "admin": {
        "password": "123456"
    },
    "settings": {
        "timeout": "10",
        "auto_prolong_hours": "23"
    }
}

class ConfigManager:
    """配置文件管理器"""

    def __init__(self, config_file: str = CONFIG_FILE):
        self.config_file = config_file
        self.config = configparser.ConfigParser()
        self._load_or_create_config()

    def _load_or_create_config(self):
        """加载或创建配置文件"""
        if os.path.exists(self.config_file):
            self.config.read(self.config_file, encoding='utf-8')
        else:
            # 创建默认配置
            for section, options in DEFAULT_CONFIG.items():
                self.config.add_section(section)
                for key, value in options.items():
                    self.config.set(section, key, value)
            self._save_config()

    def _save_config(self):
        """保存配置文件"""
        with open(self.config_file, 'w', encoding='utf-8') as f:
            self.config.write(f)

    def get(self, section: str, key: str, fallback: str = None) -> str:
        """获取配置值"""
        return self.config.get(section, key, fallback=fallback)

    def get_int(self, section: str, key: str, fallback: int = 0) -> int:
        """获取整数配置值"""
        return self.config.getint(section, key, fallback=fallback)

    def set(self, section: str, key: str, value: str):
        """设置配置值"""
        if not self.config.has_section(section):
            self.config.add_section(section)
        self.config.set(section, key, value)
        self._save_config()

    def get_all_api_settings(self) -> Dict:
        """获取所有 API 设置"""
        return {
            "base_url": self.get("api", "base_url"),
            "login_path": self.get("api", "login_path"),
            "prolong_path": self.get("api", "prolong_path"),
            "crossing_path": self.get("api", "crossing_path"),
            "mmo_list_path": self.get("api", "mmo_list_path"),
            "timeout": self.get_int("settings", "timeout", 10),
            "auto_prolong_hours": self.get_int("settings", "auto_prolong_hours", 23)
        }

    def get_admin_password(self) -> str:
        """获取管理员密码"""
        return self.get("admin", "password", fallback="123456")

    def set_admin_password(self, new_password: str):
        """设置管理员密码"""
        self.set("admin", "password", new_password)

    def update_api_settings(self, settings: Dict):
        """批量更新 API 设置"""
        for key, value in settings.items():
            if key == "base_url":
                self.set("api", "base_url", value)
            elif key == "login_path":
                self.set("api", "login_path", value)
            elif key == "prolong_path":
                self.set("api", "prolong_path", value)
            elif key == "crossing_path":
                self.set("api", "crossing_path", value)
            elif key == "mmo_list_path":
                self.set("api", "mmo_list_path", value)
            elif key == "timeout":
                self.set("settings", "timeout", str(value))
            elif key == "auto_prolong_hours":
                self.set("settings", "auto_prolong_hours", str(value))


# ==================== 数据库操作 ====================
DB_FILE = "offline_data.db"

class OfflineDatabase:
    """离线数据 SQLite 数据库管理"""

    def __init__(self, db_path: str = DB_FILE):
        self.db_path = db_path
        self._init_db()

    def _init_db(self):
        """初始化数据库表"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS scan_records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                device_code TEXT NOT NULL,
                mmo_code TEXT NOT NULL,
                label TEXT NOT NULL,
                qty INTEGER DEFAULT 1,
                result INTEGER DEFAULT 10,
                scanned_time TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                uploaded INTEGER DEFAULT 0
            )
        ''')
        conn.commit()
        conn.close()

    def add_scan_record(self, device_code: str, mmo_code: str, label: str,
                        qty: int = 1, result: int = 10):
        """添加扫码记录"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO scan_records (device_code, mmo_code, label, qty, result)
            VALUES (?, ?, ?, ?, ?)
        ''', (device_code, mmo_code, label, qty, result))
        conn.commit()
        conn.close()

    def get_unuploaded_records(self) -> List[Dict]:
        """获取未上传的记录"""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM scan_records WHERE uploaded = 0 ORDER BY scanned_time')
        records = [dict(row) for row in cursor.fetchall()]
        conn.close()
        return records

    def mark_as_uploaded(self, record_ids: List[int]):
        """标记记录为已上传"""
        if not record_ids:
            return
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        placeholders = ','.join('?' * len(record_ids))
        cursor.execute(f'UPDATE scan_records SET uploaded = 1 WHERE id IN ({placeholders})', record_ids)
        conn.commit()
        conn.close()

    def get_all_records(self) -> List[Dict]:
        """获取所有记录（用于导出）"""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM scan_records ORDER BY scanned_time DESC')
        records = [dict(row) for row in cursor.fetchall()]
        conn.close()
        return records

    def clear_uploaded_records(self):
        """清除已上传的记录"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('DELETE FROM scan_records WHERE uploaded = 1')
        conn.commit()
        conn.close()

    def get_record_count(self) -> int:
        """获取未上传记录数量"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('SELECT COUNT(*) FROM scan_records WHERE uploaded = 0')
        count = cursor.fetchone()[0]
        conn.close()
        return count


# ==================== MES API 客户端 ====================
class MESClient:
    """MES API 客户端"""

    def __init__(self, config: ConfigManager):
        self.config = config
        self.token: Optional[str] = None
        self.session = requests.Session()

    @property
    def base_url(self) -> str:
        return self.config.get("api", "base_url")

    @property
    def timeout(self) -> int:
        return self.config.get_int("settings", "timeout", 10)

    def _get_headers(self) -> Dict:
        """获取请求头"""
        headers = {"Content-Type": "application/json"}
        if self.token:
            headers["Jwt"] = self.token
        return headers

    def _get_full_url(self, path: str) -> str:
        """获取完整 URL"""
        base = self.base_url.rstrip('/')
        clean_path = path.lstrip('/')
        return f"{base}/{clean_path}"

    def login(self, username: str, password: str, device_code: str) -> Dict:
        """设备登录"""
        url = self._get_full_url(self.config.get("api", "login_path"))
        payload = {
            "userName": username,
            "password": password,
            "deviceCode": device_code
        }
        try:
            resp = self.session.post(url, json=payload, headers=self._get_headers(), timeout=self.timeout)
            resp.raise_for_status()
            data = resp.json()

            if data.get("state") == 200:
                self.token = data.get("data")
                return {"success": True, "token": self.token}
            else:
                return {"success": False, "msg": data.get("msg", "登录失败")}
        except requests.exceptions.Timeout:
            return {"success": False, "msg": "请求超时"}
        except requests.exceptions.ConnectionError:
            return {"success": False, "msg": "无法连接到服务器"}
        except Exception as e:
            return {"success": False, "msg": str(e)}

    def prolong(self) -> Dict:
        """延期 token(24h)"""
        url = self._get_full_url(self.config.get("api", "prolong_path"))
        try:
            resp = self.session.post(url, headers=self._get_headers(), timeout=self.timeout)
            resp.raise_for_status()
            data = resp.json()

            if data.get("state") == 200:
                new_token = data.get("data")
                if new_token:
                    self.token = new_token
                return {"success": True, "token": self.token}
            else:
                return {"success": False, "msg": data.get("msg", "延期失败")}
        except requests.exceptions.Timeout:
            return {"success": False, "msg": "请求超时"}
        except requests.exceptions.ConnectionError:
            return {"success": False, "msg": "无法连接到服务器"}
        except Exception as e:
            return {"success": False, "msg": str(e)}

    def get_mmo_list(self, device_code: str) -> Dict:
        """获取绑定设备的工单列表"""
        url = self._get_full_url(self.config.get("api", "mmo_list_path"))
        params = {"deviceCode": device_code}
        try:
            resp = self.session.get(url, params=params, headers=self._get_headers(), timeout=self.timeout)
            resp.raise_for_status()
            data = resp.json()

            if data.get("state") == 200:
                # data 是一个对象，包含 source 数组
                data_obj = data.get("data", {})
                if isinstance(data_obj, dict):
                    source_list = data_obj.get("source", [])
                else:
                    source_list = data_obj if isinstance(data_obj, list) else []
                return {"success": True, "mmoList": source_list}
            else:
                return {"success": False, "msg": data.get("msg", "获取工单列表失败")}
        except requests.exceptions.Timeout:
            return {"success": False, "msg": "请求超时"}
        except requests.exceptions.ConnectionError:
            return {"success": False, "msg": "无法连接到服务器"}
        except Exception as e:
            return {"success": False, "msg": str(e)}

    def crossing(self, device_code: str, mmo_code: str, labels: List[Dict]) -> Dict:
        """出站接口"""
        url = self._get_full_url(self.config.get("api", "crossing_path"))
        payload = {
            "deviceCode": device_code,
            "mmoCode": mmo_code,
            "labels": labels
        }
        try:
            resp = self.session.post(url, json=payload, headers=self._get_headers(), timeout=self.timeout)
            resp.raise_for_status()
            data = resp.json()

            if data.get("state") == 200:
                return {"success": True, "data": data.get("data")}
            else:
                return {"success": False, "msg": data.get("msg", "出站失败"), "code": data.get("code")}
        except requests.exceptions.Timeout:
            return {"success": False, "msg": "请求超时"}
        except requests.exceptions.ConnectionError:
            return {"success": False, "msg": "无法连接到服务器"}
        except Exception as e:
            return {"success": False, "msg": str(e)}

    def logout(self):
        """登出（清空 token）"""
        self.token = None


# ==================== 设置对话框 ====================
class SettingsDialog:
    """API 设置对话框"""

    def __init__(self, parent, config: ConfigManager):
        self.parent = parent
        self.config = config
        self.result = False

    def show(self):
        """显示设置对话框"""
        dialog = tk.Toplevel(self.parent)
        dialog.title("API 设置")
        dialog.geometry("600x500")
        dialog.transient(self.parent)
        dialog.grab_set()

        # 标题
        title_label = ttk.Label(dialog, text="API 接口配置", font=("Microsoft YaHei", 14, "bold"))
        title_label.pack(pady=15)

        # 表单框架
        form_frame = ttk.Frame(dialog, padding="20")
        form_frame.pack(fill=tk.BOTH, expand=True)

        # 配置项
        settings = self.config.get_all_api_settings()

        entries = {}

        # 基础 URL
        ttk.Label(form_frame, text="基础 URL:", font=("Microsoft YaHei", 11)).grid(row=0, column=0, sticky=tk.W, pady=8)
        entries["base_url"] = ttk.Entry(form_frame, width=50, font=("Microsoft YaHei", 11))
        entries["base_url"].grid(row=0, column=1, padx=10, pady=8)
        entries["base_url"].insert(0, settings["base_url"])

        # 登录接口路径
        ttk.Label(form_frame, text="登录接口路径:", font=("Microsoft YaHei", 11)).grid(row=1, column=0, sticky=tk.W, pady=8)
        entries["login_path"] = ttk.Entry(form_frame, width=50, font=("Microsoft YaHei", 11))
        entries["login_path"].grid(row=1, column=1, padx=10, pady=8)
        entries["login_path"].insert(0, settings["login_path"])

        # 延期接口路径
        ttk.Label(form_frame, text="延期接口路径:", font=("Microsoft YaHei", 11)).grid(row=2, column=0, sticky=tk.W, pady=8)
        entries["prolong_path"] = ttk.Entry(form_frame, width=50, font=("Microsoft YaHei", 11))
        entries["prolong_path"].grid(row=2, column=1, padx=10, pady=8)
        entries["prolong_path"].insert(0, settings["prolong_path"])

        # 工单列表接口路径
        ttk.Label(form_frame, text="工单列表接口路径:", font=("Microsoft YaHei", 11)).grid(row=3, column=0, sticky=tk.W, pady=8)
        entries["mmo_list_path"] = ttk.Entry(form_frame, width=50, font=("Microsoft YaHei", 11))
        entries["mmo_list_path"].grid(row=3, column=1, padx=10, pady=8)
        entries["mmo_list_path"].insert(0, settings["mmo_list_path"])

        # 出站接口路径
        ttk.Label(form_frame, text="出站接口路径:", font=("Microsoft YaHei", 11)).grid(row=4, column=0, sticky=tk.W, pady=8)
        entries["crossing_path"] = ttk.Entry(form_frame, width=50, font=("Microsoft YaHei", 11))
        entries["crossing_path"].grid(row=4, column=1, padx=10, pady=8)
        entries["crossing_path"].insert(0, settings["crossing_path"])

        # 请求超时时间
        ttk.Label(form_frame, text="请求超时时间 (秒):", font=("Microsoft YaHei", 11)).grid(row=5, column=0, sticky=tk.W, pady=8)
        entries["timeout"] = ttk.Entry(form_frame, width=20, font=("Microsoft YaHei", 11))
        entries["timeout"].grid(row=5, column=1, padx=10, pady=8, sticky=tk.W)
        entries["timeout"].insert(0, str(settings["timeout"]))

        # 自动延期小时数
        ttk.Label(form_frame, text="自动延期间隔 (小时):", font=("Microsoft YaHei", 11)).grid(row=6, column=0, sticky=tk.W, pady=8)
        entries["auto_prolong_hours"] = ttk.Entry(form_frame, width=20, font=("Microsoft YaHei", 11))
        entries["auto_prolong_hours"].grid(row=6, column=1, padx=10, pady=8, sticky=tk.W)
        entries["auto_prolong_hours"].insert(0, str(settings["auto_prolong_hours"]))

        # 按钮框架
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=20)

        # 重置默认按钮
        def on_reset():
            defaults = DEFAULT_CONFIG["api"]
            entries["base_url"].delete(0, tk.END)
            entries["base_url"].insert(0, defaults["base_url"])
            entries["login_path"].delete(0, tk.END)
            entries["login_path"].insert(0, defaults["login_path"])
            entries["prolong_path"].delete(0, tk.END)
            entries["prolong_path"].insert(0, defaults["prolong_path"])
            entries["mmo_list_path"].delete(0, tk.END)
            entries["mmo_list_path"].insert(0, defaults["mmo_list_path"])
            entries["crossing_path"].delete(0, tk.END)
            entries["crossing_path"].insert(0, defaults["crossing_path"])
            entries["timeout"].delete(0, tk.END)
            entries["timeout"].insert(0, DEFAULT_CONFIG["settings"]["timeout"])
            entries["auto_prolong_hours"].delete(0, tk.END)
            entries["auto_prolong_hours"].insert(0, DEFAULT_CONFIG["settings"]["auto_prolong_hours"])

        ttk.Button(btn_frame, text="恢复默认", command=on_reset).pack(side=tk.LEFT, padx=10)

        # 取消按钮
        ttk.Button(btn_frame, text="取消", command=lambda: dialog.destroy()).pack(side=tk.RIGHT, padx=10)

        # 保存按钮
        def on_save():
            try:
                new_settings = {
                    "base_url": entries["base_url"].get().strip(),
                    "login_path": entries["login_path"].get().strip(),
                    "prolong_path": entries["prolong_path"].get().strip(),
                    "mmo_list_path": entries["mmo_list_path"].get().strip(),
                    "crossing_path": entries["crossing_path"].get().strip(),
                    "timeout": int(entries["timeout"].get().strip()),
                    "auto_prolong_hours": int(entries["auto_prolong_hours"].get().strip())
                }

                if not new_settings["base_url"]:
                    messagebox.showerror("错误", "基础 URL 不能为空", parent=dialog)
                    return
                if not new_settings["base_url"].startswith("http"):
                    messagebox.showerror("错误", "基础 URL 必须以 http://或 https://开头", parent=dialog)
                    return

                self.config.update_api_settings(new_settings)
                messagebox.showinfo("成功", "设置已保存", parent=dialog)
                self.result = True
                dialog.destroy()

            except ValueError as e:
                messagebox.showerror("错误", "超时时间和延期间隔必须是数字", parent=dialog)

        ttk.Button(btn_frame, text="保存", command=on_save).pack(side=tk.RIGHT, padx=10)

        # 等待对话框关闭
        self.parent.wait_window(dialog)
        return self.result


class AdminPasswordDialog:
    """管理员密码修改对话框"""

    def __init__(self, parent, config: ConfigManager):
        self.parent = parent
        self.config = config

    def show(self):
        """显示密码修改对话框"""
        dialog = tk.Toplevel(self.parent)
        dialog.title("修改管理员密码")
        dialog.geometry("400x250")
        dialog.transient(self.parent)
        dialog.grab_set()

        # 标题
        title_label = ttk.Label(dialog, text="修改管理员密码", font=("Microsoft YaHei", 14, "bold"))
        title_label.pack(pady=15)

        # 表单框架
        form_frame = ttk.Frame(dialog, padding="20")
        form_frame.pack(fill=tk.BOTH, expand=True)

        # 当前密码
        ttk.Label(form_frame, text="当前密码:", font=("Microsoft YaHei", 11)).grid(row=0, column=0, sticky=tk.W, pady=8)
        current_pwd = ttk.Entry(form_frame, width=30, font=("Microsoft YaHei", 11), show="*")
        current_pwd.grid(row=0, column=1, padx=10, pady=8)

        # 新密码
        ttk.Label(form_frame, text="新密码:", font=("Microsoft YaHei", 11)).grid(row=1, column=0, sticky=tk.W, pady=8)
        new_pwd = ttk.Entry(form_frame, width=30, font=("Microsoft YaHei", 11), show="*")
        new_pwd.grid(row=1, column=1, padx=10, pady=8)

        # 确认新密码
        ttk.Label(form_frame, text="确认新密码:", font=("Microsoft YaHei", 11)).grid(row=2, column=0, sticky=tk.W, pady=8)
        confirm_pwd = ttk.Entry(form_frame, width=30, font=("Microsoft YaHei", 11), show="*")
        confirm_pwd.grid(row=2, column=1, padx=10, pady=8)

        # 按钮框架
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=20)

        ttk.Button(btn_frame, text="取消", command=lambda: dialog.destroy()).pack(side=tk.RIGHT, padx=10)

        def on_save():
            current = current_pwd.get()
            new = new_pwd.get()
            confirm = confirm_pwd.get()

            if current != self.config.get_admin_password():
                messagebox.showerror("错误", "当前密码不正确", parent=dialog)
                return

            if len(new) < 1:
                messagebox.showerror("错误", "新密码不能为空", parent=dialog)
                return

            if new != confirm:
                messagebox.showerror("错误", "两次输入的新密码不一致", parent=dialog)
                return

            self.config.set_admin_password(new)
            messagebox.showinfo("成功", "管理员密码已修改", parent=dialog)
            dialog.destroy()

        ttk.Button(btn_frame, text="保存", command=on_save).pack(side=tk.RIGHT, padx=10)

        self.parent.wait_window(dialog)


# ==================== 主界面 ====================
class MESTerminalApp:
    """MES 终端主界面"""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("MES 终端扫描系统")
        self.root.geometry("900x700")

        # 初始化配置管理
        self.config = ConfigManager()

        # 初始化
        self.mes_client = MESClient(self.config)
        self.db = OfflineDatabase(DB_FILE)
        self.is_logged_in = False
        self.is_offline_mode = False
        self.current_device_code = ""
        self.current_username = ""
        self.current_mmo_code = ""
        self.scan_count = 0
        self.auto_prolong_running = False
        self.scan_var = tk.StringVar()
        self.scan_entry = None

        # 创建菜单
        self._create_menu()

        # 创建界面
        self._create_login_frame()
        self._create_main_frame()

    def _create_menu(self):
        """创建菜单栏"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # 系统菜单
        system_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="系统", menu=system_menu)
        system_menu.add_command(label="API 设置", command=self._open_api_settings)
        system_menu.add_command(label="修改管理员密码", command=self._open_password_settings)
        system_menu.add_separator()
        system_menu.add_command(label="退出", command=self.on_closing)

        # 帮助菜单
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="帮助", menu=help_menu)
        help_menu.add_command(label="关于", command=self._show_about)

    def _open_api_settings(self):
        """打开 API 设置"""
        # 验证管理员密码
        password = simpledialog.askstring("验证", "请输入管理员密码:", parent=self.root, show="*")
        if password is None:
            return
        if password != self.config.get_admin_password():
            messagebox.showerror("错误", "管理员密码错误")
            return

        # 显示设置对话框
        settings_dlg = SettingsDialog(self.root, self.config)
        if settings_dlg.show():
            # 设置已保存，重新初始化客户端
            self.mes_client = MESClient(self.config)
            self._log("API 设置已更新")

    def _open_password_settings(self):
        """打开密码修改"""
        pwd_dlg = AdminPasswordDialog(self.root, self.config)
        pwd_dlg.show()

    def _show_about(self):
        """显示关于对话框"""
        messagebox.showinfo("关于", "MES 终端扫描系统\n\n版本：1.0\n功能：扫码出站、离线模式、自动延期")

    def _create_login_frame(self):
        """创建登录框架"""
        self.login_frame = ttk.Frame(self.root, padding="20")
        self.login_frame.pack(fill=tk.BOTH, expand=True)

        # 标题
        title_label = ttk.Label(self.login_frame, text="MES 终端扫描系统",
                                font=("Microsoft YaHei", 18, "bold"))
        title_label.pack(pady=20)

        # 登录表单
        form_frame = ttk.Frame(self.login_frame)
        form_frame.pack(pady=30)

        # 用户名
        ttk.Label(form_frame, text="用户名:", font=("Microsoft YaHei", 12)).grid(row=0, column=0, sticky=tk.W, pady=10)
        self.username_entry = ttk.Entry(form_frame, width=30, font=("Microsoft YaHei", 12))
        self.username_entry.grid(row=0, column=1, padx=10, pady=10)

        # 密码
        ttk.Label(form_frame, text="密码:", font=("Microsoft YaHei", 12)).grid(row=1, column=0, sticky=tk.W, pady=10)
        self.password_entry = ttk.Entry(form_frame, width=30, font=("Microsoft YaHei", 12), show="*")
        self.password_entry.grid(row=1, column=1, padx=10, pady=10)

        # 设备编码
        ttk.Label(form_frame, text="设备编码:", font=("Microsoft YaHei", 12)).grid(row=2, column=0, sticky=tk.W, pady=10)
        self.device_code_entry = ttk.Entry(form_frame, width=30, font=("Microsoft YaHei", 12))
        self.device_code_entry.grid(row=2, column=1, padx=10, pady=10)
        self.device_code_entry.insert(0, "P10-10")  # 默认值

        # 离线模式复选框
        self.offline_var = tk.BooleanVar()
        offline_check = ttk.Checkbutton(form_frame, text="离线模式",
                                         variable=self.offline_var)
        offline_check.grid(row=3, column=0, columnspan=2, pady=20)

        # 登录按钮
        self.login_btn = ttk.Button(form_frame, text="登录", command=self._do_login)
        self.login_btn.grid(row=4, column=0, columnspan=2, pady=20)

        # 状态标签
        self.login_status_label = ttk.Label(form_frame, text="", foreground="red",
                                            font=("Microsoft YaHei", 10))
        self.login_status_label.grid(row=5, column=0, columnspan=2)

    def _create_main_frame(self):
        """创建主界面框架（初始隐藏）"""
        self.main_frame = ttk.Frame(self.root, padding="10")

        # 顶部状态栏
        self._create_status_bar()

        # 工单选择区域
        self._create_order_select_frame()

        # 扫码区域
        self._create_scan_frame()

        # 底部操作区
        self._create_bottom_frame()

    def _create_status_bar(self):
        """创建状态栏"""
        status_frame = ttk.LabelFrame(self.main_frame, text="状态信息", padding="10")
        status_frame.pack(fill=tk.X, pady=5)

        # 用户信息
        self.user_info_label = ttk.Label(status_frame, text="", font=("Microsoft YaHei", 10))
        self.user_info_label.pack(side=tk.LEFT, padx=10)

        # Token 状态
        self.token_status_label = ttk.Label(status_frame, text="", font=("Microsoft YaHei", 10))
        self.token_status_label.pack(side=tk.LEFT, padx=10)

        # 离线模式状态
        self.offline_status_label = ttk.Label(status_frame, text="", font=("Microsoft YaHei", 10), foreground="orange")
        self.offline_status_label.pack(side=tk.RIGHT, padx=10)

        # 未上传数据数量
        self.offline_count_label = ttk.Label(status_frame, text="", font=("Microsoft YaHei", 10), foreground="blue")
        self.offline_count_label.pack(side=tk.RIGHT, padx=10)

    def _create_order_select_frame(self):
        """创建工单选择区域"""
        order_frame = ttk.LabelFrame(self.main_frame, text="工单选择", padding="10")
        order_frame.pack(fill=tk.X, pady=5)

        # 工单下拉框
        ttk.Label(order_frame, text="选择工单:", font=("Microsoft YaHei", 11)).pack(side=tk.LEFT)

        self.mmo_var = tk.StringVar()
        self.mmo_combo = ttk.Combobox(order_frame, textvariable=self.mmo_var,
                                       width=40, font=("Microsoft YaHei", 11), state="readonly")
        self.mmo_combo.pack(side=tk.LEFT, padx=10)
        self.mmo_combo.bind("<<ComboboxSelected>>", self._on_mmo_selected)

        # 刷新按钮
        refresh_btn = ttk.Button(order_frame, text="刷新", command=self._refresh_mmo_list)
        refresh_btn.pack(side=tk.LEFT, padx=10)

    def _create_scan_frame(self):
        """创建扫码区域"""
        scan_frame = ttk.LabelFrame(self.main_frame, text="扫码操作", padding="15")
        scan_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        # 扫码计数显示
        count_frame = ttk.Frame(scan_frame)
        count_frame.pack(fill=tk.X, pady=5)

        ttk.Label(count_frame, text="当前扫码数量:", font=("Microsoft YaHei", 14)).pack(side=tk.LEFT, padx=5)
        self.scan_count_label = ttk.Label(count_frame, text="0", font=("Microsoft YaHei", 14, "bold"), foreground="green")
        self.scan_count_label.pack(side=tk.LEFT, padx=5)

        # 扫码输入框
        ttk.Label(scan_frame, text="扫码:", font=("Microsoft YaHei", 12)).pack(pady=5)

        self.scan_entry = ttk.Entry(scan_frame, textvariable=self.scan_var,
                                    font=("Microsoft YaHei", 14), width=50)
        self.scan_entry.pack(pady=10)
        self.scan_entry.bind("<Return>", self._on_scan)

        # 扫码状态
        self.scan_status_label = ttk.Label(scan_frame, text="", font=("Microsoft YaHei", 11))
        self.scan_status_label.pack(pady=5)

        # 开始/停止按钮
        btn_frame = ttk.Frame(scan_frame)
        btn_frame.pack(pady=10)

        self.start_scan_btn = ttk.Button(btn_frame, text="开始扫码", command=self._toggle_scanning)
        self.start_scan_btn.pack(side=tk.LEFT, padx=10)

        # 手动出站按钮（用于离线模式上传）
        self.upload_btn = ttk.Button(btn_frame, text="上传离线数据", command=self._upload_offline_data)
        self.upload_btn.pack(side=tk.LEFT, padx=10)

        # 导出按钮
        self.export_btn = ttk.Button(btn_frame, text="导出记录", command=self._export_records)
        self.export_btn.pack(side=tk.LEFT, padx=10)

        # 日志文本框
        log_frame = ttk.LabelFrame(scan_frame, text="操作日志", padding="5")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        self.log_text = tk.Text(log_frame, height=15, font=("Consolas", 9), state=tk.DISABLED)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.configure(yscrollcommand=scrollbar.set)

    def _create_bottom_frame(self):
        """创建底部操作区"""
        bottom_frame = ttk.Frame(self.main_frame)
        bottom_frame.pack(fill=tk.X, pady=10)

        # 登出按钮
        self.logout_btn = ttk.Button(bottom_frame, text="登出", command=self._do_logout)
        self.logout_btn.pack(side=tk.RIGHT, padx=10)

    def _log(self, message: str, level: str = "INFO"):
        """添加日志"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"[{timestamp}] [{level}] {message}\n"
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, log_entry)
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _do_login(self):
        """执行登录"""
        username = self.username_entry.get().strip()
        password = self.password_entry.get()
        device_code = self.device_code_entry.get().strip()

        if not username or not password or not device_code:
            self.login_status_label.config(text="请填写完整信息")
            return

        self.login_btn.config(state=tk.DISABLED)
        self.login_status_label.config(text="登录中...")

        def login_thread():
            try:
                result = self.mes_client.login(username, password, device_code)
                self.root.after(0, lambda: self._handle_login_result(result))
            except Exception as e:
                self.root.after(0, lambda: self._handle_login_result({"success": False, "msg": str(e)}))

        threading.Thread(target=login_thread, daemon=True).start()

    def _handle_login_result(self, result: Dict):
        """处理登录结果"""
        self.login_btn.config(state=tk.NORMAL)

        if result["success"]:
            self.is_logged_in = True
            self.current_username = self.username_entry.get()
            self.current_device_code = self.device_code_entry.get()
            self.is_offline_mode = self.offline_var.get()

            # 切换界面
            self.login_frame.pack_forget()
            self.main_frame.pack(fill=tk.BOTH, expand=True)

            # 更新状态
            self._update_status_bar()

            # 获取工单列表
            self._refresh_mmo_list()

            # 启动自动延期
            self._start_auto_prolong()

            self._log(f"用户 {self.current_username} 登录成功，设备：{self.current_device_code}")

            if self.is_offline_mode:
                self._log("离线模式已启用", "WARN")
                self._update_offline_count()
        else:
            self.login_status_label.config(text=result["msg"])
            self._log(f"登录失败：{result['msg']}", "ERROR")

    def _update_status_bar(self):
        """更新状态栏"""
        self.user_info_label.config(text=f"用户：{self.current_username} | 设备：{self.current_device_code}")
        self.token_status_label.config(text="Token: 有效", foreground="green")

        if self.is_offline_mode:
            self.offline_status_label.config(text="【离线模式】")
            self._update_offline_count()
        else:
            self.offline_status_label.config(text="")
            self.offline_count_label.config(text="")

    def _update_offline_count(self):
        """更新离线数据计数"""
        if self.is_offline_mode:
            count = self.db.get_record_count()
            self.offline_count_label.config(text=f"待上传：{count}条")

    def _start_auto_prolong(self):
        """启动自动延期（默认每 23 小时延期一次，确保 token 有效）"""
        self.auto_prolong_running = True
        hours = self.config.get_int("settings", "auto_prolong_hours", 23)

        def prolong_loop():
            while self.auto_prolong_running and self.is_logged_in:
                try:
                    # 等待指定小时数
                    for _ in range(hours * 60 * 60):  # 每秒检查一次
                        if not self.auto_prolong_running or not self.is_logged_in:
                            return
                        threading.Event().wait(1)

                    if self.is_logged_in and self.mes_client.token:
                        result = self.mes_client.prolong()
                        if result["success"]:
                            self.root.after(0, lambda: self._log("Token 延期成功"))
                        else:
                            self.root.after(0, lambda: self._log(f"Token 延期失败：{result['msg']}", "ERROR"))
                except Exception as e:
                    self.root.after(0, lambda: self._log(f"自动延期异常：{str(e)}", "ERROR"))

        threading.Thread(target=prolong_loop, daemon=True).start()

    def _refresh_mmo_list(self):
        """刷新工单列表"""
        if not self.is_logged_in:
            return

        def fetch_thread():
            try:
                result = self.mes_client.get_mmo_list(self.current_device_code)
                self.root.after(0, lambda: self._handle_mmo_list_result(result))
            except Exception as e:
                self.root.after(0, lambda: self._handle_mmo_list_result({"success": False, "msg": str(e)}))

        threading.Thread(target=fetch_thread, daemon=True).start()

    def _handle_mmo_list_result(self, result: Dict):
        """处理工单列表结果"""
        if result["success"]:
            mmo_list = result.get("mmoList", [])
            if mmo_list:
                # 从 source 数组的每个对象中提取 code 字段作为工单号
                display_list = []
                for item in mmo_list:
                    if isinstance(item, dict):
                        code = item.get("code", "")
                        if code:
                            display_list.append(code)
                    else:
                        display_list.append(str(item))

                self.mmo_combo["values"] = display_list
                if display_list:
                    self.mmo_combo.current(0)
                    self.current_mmo_code = display_list[0]
                self._log(f"获取到 {len(display_list)} 个工单")
            else:
                self.mmo_combo["values"] = []
                self._log("未找到绑定的工单", "WARN")
        else:
            messagebox.showerror("错误", f"获取工单列表失败:\n{result['msg']}")
            self._log(f"获取工单列表失败：{result['msg']}", "ERROR")

    def _on_mmo_selected(self, event):
        """工单选择变更"""
        self.current_mmo_code = self.mmo_var.get()
        self._log(f"选择工单：{self.current_mmo_code}")

    def _toggle_scanning(self):
        """切换扫码状态"""
        if not self.current_mmo_code:
            messagebox.showwarning("警告", "请先选择工单")
            return

        if self.start_scan_btn.cget("text") == "开始扫码":
            self.start_scan_btn.config(text="停止扫码")
            self.scan_entry.focus_set()
            self.scan_status_label.config(text="● 扫码中...", foreground="green")
            self._log("开始扫码")
        else:
            self.start_scan_btn.config(text="开始扫码")
            self.scan_status_label.config(text="○ 已停止", foreground="red")
            self._log("停止扫码")

    def _on_scan(self, event=None):
        """处理扫码"""
        label = self.scan_var.get().strip()
        if not label:
            return

        # 清空输入框
        self.scan_var.set("")

        # 保持焦点
        if self.scan_entry:
            self.scan_entry.focus_set()

        if not self.is_logged_in:
            messagebox.showerror("错误", "请先登录")
            return

        if not self.current_mmo_code:
            messagebox.showwarning("警告", "请先选择工单")
            return

        # 播放扫描提示音
        self._play_scan_sound(True)

        if self.is_offline_mode:
            # 离线模式：保存到本地
            self._handle_offline_scan(label)
        else:
            # 在线模式：直接调用 API
            self._handle_online_scan(label)

    def _handle_offline_scan(self, label: str):
        """处理离线扫码"""
        self.db.add_scan_record(
            device_code=self.current_device_code,
            mmo_code=self.current_mmo_code,
            label=label
        )
        self.scan_count += 1
        self.scan_count_label.config(text=str(self.scan_count))
        self.scan_status_label.config(text=f"✓ {label} (离线)", foreground="blue")
        self._log(f"[离线] 扫码：{label}")
        self._update_offline_count()
        self._play_scan_sound(True)

    def _handle_online_scan(self, label: str):
        """处理在线扫码"""
        self.start_scan_btn.config(state=tk.DISABLED)

        def scan_thread():
            try:
                labels_data = [{"label": label, "qty": 1, "result": 10}]
                result = self.mes_client.crossing(
                    device_code=self.current_device_code,
                    mmo_code=self.current_mmo_code,
                    labels=labels_data
                )
                self.root.after(0, lambda: self._handle_scan_result(result, label))
            except Exception as e:
                self.root.after(0, lambda: self._handle_scan_result({"success": False, "msg": str(e)}, label))

        threading.Thread(target=scan_thread, daemon=True).start()

    def _handle_scan_result(self, result: Dict, label: str):
        """处理扫码结果"""
        self.start_scan_btn.config(state=tk.NORMAL)

        # 保持焦点在扫码框
        if self.start_scan_btn.winfo_exists():
            self.scan_entry.focus_set()

        if result["success"]:
            self.scan_count += 1
            self.scan_count_label.config(text=str(self.scan_count))
            self.scan_status_label.config(text=f"✓ {label}", foreground="green")
            self._log(f"[在线] 扫码成功：{label}")
            self._play_scan_sound(True)
        else:
            self.scan_status_label.config(text=f"✗ {label}", foreground="red")
            self._log(f"[在线] 扫码失败：{label} - {result.get('msg', '未知错误')}", "ERROR")
            self._play_scan_sound(False)

            # 弹窗提示
            messagebox.showerror("扫码失败", f"条码：{label}\n错误：{result.get('msg', '未知错误')}")

            # 停止扫码
            self.start_scan_btn.config(text="开始扫码")
            self.scan_status_label.config(text="○ 已停止（扫码失败）", foreground="red")

    def _play_scan_sound(self, success: bool):
        """播放扫码提示音"""
        if success:
            # 成功提示音（高频短音）
            winsound.Beep(800, 100)
        else:
            # 失败提示音（低频长音）
            winsound.Beep(400, 300)

    def _upload_offline_data(self):
        """上传离线数据"""
        if not self.is_logged_in:
            messagebox.showerror("错误", "请先登录")
            return

        records = self.db.get_unuploaded_records()
        if not records:
            messagebox.showinfo("提示", "没有待上传的数据")
            return

        self._log(f"开始上传 {len(records)} 条离线数据...")

        success_count = 0
        failed_records = []

        def upload_thread():
            nonlocal success_count, failed_records

            for record in records:
                try:
                    labels_data = [{
                        "label": record["label"],
                        "qty": record["qty"],
                        "result": record["result"]
                    }]
                    result = self.mes_client.crossing(
                        device_code=record["device_code"],
                        mmo_code=record["mmo_code"],
                        labels=labels_data
                    )

                    if result["success"]:
                        success_count += 1
                        self.db.mark_as_uploaded([record["id"]])
                        self.root.after(0, lambda r=record: self._log(f"上传成功：{r['label']}"))
                    else:
                        failed_records.append(record)
                        self.root.after(0, lambda r=record, m=result.get('msg'): self._log(f"上传失败：{r['label']} - {m}", "ERROR"))
                except Exception as e:
                    failed_records.append(record)
                    self.root.after(0, lambda r=record, e=str(e): self._log(f"上传异常：{r['label']} - {e}", "ERROR"))

            # 更新 UI
            self.root.after(0, lambda: self._finish_upload(success_count, len(records) - success_count))

        threading.Thread(target=upload_thread, daemon=True).start()

    def _finish_upload(self, success: int, failed: int):
        """完成上传"""
        self._update_offline_count()
        if failed == 0:
            messagebox.showinfo("上传完成", f"成功上传 {success} 条数据")
            self._log(f"离线数据上传完成，成功{success}条")
        else:
            messagebox.showwarning("上传完成", f"成功：{success}条\n失败：{failed}条")
            self._log(f"离线数据上传完成，成功{success}条，失败{failed}条", "WARN")

    def _export_records(self):
        """导出扫码记录"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV 文件", "*.csv"), ("所有文件", "*.*")],
            title="导出记录"
        )

        if not file_path:
            return

        records = self.db.get_all_records()
        if not records:
            messagebox.showinfo("提示", "没有可导出的记录")
            return

        try:
            with open(file_path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                writer.writerow(['ID', '设备编码', '工单编码', '条码', '数量', '结果', '扫码时间', '已上传'])
                for r in records:
                    writer.writerow([
                        r['id'],
                        r['device_code'],
                        r['mmo_code'],
                        r['label'],
                        r['qty'],
                        'OK' if r['result'] == 10 else 'NG',
                        r['scanned_time'],
                        '是' if r['uploaded'] == 1 else '否'
                    ])

            messagebox.showinfo("导出成功", f"已导出 {len(records)} 条记录到:\n{file_path}")
            self._log(f"导出记录：{len(records)}条 -> {file_path}")
        except Exception as e:
            messagebox.showerror("导出失败", str(e))
            self._log(f"导出失败：{str(e)}", "ERROR")

    def _do_logout(self):
        """执行登出"""
        if messagebox.askyesno("确认", "确定要登出吗？"):
            self.auto_prolong_running = False
            self.is_logged_in = False
            self.mes_client.logout()
            self.scan_count = 0

            # 重置界面
            self.main_frame.pack_forget()
            self.login_frame.pack(fill=tk.BOTH, expand=True)

            # 清空状态
            self.login_status_label.config(text="")
            self.scan_status_label.config(text="")
            self.scan_count_label.config(text="0")
            self.mmo_combo["values"] = []

            self._log("用户已登出")

    def on_closing(self):
        """窗口关闭处理"""
        if self.is_logged_in:
            if messagebox.askyesno("确认", "确定要退出系统吗？"):
                self.auto_prolong_running = False
                self.root.destroy()
        else:
            self.root.destroy()


# ==================== 主程序入口 ====================
def main():
    root = tk.Tk()
    app = MESTerminalApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()


if __name__ == "__main__":
    main()
