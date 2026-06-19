import os
import sys
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, filedialog
import customtkinter as ctk
import pandas as pd
from dateutil.relativedelta import relativedelta
from tkinterdnd2 import TkinterDnD, DND_FILES

# ライトモード設定
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

# 費目変換マップ
CONVERSION_MAP = {
    "授業料": "01", "在籍料": "02", "抗体検査費": "04", "聴講料": "07", "入学金": "08",
    "登録料": "09", "実験実習教育研究費": "10", "調理学実習費": "13", "調理実習費 ": "14",
    "製菓実習費": "15", "給食管理実習": "16", "集団給食調理実習費": "17", "検定料": "20",
    "若葉寮費": "21", "家庭料理技能検定料": "22", "施設費": "23", "ﾃｸﾆｯｸ実習費": "24",
    "実践調理学実習(実)": "25", "食文化調理学実習Ⅰ": "26", "食文化調理学実習Ⅲ": "27",
    "ﾜｲﾝｺｰﾃﾞｨﾈｰﾄ論実習Ⅰ": "28", "介護食士3級申請料等": "29", "健康診断・細菌検査費": "30",
    "卒業関係経費": "31", "校外調理研修費": "32", "校外製菓研修費": "33", "専門調理学実習(科)": "34",
    "保険料": "35", "応用調理学実習(養)": "36", "学用品費": "37", "食文化調理学実習Ⅱ": "38",
    "外食ﾒﾆｭｰ開発実習": "39", "ｶﾌｪﾚｽﾄﾗﾝ実習": "40", "ﾌｰﾄﾞｺｰﾃﾞｨネオート論実習": "41",
    "ﾜｲﾝｺｰﾃﾞｨﾈｰﾄ論実習Ⅱ": "42", "香友会入会費": "43", "専門調理実習(短)": "44",
    "奨学費": "45", "横巻のぶ奨学金": "46", "AL特待生": "47", "学生会会費": "48",
    "受講料(履修証明プログラム)": "56", "授業料特別減免措置": "57", "大学院修士課程特別奨学生": "58",
    "香川調理製菓専門学校特待生": "59", "栄大スカラシップ制度": "60", "北郁子奨学基金奨学金": "61",
    "浅野嘉久賞奨学金": "62", "香川綾・芳子奨励賞": "63", "岡本萌実記念奨学金": "64",
    "野口医学研究所奨学金": "65", "荒井慶子ｸﾞﾛｰﾊﾞﾙ人材育成奨学金": "66", "私費外国人留学生奨学金制度": "67",
    "食育ｲﾝｽﾄラクタ―受験料等": "77", "日本語学校費用": "78", "教職員子女減免": "81",
    "修学支援新制度": "83", "保護者会費": "85", "研修親睦・入卒費": "88",
    "資格申請・受験・検定費": "89", "預り金": "90", "仮受金収入": "91"
}

class App:
    def __init__(self):
        self.root = TkinterDnD.Tk()
        self.root.title("CPからedufeeへ変換")
        self.root.geometry("720x325") 
        self.root.resizable(False, False)
        self.root.configure(bg="#fbfbfb") 
        self.file_path = ""

        # --- ① 上部タブ ---
        self.tab_frame = tk.Frame(self.root, bg="#f0f0f0", height=40)
        self.tab_frame.pack(fill="x", side="top")
        self.tab_frame.pack_propagate(False)

        self.tab_buttons = {}
        tabs = ["新規・学年変換", "確認用出力", "徴収情報変換", "分納情報変換"]
        
        for tab in tabs:
            is_active = (tab == "徴収情報変換")
            btn = tk.Button(
                self.tab_frame, text=tab, font=("メイリオ", 11, "bold" if is_active else "normal"),
                bg="#0066cc" if is_active else "#f0f0f0", fg="#ffffff" if is_active else "#000000",
                relief="flat", bd=0, width=16, activebackground="#0055aa", activeforeground="#ffffff",
                command=lambda t=tab: self.change_tab(t)
            )
            btn.pack(side="left", fill="y")
            self.tab_buttons[tab] = btn

        # メインコンテナ
        self.main_container = tk.Frame(self.root, bg="#fbfbfb")
        self.main_container.pack(fill="both", expand=True, padx=20, pady=10)

        # --- ② 1. データ読み込みエリア ---
        self.group1_label = tk.Label(self.main_container, text="1. データ読み込み", font=("メイリオ", 12, "bold"), bg="#fbfbfb", fg="#000000")
        self.group1_label.pack(anchor="w", pady=(5, 5))

        self.drop_canvas = tk.Canvas(self.main_container, bg="#f5f7f8", highlightthickness=0, height=95)
        self.drop_canvas.pack(fill="x", pady=2)
        self.drop_canvas.bind("<Configure>", lambda e: self.draw_canvas_border())
        
        self.icon_label = tk.Label(self.drop_canvas, text="[CSV]", font=("Arial", 14, "bold"), fg="#2ea44f", bg="#f5f7f8")
        self.icon_id = self.drop_canvas.create_window(340, 25, window=self.icon_label)
        
        self.text_id = self.drop_canvas.create_text(
            340, 65,
