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
    "外食ﾒﾆｭｰ開発実習": "39", "ｶﾌｪﾚｽﾄﾗﾝ実習": "40", "ﾌｰﾄﾞｺｰﾃﾞｨﾈｰﾄ論実習": "41",
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
            340, 65, text="元データExcel / CSVファイルをここにドラッグ＆ドロップ\n(またはここをクリックしてファイルを選択)",
            font=("メイリオ", 11, "bold"), fill="#333333", justify="center"
        )
        
        self.drop_canvas.bind("<Button-1>", lambda e: self.browse_file())
        self.drop_canvas.drop_target_register(DND_FILES)
        self.drop_canvas.dnd_bind('<<Drop>>', self.handle_drop)

        self.status_frame = tk.Frame(self.main_container, bg="#fbfbfb")
        self.status_frame.pack(fill="x", pady=5)
        
        self.check_icon = tk.Label(self.status_frame, text="✔", font=("Arial", 12, "bold"), fg="#999999", bg="#fbfbfb")
        self.check_icon.pack(side="left", padx=(5, 2))
        
        self.path_label = tk.Label(self.status_frame, text="警告: ファイルが選択されていません。", font=("メイリオ", 10), fg="#666666", bg="#fbfbfb")
        self.path_label.pack(side="left")

        # --- ③ 2. アップロード用ファイル生成エリア ---
        self.group2_label = tk.Label(self.main_container, text="2. アップロード用ファイル生成", font=("メイリオ", 12, "bold"), bg="#fbfbfb", fg="#000000")
        self.group2_label.pack(anchor="w", pady=(5, 5))

        self.control_frame = tk.Frame(self.main_container, bg="#fbfbfb")
        self.control_frame.pack(fill="x", pady=2)

        self.run_btn = tk.Button(self.control_frame, text="📄  変換実行", font=("メイリオ", 13, "bold"), bg="#22863a", fg="#ffffff", relief="raised", bd=1, activebackground="#1b5e20", activeforeground="#ffffff", command=self.process_data)
        self.run_btn.pack(fill="both", expand=True, padx=2, pady=2, ipady=8)

    def change_tab(self, selected_tab):
        for tab_name, btn in self.tab_buttons.items():
            if tab_name == selected_tab:
                btn.configure(bg="#0066cc", fg="#ffffff", font=("メイリオ", 11, "bold"))
            else:
                btn.configure(bg="#f0f0f0", fg="#000000", font=("メイリオ", 11, "normal"))
        if selected_tab != "徴収情報変換":
            messagebox.showinfo("ご案内", f"「{selected_tab}」機能は現在システム準備中です。\nデータ変換は「徴収情報変換」タブで行ってください。")
            self.change_tab("徴収情報変換")

    def draw_canvas_border(self):
        self.drop_canvas.delete("border")
        w = self.drop_canvas.winfo_width()
        h = self.drop_canvas.winfo_height()
        self.drop_canvas.create_rectangle(2, 2, w-2, h-2, dash=(4, 4), outline="#999999", width=1, tags="border")
        self.drop_canvas.coords(self.icon_id, w/2, 25)
        self.drop_canvas.coords(self.text_id, w/2, 65)

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel/CSV Files", "*.xlsx *.xls *.csv")])
        if file_path: self.set_file(file_path)

    def handle_drop(self, event):
        file_path = event.data.strip('{}')
        self.set_file(file_path)

    def set_file(self, path):
        self.file_path = path
        filename = os.path.basename(path)
        self.path_label.configure(text=f"ファイルの読み込みが完了しました。 (ファイル名: {filename})", fg="#2ea44f")
        self.check_icon.configure(text="✔", fg="#2ea44f")

    def process_data(self):
        if not self.file_path:
            messagebox.showwarning("警告", "ファイルを選択してください。")
            return
        try:
            offset_val = -1
            df_src = None
            try:
                df_src = pd.read_csv(self.file_path, header=None, encoding="cp932")
            except Exception:
                try:
                    df_src = pd.read_excel(self.file_path, header=None)
                except Exception:
                    df_src = pd.read_csv(self.file_path, header=None, encoding="utf-8")

            if df_src is None: raise ValueError("ファイルの読み込みに失敗しました。")

            df_dest = pd.DataFrame("", index=range(len(df_src)), columns=range(36))
            df_dest[0] = df_src[0]
            if len(df_src.columns) > 13: df_dest[1] = df_src[13]
            df_dest[2] = "01"
            
            if len(df_src.columns) > 14:
                df_dest[3] = df_src[14].astype(str).str.replace(r'\.0$', '', regex=True)
                val_map = {"1": "01", "3": "03", "11": "11", "8": "08"}
                df_dest[3] = df_dest[3].map(lambda x: val_map.get(x, x))
            
            if len(df_src.columns) > 15: df_dest[6] = df_src[15]

            if len(df_src.columns) > 16:
                for idx, val in df_src[16].items():
                    val_str = str(val).split('.')[0].strip()
                    if val_str.isdigit() and len(val_str) == 8:
                        dt = datetime.strptime(val_str, "%Y%m%d")
                        df_dest.at[idx, 8] = dt.strftime("%Y/%m/%d 23:59")
                        # 🕒 H列(徴収開始日)のフォーマットを「秒」単位まで表示するように修正しました
                        # 時・分・秒を「 00:00:01」に完全固定する場合
                        df_dest.at[idx, 7] = (dt - relativedelta(months=abs(offset_val)*2)).strftime("%Y/%m/%d 00:00:01")
                        df_dest.at[idx, 9] = (dt + relativedelta(years=1)).strftime("%Y/%m/%d 23:59")
                    else:
                        df_dest.at[idx, 8] = df_dest.at[idx, 7] = df_dest.at[idx, 9] = ""

            df_dest[10] = df_dest[0].astype(str).map(lambda x: f"000{x.split('.')[0]}" if x and x != "nan" else "")

            headers_src = df_src.iloc[0].tolist()
            target_cols = [17, 19, 21, 23, 25, 27, 29, 31, 33, 35]
            header_cols = [16, 18, 20, 22, 24, 26, 28, 30, 32, 34]

            for idx in range(1, len(df_src)):
                current_col_idx = 0
                for col_idx in range(26, min(90, len(df_src.columns))):
                    val = df_src.iat[idx, col_idx]
                    if pd.notna(val) and str(val).strip() != "":
                        if current_col_idx < len(target_cols):
                            df_dest.iat[idx, target_cols[current_col_idx]] = val
                            src_header_name = str(headers_src[col_idx]).strip()
                            df_dest.iat[idx, header_cols[current_col_idx]] = CONVERSION_MAP.get(src_header_name, src_header_name)
                            current_col_idx += 1
                if current_col_idx >= len(target_cols): break

            excel_headers = [
                "学籍番号", "年度", "徴収種別コード", "徴収名目コード", "説明文", "説明文（英字）", 
                "徴収金額", "徴収開始日", "徴収期限", "表示期限", "整理番号", 
                "システム連携用Key1", "システム連携用Key2", "システム連携用Key3", "システム連携用Key4", "システム連携用Key5", 
                "明細コード1", "明細金額1", "明細コード2", "明細金額2", "明細コード3", "明細金額3", 
                "明細コード4", "明細金額4", "明細コード5", "明細金額5", "明細コード6", "明細金額6", 
                "明細コード7", "明細金額7", "明細コード8", "明細金額8", "明細コード9", "明細金額9", 
                "明細コード10", "明細金額10"
            ]

            # 不要な1行目をカットし、正規のヘッダーを設定
            df_final = df_dest.iloc[1:].copy()
            df_final.columns = excel_headers

            # 📊 ご指定の列を数値型に変換
            numeric_cols = ["学籍番号", "年度", "徴収金額", "明細金額1", "明細金額2", "明細金額3", "明細金額4", "明細金額5", "明細金額6", "明細金額7", "明細金額8", "明細金額9", "明細金額10"]
            for col in numeric_cols:
                df_final[col] = pd.to_numeric(df_final[col].astype(str).str.strip(), errors='coerce')

            # 🔒 コード類にプラスして「整理番号(K列)」も完全に文字列として固定
            string_cols = ["整理番号", "徴収種別コード", "徴収名目コード", "明細コード1", "明細コード2", "明細コード3", "明細コード4", "明細コード5", "明細コード6", "明細コード7", "明細コード8", "明細コード9", "明細コード10"]
            for col in string_cols:
                df_final[col] = df_final[col].astype(str).str.strip()

            home_dir = os.path.expanduser("~")
            desktop_path = os.path.join(home_dir, "OneDrive", "デスクトップ")
            if not os.path.exists(desktop_path): desktop_path = os.path.join(home_dir, "Desktop")

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            default_filename = f"{timestamp}_データ.xlsx"
            
            save_path = filedialog.asksaveasfilename(
                initialdir=desktop_path, initialfile=default_filename,
                defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")]
            )
            if not save_path: return

            # 🛠️ 指定した列にExcelの書式（文字列・日付）を適用
            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, header=True)
                
                workbook = writer.book
                worksheet = writer.sheets['Sheet1']
                
            # コード類の列を「文字列(@)」に固定
                for col_name in string_cols:
                    if col_name in excel_headers:
                        col_idx = excel_headers.index(col_name) + 1
                        for row in range(2, worksheet.max_row + 1):
                            cell = worksheet.cell(row=row, column=col_idx)
                            cell.number_format = '@'

             # 📅 徴収開始日(H列)を、ユーザー定義ではなく「日付(yyyy/mm/dd)」の書式に完全固定
                if "徴収開始日" in excel_headers:
                    h_col_idx = excel_headers.index("徴収開始日") + 1
                    for row in range(2, worksheet.max_row + 1):
                        cell = worksheet.cell(row=row, column=h_col_idx)
                        cell.number_format = 'yyyy/mm/dd'  # 標準の日付書式を指定

            messagebox.showinfo("保存完了", f"ファイルを保存しました！\n\n{os.path.basename(save_path)}")
        except Exception as e:
            messagebox.showerror("エラー", f"処理中にエラーが発生しました:\n{str(e)}")

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = App()
    app.run()
