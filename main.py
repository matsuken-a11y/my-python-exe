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

# 費目変換マップ（徴収情報変換用）
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
        self.current_tab = "徴収情報変換"  # 初期タブ

        # --- ① 上部タブ ---
        self.tab_frame = tk.Frame(self.root, bg="#f0f0f0", height=40)
        self.tab_frame.pack(fill="x", side="top")
        self.tab_frame.pack_propagate(False)

        self.tab_buttons = {}
        tabs = ["新規・学年変換", "確認用出力", "徴収情報変換", "分納情報変換"]
        
        for tab in tabs:
            is_active = (tab == self.current_tab)
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

        self.run_btn = tk.Button(self.control_frame, text="📄   変換実行", font=("メイリオ", 13, "bold"), bg="#22863a", fg="#ffffff", relief="raised", bd=1, activebackground="#1b5e20", activeforeground="#ffffff", command=self.process_data)
        self.run_btn.pack(fill="both", expand=True, padx=2, pady=2, ipady=8)

    def change_tab(self, selected_tab):
        if selected_tab in ["徴収情報変換", "確認用出力"]:
            self.current_tab = selected_tab
            for tab_name, btn in self.tab_buttons.items():
                if tab_name == selected_tab:
                    btn.configure(bg="#0066cc", fg="#ffffff", font=("メイリオ", 11, "bold"))
                else:
                    btn.configure(bg="#f0f0f0", fg="#000000", font=("メイリオ", 11, "normal"))
            
            # ラベルテキストの動的切り替え
            if selected_tab == "確認用出力":
                self.group2_label.configure(text="2. 確認用ファイル生成")
            else:
                self.group2_label.configure(text="2. アップロード用ファイル生成")
        else:
            messagebox.showinfo("ご案内", f"「{selected_tab}」機能は現在システム準備中です。\nデータ変換は「徴収情報変換」または「確認用出力」タブで行ってください。")

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
        
        # タブの状態に応じて呼び出すロジックを分岐
        if self.current_tab == "徴収情報変換":
            self.process_collection_data()
        elif self.current_tab == "確認用出力":
            self.process_verification_data()

    # 1. 徴収情報変換（edufee用）ロジック
    def process_collection_data(self):
        try:
            offset_val = -1
            df_src = self.load_source_file()
            if df_src is None: raise ValueError("ファイルの読み込みに失敗しました。")

            df_dest = pd.DataFrame(None, index=range(len(df_src)), columns=range(36))
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
                        calc_dt = dt - relativedelta(months=abs(offset_val)*2)
                        df_dest.at[idx, 7] = pd.Timestamp(calc_dt)
                        df_dest.at[idx, 9] = (dt + relativedelta(years=1)).strftime("%Y/%m/%d 23:59")
                    else:
                        df_dest.at[idx, 8] = df_dest.at[idx, 7] = df_dest.at[idx, 9] = None

            df_dest[10] = df_dest[0].astype(str).map(lambda x: f"000{x.split('.')[0]}" if x and x != "nan" else None)

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

            df_final = df_dest.iloc[1:].copy()
            df_final.columns = excel_headers

            for i in range(1, 11):
                code_col = f"明細コード{i}"
                amt_col = f"明細金額{i}"
                if code_col in df_final.columns and amt_col in df_final.columns:
                    mask = df_final[code_col].notna() & (df_final[code_col].astype(str).str.strip() != "") & (df_final[amt_col].isna() | (df_final[amt_col].astype(str).str.strip() == ""))
                    df_final.loc[mask, amt_col] = 0

            numeric_cols = ["学籍番号", "年度", "徴収金額", "明細金額1", "明細金額2", "明細金額3", "明細金額4", "明細金額5", "明細金額6", "明細金額7", "明細金額8", "明細金額9", "明細金額10"]
            for col in numeric_cols:
                df_final[col] = pd.to_numeric(df_final[col].astype(str).str.strip(), errors='coerce')

            string_cols = ["整理番号", "徴収種別コード", "徴収名目コード", "明細コード1", "明細コード2", "明細コード3", "明細コード4", "明細コード5", "明細コード6", "明細コード7", "明細コード8", "明細コード9", "明細コード10"]
            for col in string_cols:
                df_final[col] = df_final[col].map(lambda x: str(x).strip() if pd.notna(x) and str(x).strip() != "None" and str(x).strip() != "" else None)

            save_path = self.get_save_path("_データ")
            if not save_path: return

            with pd.ExcelWriter(save_path, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False, header=True)
                workbook = writer.book
                worksheet = writer.sheets['Sheet1']
                string_format = workbook.add_format({'num_format': '@'})
                datetime_format = workbook.add_format({'num_format': 'yyyy/mm/dd hh:mm:ss'})
                
                for col_name in string_cols:
                    if col_name in excel_headers:
                        col_idx = excel_headers.index(col_name)
                        for row_idx in range(1, len(df_final) + 1):
                            val = df_final.iloc[row_idx - 1][col_name]
                            if pd.notna(val): worksheet.write(row_idx, col_idx, str(val), string_format)

                if "徴収開始日" in excel_headers:
                    h_col_idx = excel_headers.index("徴収開始日")
                    for row_idx in range(1, len(df_final) + 1):
                        val = df_final.iloc[row_idx - 1]["徴収開始日"]
                        if pd.notna(val): worksheet.write_datetime(row_idx, h_col_idx, val, datetime_format)

            messagebox.showinfo("保存完了", f"ファイルを保存しました！\n\n{os.path.basename(save_path)}")
        except Exception as e:
            messagebox.showerror("エラー", f"処理中にエラーが発生しました:\n{str(e)}")

    # 2. 確認用出力ロジック (VBAマクロを完全再現)
    def process_verification_data(self):
        try:
            df_src = self.load_source_file()
            if df_src is None: raise ValueError("ファイルの読み込みに失敗しました。")

            # 1行目をヘッダーとして設定
            df_src.columns = df_src.iloc[0]
            df_data = df_src.iloc[1:].copy()

            # ① 2行目以降でW列（学生納付情報.消込日）が空白でない（値がある）行を削除
            # インデックスによる指定安全策：W列は22番目（0始まり）
            clear_day_col = df_data.columns[22] 
            df_data = df_data[df_data[clear_day_col].isna() | (df_data[clear_day_col].astype(str).str.strip() == "")]

            # ② AA列(26)からCJ列(87)までの空白列を自動削除するロジック
            # まず、AA列以降の費目列のリストを取得
            fee_cols = list(df_data.columns[26:])
            active_fee_cols = []
            for col in fee_cols:
                # 列が全て空欄、または0、または空文字列でないかチェック
                series = df_data[col].dropna().astype(str).str.strip()
                series = series[series != ""]
                if not series.empty:
                    active_fee_cols.append(col)

            # ③ 固定削除する列（アルファベット指定を列名にマッピング）
            # マクロのdeleteCols = Array("Z", "Y", "X", "W", "V", "U", "T", "S", "N", "M", "L", "K", "J", "I", "G", "D", "B")
            # 0始まりのインデックス: B(1), D(3), G(6), I(8), J(9), K(10), L(11), M(12), N(13), S(18), T(19), U(20), V(21), W(22), X(23), Y(24), Z(25)
            all_cols = list(df_src.columns)
            indices_to_delete = [1, 3, 6, 8, 9, 10, 11, 12, 13, 18, 19, 20, 21, 22, 23, 24, 25]
            delete_names = [all_cols[i] for i in indices_to_delete if i < len(all_cols)]
            
            # 残すべき基本情報列を抽出
            # 残るのは：A(学籍番号), C(氏名), E(在学区分名称), F(所属コード), H(学年１) などのうち削除から外れたもの
            base_cols = [col for idx, col in enumerate(df_src.columns[:26]) if idx not in indices_to_delete]
            
            # 最終的に残す列をドッキング
            final_cols = base_cols + active_fee_cols
            df_result = df_data[final_cols].copy()

            # ④ セルに指定の値を入力（ヘッダーの上書き）
            # マクロの対応関係：C1="区分", E1="学年", F1="納付回数", G1="納付金額", H1="納期", I1="延納期限"
            # 列追加・削除後の位置に合わせ、元ヘッダー名ベースで確実に置換
            rename_map = {
                "在学区分名称": "区分",
                "学年１": "学年",
                "学生納付情報.納付回数": "納付回数",
                "学生納付情報.納付金額": "納付金額",
                "学生納付情報.納期": "納期",
                "学生納付情報.延納期限": "延納期限"
            }
            df_result.rename(columns=rename_map, inplace=True)

            # ⑤ 独自の学生数集計：B列（学籍番号）の重複を除いた数をカウント
            unique_student_count = df_result["学籍番号"].nunique()

            # ⑥ A列を追加して連番（No）を入力
            df_result.insert(0, "No", range(1, len(df_result) + 1))

            # 保存用画面を起動
            save_path = self.get_save_path("_確認用出力")
            if not save_path: return

            # ⑦ xlsxwriterによるデザインの適用（格子・水色ヘッダー・カンマ・幅調整）
            with pd.ExcelWriter(save_path, engine='xlsxwriter') as writer:
                df_result.to_excel(writer, index=False, header=True)
                
                workbook = writer.book
                worksheet = writer.sheets['Sheet1']
                
                # デザイン定義
                header_format = workbook.add_format({
                    'bg_color': '#E0FFFF',  # 薄い水色 (RGB: 224, 255, 255)
                    'border': 1,
                    'bold': True,
                    'align': 'center',
                    'valign': 'vcenter'
                })
                data_format = workbook.add_format({'border': 1})
                comma_format = workbook.add_format({'border': 1, 'num_format': '#,##0'})
                string_format = workbook.add_format({'border': 1, 'num_format': '@'})

                # 全セルに格子を追加しつつ書式適用
                headers = list(df_result.columns)
                for col_idx, col_name in enumerate(headers):
                    # ヘッダーを水色に変更
                    worksheet.write(0, col_idx, col_name, header_format)
                    
                    # データの書き込みと書式化
                    for row_idx in range(1, len(df_result) + 1):
                        val = df_result.iloc[row_idx - 1][col_name]
                        
                        # 金額を扱う列はカンマ区切りにする (G列相当＝納付金額、および費目列)
                        if col_name == "納付金額" or col_name in active_fee_cols:
                            if pd.notna(val) and str(val).strip() != "":
                                try:
                                    worksheet.write_number(row_idx, col_idx, float(val), comma_format)
                                except ValueError:
                                    worksheet.write(row_idx, col_idx, val, data_format)
                            else:
                                worksheet.write(row_idx, col_idx, "", data_format)
                        # 学籍番号などは文字列として固定
                        elif col_name == "学籍番号":
                            worksheet.write(row_idx, col_idx, str(val) if pd.notna(val) else "", string_format)
                        else:
                            if pd.notna(val):
                                worksheet.write(row_idx, col_idx, val, data_format)
                            else:
                                worksheet.write(row_idx, col_idx, "", data_format)

                # 重複を除いた学生数を、C列の最終行の2行下に書き込み
                # NoがA列(0), 学籍番号がB列(1), 氏名がC列(2)
                summary_row = len(df_result) + 2
                worksheet.write(summary_row, 2, unique_student_count, data_format)

                # 列幅の自動調整 & J列(9)以降は幅12に固定
                for col_idx, col_name in enumerate(headers):
                    if col_idx >= 9:  # J列以降
                        worksheet.set_column(col_idx, col_idx, 12)
                    else:
                        max_len = max(
                            df_result[col_name].astype(str).map(len).max() if not df_result.empty else 0,
                            len(col_name)
                        ) + 3
                        worksheet.set_column(col_idx, col_idx, min(max_len, 30))

            messagebox.showinfo("保存完了", f"確認用出力ファイルの生成が完了しました！\n\n{os.path.basename(save_path)}")
        except Exception as e:
            messagebox.showerror("エラー", f"処理中にエラーが発生しました:\n{str(e)}")

    # 共通ファイル読み込み処理
    def load_source_file(self):
        df = None
        try:
            df = pd.read_csv(self.file_path, header=None, encoding="cp932")
        except Exception:
            try:
                df = pd.read_excel(self.file_path, header=None)
            except Exception:
                df = pd.read_csv(self.file_path, header=None, encoding="utf-8")
        return df

    # 共通「名前を付けて保存」画面処理
    def get_save_path(self, suffix):
        home_dir = os.path.expanduser("~")
        desktop_path = os.path.join(home_dir, "OneDrive", "デスクトップ")
        if not os.path.exists(desktop_path): desktop_path = os.path.join(home_dir, "Desktop")

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f"{timestamp}{suffix}.xlsx"
        
        return filedialog.asksaveasfilename(
            initialdir=desktop_path, initialfile=default_filename,
            defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")]
        )

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = App()
    app.run()
