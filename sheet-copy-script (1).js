import tkinter as tk
from tkinter import messagebox, filedialog
import pandas as pd
import os
from openpyxl import load_workbook

# メインウィンドウの作成
root = tk.Tk()
root.title("Excelファイル選択")
root.geometry("500x300")

# 選択されたファイルパスを保存する変数
selected_file_path = tk.StringVar()

# 設定ファイルの情報
config_excel = "C:/path/to/config.xlsx"  # 設定ファイルのパス
sheet_name = "Sheet1"                    # シート名
cell = "A1"                             # ディレクトリパスが書かれているセル

# データフレームを格納する変数
df = None

def select_file():
    try:
        # 初期ディレクトリの取得
        wb = load_workbook(config_excel)
        ws = wb[sheet_name]
        initial_dir = ws[cell].value
        wb.close()
    except:
        initial_dir = os.path.expanduser("~")
    
    # ファイル選択ダイアログを表示
    file_path = filedialog.askopenfilename(
        initialdir=initial_dir,
        title="Excelファイルを選択してください",
        filetypes=(("Excelファイル", "*.xlsx *.xls"),)
    )
    
    if file_path:
        selected_file_path.set(file_path)

def read_excel_data():
    global df
    file_path = selected_file_path.get()
    
    if not file_path:
        messagebox.showwarning("警告", "ファイルが選択されていません")
        return
    
    try:
        # まず列名を取得
        temp_df = pd.read_excel(
            file_path,
            sheet_name='Sheet1',
            header=1,
            usecols='E:AA',
            nrows=0  # ヘッダーだけ読む
        )
        columns = temp_df.columns.tolist()
        
        # H列、V列、Z列のインデックスを取得
        h_col = columns[3]  # E列から数えて4番目（H列）の列名
        v_col = columns[17]  # E列から数えて18番目（V列）の列名
        z_col = columns[21]  # E列から数えて22番目（Z列）の列名
        
        # Excelファイルを読み込む
        df = pd.read_excel(
            file_path,
            sheet_name='Sheet1',
            header=1,
            usecols='E:AA',
            engine='openpyxl',
            thousands=',',  # 3桁区切りのカンマを解釈
            converters={
                h_col: lambda x: pd.to_datetime(x, format='%Y年%m月'),  # H列の日付変換
                v_col: lambda x: float(str(x).replace(',', '')) if pd.notna(x) else None,  # V列の金額変換
                z_col: lambda x: float(str(x).replace(',', '')) if pd.notna(x) else None   # Z列の金額変換
            }
        )
        
        # 読み込んだデータの型を確認
        dtypes_info = df.dtypes.to_string()
        
        # データサンプルを作成
        sample_data = f"\n\n日付列(H列)サンプル:\n{df[h_col].head()}"
        sample_data += f"\n\n金額列(V列)サンプル:\n{df[v_col].head()}"
        sample_data += f"\n\n金額列(Z列)サンプル:\n{df[z_col].head()}"
        
        # 読み込み成功のメッセージを表示
        rows, cols = df.shape
        messagebox.showinfo(
            "成功",
            f"データを読み込みました\n"
            f"行数: {rows}\n"
            f"列数: {cols}\n"
            f"列名と型:\n{dtypes_info}\n"
            f"{sample_data}"
        )
        
        # ウィンドウを閉じる
        root.quit()
        root.destroy()
        
    except Exception as e:
        messagebox.showerror("エラー", f"データの読み込みに失敗しました: {str(e)}")

# メインフレームの作成
frame = tk.Frame(root, padx=20, pady=20)
frame.pack(fill=tk.BOTH, expand=True)

# タイトルラベル
title_label = tk.Label(
    frame, 
    text="Excelファイルを選択してください", 
    font=("Helvetica", 12, "bold"),
    pady=10
)
title_label.pack()

# ファイル選択ボタン
select_button = tk.Button(
    frame,
    text="Excelファイル選択",
    command=select_file,
    width=20,
    height=2
)
select_button.pack(pady=10)

# 選択されたファイルパスを表示するラベル
path_label = tk.Label(
    frame,
    textvariable=selected_file_path,
    wraplength=400,
    pady=10
)
path_label.pack()

# データ読み込みボタン
read_button = tk.Button(
    frame,
    text="データを読み込む",
    command=read_excel_data,
    width=20,
    height=2
)
read_button.pack(pady=10)

# メインループの開始
root.mainloop()

# メインループが終了した後のデータ処理
if df is not None:
    # H列（日付）の確認
    date_col = df.columns[3]  # E列から数えて4番目（H列）
    print(f"\n日付列（{date_col}）のサンプル:")
    print(df[date_col].head())
    print(f"\n日付列の型: {df[date_col].dtype}")
    
    # V列とZ列（金額）の確認
    amount_col_1 = df.columns[17]  # E列から数えて18番目（V列）
    amount_col_2 = df.columns[21]  # E列から数えて22番目（Z列）
    print(f"\n金額列（{amount_col_1}）のサンプル:")
    print(df[amount_col_1].head())
    print(f"\n金額列（{amount_col_2}）のサンプル:")
    print(df[amount_col_2].head())
    
    # 基本統計量の確認
    print("\n金額列の基本統計量:")
    print(df[[amount_col_1, amount_col_2]].describe())
