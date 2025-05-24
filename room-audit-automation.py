import pandas as pd
import re
import pyperclip
import os
import shutil
import tkinter as tk
from tkinter import messagebox
import xlwings as xw
from datetime import datetime
import sys

def get_base_path():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

def write_log(message):
    base_path = get_base_path()
    log_dir = os.path.join(base_path, "logs")
    os.makedirs(log_dir, exist_ok=True)
    date_str = datetime.now().strftime("%Y%m%d")
    log_file = os.path.join(log_dir, f"実行履歴_{date_str}.log")
    timestamp = datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
    with open(log_file, "a", encoding="utf-8") as f:
        f.write(f"{timestamp} {message}\n")

def run_process():
    base_path = get_base_path()
    csv_folder = os.path.join(base_path, "CSV")
    koji_path = os.path.join(csv_folder, "工事.csv")
    heya_path = os.path.join(csv_folder, "部屋.csv")
    paste_path = os.path.join(base_path, "部屋作成.xlsm")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = os.path.join(base_path, f"部屋作成_バックアップ_{timestamp}.xlsm")

    write_log("実行開始")

    if not os.path.exists(koji_path) or not os.path.exists(heya_path):
        messagebox.showerror("エラー", "CSVフォルダ内の工事.csvまたは部屋.csvが見つかりません。")
        write_log("❌ 必要なCSVファイルが見つかりませんでした。処理中止")
        return
    if not os.path.exists(paste_path):
        messagebox.showerror("エラー", "部屋作成.xlsmが見つかりません。")
        write_log("❌ 部屋作成.xlsmが見つかりませんでした。処理中止")
        return

    try:
        df_kouji = pd.read_csv(koji_path, encoding="cp932")
        df_heya = pd.read_csv(heya_path, encoding="cp932")
        write_log("✅ CSVファイル読み込み成功")

        df_kouji["マンション番号"] = df_kouji["マンション番号"].astype(str).str.strip()
        df_heya["マンション番号"] = df_heya["マンション番号"].astype(str).str.strip()

        count_map = df_heya["マンション番号"].value_counts().to_dict()
        df_kouji["(提供可能)戸数"] = pd.to_numeric(df_kouji["(提供可能)戸数"], errors='coerce').fillna(0).astype(int)

        output_lines = []
        for _, row in df_kouji.iterrows():
            mansion = row["マンション番号"]
            available = row["(提供可能)戸数"]
            created = count_map.get(mansion, 0)
            if available > created:
                remarks = str(row.get("埋込WiFi備考", "")).strip()
                remarks = re.sub(r"[\t\n\r\-]", " ", remarks)
                remarks = re.sub(r" +", " ", remarks)
                rooms = [room.strip() for room in remarks.split(" ") if room.strip()]
                for room in rooms:
                    output_lines.append(f"{mansion}\t{mansion}\t{room}")

        if not output_lines:
            messagebox.showinfo("結果", "精査の結果作成が必要な案件はありませんでした。")
            write_log("⚠ 出力対象なし（0件）")
            return

        pyperclip.copy("\n".join(output_lines))
        shutil.copy2(paste_path, backup_path)
        write_log(f"📋 出力対象：{len(output_lines)} 件")
        write_log(f"📁 バックアップファイル作成：{os.path.basename(backup_path)}")

        app = xw.App(visible=True)
        wb = app.books.open(backup_path)
        for book in app.books:
            if book.name == "Book1":
                book.close()

        sheet = wb.sheets[0]
        sheet.range("A2").value = [line.split("\t") for line in output_lines]
        wb.save()

        write_log("✅ Excelに貼り付け完了")
        write_log("✔ 実行完了\n")

        messagebox.showinfo("完了", f"{len(output_lines)} 件のデータを『{os.path.basename(backup_path)}』に貼り付けました。")

    except Exception as e:
        write_log(f"❌ エラー発生: {str(e)}")
        messagebox.showerror("エラー発生", str(e))

# GUI構築
root = tk.Tk()
root.title("部屋精査自動ツール")
root.geometry("300x150")

tk.Label(root, text="CSV/工事.csv・部屋.csvを参照します", pady=10).pack()
tk.Button(root, text="実行", command=run_process, height=2, width=15, bg="#4CAF50", fg="white").pack()

root.mainloop()