# room-audit-automation.py
Automates audit of apartment project CSVs and outputs validated room list via GUI（アパート案件のCSVを自動で精査し、部屋リストをGUI出力）
# Room Audit Automation（部屋精査自動ツール・匿名化版）

This tool automates the validation of apartment project data by checking CSV files and exporting eligible room entries to an Excel template via GUI.  
（CSVから部屋データを精査し、提供可能な部屋リストをGUI操作でExcelに貼り付けるツールです）

---

## ✅ Features / 特徴
- 📄 工事案件CSV（工事.csv・部屋.csv）の読み込み・精査
- ✅ 提供可能戸数と作成済戸数を比較し、差異のある物件のみ抽出
- 🖥️ Tkinter GUI で簡単操作
- 📋 クリップボード自動コピー＆Excelテンプレに貼り付け
- 💾 自動バックアップ（xlsmファイル）
- 📜 実行ログ生成（logフォルダ）

---

## 🛠️ Requirements / 動作環境

- Windows OS
- Python 3.10+
- Microsoft Excel（xlwings対応）
- インストール必要ライブラリ：

```bash
pip install pandas xlwings pyperclip


## 🚀 How to Use / 使い方

1. このリポジトリをクローン or ZIPダウンロードします
2. `CSV/` フォルダ内に以下の2つのCSVを格納します：
   - `工事.csv`
   - `部屋.csv`
3. Excelテンプレート `部屋作成.xlsm` をルートに配置します
4. 以下のコマンドでツールを起動します：

```bash
python room-audit-automation.py
