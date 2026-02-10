================================================================================
Comment Sheet Aggregator / コメントシート集計ツール
================================================================================

[日本語 / Japanese]

■ 1. 概要
このツールは、教務システム等からダウンロードした「日ごとの学生コメントシート (Excel)」を読み込み、
学生ごとに1行にまとめた「集計結果 (Excel)」を作成するツールです。

■ 2. 使い方 (3つの方法)

   【A. Web版を使う (推奨)】
   インストール不要で、スマホやMacでも使えます。
   URL: https://ksgadget-csa.streamlit.app/

   【B. Windowsアプリを使う】
   フォルダ内の `コメントシート集計ツール.exe` をダブルクリックして起動します。
   1. 「1. コメントシート選択」: 集計したいExcelファイルを全て選びます。
   2. 「2. 出席簿を選択 (任意)」: 出席簿があれば選択します（学生ID順に並び変えるため）。
   3. 「対象年度 (任意)」: "2025" のように入力すると、その年度の学生だけ抽出します。
   4. 「集計開始」ボタンを押すと、結果の保存先を聞かれます。

   【C. Web版をローカルで動かす (開発者向け)】
   `run_web_app.bat` をダブルクリックすると、自分のパソコン上でWeb版が起動します。

■ 3. メンテナンス・修正ガイド (重要)
もし大学のシステムが変わり、Excelの「列」や「行」の位置が変わった場合は、
以下の手順で `src/aggregator.py` を修正してください。

   【修正するファイル】
   `src/aggregator.py` (メモ帳やテキストエディタで開いてください)

   【修正方法】
   ファイルの一番上に「CONFIGURATION (設定)」というエリアがあります。
   ここの数字を書き換えるだけで対応できます。
   
   ------------------------------------------------------------
   CONFIG = {
       # --- 1. Comment Sheet (コメントシート) ---
       "COL_SUB_ID": 0,    # A列
       "COL_COURSE": 2,    # C列
       "COL_NAME": 4,      # E列 (氏名)
       "COL_ID": 5,        # F列 (学籍番号)
       "COL_COMMENT": 6,   # G列 (コメント)
       
       # --- 2. Attendance Sheet (出席簿) ---
       "ATT_SKIP_ROWS": 6, # 上から読み飛ばす行数
       "ATT_COL_ID": 1,    # B列 (学籍番号)
       "ATT_COL_NAME": 2,  # C列 (氏名)
   }
   ------------------------------------------------------------
   例えば「氏名」がC列に移動したら、`"COL_NAME": 2` に書き換えてください。
   (※ A列=0, B列=1, C列=2 ... というルールです)

   【修正後の反映方法】
   修正が終わったら、フォルダにある `build_exe.bat` をダブルクリックしてください。
   自動的に新しい `コメントシート集計ツール.exe` が作成されます。

■ 4. フォルダの整理 (削除しても良いファイル)
`build_exe.bat` を実行した後、自動的に掃除されますが、もし残っていたら以下は削除して大丈夫です。

   * `build/` フォルダ (作業用の一時ファイル)
   * `dist/` フォルダ (作成中のexe置き場)
   * `*.spec` ファイル (ビルド設定ファイル)
   * `__pycache__/` フォルダ

--------------------------------------------------------------------------------

[English]

■ 1. Overview
This tool aggregates multiple daily "Student Comment Sheets" (Excel) into a single summary Excel file, grouped by student.

■ 2. How to Use

   【A. Web Version (Recommended)】
   No installation required. Works on Mobile/Mac.
   URL: https://ksgadget-csa.streamlit.app/

   【B. Windows Application】
   Double-click `コメントシート集計ツール.exe`.
   1. Select input Excel files.
   2. (Optional) Select Attendance Sheet for sorting.
   3. (Optional) Enter Target Year (e.g., 2025).
   4. Click "Run" to save the summary.

   【C. Local Web Version】
   Double-click `run_web_app.bat` to launch the Streamlit app locally on your PC.

■ 3. Maintenance Guide
If the university's Excel format changes (e.g., column positions change), edit `src/aggregator.py`.

   【File to Edit】
   `src/aggregator.py`

   【How to Edit】
   Look for the "CONFIGURATION" section at the top of the file.
   Change the numbers to match the new Excel layout.

   ------------------------------------------------------------
   CONFIG = {
       "COL_NAME": 4,      # Column E (Name)
       "COL_ID": 5,        # Column F (Student ID)
       "ATT_SKIP_ROWS": 6, # Rows to skip
       ...
   }
   ------------------------------------------------------------
   * Columns are 0-indexed (A=0, B=1, C=2...).

   【How to Rebuild EXE】
   After editing, double-click `build_exe.bat`. It will automatically generate a new `コメントシート集計ツール.exe`.

■ 4. Cleanup
The following files/folders are temporary and can be safely deleted:
   * `build/` folder
   * `dist/` folder
   * `*.spec` files
   * `__pycache__/` folder
