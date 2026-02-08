================================================================================
README.txt
コメントシート集計ツール / Comment Sheet Aggregator
================================================================================

[日本語 / Japanese]

--------------------------------------------------------------------------------
1. 使用方法 (How to Use)
--------------------------------------------------------------------------------
このツールは、複数の学生からのコメントシート Excel ファイルを読み込み、
一つの Excel ファイルに集計・整理するためのものです。

【手順】
1.  `CommentAggregatorUI.exe` をダブルクリックして起動します。
2.  「ファイル選択」ボタンをクリックし、集計したいコメントシートの Excel ファイル（複数可）を選択します。
3.  （任意）「対象年度」を入力します（例: 2025）。入力すると、C列（科目名など）がその年度で始まる行のみを抽出します。
4.  （任意）「出席表選択」ボタンをクリックし、出席簿の Excel ファイルを選択します。
    *   出席簿があると、欠席者の確認や、学籍番号順の並び替え、氏名の正規化（漢字表記への統一）が自動で行われます。
5.  「実行」ボタンをクリックします。
6.  処理が完了すると、`output` フォルダに `summary_YYYYMMDD_HHMMSS.xlsx` という名前で集計結果が保存されます。

【主な機能】
*   **重複排除**: 同一学生が同じ日に複数回提出している場合、提出番号（A列）が最大のものを採用します。
*   **名前寄せ (名寄せ)**: 学籍番号の入力ミスがあっても、氏名が一致していれば出席簿の情報に基づいて修正・集計します。
*   **色情報の保持**: コメントシートで黄色などに塗りつぶされたセルは、集計後もそのままの色で出力されます。
*   **未回答のハイライト**: 未提出または未回答のセルは「薄い赤色」でハイライトされます。

--------------------------------------------------------------------------------
2. 将来、Excelの形式が変わった場合の修正方法 (How to Modify Source Code)
--------------------------------------------------------------------------------
★重要★
将来、教务システムの変更により、コメントシートの列の並び順が変わった場合（例：氏名がE列からF列に移動した等）、
Pythonがわかる学生に以下の修正を依頼してください。

【修正対象ファイル】
`src/aggregator.py`

【修正手順】
1.  `src/aggregator.py` をテキストエディタ（VS Code, メモ帳など）で開きます。
2.  `WS_IN.ITER_ROWS` (小文字で `ws_in.iter_rows`) という箇所を探し、その直下にあるデータ読み込み部分を見つけます。
    
    ```python
    # 現在のコード（例）
    sub_id_col = get_val(row[0])  # A列 (0番目)
    course_col = get_val(row[2])  # C列 (2番目)
    name_col   = get_val(row[4])  # E列 (4番目)
    id_col     = get_val(row[5])  # F列 (5番目)
    comment_col = get_val(row[6]) # G列 (6番目)
    ```

3.  上記の数字 `[]` の中身を変更してください。
    *   Excelの列は 0 から始まります（A列=0, B列=1, C列=2, D列=3, E列=4, F列=5, G列=6 ...）。
    *   例えば、氏名が「F列」に移動した場合は、`name_col = get_val(row[4])` を `get_val(row[5])` に変更します。

4.  ファイルを保存します。
5.  修正を反映させるには、Python環境で直接実行するか、`PyInstaller` を使って exe ファイルを作り直す必要があります。
    *   直接実行する場合: `python src/gui_app.py`
    *   exe再作成する場合: `pyinstaller --onefile --noconsole --name "CommentAggregatorUI" --clean --hidden-import=xlrd src/gui_app.py`

--------------------------------------------------------------------------------
3. Macでのインストールと実行 (How to Install on Mac)
--------------------------------------------------------------------------------
1.  ターミナル (Terminal) を開きます。
2.  このフォルダ (`comment_sheet_aggregator`) に移動します。
3.  以下のコマンドを実行して、セットアップとアプリの作成を行います。
    sh setup_mac.sh
4.  処理が完了すると、`dist` フォルダの中に `CommentAggregator.app` が作成されます。
5.  これをアプリケーションフォルダに移動するか、ダブルクリックして実行してください。

================================================================================

[English]

--------------------------------------------------------------------------------
1. How to Use
--------------------------------------------------------------------------------
This tool aggregates comment sheets from multiple student Excel files into a single summary Excel file.

【Steps】
1.  Double-click `CommentAggregatorUI.exe` to launch the application.
2.  Click "Select Files" to choose the student Excel files you want to process.
3.  (Optional) Enter "Target Year" (e.g., 2025). This filters entries based on the Course Name column (starts with the year).
4.  (Optional) Click "Select Attendance Sheet" to load a master student list.
    *   This enables sorting by ID, identifying missing students, and normalizing names.
5.  Click "Run".
6.  The output file will be saved in the `output` folder as `summary_YYYYMMDD_HHMMSS.xlsx`.

【Key Features】
*   **Deduplication**: Keeps the latest submission if a student submits multiple files for the same date.
*   **Fuzzy Name Matching**: Corrects student IDs if they are wrong but the name matches the attendance sheet.
*   **Color Preservation**: Preserves cell background colors (e.g., yellow highlights) from input files.
*   **Unanswered Highlight**: Missing entries are highlighted in light red.

--------------------------------------------------------------------------------
2. How to Modify the Source Code (For Maintenance)
--------------------------------------------------------------------------------
★IMPORTANT★
If the university's Excel format changes in the future (e.g., the "Name" column moves from Column E to Column F),
please ask a student who knows Python to modify the code as follows.

【File to Modify】
`src/aggregator.py`

【Instruction】
1.  Open `src/aggregator.py` with a text editor (VS Code, Notepad, etc.).
2.  Search for the section inside the loop `for row in ws_in.iter_rows():`.
    You will see lines that extract data such as:

    ```python
    # Current Code Example
    sub_id_col = get_val(row[0])  # Column A (Index 0)
    course_col = get_val(row[2])  # Column C (Index 2)
    name_col   = get_val(row[4])  # Column E (Index 4) -> Name
    id_col     = get_val(row[5])  # Column F (Index 5) -> Student ID
    comment_col = get_val(row[6]) # Column G (Index 6) -> Comment
    ```

3.  Update the index number inside `[]`.
    *   Excel columns are 0-indexed (A=0, B=1, ... E=4, F=5, G=6).
    *   Example: If the Name moves to Column F, change `row[4]` to `row[5]`.

4.  Save the file.
5.  To apply changes, you must run the script directly with Python or rebuild the exe.
    *   Run directly: `python src/gui_app.py`
    *   Rebuild exe: `pyinstaller --onefile --noconsole --name "CommentAggregatorUI" --clean --hidden-import=xlrd src/gui_app.py`

--------------------------------------------------------------------------------
3. How to Install on Mac
--------------------------------------------------------------------------------
1.  Open Terminal.
2.  Navigate to this folder (`comment_sheet_aggregator`).
3.  Run the following command to set up and build the app:
    sh setup_mac.sh
4.  When finished, `CommentAggregator.app` will be created in the `dist` folder.
5.  Move it to your Applications folder or double-click to run.
--------------------------------------------------------------------------------
4. Web版 (Streamlit) の実行 (How to Run Web Version)
--------------------------------------------------------------------------------
Webブラウザ上で動作するバージョンも用意しました。

【Windows】
フォルダ内にある `run_web_app.bat` をダブルクリックするだけで起動します。

【Mac】
1.  セットアップがまだの場合は、一度だけ `sh setup_mac.sh` を実行してください。
2.  そのあとは、`run_web_app_mac.command` をダブルクリックするだけで起動します。
    
    ⚠️ **「権限がありません」や「開けません」と言われる場合**:
    ターミナルで以下のコマンドを**1回だけ**実行してください（ファイルの実行権限を付与します）：
    chmod +x run_web_app_mac.command
    
    (※初回のみ、「開発元が未確認」等の警告が出る場合があります。その場合は右クリック→「開く」を選択してください)
