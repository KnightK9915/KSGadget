import sys
import os

# Ensure the current directory (src) is in sys.path so we can import aggregator
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.append(current_dir)

import streamlit as st
import tempfile
import shutil
import pandas as pd
from datetime import datetime

try:
    from aggregator import process_files
except ImportError as e:
    # Try fallback but keep original error if both fail
    try:
        from src.aggregator import process_files
    except ImportError as e2:
        st.error(f"Critical Error: Could not import 'aggregator'.")
        st.error(f"Attempt 1 (Direct): {e}")
        st.error(f"Attempt 2 (src.aggregator): {e2}")
        st.write("Current sys.path:", sys.path)
        st.write("Current files in src:", os.listdir(os.path.dirname(os.path.abspath(__file__))))
        st.stop()

def main():
    st.set_page_config(
        page_title="コメントシート集計ツール", 
        page_icon="📑", 
        layout="centered",
        initial_sidebar_state="expanded"
    )

    # --- Custom CSS for Mobile Optimization ---
    st.markdown("""
        <style>
        .stButton>button {
            width: 100%;
            border-radius: 20px;
            font-weight: bold;
            height: 3em;
        }
        .main .block-container {
            padding-top: 2rem;
            padding-bottom: 2rem;
        }
        h1 {
            font-size: 1.8rem !important;
        }
        .upload-box {
            border: 2px dashed #ccc;
            padding: 20px;
            text-align: center;
            border-radius: 10px;
        }
        </style>
    """, unsafe_allow_html=True)
    
    # --- Header ---
    st.title("📑 コメントシート集計")
    st.caption("Excelファイルをアップロードして、一人一行にまとめます。")

    with st.expander("ℹ️ 使い方 (How to use)", expanded=False):
        st.markdown("""
        1. **「コメントシート」** (複数可) をアップロードします。
        2. (任意) **「出席簿」** をアップロードすると、学籍番号順に並び替えられます。
        3. サイドバーで **「対象年度」** を指定できます。
        4. **「集計開始」** ボタンを押すと、結果がダウンロードできます。
        """)
    
    # --- Sidebar / Settings ---
    with st.sidebar:
        st.header("⚙️ 設定 (Settings)")
        target_year = st.text_input("対象年度 (例: 2025)", placeholder="空白なら全期間")
        st.info("※ 空白の場合は、すべてのデータを集計します。")
    
    # --- Layout ---
    
    # Step 1: Input Files
    st.subheader("1️⃣ コメントシート (必須)")
    uploaded_files = st.file_uploader(
        "ここにExcelファイルをドラッグ＆ドロップ", 
        type=["xlsx", "xls"], 
        accept_multiple_files=True,
        key="comments"
    )
    
    if uploaded_files:
        st.success(f"{len(uploaded_files)} 個のファイルを選択中")

    st.divider()

    # Step 2: Attendance (Optional)
    st.subheader("2️⃣ 出席簿 (オプション)")
    attendance_file = st.file_uploader(
        "出席簿 (Excel) がある場合はここで選択", 
        type=["xlsx", "xls"],
        key="attendance"
    )

    st.divider()

    # --- Processing ---
    if st.button("🚀 集計開始 (Start Aggregation)", type="primary"):
        if not uploaded_files:
            st.warning("⚠️ まずはコメントシートを選択してください。")
            return

        # Create a temporary directory to store uploaded files
        # because aggregator.py expects file paths, not file objects.
        with tempfile.TemporaryDirectory() as temp_dir:
            input_dir = os.path.join(temp_dir, "input")
            os.makedirs(input_dir, exist_ok=True)
            
            # Progress Bar
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            status_text.text("📂 ファイルを読み込んでいます...")
            
            # Save Input Files
            input_paths = []
            for i, uploaded_file in enumerate(uploaded_files):
                file_path = os.path.join(input_dir, uploaded_file.name)
                with open(file_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                input_paths.append(file_path)
                progress_bar.progress((i + 1) / len(uploaded_files) * 0.3) # First 30%
            
            # Save Attendance File
            attendance_path = None
            if attendance_file:
                attendance_path = os.path.join(temp_dir, attendance_file.name)
                with open(attendance_path, "wb") as f:
                    f.write(attendance_file.getbuffer())
                status_text.text("📋 出席簿を処理中...")

            # Define Output Path
            output_filename = f"summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            output_path = os.path.join(temp_dir, output_filename)
            
            # Run Aggregation logic
            status_text.text("⏳ 集計処理を実行中... (これには数秒かかる場合があります)")
            progress_bar.progress(0.6)
            
            try:
                # Run the actual aggregation
                success, msg = process_files(input_paths, output_path, target_year, attendance_path)
                
                progress_bar.progress(1.0)
                
                if success and os.path.exists(output_path):
                    status_text.success("✅ 集計完了！")
                    st.balloons()
                    
                    # Read the result file for download
                    with open(output_path, "rb") as f:
                        file_data = f.read()
                        
                    st.download_button(
                        label="📥 結果をダウンロード (Download Result)",
                        data=file_data,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
                else:
                    status_text.error(f"❌ エラー: {msg}")
                    
            except Exception as e:
                status_text.error(f"❌ 予期せぬエラー: {e}")

    # --- Footer ---
    st.markdown("---")
    st.markdown(
        """
        <div style='text-align: center; color: gray; font-size: 0.8em;'>
        制作：2025年度院生（有志）<br>
        Developed by 2025 Graduate Students
        </div>
        """, 
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()
