import streamlit as st
import os
import tempfile
import shutil
import pandas as pd
from datetime import datetime
from aggregator import process_files

def main():
    st.set_page_config(page_title="コメントシート集計ツール", layout="centered")
    
    st.title("📑 コメントシート集計ツール (Web版)")
    st.markdown("""
    複数のコメントシート(Excel)をアップロードして、一つにまとめます。
    既存のWindows/Macアプリと同じロジックを使用しています。
    """)
    
    # --- Sidebar / Settings ---
    st.sidebar.header("設定 (Settings)")
    target_year = st.sidebar.text_input("対象年度 (例: 2025)", placeholder="2025")
    
    # --- File Upload ---
    st.subheader("1. ファイル選択 (Input Files)")
    uploaded_files = st.file_uploader(
        "学生のコメントシートを選択してください (複数可)", 
        type=["xlsx", "xls"], 
        accept_multiple_files=True
    )
    
    st.subheader("2. 出席簿 (Attendance Sheet) [任意]")
    attendance_file = st.file_uploader(
        "出席簿を選択してください (オプション)", 
        type=["xlsx", "xls"]
    )

    # --- Processing ---
    if st.button("集計開始 (Run Aggregation)", type="primary"):
        if not uploaded_files:
            st.error("⚠️ コメントシートを選択してください。")
            return

        # Create a temporary directory to store uploaded files
        # because aggregator.py expects file paths, not file objects.
        with tempfile.TemporaryDirectory() as temp_dir:
            input_dir = os.path.join(temp_dir, "input")
            os.makedirs(input_dir, exist_ok=True)
            
            # Save Input Files
            input_paths = []
            for uploaded_file in uploaded_files:
                file_path = os.path.join(input_dir, uploaded_file.name)
                with open(file_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                input_paths.append(file_path)
            
            st.info(f"📄 {len(input_paths)} 個のファイルを読み込みました。")
            
            # Save Attendance File
            attendance_path = None
            if attendance_file:
                attendance_path = os.path.join(temp_dir, attendance_file.name)
                with open(attendance_path, "wb") as f:
                    f.write(attendance_file.getbuffer())
                st.info(f"📋 出席簿: {attendance_file.name}")

            # Define Output Path
            output_filename = f"summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            output_path = os.path.join(temp_dir, output_filename)
            
            # Run Aggregation logic
            status_text = st.empty()
            status_text.text("⏳ 集計中... (Processing...)")
            
            try:
                # Redirect print output to capture logs if needed, but for now just run it
                # The aggregator returns (Success, Message)
                # But looking at src/aggregator.py, process_files returns nothing? 
                # Let's double check aggregator.py signature. It prints to console.
                # Assuming it works if no exception.
                
                process_files(input_paths, output_path, target_year, attendance_path)
                
                if os.path.exists(output_path):
                    status_text.success("✅ 集計完了！ (Done!)")
                    
                    # Read the result file for download
                    with open(output_path, "rb") as f:
                        file_data = f.read()
                        
                    st.download_button(
                        label="📥 結果をダウンロード (Download Result)",
                        data=file_data,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    status_text.error("❌ エラー: 出力ファイルが生成されませんでした。")
                    
            except Exception as e:
                status_text.error(f"❌ エラーが発生しました: {e}")

if __name__ == "__main__":
    main()
