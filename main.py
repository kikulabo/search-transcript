import os
import pandas as pd
import glob
import sys
import traceback
import warnings

warnings.filterwarnings('ignore', category=UserWarning)

def search_in_excel(search_string, directory):
    """
    指定されたディレクトリ内のエクセルファイルで文字列を検索する関数
    
    Args:
        search_string (str): 検索する文字列
        directory (str): 検索対象のディレクトリパス
    """
    excel_patterns = ['*.xlsx', '*.xls']
    error_files = []
    found_matches = False
    
    for pattern in excel_patterns:
        excel_files = glob.glob(os.path.join(directory, '**', pattern), recursive=True)
        
        for excel_file in excel_files:
            try:
                df = pd.read_excel(excel_file, engine='openpyxl')
                
                for column in df.columns:
                    if df[column].dtype == 'object':
                        matches = df[df[column].astype(str).str.contains(search_string, na=False)]
                        
                        if not matches.empty:
                            if not found_matches:
                                print("\n=== 検索結果 ===")
                                found_matches = True
                            
                            print(f"\n【ファイル】: {os.path.basename(excel_file)}")
                            print(f"【パス】: {excel_file}")
                            print(f"【列名】: {column}")
                            print("【該当箇所】:")
                            for idx, row in matches.iterrows():
                                print(f"  行 {idx + 1}: {row[column]}")
                            
            except Exception as e:
                error_msg = f"エラー: {excel_file} の処理中にエラーが発生しました:\n{traceback.format_exc()}"
                print(error_msg)
                error_files.append(excel_file)
    
    if not found_matches:
        print("\n検索文字列を含む箇所は見つかりませんでした。")
    
    if error_files:
        print("\n【エラーが発生したファイル一覧】:")
        for file in error_files:
            print(f"- {file}")

def main():
    if len(sys.argv) != 2:
        print("使用方法: python search_transcript.py <検索文字列>")
        sys.exit(1)
    
    search_string = sys.argv[1]
    directory = os.environ["SEARCH_DIR"]
    
    print(f"検索文字列: {search_string}")
    print(f"検索ディレクトリ: {directory}")
    
    search_in_excel(search_string, directory)

if __name__ == "__main__":
    main() 
