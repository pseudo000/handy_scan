import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os

txt_files = []  # 선택한 텍스트 파일들을 저장하는 리스트

def search_text_file():
    file_paths = filedialog.askopenfilenames(filetypes=[("Text Files", "*.txt")])
    if file_paths:
        txt_files.extend(file_paths)
        txt_file_entry.delete(0, tk.END)
        txt_file_entry.insert(tk.END, ", ".join(txt_files))

def start_processing():
    if txt_files:
        output = pd.DataFrame(columns=['申告番号', 'AWB番号', '区分'])

        for txt_file_path in txt_files:
            with open(txt_file_path, "r") as file:
                lines = file.readlines()

            row_index = 10
            row_gap = [1, 4, 14]
            row_data = []

            while row_index <= len(lines):
                for i in range(len(row_gap)):
                    row = lines[row_index - 1].strip()
                    row_data.append(row)

                    if len(row_data) == 3:
                        if len(output) == 0:  # 첫 번째 파일인 경우에는 컬럼 생성
                            output = pd.concat([output, pd.DataFrame([row_data], columns=['申告番号', 'AWB番号', '区分'])], ignore_index=True)
                        else:  # 두 번째 파일부터는 기존 행의 다음 행부터 데이터 추가
                            output.loc[len(output)] = row_data

                        row_data = []

                    row_index += row_gap[i]

            if row_data:
                if len(output) == 0:  # 첫 번째 파일인 경우에는 컬럼 생성
                    output = pd.concat([output, pd.DataFrame([row_data], columns=['申告番号', 'AWB番号', '区分'])], ignore_index=True)
                else:  # 두 번째 파일부터는 기존 행의 다음 행부터 데이터 추가
                    output.loc[len(output)] = row_data

        output_file_path = 'output.xlsx'
        output.to_excel(output_file_path, index=False)
        result_label.config(text="SAVEしました。")

        # 저장된 엑셀 파일 열기
        os.startfile(output_file_path)
    else:
        result_label.config(text="TXTファイルを選択して下さい。")

# GUI 생성
root = tk.Tk()
root.title("エア。。。。")

# 텍스트 파일 경로 입력 필드
txt_file_entry = tk.Entry(root, width=50)
txt_file_entry.pack(pady=10)

# "TXT 파일 검색" 버튼
search_button = tk.Button(root, text="TXTファイル選択", command=search_text_file)
search_button.pack(pady=5)

# "Start" 버튼
start_button = tk.Button(root, text="Start", command=start_processing)
start_button.pack(pady=10)

# 결과 레이블
result_label = tk.Label(root, text="")
result_label.pack(pady=5)

root.mainloop()



