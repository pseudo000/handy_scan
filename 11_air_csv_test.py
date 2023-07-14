import pandas as pd

# 행 번호
rows = [10, 11, 15, 29, 30, 34, 48, 49, 53, 67, 68, 72, 86, 87, 91]

# 데이터 추출하여 엑셀에 저장
output = pd.DataFrame(columns=['A', 'B', 'C'])  # 열 이름 설정

# TXT 파일 읽기
with open("C:\\Users\\swwoo\\Desktop\\air_csv\\AIR3.txt", "r") as file:
    lines = file.readlines()

# 행 번호에 해당하는 데이터 엑셀에 저장
for i, row in enumerate(rows):
    column_name = chr(65 + (i % 3))  # 열 이름 할당 (A, B, C)

    if i % 3 == 0:  # 첫 번째 숫자인 경우
        output.loc[i // 3, column_name] = lines[row - 1].strip()
    elif i % 3 == 1:  # 두 번째 숫자인 경우
        output.loc[i // 3, column_name] = lines[row - 1].strip()
    else:  # 세 번째 숫자인 경우
        output.loc[i // 3, column_name] = lines[row - 1].strip()

# 엑셀 파일로 저장
output.to_excel('output.xlsx', index=False)

