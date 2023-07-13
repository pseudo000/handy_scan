import sys
import ctypes

if getattr(sys, 'frozen', False):
    # 실행 파일의 경로 추출
    application_path = sys.executable
else:
    # 개발 모드에서 스크립트의 경로 추출
    application_path = sys.argv[0]

# 아이콘 파일 경로
icon_path = 'C:\\Users\\swwoo\\Desktop\\handy_csv\\ICON.ico'  # ICON.ico를 실제 아이콘 파일의 경로로 대체해주세요.

# 아이콘 적용
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(application_path)
ctypes.windll.shell32.SetShortcutIcon(icon_path, application_path)
