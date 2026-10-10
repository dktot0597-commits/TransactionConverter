TransactionConverter v4.4 - corrected Windows EXE build kit

1. ZIP 압축을 전체 해제하세요. BAT 파일만 따로 실행하면 안 됩니다.
2. Build_EXE.bat를 더블클릭합니다 (빌드 컴퓨터에는 Python 3 및 인터넷 필요).
3. dist\TransactionConverter_v4_4.exe 생성 여부를 확인하세요.
4. 프로그램 프리셋을 유지하려면 TransactionConverter_Settings.json을 EXE 옆에 복사하세요.

v4.4 원본 Python 소스는 변경하지 않았습니다.
이번 수정은 빌드 명령에 대상 .py 파일이 누락되는 현상을 피하기 위해
별도 build_windows.py가 인수 목록을 구성하여 PyInstaller를 실행하도록 한 것입니다.

빌드 대상 파일이 없으면 실행 전에 오류를 표시하고 중단합니다.
빌드 단계가 실패하면 해당 오류를 표시합니다.
Windows용 EXE는 Windows 환경에서 빌드해야 합니다.
