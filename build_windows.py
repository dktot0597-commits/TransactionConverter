"""Build TransactionConverter v4.4 as a single Windows executable."""
from pathlib import Path
import os
import subprocess
import sys

ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / 'TransactionConverter_GUI_v4_4.py'
EXE_NAME = 'TransactionConverter_v4_4'

def run(args):
    print('> ' + subprocess.list2cmdline([str(a) for a in args]), flush=True)
    subprocess.run([str(a) for a in args], cwd=ROOT, check=True)

def main():
    if os.name != 'nt':
        raise RuntimeError('Windows PC에서만 Windows EXE를 빌드할 수 있습니다.')
    if not SOURCE.is_file():
        raise FileNotFoundError(f'빌드 대상 Python 파일이 없습니다: {SOURCE}')
    print('[1/3] Installing build dependencies...', flush=True)
    run([sys.executable, '-m', 'pip', 'install', 'pyinstaller', 'openpyxl', 'tkinterdnd2'])
    print('[2/3] Building Windows EXE...', flush=True)
    cmd = [
        sys.executable, '-m', 'PyInstaller',
        '--noconfirm', '--clean', '--onefile', '--windowed',
        '--name', EXE_NAME,
        '--collect-all', 'tkinterdnd2',
        '--hidden-import', 'openpyxl',
        '--distpath', str(ROOT / 'dist'),
        '--workpath', str(ROOT / 'build'),
        '--specpath', str(ROOT),
        str(SOURCE),  # required positional scriptname
    ]
    assert cmd[-1] == str(SOURCE) and SOURCE.exists()
    run(cmd)
    output = ROOT / 'dist' / (EXE_NAME + '.exe')
    if not output.is_file():
        raise FileNotFoundError(f'빌드 완료 파일을 찾을 수 없습니다: {output}')
    print(f'\n[3/3] SUCCESS: {output}', flush=True)
    print('기존 TransactionConverter_Settings.json 파일은 EXE 옆에 복사하세요.', flush=True)

if __name__ == '__main__':
    try:
        main()
    except (RuntimeError, FileNotFoundError, subprocess.CalledProcessError) as exc:
        print(f'\n[ERROR] {exc}', file=sys.stderr, flush=True)
        sys.exit(1)
