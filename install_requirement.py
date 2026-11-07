import subprocess, sys

packages = [
    "selenium>=4.10.0",
    "requests",
    "openpyxl",
]

for pkg in packages:
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg])
    except subprocess.CalledProcessError:
        print(f"安裝失敗: {pkg}")

input("\n按 Enter 結束...")
