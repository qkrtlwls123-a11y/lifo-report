# -*- coding: utf-8 -*-
"""LIFO 웹 앱 실행 스크립트"""
import subprocess
import sys
import os

os.chdir(os.path.dirname(os.path.abspath(__file__)))

# 패키지 설치
print("[1/2] Installing packages...")
subprocess.check_call(
    [sys.executable, "-m", "pip", "install", "-r", "requirements.txt", "-q"],
    stdout=subprocess.DEVNULL
)

# 서버 시작
print("[2/2] Starting server...")
print()
print("  >>> http://localhost:5000 <<<")
print()
print("  Ctrl+C to stop")
print()

subprocess.call([sys.executable, "app.py"])
