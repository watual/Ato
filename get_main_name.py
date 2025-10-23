#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MAIN_NAME을 읽어오는 스크립트
한글 인코딩 문제를 해결하기 위해 별도 파일로 분리
"""

import sys
import os

# 현재 디렉토리를 Python 경로에 추가
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

try:
    from pdf_email_sender_gui import MAIN_NAME
    if MAIN_NAME:
        sys.stdout.buffer.write(MAIN_NAME.encode('utf-8'))
    else:
        sys.stdout.buffer.write("PDF 이메일 자동발송".encode('utf-8'))
except Exception as e:
    sys.stdout.buffer.write("PDF 이메일 자동발송".encode('utf-8'))
