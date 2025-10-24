"""
PDF 자동 이메일 발송 프로그램 (GUI 버전 - 내장 설정)
모든 설정은 프로그램 내부에 저장됩니다.
"""

import copy
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import socket
import smtplib
from datetime import datetime
import re
import json
import webbrowser
import logging
from pathlib import Path
import time
import os
import sys
import threading
from tkinter import ttk, scrolledtext, messagebox, filedialog
import tkinter as tk
import ctypes


def get_version_from_release_notes():
    """RELEASE.md에서 버전 정보를 읽어옴"""
    try:
        # 실행 파일의 경로 확인
        if getattr(sys, 'frozen', False):
            base_dir = Path(sys.executable).parent
        else:
            base_dir = Path(__file__).parent
        
        release_notes_path = base_dir / "RELEASE.md"
        
        if release_notes_path.exists():
            with open(release_notes_path, 'r', encoding='utf-8') as f:
                content = f.read()
                # "[-[VERSION:1.0.1]-]" 형식에서 버전 추출
                match = re.search(r'\[-\[VERSION:(\d+\.\d+\.\d+)\]-\]', content)
                if match:
                    return match.group(1)
    except Exception:
        pass
    
    # 기본값 (RELEASE.md를 찾을 수 없을 때)
    return "0.0.1"


VERSION = get_version_from_release_notes()  # 프로그램 버전
MAIN_NAME = "Ato"
# 파일 및 폴더명에 사용할 prefix (MAIN_NAME이 있으면 "MAIN_NAME_", 없으면 빈 문자열)
NAME_PREFIX = f"{MAIN_NAME}_" if MAIN_NAME else ""
GLOBAL_VARIABLE = { 'test1': None }


def get_main_name():
    """MAIN_NAME을 반환하는 함수 (배치 파일에서 호출용)"""
    return MAIN_NAME


class ConfigManager:
    """프로그램 내부 설정 관리"""

    def __init__(self, log_func=None):
        self.log_func = log_func  # GUI 로그 함수 저장

        # 설정 파일 경로 (프로그램과 같은 위치에 저장 - USB 이동 가능)
        if getattr(sys, 'frozen', False):
            # exe로 실행 중
            base_dir = Path(sys.executable).parent
        else:
            # Python 스크립트로 실행 중
            base_dir = Path(__file__).parent

        # NAME_PREFIX를 사용하여 설정 파일명 생성
        self.config_file = base_dir / f'{NAME_PREFIX}settings.json'

        # 기본 설정
        # NAME_PREFIX를 사용하여 폴더명 생성
        pdf_folder_name = f'{NAME_PREFIX}전송할PDF'
        completed_folder_name = f'{NAME_PREFIX}전송완료'

        self.default_config = {
            'email': {
                'smtp_server': 'smtp.gmail.com',
                'smtp_port': 587,
                'sender_email': '',
                'sender_password': ''
            },
            'pattern': r'^([가-힣A-Za-z0-9\s]+?)(?:___|\.pdf$)',
            'auto_select_timeout': 10,
            'auto_send_timeout': 10,
            'email_send_timeout': 180,  # 이메일 발송 최대 대기 시간 (초)
            'debug_mode': False,
            'create_folders': False,
            'pdf_folder': str(Path.cwd()),
            'completed_folder': str(Path.cwd()),
            'companies': {},  # {회사명: {'emails': [], 'template': 'A'}}
            'custom_variables': {},  # {변수명: '값'} 예: {'이름': '홍길동', '담당자1': '김철수'}
            'email_templates': {
                '공식 보고서': {
                    'subject': '[{회사명}] {날짜} 업무 보고',
                    'body': '''안녕하십니까.

{회사명} 담당자님께 업무 관련 자료를 송부드립니다.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📋 전달 자료
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
▪ 파일명: {파일명}
▪ 발송일시: {날짜} {시간}

첨부 파일을 확인하신 후, 검토 부탁드리겠습니다.
문의사항이 있으시면 언제든 연락 주시기 바랍니다.

감사합니다.'''
                },
                '간결한 전달': {
                    'subject': '{회사명} 자료 전달',
                    'body': '''{회사명} 담당자님, 안녕하세요.

요청하신 자료를 첨부하여 보내드립니다.

📎 {파일명}

확인 후 회신 부탁드립니다.

감사합니다.
{날짜} {시간}'''
                },
                '정중한 공문': {
                    'subject': '[{회사명} 귀중] 문서 송부의 건',
                    'body': '''귀사의 무궁한 발전을 기원합니다.

{회사명} 담당자님께 아래와 같이 관련 문서를 송부하오니
검토하여 주시기 바랍니다.

                        - 아    래 -

1. 송부 문서: {파일명}
2. 발송 일시: {날짜} {시간}
3. 비고: 첨부파일 참조

궁금하신 사항이나 추가 자료가 필요하신 경우
연락 주시면 즉시 대응하겠습니다.

감사합니다.'''
                }
            }
        }

        self.config = self.load_config()

    def load_config(self):
        """설정 로드"""
        try:
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    loaded = json.load(f)
                    # 기본 설정과 병합 (deep copy 사용)
                    config = copy.deepcopy(self.default_config)
                    self._deep_update(config, loaded)
                    if self.log_func:
                        self.log_func(
                            f"✓ 설정 파일 로드 완료: {self.config_file.name}")
                    return config
        except Exception as e:
            if self.log_func:
                self.log_func(f"❌ 설정 로드 오류: {e}")

        if self.log_func:
            self.log_func("기본 설정으로 초기화")
        return copy.deepcopy(self.default_config)

    def _deep_update(self, base_dict, update_dict):
        """딕셔너리 깊은 업데이트"""
        for key, value in update_dict.items():
            if key in base_dict and isinstance(base_dict[key], dict) and isinstance(value, dict):
                self._deep_update(base_dict[key], value)
            else:
                base_dict[key] = value

    def save_config(self):
        """설정 저장"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=2, ensure_ascii=False)
            if self.log_func:
                self.log_func(f"✓ 설정 파일 저장 완료: {self.config_file.name}")
            return True
        except Exception as e:
            if self.log_func:
                self.log_func(f"❌ 설정 저장 오류: {e}")
            return False

    def get(self, key, default=None):
        """설정 값 가져오기"""
        keys = key.split('.')
        value = self.config
        for k in keys:
            if isinstance(value, dict):
                value = value.get(k, default)
            else:
                return default
        return value

    def set(self, key, value):
        """설정 값 설정하기"""
        keys = key.split('.')
        config = self.config
        for k in keys[:-1]:
            if k not in config:
                config[k] = {}
            config = config[k]
        config[keys[-1]] = value
        self.save_config()

    def reload(self):
        """설정 다시 로드"""
        self.config = self.load_config()


def center_window(child_window, parent_window, width=None, height=None):
    """창을 부모 창의 중앙에 위치시키는 함수"""
    child_window.update_idletasks()
    parent_window.update_idletasks()
    
    # 부모 창의 위치와 크기 가져오기
    parent_x = parent_window.winfo_x()
    parent_y = parent_window.winfo_y()
    parent_width = parent_window.winfo_width()
    parent_height = parent_window.winfo_height()
    
    # 자식 창 크기 결정
    if width is None:
        width = child_window.winfo_width()
    if height is None:
        height = child_window.winfo_height()
    
    # 중앙 위치 계산
    x = parent_x + (parent_width - width) // 2
    y = parent_y + (parent_height - height) // 2
    
    # 화면 밖으로 나가지 않도록 조정
    if x < 0:
        x = 0
    if y < 0:
        y = 0
    
    child_window.geometry(f"{width}x{height}+{x}+{y}")


class SettingsDialog:
    """설정 대화상자"""

    def __init__(self, parent, config_manager, parent_gui=None):
        self.result = None
        self.config_manager = config_manager
        self.parent_gui = parent_gui

        # 임시 데이터 저장 (취소 시 복원용)
        import copy
        self.temp_companies = copy.deepcopy(
            config_manager.get('companies', {}))
        self.temp_templates = copy.deepcopy(
            config_manager.get('email_templates', {}))

        self.dialog = tk.Toplevel(parent)
        self.dialog.title("설정")
        self.dialog.resizable(True, True)
        self.dialog.transient(parent)

        # 창을 부모 중앙에 위치 (setup_ui 전에 위치 설정)
        self.dialog.update_idletasks()
        parent.update_idletasks()

        # 중앙 위치 설정
        center_window(self.dialog, parent, 700, 650)
        self.dialog.grab_set()

        self.setup_ui()

        # 창 닫기 이벤트 처리 (취소 시 복원)
        self.dialog.protocol("WM_DELETE_WINDOW", self.on_cancel)

    def setup_ui(self):
        """UI 구성"""
        try:
            if self.parent_gui:
                self.parent_gui.log("🔧 설정 대화상자 UI 구성 시작", is_debug=True)

            # 버튼 프레임 먼저 생성 (하단 고정)
            if self.parent_gui:
                self.parent_gui.log("  - 저장/취소 버튼 생성 중...", is_debug=True)
            button_frame = ttk.Frame(self.dialog)
            button_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=10)

            ttk.Button(button_frame, text="💾 저장", command=self.save_settings, width=10).pack(
                side=tk.RIGHT, padx=5)
            ttk.Button(button_frame, text="❌ 취소",
                       command=self.on_cancel, width=10).pack(side=tk.RIGHT)

            if self.parent_gui:
                self.parent_gui.log("  ✓ 저장/취소 버튼 완료 (하단 고정)", is_debug=True)

            # 노트북 (탭) - 나머지 공간 차지
            notebook = ttk.Notebook(self.dialog)
            notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=(10, 0))

            if self.parent_gui:
                self.parent_gui.log("  ✓ 노트북 생성 완료", is_debug=True)

            # 이메일 설정 탭 (1순위)
            try:
                if self.parent_gui:
                    self.parent_gui.log("  - 이메일 설정 탭 생성 중...", is_debug=True)
                
                # 스크롤 가능한 프레임 생성
                email_container = tk.Frame(notebook)
                notebook.add(email_container, text="📧 이메일 설정")
                
                # Canvas와 Scrollbar - 가로/세로 스크롤
                email_canvas = tk.Canvas(email_container, highlightthickness=0)
                email_v_scrollbar = ttk.Scrollbar(email_container, orient="vertical", command=email_canvas.yview)
                email_h_scrollbar = ttk.Scrollbar(email_container, orient="horizontal", command=email_canvas.xview)
                
                # 스크롤 가능한 프레임
                self.email_frame = tk.Frame(email_canvas)
                
                # 프레임을 Canvas에 추가
                email_canvas_window = email_canvas.create_window((0, 0), window=self.email_frame, anchor="nw")
                
                # Frame 크기가 변경될 때 스크롤 영역과 캔버스 윈도우 크기 동기화
                def email_configure_scroll_region(event=None):
                    # 프레임이 완전히 렌더링될 때까지 대기
                    self.email_frame.update_idletasks()
                    # 요구 크기 계산
                    req_w = self.email_frame.winfo_reqwidth()
                    req_h = self.email_frame.winfo_reqheight()
                    # 현재 캔버스 가시 크기 계산
                    canvas_width = email_canvas.winfo_width()
                    canvas_height = email_canvas.winfo_height()
                    # 캔버스 윈도우 크기를 현재 캔버스 크기와 요구 크기 중 큰 값으로 설정
                    email_canvas.itemconfigure(email_canvas_window, 
                                             width=max(canvas_width, req_w), 
                                             height=max(canvas_height, req_h))
                    # 스크롤 영역 갱신
                    email_canvas.configure(scrollregion=email_canvas.bbox("all"))
                
                # 클래스 메서드로 저장하여 다른 함수에서 접근 가능하도록 함
                self.email_configure_scroll_region = email_configure_scroll_region
                
                # Canvas 크기가 변경될 때 Frame 크기와 동기화
                def email_configure_canvas_size(event):
                    # 프레임이 완전히 렌더링될 때까지 대기
                    self.email_frame.update_idletasks()
                    # 요구 크기 재계산
                    req_w = self.email_frame.winfo_reqwidth()
                    req_h = self.email_frame.winfo_reqheight()
                    # 캔버스 윈도우 크기를 이벤트 크기와 요구 크기 중 큰 값으로 설정
                    email_canvas.itemconfigure(email_canvas_window, 
                                             width=max(event.width, req_w), 
                                             height=max(event.height, req_h))
                    # 스크롤 영역 재설정
                    email_canvas.configure(scrollregion=email_canvas.bbox("all"))
                
                self.email_frame.bind("<Configure>", email_configure_scroll_region)
                email_canvas.bind("<Configure>", email_configure_canvas_size)
                email_canvas.configure(yscrollcommand=email_v_scrollbar.set, xscrollcommand=email_h_scrollbar.set)
                
                # 마우스 휠 스크롤
                def email_on_mousewheel(e):
                    email_canvas.yview_scroll(int(-1*(e.delta/120)), "units")
                def email_on_shift_mousewheel(e):
                    email_canvas.xview_scroll(int(-1*(e.delta/120)), "units")
                
                def email_bind_wheel(event):
                    email_canvas.bind_all("<MouseWheel>", email_on_mousewheel)
                    email_canvas.bind_all("<Shift-MouseWheel>", email_on_shift_mousewheel)
                
                def email_unbind_wheel(event):
                    email_canvas.unbind_all("<MouseWheel>")
                    email_canvas.unbind_all("<Shift-MouseWheel>")
                
                email_canvas.bind("<Enter>", email_bind_wheel)
                email_canvas.bind("<Leave>", email_unbind_wheel)
                
                # 배치
                email_canvas.grid(row=0, column=0, sticky="nsew")
                email_v_scrollbar.grid(row=0, column=1, sticky="ns")
                email_h_scrollbar.grid(row=1, column=0, sticky="ew")
                
                email_container.grid_rowconfigure(0, weight=1)
                email_container.grid_columnconfigure(0, weight=1)
                
                self.setup_email_tab(self.email_frame)
                if self.parent_gui:
                    self.parent_gui.log("  ✓ 이메일 설정 탭 완료", is_debug=True)
            except Exception as e:
                error_msg = f"❌ 이메일 설정 탭 UI 생성 오류: {e}"
                if self.parent_gui:
                    self.parent_gui.log(error_msg)
                logging.error(error_msg)
                import traceback
                tb = traceback.format_exc()
                if self.parent_gui:
                    self.parent_gui.log(tb)
                traceback.print_exc()
                messagebox.showerror(
                    "UI 오류", f"이메일 설정 탭 생성 실패:\n{e}", parent=self.dialog)

            # 이메일 양식 탭 (2순위)
            try:
                if self.parent_gui:
                    self.parent_gui.log("  - 이메일 양식 탭 생성 중...", is_debug=True)
                
                # 스크롤 가능한 프레임 생성
                template_container = tk.Frame(notebook)
                notebook.add(template_container, text="📝 이메일 양식")
                
                # Canvas와 Scrollbar
                template_canvas = tk.Canvas(template_container, highlightthickness=0)
                template_v_scrollbar = ttk.Scrollbar(template_container, orient="vertical", command=template_canvas.yview)
                template_h_scrollbar = ttk.Scrollbar(template_container, orient="horizontal", command=template_canvas.xview)
                
                # 스크롤 가능한 프레임
                template_frame = tk.Frame(template_canvas)
                
                # 프레임을 Canvas에 추가
                template_canvas_window = template_canvas.create_window((0, 0), window=template_frame, anchor="nw")
                
                # Frame 크기가 변경될 때 스크롤 영역과 캔버스 윈도우 크기 동기화
                def template_configure_scroll_region(event=None):
                    # 프레임이 완전히 렌더링될 때까지 대기
                    template_frame.update_idletasks()
                    # 요구 크기 계산
                    req_w = template_frame.winfo_reqwidth()
                    req_h = template_frame.winfo_reqheight()
                    # 현재 캔버스 가시 크기 계산
                    canvas_width = template_canvas.winfo_width()
                    canvas_height = template_canvas.winfo_height()
                    # 캔버스 윈도우 크기를 현재 캔버스 크기와 요구 크기 중 큰 값으로 설정
                    template_canvas.itemconfigure(template_canvas_window, 
                                                width=max(canvas_width, req_w), 
                                                height=max(canvas_height, req_h))
                    # 스크롤 영역 갱신
                    template_canvas.configure(scrollregion=template_canvas.bbox("all"))
                
                # Canvas 크기가 변경될 때 Frame 크기와 동기화
                def template_configure_canvas_size(event):
                    # 프레임이 완전히 렌더링될 때까지 대기
                    template_frame.update_idletasks()
                    # 요구 크기 재계산
                    req_w = template_frame.winfo_reqwidth()
                    req_h = template_frame.winfo_reqheight()
                    # 캔버스 윈도우 크기를 이벤트 크기와 요구 크기 중 큰 값으로 설정
                    template_canvas.itemconfigure(template_canvas_window, 
                                                width=max(event.width, req_w), 
                                                height=max(event.height, req_h))
                    # 스크롤 영역 재설정
                    template_canvas.configure(scrollregion=template_canvas.bbox("all"))
                
                template_frame.bind("<Configure>", template_configure_scroll_region)
                template_canvas.bind("<Configure>", template_configure_canvas_size)
                template_canvas.configure(yscrollcommand=template_v_scrollbar.set, xscrollcommand=template_h_scrollbar.set)
                
                # 마우스 휠 스크롤
                def template_on_mousewheel(e):
                    template_canvas.yview_scroll(int(-1*(e.delta/120)), "units")
                def template_on_shift_mousewheel(e):
                    template_canvas.xview_scroll(int(-1*(e.delta/120)), "units")
                
                def template_bind_wheel(event):
                    template_canvas.bind_all("<MouseWheel>", template_on_mousewheel)
                    template_canvas.bind_all("<Shift-MouseWheel>", template_on_shift_mousewheel)
                
                def template_unbind_wheel(event):
                    template_canvas.unbind_all("<MouseWheel>")
                    template_canvas.unbind_all("<Shift-MouseWheel>")
                
                template_canvas.bind("<Enter>", template_bind_wheel)
                template_canvas.bind("<Leave>", template_unbind_wheel)
                
                # 배치
                template_canvas.grid(row=0, column=0, sticky="nsew")
                template_v_scrollbar.grid(row=0, column=1, sticky="ns")
                template_h_scrollbar.grid(row=1, column=0, sticky="ew")
                
                template_container.grid_rowconfigure(0, weight=1)
                template_container.grid_columnconfigure(0, weight=1)
                
                self.setup_template_tab(template_frame)
                if self.parent_gui:
                    self.parent_gui.log("  ✓ 이메일 양식 탭 완료", is_debug=True)
            except Exception as e:
                error_msg = f"❌ 이메일 양식 탭 UI 생성 오류: {e}"
                if self.parent_gui:
                    self.parent_gui.log(error_msg)
                logging.error(error_msg)
                import traceback
                tb = traceback.format_exc()
                if self.parent_gui:
                    self.parent_gui.log(tb)
                traceback.print_exc()
                messagebox.showerror(
                    "UI 오류", f"이메일 양식 탭 생성 실패:\n{e}", parent=self.dialog)

            # 회사 정보 탭 (3순위)
            try:
                if self.parent_gui:
                    self.parent_gui.log("  - 회사 정보 탭 생성 중...", is_debug=True)
                
                # 스크롤 가능한 프레임 생성
                company_container = ttk.Frame(notebook)
                notebook.add(company_container, text="🏢 회사 정보")
                
                # Canvas와 Scrollbar
                company_canvas = tk.Canvas(company_container, highlightthickness=0)
                company_v_scrollbar = ttk.Scrollbar(company_container, orient="vertical", command=company_canvas.yview)
                company_h_scrollbar = ttk.Scrollbar(company_container, orient="horizontal", command=company_canvas.xview)
                
                # 스크롤 가능한 프레임
                company_frame = ttk.Frame(company_canvas, padding="10")
                
                # 프레임을 Canvas에 추가
                company_canvas_window = company_canvas.create_window((0, 0), window=company_frame, anchor="nw")
                
                # Frame 크기가 변경될 때 스크롤 영역과 캔버스 윈도우 크기 동기화
                def company_configure_scroll_region(event=None):
                    # 프레임이 완전히 렌더링될 때까지 대기
                    company_frame.update_idletasks()
                    # 요구 크기 계산
                    req_w = company_frame.winfo_reqwidth()
                    req_h = company_frame.winfo_reqheight()
                    # 현재 캔버스 가시 크기 계산
                    canvas_width = company_canvas.winfo_width()
                    canvas_height = company_canvas.winfo_height()
                    # 캔버스 윈도우 크기를 현재 캔버스 크기와 요구 크기 중 큰 값으로 설정
                    company_canvas.itemconfigure(company_canvas_window, 
                                               width=max(canvas_width, req_w), 
                                               height=max(canvas_height, req_h))
                    # 스크롤 영역 갱신
                    company_canvas.configure(scrollregion=company_canvas.bbox("all"))
                
                # Canvas 크기가 변경될 때 Frame 크기와 동기화
                def company_configure_canvas_size(event):
                    # 프레임이 완전히 렌더링될 때까지 대기
                    company_frame.update_idletasks()
                    # 요구 크기 재계산
                    req_w = company_frame.winfo_reqwidth()
                    req_h = company_frame.winfo_reqheight()
                    # 캔버스 윈도우 크기를 이벤트 크기와 요구 크기 중 큰 값으로 설정
                    company_canvas.itemconfigure(company_canvas_window, 
                                               width=max(event.width, req_w), 
                                               height=max(event.height, req_h))
                    # 스크롤 영역 재설정
                    company_canvas.configure(scrollregion=company_canvas.bbox("all"))
                
                company_frame.bind("<Configure>", company_configure_scroll_region)
                company_canvas.bind("<Configure>", company_configure_canvas_size)
                company_canvas.configure(yscrollcommand=company_v_scrollbar.set, xscrollcommand=company_h_scrollbar.set)
                
                # 마우스 휠 스크롤
                def company_on_mousewheel(e):
                    company_canvas.yview_scroll(int(-1*(e.delta/120)), "units")
                def company_on_shift_mousewheel(e):
                    company_canvas.xview_scroll(int(-1*(e.delta/120)), "units")
                
                def company_bind_wheel(event):
                    company_canvas.bind_all("<MouseWheel>", company_on_mousewheel)
                    company_canvas.bind_all("<Shift-MouseWheel>", company_on_shift_mousewheel)
                
                def company_unbind_wheel(event):
                    company_canvas.unbind_all("<MouseWheel>")
                    company_canvas.unbind_all("<Shift-MouseWheel>")
                
                company_canvas.bind("<Enter>", company_bind_wheel)
                company_canvas.bind("<Leave>", company_unbind_wheel)
                
                # 배치
                company_canvas.grid(row=0, column=0, sticky="nsew")
                company_v_scrollbar.grid(row=0, column=1, sticky="ns")
                company_h_scrollbar.grid(row=1, column=0, sticky="ew")
                
                company_container.grid_rowconfigure(0, weight=1)
                company_container.grid_columnconfigure(0, weight=1)
                
                self.setup_company_tab(company_frame)
                if self.parent_gui:
                    self.parent_gui.log("  ✓ 회사 정보 탭 완료", is_debug=True)
            except Exception as e:
                error_msg = f"❌ 회사 정보 탭 UI 생성 오류: {e}"
                if self.parent_gui:
                    self.parent_gui.log(error_msg)
                logging.error(error_msg)
                import traceback
                tb = traceback.format_exc()
                if self.parent_gui:
                    self.parent_gui.log(tb)
                traceback.print_exc()
                messagebox.showerror(
                    "UI 오류", f"회사 정보 탭 생성 실패:\n{e}", parent=self.dialog)

            # 고급 설정 탭 (4순위)
            try:
                if self.parent_gui:
                    self.parent_gui.log("  - 고급 설정 탭 생성 중...", is_debug=True)
                
                # 스크롤 가능한 프레임 생성
                advanced_container = ttk.Frame(notebook)
                notebook.add(advanced_container, text="⚙️ 고급 설정")
                
                # Canvas와 Scrollbar
                canvas = tk.Canvas(advanced_container, highlightthickness=0)
                self.advanced_canvas = canvas  # 스크롤 영역 업데이트를 위해 참조 저장
                v_scrollbar = ttk.Scrollbar(advanced_container, orient="vertical", command=canvas.yview)
                h_scrollbar = ttk.Scrollbar(advanced_container, orient="horizontal", command=canvas.xview)
                
                # 스크롤 가능한 프레임
                self.scrollable_frame = tk.Frame(canvas, padx=10, pady=10)
                
                # 프레임을 Canvas에 추가
                canvas_window = canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
                
                # Frame 크기가 변경될 때 스크롤 영역과 캔버스 윈도우 크기 동기화
                def configure_scroll_region(event=None):
                    # 프레임이 완전히 렌더링될 때까지 대기
                    self.scrollable_frame.update_idletasks()
                    # 요구 크기 계산
                    req_w = self.scrollable_frame.winfo_reqwidth()
                    req_h = self.scrollable_frame.winfo_reqheight()
                    # 현재 캔버스 가시 크기 계산
                    canvas_width = canvas.winfo_width()
                    canvas_height = canvas.winfo_height()
                    # 캔버스 윈도우 크기를 현재 캔버스 크기와 요구 크기 중 큰 값으로 설정
                    canvas.itemconfigure(canvas_window, 
                                      width=max(canvas_width, req_w), 
                                      height=max(canvas_height, req_h))
                    # 스크롤 영역 갱신
                    canvas.configure(scrollregion=canvas.bbox("all"))
                
                # 클래스 메서드로 저장하여 다른 함수에서 접근 가능하도록 함
                self.advanced_configure_scroll_region = configure_scroll_region
                
                # Canvas 크기가 변경될 때 Frame 크기와 동기화
                def configure_canvas_size(event):
                    # 프레임이 완전히 렌더링될 때까지 대기
                    self.scrollable_frame.update_idletasks()
                    # 요구 크기 재계산
                    req_w = self.scrollable_frame.winfo_reqwidth()
                    req_h = self.scrollable_frame.winfo_reqheight()
                    # 캔버스 윈도우 크기를 이벤트 크기와 요구 크기 중 큰 값으로 설정
                    canvas.itemconfigure(canvas_window, 
                                      width=max(event.width, req_w), 
                                      height=max(event.height, req_h))
                    # 스크롤 영역 재설정
                    canvas.configure(scrollregion=canvas.bbox("all"))
                
                self.scrollable_frame.bind("<Configure>", configure_scroll_region)
                canvas.bind("<Configure>", configure_canvas_size)
                canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
                
                # 마우스 휠 스크롤
                def _on_mousewheel(e):
                    canvas.yview_scroll(int(-1*(e.delta/120)), "units")
                def _on_shift_mousewheel(e):
                    canvas.xview_scroll(int(-1*(e.delta/120)), "units")
                
                def bind_wheel(event):
                    canvas.bind_all("<MouseWheel>", _on_mousewheel)
                    canvas.bind_all("<Shift-MouseWheel>", _on_shift_mousewheel)
                
                def unbind_wheel(event):
                    canvas.unbind_all("<MouseWheel>")
                    canvas.unbind_all("<Shift-MouseWheel>")
                
                canvas.bind("<Enter>", bind_wheel)
                canvas.bind("<Leave>", unbind_wheel)
                
                # 배치
                canvas.grid(row=0, column=0, sticky="nsew")
                v_scrollbar.grid(row=0, column=1, sticky="ns")
                h_scrollbar.grid(row=1, column=0, sticky="ew")
                
                advanced_container.grid_rowconfigure(0, weight=1)
                advanced_container.grid_columnconfigure(0, weight=1)
                
                self.setup_advanced_tab(self.scrollable_frame)
                if self.parent_gui:
                    self.parent_gui.log("  ✓ 고급 설정 탭 완료", is_debug=True)
            except Exception as e:
                error_msg = f"❌ 고급 설정 탭 UI 생성 오류: {e}"
                if self.parent_gui:
                    self.parent_gui.log(error_msg)
                logging.error(error_msg)
                import traceback
                tb = traceback.format_exc()
                if self.parent_gui:
                    self.parent_gui.log(tb)
                traceback.print_exc()
                messagebox.showerror(
                    "UI 오류", f"고급 설정 탭 생성 실패:\n{e}", parent=self.dialog)

            if self.parent_gui:
                self.parent_gui.log("✅ 설정 대화상자 UI 구성 완료!", is_debug=True)

        except Exception as e:
            error_msg = f"❌ 설정 UI 생성 오류: {e}"
            if self.parent_gui:
                self.parent_gui.log(error_msg)
            logging.error(error_msg)
            import traceback
            tb = traceback.format_exc()
            if self.parent_gui:
                self.parent_gui.log(tb)
            traceback.print_exc()
            messagebox.showerror(
                "UI 오류", f"설정 창 생성 실패:\n{e}", parent=self.dialog)

    def setup_email_tab(self, parent):
        """이메일 설정 탭"""
        # parent에 최소 너비 설정
        parent.update_idletasks()
        
        # 입력 필드 프레임
        input_frame = tk.Frame(parent)
        input_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

        # 이메일 서비스 선택
        tk.Label(input_frame, text="이메일 서비스:").grid(
            row=0, column=0, sticky=tk.W, pady=5, padx=(0, 10))
        
        self.email_service_var = tk.StringVar()
        email_service_combo = ttk.Combobox(
            input_frame, 
            textvariable=self.email_service_var,
            values=["Gmail (TLS)", "Gmail (SSL)", "Naver", "Daum", "Outlook", "직접 입력"],
            state='readonly',
            width=10
        )
        email_service_combo.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5)
        email_service_combo.bind('<<ComboboxSelected>>', self.on_email_service_changed)
        
        # 현재 설정값으로 초기화
        current_server = self.config_manager.get('email.smtp_server', '')
        current_port = self.config_manager.get('email.smtp_port', 587)
        self.detect_email_service(current_server, current_port)

        # SMTP 서버
        tk.Label(input_frame, text="SMTP 서버:").grid(
            row=1, column=0, sticky=tk.W, pady=5, padx=(0, 10))
        self.smtp_server_var = tk.StringVar(value=current_server)
        self.smtp_server_entry = ttk.Entry(input_frame, textvariable=self.smtp_server_var, state='readonly')
        self.smtp_server_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5)

        # SMTP 포트
        tk.Label(input_frame, text="SMTP 포트:").grid(
            row=2, column=0, sticky=tk.W, pady=5, padx=(0, 10))
        self.smtp_port_var = tk.StringVar(value=str(current_port))
        self.smtp_port_entry = ttk.Entry(input_frame, textvariable=self.smtp_port_var, state='readonly')
        self.smtp_port_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5)

        # 발신 이메일
        tk.Label(input_frame, text="발신 이메일:").grid(
            row=3, column=0, sticky=tk.W, pady=5, padx=(0, 10))
        self.sender_email_var = tk.StringVar(
            value=self.config_manager.get('email.sender_email', ''))
        ttk.Entry(input_frame, textvariable=self.sender_email_var).grid(
            row=3, column=1, sticky=(tk.W, tk.E), pady=5)

        # 앱 비밀번호
        tk.Label(input_frame, text="앱 비밀번호:").grid(
            row=4, column=0, sticky=tk.W, pady=5, padx=(0, 10))
        self.sender_password_var = tk.StringVar(
            value=self.config_manager.get('email.sender_password', ''))
        password_entry = ttk.Entry(
            input_frame, textvariable=self.sender_password_var, show='*')
        password_entry.grid(row=4, column=1, sticky=(tk.W, tk.E), pady=5)

        # 비밀번호 표시
        show_pass_var = tk.BooleanVar()
        ttk.Checkbutton(input_frame, text="비밀번호 표시", variable=show_pass_var,
                       command=lambda: password_entry.config(show='' if show_pass_var.get() else '*')).grid(row=5, column=1, sticky=tk.W)

        # 그리드 컬럼 가중치 설정 (두 번째 컬럼이 늘어나도록)
        input_frame.columnconfigure(1, weight=1)

        # 버튼 프레임
        btn_frame = tk.Frame(parent)
        btn_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)

        ttk.Button(btn_frame, text="🔌 연결 테스트",
                   command=self.test_connection).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="🔄 이메일 설정 초기화",
                   command=self.reset_email_settings).pack(side=tk.LEFT, padx=5)

        # 연결 테스트 결과 표시 레이블
        self.test_result_label = tk.Label(parent, text="", font=(
            '맑은 고딕', 9, 'bold'), anchor=tk.W, justify=tk.LEFT,
            bg="#f0f0f0", relief="sunken", bd=1)
        self.test_result_label.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 10))

        # 도움말 프레임
        help_frame = tk.LabelFrame(parent, text="📖 설정 도움말")
        help_frame.grid(row=3, column=0, sticky="nsew", padx=10, pady=10)

        help_text = scrolledtext.ScrolledText(
            help_frame, wrap=tk.WORD, font=('맑은 고딕', 9), cursor="arrow")
        help_text.pack(fill=tk.BOTH, expand=True)
        help_text.pack_propagate(False)
        

        # parent 컨테이너의 그리드 설정
        parent.grid_rowconfigure(0, weight=0)  # input_frame
        parent.grid_rowconfigure(1, weight=0)  # btn_frame  
        parent.grid_rowconfigure(2, weight=0)  # test_result_label
        parent.grid_rowconfigure(3, weight=1, minsize=100)  # help_frame (확장)
        parent.grid_columnconfigure(0, weight=1, minsize=50)

        help_content = """📧 SMTP 설정이란?

SMTP는 "Simple Mail Transfer Protocol"의 줄임말로, 이메일을 자동으로 보내기 위해 반드시 필요한 설정입니다.

🎯 빠른 시작 가이드
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

1. 위에서 "이메일 서비스"를 선택하세요 (Gmail, Naver, Daum, Outlook 등)
2. 발신 이메일 주소를 입력하세요
3. 앱 비밀번호를 입력하세요 (아래 서비스별 가이드 참고)
4. "연결 테스트" 버튼을 눌러 확인하세요
5. 성공하면 "저장" 버튼을 누르세요!

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 📧 Gmail (지메일) 설정 방법
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

■ 1단계: 2단계 인증 활성화
   → https://myaccount.google.com/security
  → "2단계 인증" 클릭하여 활성화
  → 보안을 위해 반드시 필요합니다

■ 2단계: 앱 비밀번호 생성
   → https://myaccount.google.com/apppasswords
  → "앱 선택" → "기타(맞춤 이름)" 선택
  → "PDF 이메일 발송 프로그램" 등으로 이름 입력
  → 16자리 비밀번호가 생성됩니다

■ 3단계: 프로그램에 입력
   • 이메일 서비스: "Gmail (TLS)" 또는 "Gmail (SSL)" 선택
   • 발신 이메일: your-email@gmail.com
   • 앱 비밀번호: 위에서 생성한 16자리 코드 입력
   • "연결 테스트" 버튼으로 확인!

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 📧 Naver (네이버) 설정 방법
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

■ 1단계: SMTP 설정 활성화
   → 네이버 메일(https://mail.naver.com) 접속
   → 우측 상단 톱니바퀴(⚙️) → "환경설정" 클릭
   → 좌측 메뉴에서 "POP3/IMAP 설정" 탭 선택
   → 상단의 "IMAP/SMTP 설정" 탭으로 이동
   → "IMAP/SMTP 사용" → "사용함" 선택
   → 하단의 "저장" 버튼 클릭

■ 2단계: 2단계 인증 설정 (필수)
   → 네이버 보안설정(https://nid.naver.com/user2/help/myInfoV2) 접속
   → "2단계 인증" 설정 활성화
   → SMS 또는 OTP 앱으로 인증 설정 완료

■ 3단계: 프로그램에 입력
   • 이메일 서비스: "Naver" 선택
   • 발신 이메일: your-id@naver.com
   • 앱 비밀번호: 네이버 계정 비밀번호 입력
   • "연결 테스트" 버튼으로 확인!

💡 팁: 네이버는 일반 계정 비밀번호를 사용하며, 2단계 인증 설정이 필수입니다!

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 📧 Daum (다음/한메일) 설정 방법
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

■ 1단계: SMTP 설정 활성화
   → 다음 메일(https://mail.daum.net) 접속
   → 우측 상단 "설정" 클릭
   → 좌측 메뉴에서 "IMAP/POP3" 탭 선택
   → "IMAP" 탭으로 이동
   → "IMAP / SMTP 사용" 토글을 켜기
   → 2단계 인증 안내 팝업이 표시됩니다

■ 2단계: 2단계 인증 설정 (필수)
   → 팝업에서 "앱 비밀번호 확인하기" 클릭 (또는 아래 링크)
   → 카카오 계정 보안(https://accounts.kakao.com/weblogin/account/security) 접속
   → 좌측 메뉴에서 "계정 보안" → "2단계 인증" 클릭
   → 2단계 인증 활성화 (SMS 또는 OTP)

■ 3단계: 앱 비밀번호 생성
   → 카카오 계정 보안 페이지에서 "비밀번호 변경" 클릭
   → 하단의 "앱 비밀번호" 메뉴 선택
   → "앱 이름"에 "PDF발송프로그램" 등 입력
   → "생성" 버튼 클릭
   → 생성된 앱 비밀번호 복사 (한 번만 표시됨!)

■ 4단계: 프로그램에 입력
   • 이메일 서비스: "Daum" 선택
   • 발신 이메일: your-id@hanmail.net 또는 your-id@daum.net
   • 앱 비밀번호: 위에서 생성한 앱 비밀번호 입력
   • "연결 테스트" 버튼으로 확인!

💡 팁: 다음은 반드시 앱 비밀번호를 사용해야 하며, 2단계 인증 설정이 필수입니다!

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 📧 Outlook (아웃룩/Hotmail) 설정 방법
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

■ 1단계: 2단계 인증 활성화 (선택사항이지만 권장)
   → https://account.microsoft.com/security
   → "2단계 인증" 설정

■ 2단계: 앱 비밀번호 생성 (2단계 인증 활성화 시)
   → https://account.microsoft.com/security
   → "앱 암호" 클릭
   → 새 앱 암호 생성
   → 생성된 비밀번호 복사

■ 3단계: 프로그램에 입력
   • 이메일 서비스: "Outlook" 선택
   • 발신 이메일: your-email@outlook.com 또는 @hotmail.com
   • 앱 비밀번호: 앱 암호 (또는 계정 비밀번호)
   • "연결 테스트" 버튼으로 확인!

💡 팁: 2단계 인증을 사용하지 않으면 일반 계정 비밀번호를 사용하세요!

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 ⚙️ 각 서비스별 설정값 요약
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  Gmail (TLS):    smtp.gmail.com:587    (앱 비밀번호 필수, 2단계 인증 필수)
  Gmail (SSL):    smtp.gmail.com:465    (앱 비밀번호 필수, 2단계 인증 필수)
  Naver:          smtp.naver.com:587    (계정 비밀번호, 2단계 인증 필수)
  Daum:           smtp.daum.net:465     (앱 비밀번호 필수, 2단계 인증 필수)
  Outlook:        smtp-mail.outlook.com:587  (앱 암호 권장, 2단계 인증 권장)

💡 이메일 서비스를 선택하면 서버와 포트가 자동으로 설정됩니다!

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 🔧 문제 해결
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

❌ "인증 실패" 오류가 나요
   → 이메일 주소를 다시 확인하세요
   → Gmail/Daum: 반드시 앱 비밀번호를 사용해야 합니다
   → Naver: 일반 계정 비밀번호를 사용하세요
   → 모든 서비스: 2단계 인증이 설정되어 있는지 확인하세요
   → Naver/Daum: SMTP 사용 설정이 활성화되어 있는지 확인하세요

❌ "연결 시간 초과" 오류가 나요
   → 인터넷 연결을 확인하세요
   → 방화벽에서 프로그램을 허용했는지 확인하세요
   → 이메일 서비스 선택이 올바른지 확인하세요

❌ "서버를 찾을 수 없음" 오류가 나요
   → 이메일 서비스를 다시 선택해 보세요
   → "직접 입력"을 선택한 경우 서버 주소를 확인하세요

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 ⚠️ 보안 주의사항
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

• 앱 비밀번호는 한 번만 표시되니 반드시 복사해서 저장하세요
• 앱 비밀번호를 잃어버리면 새로 생성해야 합니다
• 비밀번호는 안전한 곳에 보관하세요
• 설정을 변경한 후에는 반드시 "저장" 버튼을 눌러주세요
• 이 프로그램은 비밀번호를 암호화하여 로컬에만 저장합니다"""

        help_text.insert('1.0', help_content)

        # 링크 태그 설정
        help_text.tag_config('link', foreground='blue', underline=True)

        # URL 찾아서 태그 추가 및 클릭 이벤트 바인딩
        import re
        for match in re.finditer(r'https?://[^\s]+', help_content):
            start_idx = f"1.0+{match.start()}c"
            end_idx = f"1.0+{match.end()}c"
            help_text.tag_add('link', start_idx, end_idx)
            help_text.tag_bind('link', '<Button-1>', lambda e,
                               url=match.group(): webbrowser.open(url))
            help_text.tag_bind('link', '<Enter>',
                               lambda e: help_text.config(cursor='hand2'))
            help_text.tag_bind('link', '<Leave>',
                               lambda e: help_text.config(cursor='arrow'))

        help_text.config(state=tk.DISABLED)

    def setup_company_tab(self, parent):
        """회사 정보 탭"""
        # 설명
        ttk.Label(parent, text="회사별 이메일 주소와 사용할 양식을 설정하세요.",
                 font=('맑은 고딕', 10, 'bold')).grid(row=0, column=0, pady=10)

        # 회사 리스트
        list_frame = ttk.Frame(parent)
        list_frame.grid(row=1, column=0, sticky="nsew", pady=10)

        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.company_listbox = tk.Listbox(
            list_frame, yscrollcommand=scrollbar.set, height=15)
        self.company_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.company_listbox.yview)

        # 회사 목록 로드
        self.refresh_company_list()

        # 버튼
        btn_frame = ttk.Frame(parent)
        btn_frame.grid(row=2, column=0, sticky="nsew")

        ttk.Button(btn_frame, text="추가", command=self.add_company).pack(
            side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="수정", command=self.edit_company).pack(
            side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="삭제", command=self.delete_company).pack(
            side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="📖 사용방법", command=self.show_company_help).pack(
            side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="🔄 회사 정보 초기화",
                   command=self.reset_company_info).pack(side=tk.RIGHT, padx=5)
        
        # parent 컨테이너의 그리드 설정
        parent.grid_rowconfigure(0, weight=0)  # 설명 (고정)
        parent.grid_rowconfigure(1, weight=1)  # list_frame (확장)
        parent.grid_rowconfigure(2, weight=0)  # btn_frame (고정)
        parent.grid_columnconfigure(0, weight=1)

    def setup_template_tab(self, parent):
        """이메일 양식 탭"""
        # 양식 리스트
        list_frame = tk.Frame(parent)
        list_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.template_listbox = tk.Listbox(
            list_frame, yscrollcommand=scrollbar.set, height=8)
        self.template_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.template_listbox.yview)
        self.template_listbox.bind(
            '<<ListboxSelect>>', lambda e: self.on_template_select())

        # 양식 목록 로드
        self.refresh_template_list()

        # 버튼
        btn_frame = tk.Frame(parent)
        btn_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)

        ttk.Button(btn_frame, text="➕ 추가", command=self.add_template).pack(
            side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="✏️ 수정", command=self.edit_template).pack(
            side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="❌ 삭제", command=self.delete_template).pack(
            side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="🔧 커스텀 변수", command=self.manage_custom_variables).pack(
            side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="📖 사용방법", command=self.show_template_help).pack(
            side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="🔄 초기화", command=self.reset_templates).pack(
            side=tk.RIGHT, padx=5)

        # 미리보기
        preview_frame = tk.LabelFrame(parent, text="미리보기")
        preview_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=5)

        ttk.Label(preview_frame, text="제목:").pack(anchor=tk.W)
        self.preview_subject = tk.Text(
            preview_frame, height=2, wrap=tk.WORD, state='disabled')
        self.preview_subject.pack(fill=tk.X, pady=2)

        ttk.Label(preview_frame, text="본문:").pack(anchor=tk.W, pady=(5, 0))
        self.preview_body = scrolledtext.ScrolledText(
            preview_frame, height=6, wrap=tk.WORD, state='disabled')
        self.preview_body.pack(fill=tk.BOTH, expand=True, pady=2)

        # 변수 안내 (토글 가능)
        self.setup_variable_info(parent, row=3)
        
        # parent 컨테이너의 그리드 설정
        parent.grid_rowconfigure(0, weight=1)  # list_frame (확장)
        parent.grid_rowconfigure(1, weight=0)  # btn_frame (고정)
        parent.grid_rowconfigure(2, weight=1)  # preview_frame (확장)
        parent.grid_rowconfigure(3, weight=0)  # variable_info (고정)
        parent.grid_columnconfigure(0, weight=1)

    def setup_variable_info(self, parent, row):
        """변수 안내 설정"""
        # 변수 확인 프레임 (외부 컨테이너)
        var_check_frame = ttk.LabelFrame(parent, text="사용 가능한 변수")
        var_check_frame.grid(row=row, column=0, sticky="nsew", padx=10, pady=5)
        
        # 프레임 내부 그리드 설정 (좌우 배치)
        var_check_frame.grid_columnconfigure(0, weight=0)  # 버튼 (고정 크기)
        var_check_frame.grid_columnconfigure(1, weight=1)  # 사용 가능한 변수 프레임 (확장 가능)
        var_check_frame.grid_rowconfigure(0, weight=1)
        
        # 모든 변수 확인하기 버튼 (왼쪽)
        check_vars_btn = ttk.Button(var_check_frame, text="모든 변수\n확인하기", 
                                  command=self.show_all_variables)
        check_vars_btn.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)

        # 사용 가능한 변수 프레임 (오른쪽)
        var_frame = ttk.LabelFrame(var_check_frame, text="기본 제공 변수")
        var_frame.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)
        
        # 변수 목록 텍스트 (부모 크기에 맞춰 자동 줄바꿈)
        var_content = "{회사명}, {파일명}, {날짜}, {시간}, {년}, {월}, {일}, {요일}, {요일한글}, {시}, {분}, {초}, {시간12}, {오전오후}"
        info_text = tk.Label(var_frame, text=var_content, 
                           font=('맑은 고딕', 9), fg='#666666',
                           anchor='nw', justify='left')
        info_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 창/부모 폭 변화에 맞춰 wraplength(픽셀 단위) 갱신
        def sync_wrap(event):
            # 좌우 패딩(여기서는 5+5)을 고려해서 약간 빼줌
            wrap = max(0, event.width - 10)
            info_text.configure(wraplength=wrap)
        info_text.bind("<Configure>", sync_wrap)
        
        # parent 그리드 설정 업데이트
        parent.grid_rowconfigure(row, weight=0)

    def show_all_variables(self):
        """모든 변수 확인 창 표시"""
        # 새 창 생성
        var_window = tk.Toplevel(self.dialog)
        var_window.title("모든 변수 목록")
        var_window.resizable(True, True)
        
        # 중앙 위치 설정
        center_window(var_window, self.dialog, 500, 400)
        var_window.grab_set()
        
        # 스크롤 가능한 텍스트 위젯
        text_frame = tk.Frame(var_window)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        text_widget = tk.Text(text_frame, font=('맑은 고딕', 10), fg='#333333',
                             wrap=tk.WORD, state=tk.DISABLED)
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 변수 목록 내용
        content = """📋 기본 제공 변수:

{회사명} - 회사명
{파일명} - PDF 파일명
{날짜} - 현재 날짜 (YYYY-MM-DD)
{시간} - 현재 시간 (HH:MM:SS)
{년} - 년도 (4자리)
{월} - 월 (1-12)
{일} - 일 (1-31)
{요일} - 요일 (영문)
{요일한글} - 요일 (한글)
{시} - 시 (0-23)
{분} - 분 (0-59)
{초} - 초 (0-59)
{시간12} - 12시간 형식 시간
{오전오후} - 오전/오후

🔧 커스텀 변수:
"""
        
        # 커스텀 변수 추가
        custom_vars = self.config_manager.get('custom_variables', {})
        if custom_vars:
            for var_name, var_value in custom_vars.items():
                content += f"{{{var_name}}} - {var_value}\n"
        else:
            content += "등록된 커스텀 변수가 없습니다.\n'🔧 커스텀 변수' 버튼을 눌러 추가하세요."
        
        # 텍스트 위젯에 내용 삽입
        text_widget.config(state=tk.NORMAL)
        text_widget.insert(tk.END, content)
        text_widget.config(state=tk.DISABLED)
        
        # 닫기 버튼
        close_btn = ttk.Button(var_window, text="닫기", command=var_window.destroy)
        close_btn.pack(pady=10)

    def setup_advanced_tab(self, parent):
        """고급 설정 탭"""
        # PDF 파일명 패턴
        ttk.Label(parent, text="PDF 파일명 인식 패턴 (정규식):").pack(
            anchor=tk.W, pady=5, padx=10)
        self.pattern_var = tk.StringVar(
            value=self.config_manager.get('pattern', ''))
        ttk.Entry(parent, textvariable=self.pattern_var,
                  width=70).pack(fill=tk.X, pady=5, padx=10)

        ttk.Label(parent, text="예: ^([가-힣A-Za-z0-9\\s]+?)(?:___|\.pdf$)",
                 foreground='gray').pack(anchor=tk.W, pady=2, padx=10)

        # 타임아웃 설정
        ttk.Label(parent, text="자동 실행 대기 시간 (초):").pack(
            anchor=tk.W, pady=(20, 5), padx=10)

        timeout_frame = ttk.Frame(parent)
        timeout_frame.pack(fill=tk.X, pady=5, padx=10)

        ttk.Label(timeout_frame, text="모드 선택:").pack(
            side=tk.LEFT, padx=(0, 10))
        self.auto_select_var = tk.StringVar(
            value=str(self.config_manager.get('auto_select_timeout', 10)))
        ttk.Entry(timeout_frame, textvariable=self.auto_select_var,
                  width=10).pack(side=tk.LEFT)

        ttk.Label(timeout_frame, text="발송 확인:").pack(
            side=tk.LEFT, padx=(20, 10))
        self.auto_send_var = tk.StringVar(
            value=str(self.config_manager.get('auto_send_timeout', 10)))
        ttk.Entry(timeout_frame, textvariable=self.auto_send_var,
                  width=10).pack(side=tk.LEFT)

        ttk.Label(parent, text="* 0: 즉시 실행, 음수: 비활성화",
                 foreground='gray').pack(anchor=tk.W, pady=2, padx=10)

        # 이메일 발송 타임아웃 설정
        ttk.Label(parent, text="이메일 발송 최대 대기 시간 (초):").pack(
            anchor=tk.W, pady=(20, 5), padx=10)

        email_timeout_frame = ttk.Frame(parent)
        email_timeout_frame.pack(fill=tk.X, pady=5, padx=10)

        self.email_send_timeout_var = tk.StringVar(
            value=str(self.config_manager.get('email_send_timeout', 180)))
        ttk.Entry(email_timeout_frame, textvariable=self.email_send_timeout_var,
                  width=10).pack(side=tk.LEFT)
        ttk.Label(email_timeout_frame, text="초 (기본값: 180초 = 3분)",
                 foreground='gray').pack(side=tk.LEFT, padx=(10, 0))

        ttk.Label(parent, text="* 이메일 발송이 이 시간을 초과하면 자동으로 중단됩니다",
                 foreground='gray').pack(anchor=tk.W, pady=2, padx=10)

        # 버전 정보
        ttk.Separator(parent, orient='horizontal').pack(
            fill=tk.X, pady=20, padx=10)
        version_frame = ttk.Frame(parent)
        version_frame.pack(anchor=tk.W, pady=5, padx=10)
        ttk.Label(version_frame, text="📌 버전 정보:", font=(
            '맑은 고딕', 9, 'bold')).pack(side=tk.LEFT)
        ttk.Label(version_frame, text=f"v{VERSION}",
                 foreground='blue', font=('맑은 고딕', 9)).pack(side=tk.LEFT, padx=5)

        # 디버그 모드
        ttk.Label(parent, text="개발자 옵션:").pack(
            anchor=tk.W, pady=(20, 5), padx=10)

        self.debug_mode_var = tk.BooleanVar(
            value=self.config_manager.get('debug_mode', False))
        ttk.Checkbutton(parent, text="🐛 디버그 모드 (UI 구성 로그 표시)",
                       variable=self.debug_mode_var).pack(anchor=tk.W, pady=5, padx=10)

        ttk.Label(parent, text="* 디버그 모드 활성화 시 상세한 UI 구성 로그가 표시됩니다",
                 foreground='gray').pack(anchor=tk.W, pady=2, padx=10)

        # 구분선
        ttk.Separator(parent, orient='horizontal').pack(
            fill=tk.X, pady=20, padx=10)

        # 글자 크기 설정 (최하단)
        ttk.Label(parent, text="🔤 프로그램 글자 크기:", font=('맑은 고딕', 9, 'bold')).pack(
            anchor=tk.W, pady=(5, 5), padx=10)
        
        font_size_frame = ttk.Frame(parent)
        font_size_frame.pack(fill=tk.X, pady=5, padx=10)
        
        ttk.Label(font_size_frame, text="크기:").pack(side=tk.LEFT, padx=(0, 10))
        
        self.font_size_var = tk.StringVar(
            value=str(self.config_manager.get('ui.font_size', 9)))
        
        font_size_spinbox = ttk.Spinbox(
            font_size_frame, 
            from_=8, 
            to=16, 
            textvariable=self.font_size_var,
            width=5,
            command=self.update_font_preview
        )
        font_size_spinbox.pack(side=tk.LEFT)
        font_size_spinbox.bind('<KeyRelease>', lambda e: self.update_font_preview())
        
        ttk.Label(font_size_frame, text="pt (8~16pt, 기본값: 9pt)",
                 foreground='gray').pack(side=tk.LEFT, padx=(10, 0))
        
        # 미리보기 프레임
        preview_frame = ttk.LabelFrame(parent, text="📋 미리보기", padding="10")
        preview_frame.pack(fill=tk.X, pady=10, padx=10)
        
        self.font_preview_label = tk.Label(
            preview_frame, 
            text="이 텍스트로 크기를 확인해 보세요 (The quick brown fox)",
            font=('맑은 고딕', int(self.font_size_var.get())),
            fg='#333333',
            wraplength=400,  # 텍스트 줄바꿈 설정
            justify=tk.CENTER
        )
        self.font_preview_label.pack(pady=5)
        
        ttk.Label(parent, text="* 글자 크기 변경 후 프로그램을 재시작하면 적용됩니다",
                 foreground='orange', font=('맑은 고딕', 8)).pack(anchor=tk.W, pady=2, padx=10)

        # 버튼 프레임
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill=tk.X, pady=10, padx=10)

        ttk.Button(btn_frame, text="📖 사용방법",
                   command=self.show_advanced_help).pack(side=tk.LEFT)
        ttk.Button(btn_frame, text="🔄 고급 설정 초기화",
                   command=self.reset_advanced_settings).pack(side=tk.RIGHT)
        
                    

    def update_font_preview(self):
        """글자 크기 미리보기 업데이트"""
        try:
            font_size = int(self.font_size_var.get())
            if 8 <= font_size <= 16:
                self.font_preview_label.config(font=('맑은 고딕', font_size))
                # 스크롤 영역 강제 업데이트
                self.dialog.update_idletasks()
                # 고급 설정 탭의 스크롤 영역 업데이트
                if hasattr(self, 'advanced_canvas'):
                    # scrollable_frame 크기 강제 갱신
                    self.advanced_canvas.update_idletasks()
                    self.advanced_canvas.configure(scrollregion=self.advanced_canvas.bbox("all"))
                # 동적 크기 변경 후 레이아웃 계산이 반영되도록 콜백 호출
                if hasattr(self, 'advanced_configure_scroll_region'):
                    self.scrollable_frame.after_idle(self.advanced_configure_scroll_region)
        except ValueError:
            pass  # 잘못된 값이면 무시

    def refresh_company_list(self):
        """회사 목록 새로고침"""
        self.company_listbox.delete(0, tk.END)
        companies = self.config_manager.get('companies', {})
        for company_name, info in companies.items():
            emails = ', '.join(info.get('emails', []))
            template = info.get('template', 'A')
            self.company_listbox.insert(
                tk.END, f"{company_name} | {emails} | {template}")

    def add_company(self):
        """회사 추가"""
        try:
            CompanyDialog(self.dialog, self.config_manager, None,
                          self.refresh_company_list, parent_gui=self.parent_gui)
        except Exception as e:
            logging.error(f"회사 추가 오류: {e}")
            messagebox.showerror(
                "오류", f"회사 추가 중 오류가 발생했습니다.\n\n{str(e)}", parent=self.dialog)

    def edit_company(self):
        """회사 수정"""
        try:
            selection = self.company_listbox.curselection()
            if not selection:
                messagebox.showwarning(
                    "선택 없음", "수정할 회사를 선택하세요.", parent=self.dialog)
                return

            company_name = self.company_listbox.get(
                selection[0]).split(' | ')[0]
            CompanyDialog(self.dialog, self.config_manager, company_name,
                          self.refresh_company_list, parent_gui=self.parent_gui)
        except Exception as e:
            logging.error(f"회사 수정 오류: {e}")
            messagebox.showerror(
                "오류", f"회사 수정 중 오류가 발생했습니다.\n\n{str(e)}", parent=self.dialog)

    def delete_company(self):
        """회사 삭제"""
        try:
            selection = self.company_listbox.curselection()
            if not selection:
                messagebox.showwarning(
                    "선택 없음", "삭제할 회사를 선택하세요.", parent=self.dialog)
                return

            company_name = self.company_listbox.get(
                selection[0]).split(' | ')[0]

            if messagebox.askyesno("삭제 확인", f"'{company_name}'을(를) 삭제하시겠습니까?"):
                companies = self.config_manager.get('companies', {})
                if company_name in companies:
                    del companies[company_name]
                    self.config_manager.set('companies', companies)
                    self.refresh_company_list()
        except Exception as e:
            logging.error(f"회사 삭제 오류: {e}")
            messagebox.showerror(
                "오류", f"회사 삭제 중 오류가 발생했습니다.\n\n{str(e, parent=self.dialog)}")

    def refresh_template_list(self):
        """양식 목록 새로고침"""
        self.template_listbox.delete(0, tk.END)
        templates = self.config_manager.get('email_templates', {})
        for template_name in templates.keys():
            if not template_name.startswith('_'):
                self.template_listbox.insert(tk.END, template_name)

    def on_template_select(self):
        """양식 선택 시 미리보기 업데이트"""
        try:
            selection = self.template_listbox.curselection()
            if not selection:
                return

            template_name = self.template_listbox.get(selection[0])
            templates = self.config_manager.get('email_templates', {})
            template = templates.get(template_name, {})

            # 미리보기 업데이트
            self.preview_subject.config(state='normal')
            self.preview_subject.delete('1.0', tk.END)
            self.preview_subject.insert('1.0', template.get('subject', ''))
            self.preview_subject.config(state='disabled')

            self.preview_body.config(state='normal')
            self.preview_body.delete('1.0', tk.END)
            self.preview_body.insert('1.0', template.get('body', ''))
            self.preview_body.config(state='disabled')
        except Exception as e:
            logging.error(f"양식 미리보기 오류: {e}")

    def add_template(self):
        """양식 추가"""
        try:
            TemplateDialog(self.dialog, self.config_manager, None,
                           self.refresh_template_list, parent_gui=self.parent_gui)
        except Exception as e:
            logging.error(f"양식 추가 오류: {e}")
            messagebox.showerror(
                "오류", f"양식 추가 중 오류가 발생했습니다.\n\n{str(e)}", parent=self.dialog)

    def edit_template(self):
        """양식 수정"""
        try:
            selection = self.template_listbox.curselection()
            if not selection:
                messagebox.showwarning(
                    "선택 없음", "수정할 양식을 선택하세요.", parent=self.dialog)
                return

            template_name = self.template_listbox.get(selection[0])
            TemplateDialog(self.dialog, self.config_manager, template_name,
                           self.refresh_template_list, parent_gui=self.parent_gui)
        except Exception as e:
            logging.error(f"양식 수정 오류: {e}")
            messagebox.showerror(
                "오류", f"양식 수정 중 오류가 발생했습니다.\n\n{str(e)}", parent=self.dialog)

    def delete_template(self):
        """양식 삭제"""
        try:
            selection = self.template_listbox.curselection()
            if not selection:
                messagebox.showwarning(
                    "선택 없음", "삭제할 양식을 선택하세요.", parent=self.dialog)
                return

            template_name = self.template_listbox.get(selection[0])

            if messagebox.askyesno("삭제 확인", f"'{template_name}'을(를) 삭제하시겠습니까?"):
                templates = self.config_manager.get('email_templates', {})
                if template_name in templates:
                    del templates[template_name]
                    self.config_manager.set('email_templates', templates)
                    self.refresh_template_list()
                    # 미리보기 클리어
                    self.preview_subject.config(state='normal')
                    self.preview_subject.delete('1.0', tk.END)
                    self.preview_subject.config(state='disabled')
                    self.preview_body.config(state='normal')
                    self.preview_body.delete('1.0', tk.END)
                    self.preview_body.config(state='disabled')
        except Exception as e:
            logging.error(f"양식 삭제 오류: {e}")
            messagebox.showerror(
                "오류", f"양식 삭제 중 오류가 발생했습니다.\n\n{str(e, parent=self.dialog)}")

    def load_template(self):
        """양식 로드 (하위 호환성 유지)"""
        pass

    def detect_email_service(self, server, port):
        """현재 서버/포트로 이메일 서비스 감지"""
        port = int(port) if isinstance(port, str) else port
        
        if server == 'smtp.gmail.com' and port == 587:
            self.email_service_var.set("Gmail (TLS)")
        elif server == 'smtp.gmail.com' and port == 465:
            self.email_service_var.set("Gmail (SSL)")
        elif server == 'smtp.naver.com' and port == 587:
            self.email_service_var.set("Naver")
        elif server == 'smtp.daum.net' and port == 465:
            self.email_service_var.set("Daum")
        elif server == 'smtp-mail.outlook.com' and port == 587:
            self.email_service_var.set("Outlook")
        else:
            self.email_service_var.set("직접 입력")
            # 직접 입력 모드로 전환
            self.smtp_server_entry.config(state='normal')
            self.smtp_port_entry.config(state='normal')

    def on_email_service_changed(self, event=None):
        """이메일 서비스 선택 변경 시 처리"""
        service = self.email_service_var.get()
        
        # 서비스별 설정
        email_services = {
            "Gmail (TLS)": ("smtp.gmail.com", "587"),
            "Gmail (SSL)": ("smtp.gmail.com", "465"),
            "Naver": ("smtp.naver.com", "587"),
            "Daum": ("smtp.daum.net", "465"),
            "Outlook": ("smtp-mail.outlook.com", "587"),
        }
        
        if service == "직접 입력":
            # 입력 가능하게 변경
            self.smtp_server_entry.config(state='normal')
            self.smtp_port_entry.config(state='normal')
            # 기존 값 유지하거나 비우기
            if self.smtp_server_var.get() in [s[0] for s in email_services.values()]:
                self.smtp_server_var.set('')
                self.smtp_port_var.set('')
        else:
            # 읽기 전용으로 변경
            self.smtp_server_entry.config(state='readonly')
            self.smtp_port_entry.config(state='readonly')
            # 선택한 서비스의 설정 적용
            if service in email_services:
                server, port = email_services[service]
                self.smtp_server_var.set(server)
                self.smtp_port_var.set(port)

    def test_connection(self):
        """연결 테스트"""
        # 이전 결과 초기화
        self.test_result_label.config(text="🔄 연결 테스트 중...", fg="#666666")
        self.dialog.update_idletasks()

        smtp_server = self.smtp_server_var.get().strip()
        smtp_port = self.smtp_port_var.get().strip()
        sender_email = self.sender_email_var.get().strip()
        sender_password = self.sender_password_var.get().replace(' ', '').replace('\t', '')

        if not all([smtp_server, smtp_port, sender_email, sender_password]):
            self.test_result_label.config(
                text="⚠️ 모든 항목을 입력해주세요.",
                fg="#FFA500"
            )
            return

        try:
            smtp_port = int(smtp_port)
        except ValueError:
            self.test_result_label.config(
                text="❌ 포트 번호는 숫자여야 합니다.",
                fg="#FF0000"
            )
            return

        try:
            # 포트에 따라 SSL/TLS 선택
            if smtp_port == 465:
                # SSL 연결
                server = smtplib.SMTP_SSL(smtp_server, smtp_port, timeout=10)
            else:
                # TLS 연결
                server = smtplib.SMTP(smtp_server, smtp_port, timeout=10)
            server.starttls()
            
            server.login(sender_email, sender_password)
            server.quit()

            self.test_result_label.config(
                text=f"✅ 연결 성공! ({sender_email})",
                fg="#00AA00"
            )

        except smtplib.SMTPAuthenticationError as e:
            error_msg = f"""❌ 인증 실패: 이메일 주소 또는 앱 비밀번호를 확인하세요.
💡 Gmail은 '앱 비밀번호'가 필요합니다 (일반 비밀번호 X)
   Google 계정 → 보안 → 2단계 인증 활성화 → 앱 비밀번호 생성"""
            self.test_result_label.config(text=error_msg, fg="#FF0000")

        except smtplib.SMTPConnectError as e:
            self.test_result_label.config(
                text=f"❌ 서버 연결 실패: SMTP 서버 주소와 포트를 확인하세요.\n{str(e)}",
                fg="#FF0000"
            )

        except smtplib.SMTPException as e:
            self.test_result_label.config(
                text=f"❌ SMTP 오류: {str(e)}",
                fg="#FF0000"
            )

        except socket.timeout:
            self.test_result_label.config(
                text="❌ 연결 시간 초과: 네트워크 연결과 방화벽 설정을 확인하세요.",
                fg="#FF0000"
            )

        except Exception as e:
            self.test_result_label.config(
                text=f"❌ 연결 실패: {str(e)}",
                fg="#FF0000"
            )

    def save_settings(self):
        """설정 저장"""
        try:
            # 이메일 설정 (메모리에만 반영, 파일 저장은 아직 안함)
            old_email = self.config_manager.config['email']['sender_email']
            old_password = self.config_manager.config['email']['sender_password']
            old_smtp_server = self.config_manager.config['email']['smtp_server']
            old_smtp_port = self.config_manager.config['email']['smtp_port']

            self.config_manager.config['email']['smtp_server'] = self.smtp_server_var.get(
            ).strip()
            self.config_manager.config['email']['smtp_port'] = int(
                self.smtp_port_var.get().strip())
            self.config_manager.config['email']['sender_email'] = self.sender_email_var.get(
            ).strip()
            self.config_manager.config['email']['sender_password'] = self.sender_password_var.get(
            ).replace(' ', '').replace('\t', '')

            # 고급 설정 (메모리에만 반영)
            self.config_manager.config['pattern'] = self.pattern_var.get()
            self.config_manager.config['auto_select_timeout'] = int(
                self.auto_select_var.get())
            self.config_manager.config['auto_send_timeout'] = int(
                self.auto_send_var.get())
            self.config_manager.config['email_send_timeout'] = int(
                self.email_send_timeout_var.get())
            self.config_manager.config['debug_mode'] = self.debug_mode_var.get()
            
            # 글자 크기 설정 저장
            old_font_size = self.config_manager.get('ui.font_size', 9)
            new_font_size = int(self.font_size_var.get())
            self.config_manager.set('ui.font_size', new_font_size)
            font_size_changed = (old_font_size != new_font_size)

            # 회사 정보와 이메일 양식은 이미 config_manager에 반영됨 (실시간 저장)
            # 모든 설정을 한번에 파일로 저장
            self.config_manager.save_config()

            # 이메일 설정이 변경되었는지 확인
            new_email = self.config_manager.config['email']['sender_email']
            new_password = self.config_manager.config['email']['sender_password']
            new_smtp_server = self.config_manager.config['email']['smtp_server']
            new_smtp_port = self.config_manager.config['email']['smtp_port']

            email_changed = (old_email != new_email or
                           old_password != new_password or
                           old_smtp_server != new_smtp_server or
                           old_smtp_port != new_smtp_port)

            # 이메일 설정이 변경되었으면 기존 연결 완전히 제거하고 새로 연결
            if email_changed and self.parent_gui:
                self.parent_gui.log(
                    "📧 이메일 설정이 변경되었습니다. 기존 연결을 완전히 제거합니다.", 'INFO')
                # 기존 연결 완전히 제거
                self.parent_gui.disconnect_smtp()
                self.parent_gui.stop_connection_monitor()
                # 새 연결 시도
                self.parent_gui.check_and_connect_email()

            self.dialog.focus_force()
            
            # 글자 크기 변경 시 재시작 안내
            if font_size_changed:
                messagebox.showinfo(
                    "저장 완료", 
                    "설정이 저장되었습니다!\n\n글자 크기 변경사항을 적용하려면 프로그램을 재시작해 주세요.",
                    parent=self.dialog
                )
            else:
                messagebox.showinfo("저장 완료", "설정이 저장되었습니다!", parent=self.dialog)
            
            self.result = True
            self.dialog.destroy()

        except Exception as e:
            self.dialog.focus_force()
            messagebox.showerror(
                "저장 오류", f"설정 저장 중 오류가 발생했습니다.\n\n{str(e)}", parent=self.dialog)

    def on_cancel(self):
        """취소 버튼 / 창 닫기"""
        # 임시 데이터 복원
        self.config_manager.set('companies', self.temp_companies)
        self.config_manager.set('email_templates', self.temp_templates)
        self.result = False
        self.dialog.destroy()

    def reset_email_settings(self):
        """이메일 설정 초기화"""
        try:
            if not messagebox.askyesno("초기화 확인", "이메일 설정을 기본값으로 초기화하시겠습니까?", parent=self.dialog):
                return

            self.config_manager.set('email.smtp_server', 'smtp.gmail.com')
            self.config_manager.set('email.smtp_port', 587)
            self.config_manager.set('email.sender_email', '')
            self.config_manager.set('email.sender_password', '')

            # UI 업데이트
            self.smtp_server_var.set('smtp.gmail.com')
            self.smtp_port_var.set('587')
            self.sender_email_var.set('')
            self.sender_password_var.set('')

            messagebox.showinfo(
                "초기화 완료", "이메일 설정이 초기화되었습니다.", parent=self.dialog)
        except Exception as e:
            logging.error(f"이메일 설정 초기화 오류: {e}")
            messagebox.showerror(
                "오류", f"초기화 중 오류가 발생했습니다.\n\n{str(e, parent=self.dialog)}")

    def reset_company_info(self):
        """회사 정보 초기화"""
        try:
            if not messagebox.askyesno("초기화 확인", "모든 회사 정보를 삭제하시겠습니까?\n'저장' 버튼을 눌러야 실제 반영됩니다.", parent=self.dialog):
                return

            # 임시 데이터만 초기화 (저장 시 실제 반영)
            self.temp_companies = {}
            self.config_manager.set('companies', self.temp_companies)
            self.refresh_company_list()

            messagebox.showinfo(
                "초기화 완료", "회사 정보가 초기화되었습니다.\n'저장' 버튼을 눌러 반영하세요.", parent=self.dialog)
        except Exception as e:
            logging.error(f"회사 정보 초기화 오류: {e}")
            messagebox.showerror(
                "오류", f"초기화 중 오류가 발생했습니다.\n\n{str(e, parent=self.dialog)}")

    def reset_templates(self):
        """이메일 양식 초기화"""
        try:
            if not messagebox.askyesno("초기화 확인", "모든 이메일 양식을 기본값으로 초기화하시겠습니까?\n'저장' 버튼을 눌러야 실제 반영됩니다.", parent=self.dialog):
                return

            default_templates = {
                '공식 보고서': {
                    'subject': '[{회사명}] {날짜} 업무 보고',
                    'body': '''안녕하십니까.

{회사명} 담당자님께 업무 관련 자료를 송부드립니다.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📋 전달 자료
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
▪ 파일명: {파일명}
▪ 발송일시: {날짜} {시간}

첨부 파일을 확인하신 후, 검토 부탁드리겠습니다.
문의사항이 있으시면 언제든 연락 주시기 바랍니다.

감사합니다.'''
                },
                '간결한 전달': {
                    'subject': '{회사명} 자료 전달',
                    'body': '''{회사명} 담당자님, 안녕하세요.

요청하신 자료를 첨부하여 보내드립니다.

📎 {파일명}

확인 후 회신 부탁드립니다.

감사합니다.
{날짜} {시간}'''
                },
                '정중한 공문': {
                    'subject': '[{회사명} 귀중] 문서 송부의 건',
                    'body': '''귀사의 무궁한 발전을 기원합니다.

{회사명} 담당자님께 아래와 같이 관련 문서를 송부하오니
검토하여 주시기 바랍니다.

                        - 아    래 -

1. 송부 문서: {파일명}
2. 발송 일시: {날짜} {시간}
3. 비고: 첨부파일 참조

궁금하신 사항이나 추가 자료가 필요하신 경우
연락 주시면 즉시 대응하겠습니다.

감사합니다.'''
                }
            }

            # 임시 데이터만 초기화 (저장 시 실제 반영)
            self.temp_templates = default_templates
            self.config_manager.set('email_templates', self.temp_templates)
            self.refresh_template_list()

            messagebox.showinfo(
                "초기화 완료", "이메일 양식이 초기화되었습니다.\n'저장' 버튼을 눌러 반영하세요.", parent=self.dialog)
        except Exception as e:
            logging.error(f"양식 초기화 오류: {e}")
            messagebox.showerror(
                "오류", f"초기화 중 오류가 발생했습니다.\n\n{str(e, parent=self.dialog)}")

    def reset_advanced_settings(self):
        """고급 설정 초기화"""
        try:
            if not messagebox.askyesno("초기화 확인", "고급 설정을 기본값으로 초기화하시겠습니까?", parent=self.dialog):
                return

            self.config_manager.set(
                'pattern', '^([가-힣A-Za-z0-9\\s]+?)(?:___|\.pdf$)')
            self.config_manager.set('auto_select_timeout', 10)
            self.config_manager.set('auto_send_timeout', 10)
            self.config_manager.set('email_send_timeout', 180)

            # UI 업데이트
            self.pattern_var.set('^([가-힣A-Za-z0-9\\s]+?)(?:___|\.pdf$)')
            self.auto_select_var.set('10')
            self.auto_send_var.set('10')
            self.email_send_timeout_var.set('180')

            messagebox.showinfo(
                "초기화 완료", "고급 설정이 초기화되었습니다.", parent=self.dialog)
        except Exception as e:
            logging.error(f"고급 설정 초기화 오류: {e}")
            messagebox.showerror(
                "오류", f"초기화 중 오류가 발생했습니다.\n\n{str(e, parent=self.dialog)}")

    def manage_custom_variables(self):
        """커스텀 변수 관리"""
        CustomVariableManager(
            self.dialog, self.config_manager, self.parent_gui)

    def show_company_help(self):
        """회사 정보 사용방법"""
        help_window = tk.Toplevel(self.dialog)
        help_window.title("📖 회사 정보 사용방법")
        help_window.geometry("700x600")
        help_window.transient(self.dialog)

        # 중앙 위치 설정
        center_window(help_window, self.dialog, 700, 600)

        text_widget = scrolledtext.ScrolledText(
            help_window, wrap=tk.WORD, padx=10, pady=10)
        text_widget.pack(fill=tk.BOTH, expand=True)

        help_content = """📖 회사 정보 관리 방법

안녕하세요! 😊 회사 정보를 관리하는 방법을 알려드릴게요

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📋 회사 정보란?
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

• 회사명: PDF 파일에서 자동으로 찾아지는 회사 이름입니다
• 이메일 주소: 해당 회사로 보낼 이메일 주소들입니다 (여러 개도 가능해요!)
• 사용할 양식: 그 회사에게 보낼 이메일의 모양을 정하는 것입니다

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📝 회사 추가하는 방법
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

1단계: "추가" 버튼을 눌러주세요
2단계: 회사명을 입력해 주세요 (PDF 파일명의 회사명과 똑같이 써주세요!)
3단계: 이메일 주소를 입력해 주세요 (여러 개는 쉼표로 나누어 써주세요)
4단계: 사용할 이메일 양식을 선택해 주세요
5단계: "저장" 버튼을 눌러주세요

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
✏️ 회사 수정하거나 삭제하는 방법
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

■ 수정하기:
  • 리스트에서 회사를 선택한 후 "수정" 버튼을 눌러주세요
  • 회사명, 이메일 주소, 양식을 수정할 수 있습니다

■ 삭제하기:
  • 리스트에서 회사를 선택한 후 "삭제" 버튼을 눌러주세요
  • 삭제된 회사 정보는 복구할 수 없으니 주의해 주세요

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
⚠️ 중요한 주의사항들
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🎯 회사명은 PDF 파일명에서 찾아지는 이름과 정확히 똑같아야 합니다

예시를 보여드릴게요:
• PDF 파일명: "삼성전자___2024보고서.pdf"
• 찾아지는 회사명: "삼성전자"
• 여기에 등록할 회사명: "삼성전자" (정확히 똑같이!)

📧 이메일 주소는 올바른 형식으로 써주세요

올바른 예시: example@gmail.com, test@naver.com
잘못된 예시: example@, @gmail.com, example

📝 사용할 양식은 "이메일 양식" 탭에 있는 양식이어야 합니다

양식 탭에서 먼저 양식을 만든 후 여기서 선택해 주세요

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
💡 꿀팁
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

회사를 추가한 후에는 반드시 "저장" 버튼을 눌러야 실제로 저장됩니다
"""

        text_widget.insert('1.0', help_content)
        text_widget.config(state=tk.DISABLED)

        ttk.Button(help_window, text="닫기",
                   command=help_window.destroy).pack(pady=10)

    def show_template_help(self):
        """이메일 양식 사용방법"""
        help_window = tk.Toplevel(self.dialog)
        help_window.title("📖 이메일 양식 사용방법")
        help_window.geometry("700x600")
        help_window.transient(self.dialog)

        # 중앙 위치 설정
        center_window(help_window, self.dialog, 700, 600)

        text_widget = scrolledtext.ScrolledText(
            help_window, wrap=tk.WORD, padx=10, pady=10)
        text_widget.pack(fill=tk.BOTH, expand=True)

        help_content = """📖 이메일 양식 관리 방법

안녕하세요! 😊 이메일 양식을 관리하는 방법을 알려드릴게요

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📖 사용 가능한 변수 (자동 치환 기능)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

이메일 양식에서 사용할 수 있는 특별한 변수들입니다
이 변수들을 사용하면 이메일 발송 시 자동으로 실제 값으로 바뀝니다

{회사명}  {파일명}  {날짜}  {시간}  {년}  {월}  {일}  {요일}  {요일한글}  {시}  {분}  {초}  {시간12}  {오전오후}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 📋 변수 상세 설명
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

■ {회사명} 변수:
  • PDF 파일명에서 자동으로 추출된 회사 이름이 들어갑니다
  • 예시: "삼성전자___보고서.pdf" → {회사명} = "삼성전자"
  • 이메일 제목이나 본문에서 "안녕하세요, {회사명} 담당자님" 이렇게 사용하면
    실제 발송 시 "안녕하세요, 삼성전자 담당자님"으로 바뀝니다

■ {파일명} 변수:
  • 전송되는 PDF 파일의 전체 이름이 들어갑니다
  • 예시: "삼성전자___2024년_1분기_보고서.pdf"
  • 이메일 본문에서 "첨부파일: {파일명}" 이렇게 사용하면
    실제 발송 시 "첨부파일: 삼성전자___2024년_1분기_보고서.pdf"로 바뀝니다

■ {날짜} 변수:
  • 이메일을 발송하는 날짜가 자동으로 들어갑니다
  • 형식: YYYY-MM-DD (예: 2024-01-15)
  • 이메일 본문에서 "발송일: {날짜}" 이렇게 사용하면
    실제 발송 시 "발송일: 2024-01-15"로 바뀝니다

■ {시간} 변수:
  • 이메일을 발송하는 시각이 자동으로 들어갑니다
  • 형식: HH:MM:SS (예: 14:30:25)
  • 이메일 본문에서 "발송시각: {시간}" 이렇게 사용하면
    실제 발송 시 "발송시각: 14:30:25"로 바뀝니다

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 📅 세분화된 날짜 변수들
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

■ {년} 변수:
  • 발송 연도만 표시합니다 (예: 2024)
  • "올해는 {년}년입니다" → "올해는 2024년입니다"

■ {월} 변수:
  • 발송 월만 표시합니다 (예: 01, 12)
  • "이번 달은 {월}월입니다" → "이번 달은 01월입니다"

■ {일} 변수:
  • 발송 일만 표시합니다 (예: 15, 03)
  • "오늘은 {일}일입니다" → "오늘은 15일입니다"

■ {요일} 변수:
  • 영문 요일이 표시됩니다 (예: Monday, Tuesday)
  • "오늘은 {요일}입니다" → "오늘은 Monday입니다"

■ {요일한글} 변수:
  • 한글 요일이 표시됩니다 (예: 월요일, 화요일)
  • "오늘은 {요일한글}입니다" → "오늘은 월요일입니다"

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 ⏰ 세분화된 시간 변수들
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

■ {시} 변수:
  • 발송 시각만 표시합니다 (예: 14, 09)
  • "지금은 {시}시입니다" → "지금은 14시입니다"

■ {분} 변수:
  • 발송 분만 표시합니다 (예: 30, 05)
  • "지금은 {분}분입니다" → "지금은 30분입니다"

■ {초} 변수:
  • 발송 초만 표시합니다 (예: 25, 03)
  • "지금은 {초}초입니다" → "지금은 25초입니다"

■ {시간12} 변수:
  • 12시간 형식으로 표시됩니다 (예: 02:30 PM, 09:15 AM)
  • "발송시각: {시간12}" → "발송시각: 02:30 PM"

■ {오전오후} 변수:
  • 오전/오후만 표시됩니다 (예: 오전, 오후)
  • "지금은 {오전오후}입니다" → "지금은 오후입니다"

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 💡 양식 관리 방법
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

■ 새 양식 만들기 (단계별 설명):

  1단계: "➕ 추가" 버튼을 클릭해 주세요
     → 새로운 양식을 만들 수 있는 창이 열립니다

  2단계: 양식명을 입력해 주세요
     → 예시: "정식양식", "간단양식", "긴급양식", "회사별양식" 등
     → 나중에 회사 정보에서 이 양식을 선택할 때 사용됩니다

  3단계: 이메일 제목을 입력해 주세요
     → 예시: "[{회사명}] PDF 문서 전송", "안녕하세요, {회사명}님"
     → 변수를 사용하면 자동으로 회사명이 들어갑니다

  4단계: 이메일 본문을 입력해 주세요
     → 예시: "안녕하세요, {회사명} 담당자님\n\n첨부된 PDF를 전송드립니다..."
     → 변수를 자유롭게 사용할 수 있습니다

  5단계: "💾 저장" 버튼을 클릭해 주세요
     → 양식이 저장되고 리스트에 나타납니다

■ 기존 양식 수정하기:

  1단계: 리스트에서 수정하고 싶은 양식을 선택해 주세요
     → 마우스로 클릭하면 선택됩니다

  2단계: "✏️ 수정" 버튼을 클릭해 주세요
     → 선택한 양식의 내용이 수정 창에 나타납니다

  3단계: 내용을 수정해 주세요
     → 양식명, 제목, 본문을 모두 수정할 수 있습니다

  4단계: "💾 저장" 버튼을 클릭해 주세요
     → 수정된 내용이 저장됩니다

■ 양식 삭제하기:

  1단계: 리스트에서 삭제하고 싶은 양식을 선택해 주세요
     → 마우스로 클릭하면 선택됩니다

  2단계: "❌ 삭제" 버튼을 클릭해 주세요
     → 삭제 확인 창이 나타납니다

  3단계: 확인 버튼을 눌러 주세요
     → 양식이 완전히 삭제됩니다
     → ⚠️ 삭제된 양식은 복구할 수 없으니 주의해 주세요

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 📝 양식 작성 예시
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

양식명: 정중한양식

제목:
[{회사명}] PDF 문서 전송

본문:
안녕하세요, {회사명} 담당자님

첨부된 PDF 문서를 전송드립니다

파일명: {파일명}
전송일시: {날짜} {시간}

확인 부탁드립니다

감사합니다

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

💡 Tip:
  • 양식을 수정/삭제해도 "저장" 버튼을 눌러야 실제 반영됩니다
  • 회사 정보에서 이 양식을 선택하여 사용할 수 있습니다
"""

        text_widget.insert('1.0', help_content)
        text_widget.config(state=tk.DISABLED)

        ttk.Button(help_window, text="닫기",
                   command=help_window.destroy).pack(pady=10)

    def show_advanced_help(self):
        """고급 설정 사용방법"""
        help_window = tk.Toplevel(self.dialog)
        help_window.title("📖 고급 설정 사용방법")
        help_window.geometry("700x600")
        help_window.transient(self.dialog)

        # 중앙 위치 설정
        center_window(help_window, self.dialog, 700, 600)

        text_widget = scrolledtext.ScrolledText(
            help_window, wrap=tk.WORD, padx=10, pady=10)
        text_widget.pack(fill=tk.BOTH, expand=True)

        help_content = """╔═══════════════════════════════════════════════════════════════╗
║                    PDF 파일명 인식 패턴 설정                   ║
╚═══════════════════════════════════════════════════════════════╝

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 📖 PDF 파일명 인식 패턴 (정규식)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

■ 기본 패턴:
  ^([가-힣A-Za-z0-9\\s]+?)(?:___|\.pdf$)

■ PDF 파일명 규칙:
  [회사명] + ___ + 나머지.pdf  또는  [회사명].pdf

  • 회사명: 한글, 영문, 숫자, 공백 사용 가능
  • 구분자: ___ (언더스코어 3개) 또는 생략 가능

  예시:
  ✅ 홍길동네 회사___2026_1분기.pdf → "홍길동네 회사"
  ✅ 김기사네 회사___데이터.pdf → "김기사네 회사"
  ✅ A네 회사.pdf → "A네 회사" (구분자 없이 회사명만)

■ 정규식 패턴 설명:

  • ^: 파일명의 시작
  • ([가-힣A-Za-z0-9\\s]+?): 회사명 (한글, 영문, 숫자, 공백)
  • (?:___|\.pdf$): ___ 또는 .pdf로 끝남

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 ⏱️ 자동 실행 대기 시간 설정
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

■ 모드 선택 자동 실행:

  프로그램 시작 시 자동으로 "PDF 일괄 발송" 모드를 선택합니다.

  • 양수 (예: 10): 10초 대기 후 자동 실행
  • 0: 즉시 실행
  • 음수 (예: -1): 비활성화, 수동으로 선택 필요

■ 발송 확인 자동 실행:

  이메일 발송 전 확인 화면에서 자동으로 발송합니다.

  • 양수 (예: 10): 10초 대기 후 자동 발송
  • 0: 즉시 발송
  • 음수 (예: -1): 비활성화, Enter 키를 눌러야 발송

■ 권장 설정:

  • 초보 사용자: 10초 이상 (충분한 확인 시간)
  • 숙련 사용자: 3~5초
  • 자동화 환경: 0 (즉시 실행)
  • 수동 확인 필요: -1 (비활성화)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
⏰ 이메일 발송 최대 대기 시간 설정
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

■ 이메일 발송 타임아웃:

  이메일 발송이 이 시간을 초과하면 자동으로 중단됩니다.

  • 기본값: 180초 (3분)
  • 권장 범위: 60~300초 (1~5분)
  • 네트워크가 느린 환경: 300초 이상
  • 빠른 환경: 60~120초

■ 설정 방법:

  1. 고급 설정 탭에서 "이메일 발송 최대 대기 시간" 입력
  2. 초 단위로 입력 (예: 180 = 3분)
  3. "저장" 버튼 클릭

■ 주의사항:

  • 너무 짧게 설정하면 정상 발송도 중단될 수 있습니다!
  • 너무 길게 설정하면 문제 발생 시 오래 기다려야 합니다!
  • 네트워크 상황에 맞게 조정하세요!

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 ⚠️ 주의사항
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

• 정규식 패턴을 잘못 입력하면 회사명 인식이 안 될 수 있습니다!
• 패턴 변경 시 기존 파일명 규칙을 유지하는 것이 좋습니다!
• 자동 실행 시간은 신중히 설정하세요! (실수로 발송될 수 있음)
• 이메일 발송 타임아웃은 네트워크 상황을 고려하여 설정하세요!

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

💡 Tip:
  • 설정을 변경한 후 반드시 "저장" 버튼을 눌러야 반영됩니다!
  • 기본값으로 되돌리려면 "🔄 고급 설정 초기화" 버튼을 사용하세요!
"""

        text_widget.insert('1.0', help_content)
        text_widget.config(state=tk.DISABLED)

        ttk.Button(help_window, text="닫기",
                   command=help_window.destroy).pack(pady=10)


class CustomVariableManager:
    """커스텀 변수 관리 대화상자"""

    def __init__(self, parent, config_manager, parent_gui=None):
        self.config_manager = config_manager
        self.parent_gui = parent_gui

        self.dialog = tk.Toplevel(parent)
        self.dialog.title("🔧 커스텀 변수 관리")
        self.dialog.geometry("500x400")
        self.dialog.transient(parent)
        self.dialog.grab_set()

        # 부모 창 중앙에 위치
        self.dialog.update_idletasks()
        parent.update_idletasks()

        center_window(self.dialog, parent)

        self.setup_ui()

    def setup_ui(self):
        """UI 구성"""
        try:
            if self.parent_gui:
                self.parent_gui.log("🔧 커스텀 변수 관리 UI 구성 시작", is_debug=True)

            # 설명
            info_label = ttk.Label(self.dialog, text="📝 커스텀 변수를 추가하여 이메일 양식에서 사용할 수 있습니다.\n예: {이름}, {담당자1}, {담당자2} 등",
                                 foreground='blue', font=('맑은 고딕', 9))
            info_label.pack(pady=(10, 10))

            # 변수 목록
            list_frame = ttk.LabelFrame(
                self.dialog, text="📋 커스텀 변수 목록", padding="10")
            list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

            # 리스트박스와 스크롤바
            listbox_frame = ttk.Frame(list_frame)
            listbox_frame.pack(fill=tk.BOTH, expand=True)

            self.custom_vars_listbox = tk.Listbox(listbox_frame, height=8)
            self.custom_vars_listbox.pack(
                side=tk.LEFT, fill=tk.BOTH, expand=True)

            scrollbar = ttk.Scrollbar(
                listbox_frame, orient=tk.VERTICAL, command=self.custom_vars_listbox.yview)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            self.custom_vars_listbox.config(yscrollcommand=scrollbar.set)

            # 버튼들
            button_frame = ttk.Frame(list_frame)
            button_frame.pack(fill=tk.X, pady=(10, 0))

            ttk.Button(button_frame, text="➕ 추가", command=self.add_custom_variable, width=10).pack(
                side=tk.LEFT, padx=(0, 5))
            ttk.Button(button_frame, text="✏️ 수정", command=self.edit_custom_variable, width=10).pack(
                side=tk.LEFT, padx=5)
            ttk.Button(button_frame, text="❌ 삭제", command=self.delete_custom_variable, width=10).pack(
                side=tk.LEFT, padx=5)
            ttk.Button(button_frame, text="📖 도움말",
                       command=self.show_custom_variables_help, width=10).pack(side=tk.RIGHT)

            # 변수 목록 로드
            self.load_custom_variables()

            if self.parent_gui:
                self.parent_gui.log("✅ 커스텀 변수 관리 UI 구성 완료!", is_debug=True)

        except Exception as e:
            error_msg = f"❌ 커스텀 변수 관리 UI 구성 오류: {e}"
            if self.parent_gui:
                self.parent_gui.log(error_msg)
            logging.error(error_msg)
            import traceback
            tb = traceback.format_exc()
            if self.parent_gui:
                self.parent_gui.log(tb)
            traceback.print_exc()
            messagebox.showerror(
                "UI 오류", f"커스텀 변수 관리 구성 실패:\n{e}", parent=self.dialog)

    def load_custom_variables(self):
        """커스텀 변수 목록 로드"""
        try:
            self.custom_vars_listbox.delete(0, tk.END)
            custom_vars = self.config_manager.get('custom_variables', {})

            for var_name, var_value in custom_vars.items():
                self.custom_vars_listbox.insert(
                    tk.END, f"{var_name} = {var_value}")

        except Exception as e:
            if self.parent_gui:
                self.parent_gui.log(f"❌ 커스텀 변수 로드 오류: {e}")

    def add_custom_variable(self):
        """커스텀 변수 추가"""
        CustomVariableDialog(self.dialog, self.config_manager,
                             None, self.load_custom_variables, self.parent_gui)

    def edit_custom_variable(self):
        """커스텀 변수 수정"""
        selection = self.custom_vars_listbox.curselection()
        if not selection:
            messagebox.showwarning(
                "선택 오류", "수정할 변수를 선택하세요.", parent=self.dialog)
            return

        # 선택된 항목에서 변수명 추출
        selected_text = self.custom_vars_listbox.get(selection[0])
        var_name = selected_text.split(' = ')[0]

        CustomVariableDialog(self.dialog, self.config_manager,
                             var_name, self.load_custom_variables, self.parent_gui)

    def delete_custom_variable(self):
        """커스텀 변수 삭제"""
        selection = self.custom_vars_listbox.curselection()
        if not selection:
            messagebox.showwarning(
                "선택 오류", "삭제할 변수를 선택하세요.", parent=self.dialog)
            return

        # 선택된 항목에서 변수명 추출
        selected_text = self.custom_vars_listbox.get(selection[0])
        var_name = selected_text.split(' = ')[0]

        if not messagebox.askyesno("삭제 확인", f"'{var_name}' 변수를 삭제하시겠습니까?", parent=self.dialog):
            return

        try:
            custom_vars = self.config_manager.get('custom_variables', {})
            if var_name in custom_vars:
                del custom_vars[var_name]
                self.config_manager.set('custom_variables', custom_vars)
                self.load_custom_variables()
                if self.parent_gui:
                    self.parent_gui.log(f"✓ 커스텀 변수 '{var_name}' 삭제 완료")
        except Exception as e:
            messagebox.showerror(
                "삭제 오류", f"변수 삭제 중 오류가 발생했습니다:\n{e}", parent=self.dialog)

    def show_custom_variables_help(self):
        """커스텀 변수 도움말"""
        help_window = tk.Toplevel(self.dialog)
        help_window.title("📖 커스텀 변수 사용방법")
        help_window.geometry("600x500")
        help_window.transient(self.dialog)

        # 중앙 위치 설정
        center_window(help_window, self.dialog, 600, 500)

        text_widget = scrolledtext.ScrolledText(
            help_window, wrap=tk.WORD, padx=10, pady=10)
        text_widget.pack(fill=tk.BOTH, expand=True)

        help_content = """📖 커스텀 변수 사용방법

안녕하세요! 😊 커스텀 변수를 사용하는 방법을 알려드릴게요

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
🤔 커스텀 변수란?
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

커스텀 변수는 사용자가 직접 만드는 특별한 변수입니다
이 변수들을 이메일 양식에서 사용하면 자동으로 설정한 값으로 바뀝니다

예시:
• {이름} → 홍길동
• {담당자1} → 김철수
• {담당자2} → 이영희
• {부서} → 개발팀

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📝 변수 추가하는 방법
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

1단계: "➕ 추가" 버튼을 클릭해 주세요
2단계: 변수명을 입력해 주세요 (예: 이름, 담당자1, 부서 등)
3단계: 변수값을 입력해 주세요 (예: 홍길동, 김철수, 개발팀 등)
4단계: "💾 저장" 버튼을 클릭해 주세요

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
✏️ 변수 수정하는 방법
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

1단계: 목록에서 수정하고 싶은 변수를 선택해 주세요
2단계: "✏️ 수정" 버튼을 클릭해 주세요
3단계: 변수명이나 변수값을 수정해 주세요
4단계: "💾 저장" 버튼을 클릭해 주세요

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
❌ 변수 삭제하는 방법
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

1단계: 목록에서 삭제하고 싶은 변수를 선택해 주세요
2단계: "❌ 삭제" 버튼을 클릭해 주세요
3단계: 확인 창에서 "예"를 클릭해 주세요

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
💡 이메일 양식에서 사용하는 방법
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

이메일 양식을 만들 때 커스텀 변수를 사용할 수 있습니다:

예시 이메일 양식:
제목: [{회사명}] {이름}님께 자료 전달

본문:
안녕하세요, {담당자1}님

{부서}에서 요청하신 자료를 전송드립니다.

발송자: {이름}
발송일: {날짜} {시간}

감사합니다.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
⚠️ 주의사항
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

• 변수명에는 공백이나 특수문자를 사용하지 마세요
• 변수명은 중괄호 { } 없이 입력하세요 (프로그램이 자동으로 추가합니다)
• 변수값은 이메일에서 표시될 실제 내용입니다
• 변수를 삭제하면 해당 변수를 사용하는 이메일 양식에서 오류가 발생할 수 있습니다

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

💡 꿀팁:
  • 자주 사용하는 정보(이름, 부서, 연락처 등)를 변수로 만들어 두면 편리합니다!
  • 변수명은 기억하기 쉬운 이름으로 정하세요!
"""

        text_widget.insert('1.0', help_content)
        text_widget.config(state=tk.DISABLED)

        ttk.Button(help_window, text="닫기",
                   command=help_window.destroy).pack(pady=10)


class CustomVariableDialog:
    """커스텀 변수 추가/수정 대화상자"""

    def __init__(self, parent, config_manager, var_name, callback, parent_gui=None):
        self.config_manager = config_manager
        self.var_name = var_name
        self.callback = callback
        self.parent_gui = parent_gui

        self.dialog = tk.Toplevel(parent)
        self.dialog.title("커스텀 변수 추가" if not var_name else "커스텀 변수 수정")
        self.dialog.geometry("400x200")
        self.dialog.transient(parent)
        self.dialog.grab_set()

        # 중앙 위치 설정
        center_window(self.dialog, parent)

        self.setup_ui()

    def setup_ui(self):
        """UI 구성"""
        try:
            if self.parent_gui:
                self.parent_gui.log("🔧 커스텀 변수 대화상자 UI 구성 시작", is_debug=True)

            frame = ttk.Frame(self.dialog, padding="20")
            frame.pack(fill=tk.BOTH, expand=True)

            # 변수명
            ttk.Label(frame, text="변수명:").grid(
                row=0, column=0, sticky=tk.W, pady=5)
            self.var_name_var = tk.StringVar(value=self.var_name or '')
            name_entry = ttk.Entry(
                frame, textvariable=self.var_name_var, width=40)
            name_entry.grid(row=0, column=1, pady=5, sticky=(tk.W, tk.E))

            # 변수값
            ttk.Label(frame, text="변수값:").grid(
                row=1, column=0, sticky=tk.W, pady=5)
            self.var_value_var = tk.StringVar()
            if self.var_name:
                custom_vars = self.config_manager.get('custom_variables', {})
                var_value = custom_vars.get(self.var_name, '')
                self.var_value_var.set(var_value)

            ttk.Entry(frame, textvariable=self.var_value_var, width=40).grid(
                row=1, column=1, pady=5, sticky=(tk.W, tk.E))

            frame.columnconfigure(1, weight=1)

            # 버튼
            btn_frame = ttk.Frame(frame)
            btn_frame.grid(row=2, column=0, columnspan=2, pady=20)

            ttk.Button(btn_frame, text="💾 저장", command=self.save).pack(
                side=tk.LEFT, padx=5)
            ttk.Button(btn_frame, text="❌ 취소", command=self.dialog.destroy).pack(
                side=tk.LEFT, padx=5)

            if self.parent_gui:
                self.parent_gui.log("✅ 커스텀 변수 대화상자 UI 구성 완료!", is_debug=True)

        except Exception as e:
            error_msg = f"❌ 커스텀 변수 대화상자 UI 구성 중 오류 발생: {e}"
            if self.parent_gui:
                self.parent_gui.log(error_msg)
            import traceback
            tb = traceback.format_exc()
            if self.parent_gui:
                self.parent_gui.log(tb)
            traceback.print_exc()
            messagebox.showerror(
                "UI 오류", f"커스텀 변수 대화상자 생성 중 오류:\n{e}", parent=self.dialog)

    def save(self):
        """저장"""
        var_name = self.var_name_var.get().strip()
        var_value = self.var_value_var.get().strip()

        if not var_name:
            messagebox.showwarning("입력 오류", "변수명을 입력하세요.", parent=self.dialog)
            return

        if not var_value:
            messagebox.showwarning("입력 오류", "변수값을 입력하세요.", parent=self.dialog)
            return

        try:
            custom_vars = self.config_manager.get('custom_variables', {})
            
            # 중복 체크 (새로 추가할 때만)
            if self.var_name is None and var_name in custom_vars:
                messagebox.showwarning("중복 오류", f"'{var_name}' 변수가 이미 존재합니다.\n수정하려면 기존 변수를 선택하고 '수정' 버튼을 사용하세요.", parent=self.dialog)
                return
            
            custom_vars[var_name] = var_value
            self.config_manager.set('custom_variables', custom_vars)

            if self.callback:
                self.callback()

            self.dialog.destroy()

        except Exception as e:
            messagebox.showerror(
                "저장 오류", f"변수 저장 중 오류가 발생했습니다:\n{e}", parent=self.dialog)


class CompanyDialog:
    """회사 정보 추가/수정 대화상자"""

    def __init__(self, parent, config_manager, company_name, callback, parent_gui=None):
        self.config_manager = config_manager
        self.company_name = company_name
        self.callback = callback
        self.parent_gui = parent_gui

        self.dialog = tk.Toplevel(parent)
        self.dialog.title("회사 정보 추가" if not company_name else "회사 정보 수정")
        self.dialog.transient(parent)
        self.dialog.grab_set()

        # 중앙 위치 설정
        center_window(self.dialog, parent)

        self.setup_ui()

    def setup_ui(self):
        """UI 구성"""
        try:
            if self.parent_gui:
                self.parent_gui.log("🔧 회사 정보 대화상자 UI 구성 시작", is_debug=True)
            frame = ttk.Frame(self.dialog, padding="20")
            frame.pack(fill=tk.BOTH, expand=True)
            if self.parent_gui:
                self.parent_gui.log("  ✓ Frame 생성", is_debug=True)

            # 회사명
            if self.parent_gui:
                self.parent_gui.log("  - 회사명 필드 생성 중...", is_debug=True)
            ttk.Label(frame, text="회사명:").grid(
                row=0, column=0, sticky=tk.W, pady=5)
            self.company_name_var = tk.StringVar(value=self.company_name or '')
            name_entry = ttk.Entry(
                frame, textvariable=self.company_name_var, width=40)
            name_entry.grid(row=0, column=1, pady=5, sticky=(tk.W, tk.E))
            # 회사명 수정 허용 (사용자가 잘못 입력했을 때 수정 가능)
            # if self.company_name:
            #     name_entry.config(state='readonly')
            if self.parent_gui:
                self.parent_gui.log("  ✓ 회사명 필드 완료", is_debug=True)

            # 이메일
            if self.parent_gui:
                self.parent_gui.log("  - 이메일 필드 생성 중...", is_debug=True)
            ttk.Label(frame, text="이메일 주소:").grid(
                row=1, column=0, sticky=tk.W, pady=5)
            ttk.Label(frame, text="(쉼표로 구분)", foreground='gray').grid(
                row=2, column=0, sticky=tk.W)

            self.emails_var = tk.StringVar()
            if self.company_name:
                companies = self.config_manager.get('companies', {})
                emails = companies.get(self.company_name, {}).get('emails', [])
                self.emails_var.set(', '.join(emails))

            ttk.Entry(frame, textvariable=self.emails_var, width=40).grid(
                row=1, column=1, rowspan=2, pady=5, sticky=(tk.W, tk.E))
            if self.parent_gui:
                self.parent_gui.log("  ✓ 이메일 필드 완료", is_debug=True)

            # 양식 (Combobox로 변경)
            if self.parent_gui:
                self.parent_gui.log("  - 양식 Combobox 생성 중...", is_debug=True)
            ttk.Label(frame, text="이메일 양식:").grid(
                row=3, column=0, sticky=tk.W, pady=5)

            # 모든 양식 가져오기
            templates = self.config_manager.get('email_templates', {})
            template_names = [
                name for name in templates.keys() if not name.startswith('_')]
            if self.parent_gui:
                self.parent_gui.log(
                    f"    → 양식 목록: {template_names}", is_debug=True)

            self.template_var = tk.StringVar()
            if self.company_name:
                companies = self.config_manager.get('companies', {})
                template = companies.get(self.company_name, {}).get(
                    'template', template_names[0] if template_names else '')
                self.template_var.set(template)
            else:
                self.template_var.set(
                    template_names[0] if template_names else '')

            template_combo = ttk.Combobox(
                frame, textvariable=self.template_var, values=template_names, state='readonly', width=37)
            template_combo.grid(row=3, column=1, sticky=(tk.W, tk.E), pady=5)
            if self.parent_gui:
                self.parent_gui.log("  ✓ 양식 Combobox 완료", is_debug=True)

            frame.columnconfigure(1, weight=1)

            # 버튼
            if self.parent_gui:
                self.parent_gui.log("  - 버튼 생성 중...", is_debug=True)
            btn_frame = ttk.Frame(frame)
            btn_frame.grid(row=4, column=0, columnspan=2, pady=20)

            ttk.Button(btn_frame, text="저장", command=self.save).pack(
                side=tk.LEFT, padx=5)
            ttk.Button(btn_frame, text="취소", command=self.dialog.destroy).pack(
                side=tk.LEFT, padx=5)
            if self.parent_gui:
                self.parent_gui.log("  ✓ 버튼 완료", is_debug=True)
                self.parent_gui.log("✅ 회사 정보 대화상자 UI 구성 완료!", is_debug=True)
            
            # 창 크기를 내용에 맞춰 자동 조정
            self.dialog.update_idletasks()
            self.dialog.geometry("")

        except Exception as e:
            error_msg = f"❌ 회사 정보 대화상자 UI 구성 중 오류 발생: {e}"
            if self.parent_gui:
                self.parent_gui.log(error_msg)
            import traceback
            tb = traceback.format_exc()
            if self.parent_gui:
                self.parent_gui.log(tb)
            messagebox.showerror(
                "UI 오류", f"회사 대화상자 생성 중 오류:\n{e}", parent=self.dialog)

    def save(self):
        """저장"""
        company_name = self.company_name_var.get().strip()
        emails_str = self.emails_var.get().strip()
        template = self.template_var.get()

        if not company_name or not emails_str:
            self.dialog.focus_force()
            messagebox.showwarning(
                "입력 오류", "회사명과 이메일을 입력하세요.", parent=self.dialog)
            return

        emails = [e.strip() for e in emails_str.split(',') if e.strip()]

        companies = self.config_manager.get('companies', {})
        companies[company_name] = {
            'emails': emails,
            'template': template
        }
        self.config_manager.set('companies', companies)

        self.callback()
        self.dialog.destroy()


class TemplateDialog:
    """이메일 양식 추가/수정 대화상자"""

    def __init__(self, parent, config_manager, template_name, callback, parent_gui=None):
        self.config_manager = config_manager
        self.template_name = template_name
        self.callback = callback
        self.parent_gui = parent_gui

        self.dialog = tk.Toplevel(parent)
        self.dialog.title("양식 추가" if not template_name else "양식 수정")
        self.dialog.transient(parent)
        self.dialog.grab_set()

        self.setup_ui()
        
        # 중앙 위치 설정
        center_window(self.dialog, parent, 550, 450)

    def setup_ui(self):
        """UI 구성"""
        try:
            if self.parent_gui:
                self.parent_gui.log("🔧 이메일 양식 대화상자 UI 구성 시작")
            frame = ttk.Frame(self.dialog, padding="20")
            frame.pack(fill=tk.BOTH, expand=True)
            if self.parent_gui:
                self.parent_gui.log("  ✓ Frame 생성")

            # 양식명
            if self.parent_gui:
                self.parent_gui.log("  - 양식명 필드 생성 중...")
            ttk.Label(frame, text="양식명:").grid(
                row=0, column=0, sticky=tk.W, pady=5)
            self.template_name_var = tk.StringVar(
                value=self.template_name or '')
            name_entry = ttk.Entry(
                frame, textvariable=self.template_name_var, width=40)
            name_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5)
            if self.template_name:
                name_entry.config(state='readonly')
            if self.parent_gui:
                self.parent_gui.log("  ✓ 양식명 필드 완료")

            # 제목
            if self.parent_gui:
                self.parent_gui.log("  - 제목 필드 생성 중...")
            ttk.Label(frame, text="이메일 제목:").grid(
                row=1, column=0, sticky=tk.W, pady=5)
            self.subject_var = tk.StringVar()
            if self.template_name:
                templates = self.config_manager.get('email_templates', {})
                subject = templates.get(
                    self.template_name, {}).get('subject', '')
                self.subject_var.set(subject)

            ttk.Entry(frame, textvariable=self.subject_var, width=40).grid(
                row=1, column=1, sticky=(tk.W, tk.E), pady=5)
            if self.parent_gui:
                self.parent_gui.log("  ✓ 제목 필드 완료")

            # 본문
            if self.parent_gui:
                self.parent_gui.log("  - 본문 필드 생성 중...")
            ttk.Label(frame, text="이메일 본문:").grid(
                row=2, column=0, sticky=tk.NW, pady=5)
            self.body_text = scrolledtext.ScrolledText(
                frame, wrap=tk.WORD, width=50, height=15)
            self.body_text.grid(row=2, column=1, sticky=(
                tk.W, tk.E, tk.N, tk.S), pady=5)

            if self.template_name:
                templates = self.config_manager.get('email_templates', {})
                body = templates.get(self.template_name, {}).get('body', '')
                self.body_text.insert('1.0', body)
            if self.parent_gui:
                self.parent_gui.log("  ✓ 본문 필드 완료")

            frame.columnconfigure(1, weight=1)
            frame.rowconfigure(2, weight=1)

            # 변수 안내 (개선된 표시)
            if self.parent_gui:
                self.parent_gui.log("  - 변수 안내 생성 중...")

            # 변수 안내 프레임
            var_frame = ttk.LabelFrame(frame, text="📝 사용 가능한 변수", padding="10")
            var_frame.grid(row=3, column=1, sticky=(tk.W, tk.E), pady=5)

            # 변수들을 줄바꿈으로 표시
            var_text = """{회사명}, {파일명}, {날짜}, {시간}
{년}, {월}, {일}, {요일}, {요일한글}
{시}, {분}, {초}, {시간12}, {오전오후}"""

            var_label = ttk.Label(var_frame, text=var_text, foreground='blue',
                               font=('맑은 고딕', 9), justify='left')
            var_label.pack(anchor=tk.W)

            if self.parent_gui:
                self.parent_gui.log("  ✓ 변수 안내 완료")

            # 버튼
            if self.parent_gui:
                self.parent_gui.log("  - 버튼 생성 중...")
            btn_frame = ttk.Frame(frame)
            btn_frame.grid(row=4, column=0, columnspan=2, pady=15)

            ttk.Button(btn_frame, text="💾 저장", command=self.save,
                       width=10).pack(side=tk.LEFT, padx=5)
            ttk.Button(btn_frame, text="❌ 취소", command=self.dialog.destroy, width=10).pack(
                side=tk.LEFT, padx=5)
            if self.parent_gui:
                self.parent_gui.log("  ✓ 버튼 완료")
                self.parent_gui.log("✅ 이메일 양식 대화상자 UI 구성 완료!")
            
            # 창 크기를 내용에 맞춰 자동 조정
            self.dialog.update_idletasks()
            self.dialog.geometry("")

        except Exception as e:
            error_msg = f"❌ 이메일 양식 대화상자 UI 구성 중 오류 발생: {e}"
            if self.parent_gui:
                self.parent_gui.log(error_msg)
            import traceback
            tb = traceback.format_exc()
            if self.parent_gui:
                self.parent_gui.log(tb)
            messagebox.showerror(
                "UI 오류", f"양식 대화상자 생성 중 오류:\n{e}", parent=self.dialog)

    def save(self):
        """저장"""
        template_name = self.template_name_var.get().strip()
        subject = self.subject_var.get().strip()
        body = self.body_text.get('1.0', tk.END).strip()

        if not template_name:
            self.dialog.focus_force()
            messagebox.showwarning("입력 오류", "양식명을 입력하세요.", parent=self.dialog)
            return

        if not subject or not body:
            self.dialog.focus_force()
            messagebox.showwarning(
                "입력 오류", "제목과 본문을 입력하세요.", parent=self.dialog)
            return

        templates = self.config_manager.get('email_templates', {})
        templates[template_name] = {
            'subject': subject,
            'body': body
        }
        self.config_manager.set('email_templates', templates)

        if self.callback:
            self.callback()

        self.dialog.focus_force()
        messagebox.showinfo(
            "저장 완료", f"양식 '{template_name}'이(가) 저장되었습니다.", parent=self.dialog)
        self.dialog.destroy()


# GUI 클래스
class PDFEmailSenderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title(f"{MAIN_NAME}! PDF 자동 이메일 발송 프로그램 v{VERSION}")
        self.root.geometry("900x700")

        # 초기화 중 로그 버퍼
        self.init_log_buffer = []

        # SMTP 연결 관리
        # 통합 상태 관리
        self.connection_state = {
            'server_conn': None,
            'connected': False,
            'last_activity': None,
            'check_timer': None
        }

        # 실시간 시간 표시를 위한 변수
        self.time_display_timer = None
        self.time_display_start = None

        try:
            # ConfigManager에 버퍼 로그 함수 전달
            self.config_manager = ConfigManager(log_func=self.buffer_log)
            self.current_folder = None

            self.buffer_log("🔧 프로그램 초기화 시작", is_debug=True)
            
            # 글자 크기 설정 적용
            self.apply_font_size()

            self.setup_ui()

            # 버퍼에 모인 로그 출력
            self.flush_log_buffer()

            self.log("✅ 프로그램 시작 완료", 'INFO')

            # 이메일 설정 확인 및 연결
            self.check_and_connect_email()

            # 프로그램 종료 시 연결 해제
            self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        except Exception as e:
            error_msg = f"❌ GUI 초기화 오류: {e}"
            self.init_log_buffer.append(error_msg)
            import traceback
            tb = traceback.format_exc()
            self.init_log_buffer.append(tb)

            # UI가 준비되었으면 로그 출력
            if hasattr(self, 'log_text'):
                self.flush_log_buffer()

            messagebox.showerror(
                "초기화 오류", f"프로그램 시작 중 오류가 발생했습니다.\n\n{str(e)}", parent=self.root)

    def apply_font_size(self):
        """설정된 글자 크기를 전체 프로그램에 적용"""
        try:
            font_size = self.config_manager.get('ui.font_size', 9)
            
            # tkinter 기본 폰트 설정
            import tkinter.font as tkFont
            
            # 기본 폰트 패밀리
            default_font = tkFont.nametofont("TkDefaultFont")
            default_font.configure(size=font_size, family='맑은 고딕')
            
            text_font = tkFont.nametofont("TkTextFont")
            text_font.configure(size=font_size, family='맑은 고딕')
            
            fixed_font = tkFont.nametofont("TkFixedFont")
            fixed_font.configure(size=font_size)
            
            # 추가 폰트 설정
            for font_name in ["TkMenuFont", "TkCaptionFont", "TkSmallCaptionFont", "TkIconFont", "TkTooltipFont"]:
                try:
                    font = tkFont.nametofont(font_name)
                    font.configure(size=font_size, family='맑은 고딕')
                except:
                    pass
            
            # ttk 스타일 설정 - Entry 위젯 폰트 적용
            style = ttk.Style()
            style.configure('TEntry', font=('맑은 고딕', font_size))
            style.configure('TSpinbox', font=('맑은 고딕', font_size))
            style.configure('TCombobox', font=('맑은 고딕', font_size))
            
            self.buffer_log(f"✓ 글자 크기 적용: {font_size}pt", is_debug=True)
        except Exception as e:
            self.buffer_log(f"⚠ 글자 크기 적용 실패: {e}", is_debug=True)

    def buffer_log(self, message, is_debug=False):
        """초기화 중 로그를 버퍼에 저장"""
        self.init_log_buffer.append((message, is_debug))

    def flush_log_buffer(self):
        """버퍼에 모인 로그를 GUI에 출력"""
        print(f"[DEBUG] flush_log_buffer 호출됨")
        print(f"[DEBUG] log_text 존재: {hasattr(self, 'log_text')}")
        print(f"[DEBUG] 버퍼 크기: {len(self.init_log_buffer)}")
        print(f"[DEBUG] 버퍼 내용: {self.init_log_buffer}")

        if hasattr(self, 'log_text') and self.init_log_buffer:
            print(f"[DEBUG] 버퍼 플러시 시작!")
            for item in self.init_log_buffer:
                if isinstance(item, tuple):
                    message, is_debug = item
                else:
                    # 하위 호환성 - 기존 단순 문자열 지원
                    message, is_debug = item, False

                print(f"[DEBUG] 로그 출력: {message}, is_debug={is_debug}")
                self.log(message, is_debug=is_debug)
            self.init_log_buffer.clear()

            # 이제 ConfigManager가 직접 log 사용하도록 변경
            self.config_manager.log_func = self.log
        else:
            print(
                f"[DEBUG] 플러시 실패 - log_text={hasattr(self, 'log_text')}, buffer={len(self.init_log_buffer)}")

    def setup_ui(self):
        """UI 구성"""
        try:
            self.buffer_log("🔧 UI 구성 시작", is_debug=True)

            main_frame = ttk.Frame(self.root, padding="10")
            main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

            self.root.columnconfigure(0, weight=1)
            self.root.rowconfigure(0, weight=1)
            main_frame.columnconfigure(0, weight=1)
            main_frame.rowconfigure(4, weight=1)

            # 제목 및 설정 버튼
            header_frame = ttk.Frame(main_frame)
            header_frame.grid(row=0, column=0, pady=(
                0, 10), sticky=(tk.W, tk.E))
            header_frame.columnconfigure(0, weight=1)

            ttk.Label(header_frame, text=f"📧 {MAIN_NAME}! PDF 자동 이메일 발송 프로그램",
                     font=('맑은 고딕', 16, 'bold')).grid(row=0, column=0, sticky=tk.W)

            ttk.Button(header_frame, text="📖 사용방법", command=self.show_help).grid(
                row=0, column=1, padx=5)
            ttk.Button(header_frame, text="⚙️ 설정", command=self.open_settings).grid(
                row=0, column=2, padx=5)

            # 폴더 설정
            folder_frame = ttk.LabelFrame(
                main_frame, text="📁 PDF 폴더 설정", padding="10")
            folder_frame.grid(row=1, column=0, pady=(
                0, 10), sticky=(tk.W, tk.E))
            folder_frame.columnconfigure(1, weight=1)

            # 폴더 생성 체크박스
            self.create_folders_var = tk.BooleanVar(
                value=self.config_manager.get('create_folders', False))
            ttk.Checkbutton(folder_frame, text="프로그램 옆에 '전송할PDF', '전송완료' 폴더 자동 생성",
                           variable=self.create_folders_var, command=self.toggle_folder_creation).grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=5)

            # PDF 폴더
            ttk.Label(folder_frame, text="PDF 폴더:").grid(
                row=1, column=0, sticky=tk.W, pady=5)
            self.pdf_folder_var = tk.StringVar(
                value=self.config_manager.get('pdf_folder', str(Path.cwd())))
            ttk.Entry(folder_frame, textvariable=self.pdf_folder_var, state='readonly').grid(
                row=1, column=1, sticky=(tk.W, tk.E), padx=5)
            ttk.Button(folder_frame, text="선택...", command=self.select_pdf_folder).grid(
                row=1, column=2, padx=(0, 5))
            ttk.Button(folder_frame, text="📂 열기",
                       command=self.open_pdf_folder).grid(row=1, column=3)

            # 완료 폴더
            ttk.Label(folder_frame, text="완료 폴더:").grid(
                row=2, column=0, sticky=tk.W, pady=5)
            self.completed_folder_var = tk.StringVar(
                value=self.config_manager.get('completed_folder', str(Path.cwd())))
            ttk.Entry(folder_frame, textvariable=self.completed_folder_var, state='readonly').grid(
                row=2, column=1, sticky=(tk.W, tk.E), padx=5)
            ttk.Button(folder_frame, text="선택...", command=self.select_completed_folder).grid(
                row=2, column=2, padx=(0, 5))
            ttk.Button(folder_frame, text="📂 열기",
                       command=self.open_completed_folder).grid(row=2, column=3)

            # 버튼
            button_frame = ttk.Frame(main_frame)
            button_frame.grid(row=2, column=0, pady=(
                0, 10), sticky=(tk.W, tk.E))
            for i in range(2):
                button_frame.columnconfigure(i, weight=1)

            self.scan_button = ttk.Button(button_frame, text="📂 PDF 분석하기", command=self.scan_pdfs,
                                          style='Large.TButton')
            self.scan_button.grid(row=0, column=0, padx=5,
                                  pady=5, sticky=(tk.W, tk.E))

            self.send_button = ttk.Button(button_frame, text="이메일 발송하기", command=self.send_emails,
                                          state='disabled', style='Large.TButton')
            self.send_button.grid(row=0, column=1, padx=5,
                                  pady=5, sticky=(tk.W, tk.E))

            # 로그
            log_frame = ttk.LabelFrame(
                main_frame, text="📋 실행 로그", padding="10")
            log_frame.grid(row=4, column=0, sticky=(
                tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
            log_frame.columnconfigure(0, weight=1)
            log_frame.rowconfigure(0, weight=1)

            self.log_text = scrolledtext.ScrolledText(
                log_frame, wrap=tk.WORD, width=80, height=20, font=('Consolas', 9))
            self.log_text.grid(
                row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

            self.buffer_log("✓ UI 구성 완료", is_debug=True)

            # 상태
            status_frame = ttk.Frame(main_frame)
            status_frame.grid(row=5, column=0, sticky=(tk.W, tk.E))
            status_frame.columnconfigure(1, weight=1)

            # 프로그램 상태
            ttk.Label(status_frame, text="상태:").pack(side=tk.LEFT, padx=(0, 5))
            self.status_label = ttk.Label(
                status_frame, text="대기 중...", foreground='blue')
            self.status_label.pack(side=tk.LEFT)

            # 이메일 연결 상태
            ttk.Label(status_frame, text="| 이메일:").pack(
                side=tk.LEFT, padx=(20, 5))
            self.email_status_label = ttk.Label(
                status_frame, text="연결 안됨", foreground='red')
            self.email_status_label.pack(side=tk.LEFT)

        except Exception as e:
            error_msg = f"❌ 메인 UI 생성 오류: {e}"
            self.buffer_log(error_msg)
            import traceback
            tb = traceback.format_exc()
            self.buffer_log(tb)
            messagebox.showerror(
                "UI 오류", f"UI 생성 중 오류가 발생했습니다.\n\n{str(e)}", parent=self.root)

        # 스타일
        style = ttk.Style()
        style.configure('Large.TButton', font=('맑은 고딕', 10), padding=10)

    def toggle_folder_creation(self):
        """폴더 생성 토글"""
        create = self.create_folders_var.get()
        self.config_manager.set('create_folders', create)

        if create:
            # 프로그램 위치 기준
            if getattr(sys, 'frozen', False):
                base_dir = Path(sys.executable).parent
            else:
                base_dir = Path.cwd()

            # MAIN_NAME이 있으면 폴더명에 접두사 추가
            pdf_folder = base_dir / f'{NAME_PREFIX}전송할PDF'
            completed_folder = base_dir / f'{NAME_PREFIX}전송완료'

            pdf_folder.mkdir(exist_ok=True)
            completed_folder.mkdir(exist_ok=True)

            self.pdf_folder_var.set(str(pdf_folder))
            self.completed_folder_var.set(str(completed_folder))

            self.config_manager.set('pdf_folder', str(pdf_folder))
            self.config_manager.set('completed_folder', str(completed_folder))

            self.log(f"✅ 폴더 생성됨: {pdf_folder}, {completed_folder}", 'SUCCESS')
        else:
            self.log("폴더 자동 생성 비활성화", 'INFO')

    def select_pdf_folder(self):
        """PDF 폴더 선택"""
        folder = filedialog.askdirectory(
            title="PDF 폴더 선택", initialdir=self.pdf_folder_var.get())
        if folder:
            self.pdf_folder_var.set(folder)
            self.config_manager.set('pdf_folder', folder)
            self.log(f"PDF 폴더 변경: {folder}", 'INFO')

    def select_completed_folder(self):
        """완료 폴더 선택"""
        folder = filedialog.askdirectory(
            title="완료 폴더 선택", initialdir=self.completed_folder_var.get())
        if folder:
            self.completed_folder_var.set(folder)
            self.config_manager.set('completed_folder', folder)
            self.log(f"완료 폴더 변경: {folder}", 'INFO')

    def open_pdf_folder(self):
        """PDF 폴더 열기"""
        try:
            folder = self.pdf_folder_var.get()
            if not Path(folder).exists():
                self._show_custom_message(
                    "폴더 없음", f"폴더가 존재하지 않습니다:\n{folder}", "warning")
                return

            import subprocess
            import platform

            if platform.system() == 'Windows':
                os.startfile(folder)
            elif platform.system() == 'Darwin':  # macOS
                subprocess.Popen(['open', folder])
            else:  # Linux
                subprocess.Popen(['xdg-open', folder])

            self.log(f"📂 폴더 열기: {folder}", 'INFO')
        except Exception as e:
            logging.error(f"폴더 열기 오류: {e}")
            self._show_custom_message(
                "오류", f"폴더를 열 수 없습니다.\n\n{str(e)}", "error")

    def open_completed_folder(self):
        """완료 폴더 열기"""
        try:
            folder = self.completed_folder_var.get()
            if not Path(folder).exists():
                self._show_custom_message(
                    "폴더 없음", f"폴더가 존재하지 않습니다:\n{folder}", "warning")
                return

            import subprocess
            import platform

            if platform.system() == 'Windows':
                os.startfile(folder)
            elif platform.system() == 'Darwin':  # macOS
                subprocess.Popen(['open', folder])
            else:  # Linux
                subprocess.Popen(['xdg-open', folder])

            self.log(f"📂 폴더 열기: {folder}", 'INFO')
        except Exception as e:
            logging.error(f"폴더 열기 오류: {e}")
            self._show_custom_message(
                "오류", f"폴더를 열 수 없습니다.\n\n{str(e)}", "error")

    def _show_custom_message(self, title, message, msg_type="info"):
        """커스텀 메시지 대화상자 (부모 창 중앙에 위치)"""
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry("400x200")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        # 중앙 위치 설정
        center_window(dialog, self.root, 400, 200)

        # 배경색 설정
        dialog.configure(bg='#f8f9fa')

        # 메인 프레임
        main_frame = ttk.Frame(dialog, padding="30")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 아이콘 설정
        if msg_type == "error":
            icon = "❌"
            icon_color = "red"
        elif msg_type == "warning":
            icon = "⚠️"
            icon_color = "orange"
        elif msg_type == "success":
            icon = "✅"
            icon_color = "green"
        else:  # info
            icon = "ℹ️"
            icon_color = "blue"

        # 아이콘과 제목
        icon_label = ttk.Label(main_frame, text=icon, font=('맑은 고딕', 24))
        icon_label.pack(pady=(0, 10))

        # 메시지 표시
        message_label = ttk.Label(main_frame, text=message,
                                 font=('맑은 고딕', 11),
                                 wraplength=320,
                                 justify='center')
        message_label.pack(pady=(0, 20))

        # 확인 버튼
        ok_button = ttk.Button(main_frame, text="확인",
                              command=dialog.destroy, width=12)
        ok_button.pack()

        # 창 닫기 이벤트 처리
        dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)

        # 포커스 설정
        ok_button.focus_set()

        # 창이 닫힐 때까지 대기
        dialog.wait_window()

    def _show_confirm_dialog(self, title, message):
        """커스텀 확인 대화상자 (부모 창 중앙에 위치) - 예쁜 디자인"""
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry("400x200")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        # 중앙 위치 설정
        center_window(dialog, self.root, 350, 200)

        # 배경색 설정
        dialog.configure(bg='#f8f9fa')

        # 결과 저장 변수
        result = [False]

        # 메인 프레임
        main_frame = ttk.Frame(dialog, padding="30")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 아이콘과 제목
        icon_label = ttk.Label(main_frame, text="📧", font=('맑은 고딕', 24))
        icon_label.pack(pady=(0, 10))

        # 메시지 표시 (더 예쁘게)
        message_label = ttk.Label(main_frame, text=message,
                                 font=('맑은 고딕', 11),
                                 wraplength=320,
                                 justify='center')
        message_label.pack(pady=(0, 20))

        # 버튼 프레임
        button_frame = ttk.Frame(main_frame)
        button_frame.pack()

        def on_yes():
            result[0] = True
            dialog.destroy()

        def on_no():
            result[0] = False
            dialog.destroy()

        # 예쁜 버튼들 (텍스트 길이 조정)
        yes_button = ttk.Button(button_frame, text="✅ 예, 발송",
                               command=on_yes, width=12)
        yes_button.pack(side=tk.LEFT, padx=(0, 10))

        no_button = ttk.Button(button_frame, text="❌ 아니오",
                              command=on_no, width=12)
        no_button.pack(side=tk.LEFT, padx=(10, 0))

        # 창 닫기 이벤트 처리
        dialog.protocol("WM_DELETE_WINDOW", on_no)
        
        # 엔터 키 바인딩 (발송)
        dialog.bind('<Return>', lambda e: on_yes())
        
        # ESC 키 바인딩 (취소)
        dialog.bind('<Escape>', lambda e: on_no())

        # 포커스 설정
        yes_button.focus_set()

        # 창이 닫힐 때까지 대기
        dialog.wait_window()

        return result[0]

    def show_help(self):
        """사용방법 안내 창"""
        help_window = tk.Toplevel(self.root)
        help_window.title("📖 사용방법 안내")
        help_window.geometry("800x600")
        help_window.resizable(True, True)

        # 창을 부모 창 중앙에 위치
        help_window.transient(self.root)
        help_window.grab_set()

        # 중앙 위치 설정
        center_window(help_window, self.root, 800, 600)

        # 메인 프레임
        main_frame = ttk.Frame(help_window, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 제목
        title_label = ttk.Label(main_frame, text="📖 PDF 자동 이메일 발송 프로그램 사용방법",
                               font=('맑은 고딕', 16, 'bold'))
        title_label.pack(pady=(0, 20))

        # 스크롤 가능한 텍스트 위젯
        text_frame = ttk.Frame(main_frame)
        text_frame.pack(fill=tk.BOTH, expand=True)

        text_widget = scrolledtext.ScrolledText(text_frame, wrap=tk.WORD, font=('맑은 고딕', 11),
                                               padx=15, pady=15, bg='#f8f9fa')
        text_widget.pack(fill=tk.BOTH, expand=True)

        # 사용방법 내용
        help_content = """📖 PDF 자동 이메일 발송 프로그램 사용방법

안녕하세요! 😊 PDF 자동 이메일 발송 프로그램을 사용해 주셔서 감사합니다!

이 프로그램은 PDF 파일을 자동으로 회사별로 나누어서 이메일로 보내주는 친구입니다.
마치 똑똑한 비서가 여러분을 도와주는 것처럼 말이에요! ✨

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📋 사용 방법 (쉽게 따라해 보세요!)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

[1] 첫 번째: 설정에서 이메일 주소와 비밀번호를 입력해 주세요
   • Gmail을 사용하시면 앱 비밀번호를 만들어서 사용하시면 됩니다
   • 다른 이메일도 사용할 수 있어요!

[2] 두 번째: 설정 > 이메일 양식에서 메일 양식을 만들어 주세요
   • 제목과 내용을 미리 작성해 두시면 됩니다
   • {회사명}, {파일명}, {날짜} 같은 변수도 사용할 수 있어요!

[3] 세 번째: 설정 > 회사 정보에서 회사들을 추가해 주세요
   • 추가 버튼을 눌러서 회사명과 이메일 주소를 입력해 주세요
   • 해당 회사에 어떤 양식의 메일을 보낼지 선택해 주세요
   • 양식 탭에서 양식을 먼저 만드시는 것이 좋습니다!

[4] 네 번째: PDF 폴더를 선택해 주세요
   • 보낼 PDF 파일들이 있는 폴더를 선택해 주세요
   • 완료된 파일들이 갈 폴더도 선택해 주세요

[5] 다섯 번째: PDF 분석하기 버튼을 눌러주세요
   • 프로그램이 PDF 파일들을 살펴보고 회사별로 나누어 줄 거예요
   • 분석 결과를 확인해 주세요

[6] 여섯 번째: 이메일 발송하기 버튼을 눌러주세요
   • 준비가 되면 이메일을 자동으로 보내드릴게요!

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
[중요] 파일명 규칙
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

PDF 파일 이름을 이렇게 지어주세요:
• "삼성전자___보고서.pdf"
• "네이버___계약서.pdf"
• "카카오.pdf"

___ (언더바 3개) 또는 .pdf 앞에 회사명을 써주시면 됩니다!

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
[문제해결] 문제가 생겼을 때
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Q. 이메일이 안 보내져요?
   → 설정에서 이메일 주소와 비밀번호를 다시 확인해 주세요

Q. 회사가 인식이 안 돼요?
   → 설정 > 회사 정보에서 해당 회사를 추가해 주세요

Q. 파일이 인식이 안 돼요?
   → 파일명에 회사명이 포함되어 있는지 확인해 주세요

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
🎉 이제 시작해 보세요!
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

1. ⚙️ 설정에서 이메일과 회사 정보를 입력해 주세요
2. 📂 PDF 분석하기를 눌러 파일들을 확인해 주세요
3. ✉️ 이메일 발송하기로 자동으로 보내주세요!

궁금한 점이 있으시면 언제든지 도움을 요청해 주세요! 😊"""

        text_widget.insert(tk.END, help_content)
        text_widget.config(state=tk.DISABLED)

        # 닫기 버튼
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(20, 0))

        ttk.Button(button_frame, text="😊 알겠어요!", command=help_window.destroy,
                  style='Large.TButton').pack(side=tk.RIGHT)

    def open_settings(self):
        """설정 열기"""
        # 최신 설정 다시 로드
        self.config_manager.reload()

        dialog = SettingsDialog(
            self.root, self.config_manager, parent_gui=self)
        self.root.wait_window(dialog.dialog)

        if dialog.result:
            self.log("⚙️ 설정이 업데이트되었습니다", 'SUCCESS')
            # 설정 다시 로드
            self.config_manager.reload()
            # 설정창에서 이미 연결 처리가 완료되었으므로 추가 처리 불필요

    def check_and_connect_email(self, retry_count=0):
        """이메일 설정 확인 및 SMTP 연결 (재시도 포함)"""
        email = self.config_manager.get('email.sender_email', '')
        password = self.config_manager.get('email.sender_password', '')
        smtp_server = self.config_manager.get('email.smtp_server', '')
        smtp_port = self.config_manager.get('email.smtp_port', 587)

        if not email or not password or not smtp_server:
            self.log("⚠️ 이메일 설정이 필요합니다. '⚙️ 설정' 버튼을 클릭하세요.", 'WARNING')
            self.set_status("이메일 설정 필요 ⚠️", 'orange')
            self.set_email_status("연결 안됨", 'red')
            self.connection_state['connected'] = False
            return False

        # 이미 연결되어 있으면 재사용
        if self.get_connection_state():
            self.log(f"✅ SMTP 연결 재사용: {email}", 'SUCCESS')
            self.set_status("준비 완료 ✅", 'green')
            self.set_email_status("연결됨", 'green')
            return True

        # SMTP 서버 연결 시도
        self.log(f"🔌 SMTP 서버 연결 중... ({smtp_server}:{smtp_port})", 'INFO')

        try:
            # 기존 연결이 있으면 종료
            self.disconnect_smtp()

            # 새 연결 생성 - 포트에 따라 SSL/TLS 선택
            if smtp_port == 465:
                # SSL 연결
                self.connection_state['server_conn'] = smtplib.SMTP_SSL(
                    smtp_server, smtp_port, timeout=30)
                self.connection_state['server_conn'].ehlo()
            else:
                # TLS 연결
                self.connection_state['server_conn'] = smtplib.SMTP(
                smtp_server, smtp_port, timeout=30)
            self.connection_state['server_conn'].ehlo()
            self.connection_state['server_conn'].starttls()
            self.connection_state['server_conn'].ehlo()
            
            self.connection_state['server_conn'].login(email, password)

            self.connection_state['connected'] = True
            self.connection_state['last_activity'] = time.time()
            self.log(f"✅ SMTP 서버 연결 성공: {email}", 'SUCCESS')
            self.set_status("준비 완료 ✅", 'green')
            self.set_email_status("연결됨", 'green')
            # 연결 모니터링 시작
            self.start_connection_monitor()
            return True

        except smtplib.SMTPAuthenticationError as e:
            self.connection_state['connected'] = False
            self.connection_state['server_conn'] = None
            self.log(f"❌ 인증 실패: {e}", 'ERROR')
            self.log("💡 이메일 주소와 앱 비밀번호를 확인하세요.", 'ERROR')
            self.set_status("이메일 연결 실패 ❌", 'red')
            self.set_email_status("인증 실패", 'red')
            return False

        except Exception as e:
            self.connection_state['connected'] = False
            self.connection_state['server_conn'] = None

            # 재시도 로직 (최대 2번)
            if retry_count < 2:
                retry_count += 1
                self.log(f"❌ SMTP 연결 실패 (시도 {retry_count}/3): {e}", 'ERROR')
                self.log(f"🔄 3초 후 재시도합니다...", 'WARNING')
                self.set_status(f"연결 재시도 중... ({retry_count}/3)", 'orange')
                self.set_email_status("재시도 중", 'orange')

                # 모니터링 중지 (재시도 중)
                self.stop_connection_monitor()
                # 3초 대기 후 재시도
                self.root.after(
                    3000, lambda: self.check_and_connect_email(retry_count))
                return False
            else:
                self.log(f"❌ SMTP 연결 최종 실패: {e}", 'ERROR')
                self.set_status("이메일 연결 실패 ❌", 'red')
                self.set_email_status("연결 실패", 'red')
            return False

    def disconnect_smtp(self):
        """SMTP 연결 종료"""
        if self.connection_state['server_conn']:
            try:
                self.connection_state['server_conn'].quit()
                self.log("🔌 SMTP 연결 종료", is_debug=True)
            except:
                pass

        # 상태 초기화
        self.connection_state['server_conn'] = None
        self.connection_state['connected'] = False
        self.connection_state['last_activity'] = None

        # 타이머 정지
        if self.connection_state['check_timer']:
            self.root.after_cancel(self.connection_state['check_timer'])
            self.connection_state['check_timer'] = None

        self.update_connection_status()

    def on_closing(self):
        """프로그램 종료 시 처리"""
        self.log("프로그램 종료 중...", 'INFO')
        # 연결 모니터링 중지
        self.stop_connection_monitor()
        self.disconnect_smtp()
        self.root.destroy()

    def check_email_config(self):
        """이메일 설정 확인"""
        email = self.config_manager.get('email.sender_email', '')
        password = self.config_manager.get('email.sender_password', '')

        if email and password:
            # 이메일 연결 시도
            if self.check_and_connect_email():
                self.log(f"✅ 이메일 설정 및 연결 성공: {email}", 'SUCCESS')
                self.set_status("준비 완료 ✅", 'green')
            else:
                self.log(f"⚠️ 이메일 설정됨: {email} (연결 실패)", 'WARNING')
                self.set_status("이메일 연결 실패 ❌", 'red')
        else:
            self.log("⚠️ 이메일 설정이 필요합니다. '⚙️ 설정' 버튼을 클릭하세요.", 'WARNING')
            self.set_status("이메일 설정 필요 ⚠️", 'orange')
            self.set_email_status("연결 안됨", 'red')

    def scan_pdfs(self):
        """PDF 분석"""
        self.log("\n" + "="*60, 'INFO')
        self.log("📂 PDF 파일 분석 시작", 'INFO')
        self.log("="*60 + "\n", 'INFO')

        pdf_folder = Path(self.pdf_folder_var.get())
        if not pdf_folder.exists():
            self._show_custom_message(
                "오류", f"PDF 폴더가 존재하지 않습니다:\n{pdf_folder}", "error")
            return

        # PDF 파일 검색
        pdf_files = list(pdf_folder.rglob('*.pdf'))
        self.log(f"총 {len(pdf_files)}개 PDF 파일 발견", 'INFO')

        # 회사별로 그룹화
        pattern = re.compile(self.config_manager.get('pattern', ''))
        companies = self.config_manager.get('companies', {})

        company_pdfs = {}
        unrecognized = []
        no_info = {}
        size_exceeded = {}  # 파일 크기 초과 회사들

        for pdf_path in pdf_files:
            match = pattern.search(pdf_path.name)
            if not match:
                unrecognized.append(pdf_path.name)
                continue

            company_name = match.group(1).strip()

            if company_name not in companies:
                if company_name not in no_info:
                    no_info[company_name] = []
                no_info[company_name].append(pdf_path.name)
                continue

            if company_name not in company_pdfs:
                company_pdfs[company_name] = []
            company_pdfs[company_name].append(pdf_path)

        # 결과 출력
        self.log("\n" + "="*60, 'INFO')
        self.log("📊 PDF 분석 결과", 'INFO')
        self.log("="*60, 'INFO')

        # 파일 크기 체크 및 발송 가능한 회사 분리
        valid_company_pdfs = {}
        for company_name, files in company_pdfs.items():
            total_size = sum(file.stat().st_size for file in files)
            size_mb = total_size / (1024 * 1024)
            max_size_mb = 25  # Gmail 제한

            if total_size > max_size_mb * 1024 * 1024:
                size_exceeded[company_name] = files
            else:
                valid_company_pdfs[company_name] = files

        if valid_company_pdfs:
            self.log(f"\n✅ 발송 가능한 회사 ({len(valid_company_pdfs)}개):", 'SUCCESS')
            for company_name, files in valid_company_pdfs.items():
                info = companies[company_name]
                self.log(f"   [{company_name}]", 'INFO')
                self.log(f"   받는 사람: {', '.join(info['emails'])}", 'INFO')
                self.log(f"   이메일 양식: {info['template']}", 'INFO')
                self.log(f"   첨부 파일: {len(files)}개", 'INFO')

                # 파일 크기 표시
                total_size = sum(file.stat().st_size for file in files)
                size_mb = total_size / (1024 * 1024)
                self.log(f"   📎 총 파일 크기: {size_mb:.1f}MB", 'INFO')

                for file in files:
                    self.log(f"     - {file.name}", 'INFO')

        if size_exceeded:
            self.log(
                f"\n❌ 파일 크기 초과로 발송 불가능한 회사 ({len(size_exceeded)}개):", 'ERROR')
            for company_name, files in size_exceeded.items():
                info = companies[company_name]
                total_size = sum(file.stat().st_size for file in files)
                size_mb = total_size / (1024 * 1024)
                self.log(f"   [{company_name}]", 'ERROR')
                self.log(f"   받는 사람: {', '.join(info['emails'])}", 'ERROR')
                self.log(
                    f"   ⚠️ 파일 크기 초과: {size_mb:.1f}MB (제한: 25MB)", 'ERROR')
                self.log(f"   📝 해결 방법: 파일을 분할하거나 압축하세요", 'INFO')
                for file in files:
                    self.log(f"     - {file.name}", 'ERROR')

        if unrecognized:
            self.log(f"\n⚠️ 파일명 인식 실패 ({len(unrecognized)}개):", 'WARNING')
            self.log("   📝 해결 방법:", 'INFO')
            self.log("   1. 파일명에 회사명이 포함되어 있는지 확인", 'INFO')
            self.log("   2. 파일명 패턴이 올바른지 '⚙️ 설정 > 고급 설정'에서 확인", 'INFO')
            self.log("   3. 예시: '삼성전자___보고서.pdf' 또는 '삼성전자.pdf'", 'INFO')
            self.log("   ", 'INFO')
            for f in unrecognized[:3]:
                self.log(f"   - {f}", 'WARNING')
            if len(unrecognized) > 3:
                self.log(f"   ... 외 {len(unrecognized)-3}개", 'WARNING')

        if no_info:
            self.log(f"\n❌ 회사 정보 미등록 ({len(no_info)}개):", 'ERROR')
            for company_name in list(no_info.keys())[:3]:
                self.log(f"   - {company_name}", 'ERROR')
            if len(no_info) > 3:
                self.log(f"   ... 외 {len(no_info)-3}개", 'ERROR')
            self.log("   ", 'INFO')
            self.log("   📝 해결 방법:", 'INFO')
            self.log("   1. '⚙️ 설정 > 회사 정보'에서 해당 회사 추가", 'INFO')
            self.log("   2. 이메일 주소와 사용할 양식 설정", 'INFO')
            self.log("   3. 회사명이 정확히 일치하는지 확인", 'INFO')

        # 최종 결과
        if valid_company_pdfs:
            self.send_button.config(state='normal')
            self.company_pdfs = valid_company_pdfs  # 발송 가능한 회사만 저장
            self.log(
                f"\n🎉 분석 완료! {len(valid_company_pdfs)}개 회사에 이메일을 발송할 수 있습니다.", 'SUCCESS')
            self.log("   '✉️ 이메일 발송하기' 버튼을 클릭하세요.", 'SUCCESS')
        else:
            self.log(f"\n😞 발송 가능한 PDF가 없습니다.", 'ERROR')
            if unrecognized or no_info or size_exceeded:
                self.log("   위의 해결 방법을 참고하여 문제를 해결하세요.", 'INFO')
            else:
                self.log("   PDF 폴더에 파일이 없거나, 파일명 패턴에 맞는 파일이 없습니다.", 'INFO')

    def send_emails(self):
        """이메일 발송 (별도 스레드에서 실행)"""
        if not hasattr(self, 'company_pdfs'):
            self.log("❌ company_pdfs가 설정되지 않았습니다. PDF 분석을 먼저 실행하세요.", 'ERROR')
            messagebox.showerror(
                "분석 필요", "PDF 분석을 먼저 실행해 주세요.", parent=self.root)
            return

        if not self.company_pdfs:
            self.log("❌ 발송할 PDF가 없습니다.", 'ERROR')
            messagebox.showerror(
                "PDF 없음", "발송할 PDF가 없습니다.\nPDF 분석을 먼저 실행하세요.", parent=self.root)
            return

        # 이메일 연결 상태 확인
        if not self.get_connection_state():
            messagebox.showerror(
                "연결 오류", "이메일 서버에 연결되지 않았습니다.\n'⚙️ 설정'에서 이메일 설정을 확인하세요.", parent=self.root)
            return

        # 현재 이메일 설정 확인
        current_email = self.config_manager.get('email.sender_email', '')
        if not current_email:
            self._show_custom_message(
                "설정 오류", "이메일 주소가 설정되지 않았습니다.\n'⚙️ 설정'에서 이메일 설정을 확인하세요.", "error")
            return

        # 커스텀 확인 창 생성 (부모 창 중앙에 위치)
        if not self._show_confirm_dialog("발송 확인", "이메일을 발송하시겠습니까?"):
            return

        # 발송 버튼 비활성화 (중복 실행 방지)
        self.send_button.config(state='disabled', text="📤 발송 중...")
        self.scan_button.config(state='disabled')

        # 이메일 발송 중에는 연결 모니터링 중지
        self.stop_connection_monitor()
        self.log("⏸️ 이메일 발송 중이므로 연결 모니터링을 일시 중지합니다", 'INFO')

        # 별도 스레드에서 이메일 발송 실행
        self.log("🚀 이메일 발송 스레드 시작 중...", 'INFO')
        self.send_thread = threading.Thread(
            target=self._send_emails_thread, daemon=True)
        self.send_thread.start()
        self.log("✅ 이메일 발송 스레드 시작됨", 'INFO')

        # 스레드 상태 확인을 위한 타이머 (설정된 시간 후)
        timeout_seconds = self.config_manager.get(
            'email_send_timeout', 180)  # 기본 3분
        self.thread_check_timer = self.root.after(
            timeout_seconds * 1000, self._check_thread_status)

    def _send_emails_thread(self):
        """이메일 발송 스레드 함수"""
        try:
            self._thread_safe_log("\n" + "="*60, 'INFO')
            self._thread_safe_log("✉️ 이메일 발송 시작", 'INFO')
            self._thread_safe_log("="*60 + "\n", 'INFO')
            self._thread_safe_log("🔍 스레드가 정상적으로 시작되었습니다", 'INFO')

            # company_pdfs 확인
            if not hasattr(self, 'company_pdfs') or not self.company_pdfs:
                self._thread_safe_log("❌ company_pdfs가 없거나 비어있습니다", 'ERROR')
                self.root.after(0, self._send_emails_error, "PDF 분석이 필요합니다")
                return

            self._thread_safe_log(
                f"📊 발송할 회사 수: {len(self.company_pdfs)}", 'INFO')

            # 이메일 설정 가져오기 (현재 설정)
            smtp_server = self.config_manager.get('email.smtp_server')
            smtp_port = self.config_manager.get('email.smtp_port')
            sender_email = self.config_manager.get('email.sender_email')
            sender_password = self.config_manager.get('email.sender_password')

            self._thread_safe_log(
                f"   [DEBUG] 현재 이메일 설정: {sender_email}", is_debug=True)

            companies = self.config_manager.get('companies', {})
            templates = self.config_manager.get('email_templates', {})

            success_count = 0
            fail_count = 0

            # 회사별로 이메일 발송
            for company_name, pdf_paths in self.company_pdfs.items():
                try:
                    company_info = companies[company_name]
                    to_emails = company_info['emails']
                    template_name = company_info['template']

                    # 이메일 내용 생성
                    template = templates.get(template_name, {})
                    subject = template.get('subject', '')
                    body = template.get('body', '')

                    # 변수 치환 (첫 번째 파일 이름 사용)
                    now = datetime.now()
                    filename = pdf_paths[0].name if pdf_paths else ''

                    replacements = {
                        '{회사명}': company_name,
                        '{파일명}': filename,
                        '{날짜}': now.strftime('%Y-%m-%d'),
                        '{시간}': now.strftime('%H:%M:%S'),
                        # 세분화된 날짜 변수들
                        '{년}': now.strftime('%Y'),
                        '{월}': now.strftime('%m'),
                        '{일}': now.strftime('%d'),
                        '{요일}': now.strftime('%A'),
                        '{요일한글}': ['월', '화', '수', '목', '금', '토', '일'][now.weekday()],
                        # 세분화된 시간 변수들
                        '{시}': now.strftime('%H'),
                        '{분}': now.strftime('%M'),
                        '{초}': now.strftime('%S'),
                        # 12시간 형식
                        '{시간12}': now.strftime('%I:%M %p'),
                        '{오전오후}': '오전' if now.hour < 12 else '오후'
                    }

                    # 커스텀 변수 추가
                    custom_vars = self.config_manager.get('custom_variables', {})
                    for var_name, var_value in custom_vars.items():
                        replacements[f'{{{var_name}}}'] = var_value

                    for key, value in replacements.items():
                        subject = subject.replace(key, value)
                        body = body.replace(key, value)

                    # 여러 파일인 경우 본문에 파일 목록 추가
                    if len(pdf_paths) > 1:
                        file_list = '\n'.join(
                            [f"- {pdf.name}" for pdf in pdf_paths])
                        body = body + f"\n\n[첨부 파일]\n{file_list}"

                    # 이메일 발송
                    self._thread_safe_log(
                        f"📤 [{company_name}] 발송 중...", 'INFO')
                    if self.send_email_smtp(to_emails, subject, body, pdf_paths,
                                           smtp_server, smtp_port, sender_email, sender_password):
                        self._thread_safe_log(
                                f"   ✓ 성공: {', '.join(to_emails)}", 'INFO')
                        success_count += 1
                        
                        # 발송 완료된 파일 이동
                        self.move_pdfs_to_completed(pdf_paths)
                    else:
                        self._thread_safe_log(f"   ✗ 실패", 'ERROR')
                        fail_count += 1
                        
                except Exception as e:
                    self._thread_safe_log(f"❌ [{company_name}] 오류: {e}", 'ERROR')
                    fail_count += 1
            
            # 결과 요약
            self._thread_safe_log("\n" + "="*60, 'INFO')
            self._thread_safe_log(f"📊 발송 완료: 성공 {success_count}건, 실패 {fail_count}건", 'INFO')
            self._thread_safe_log("="*60 + "\n", 'INFO')
            
            # UI 업데이트는 메인 스레드에서 실행
            self.root.after(0, self._send_emails_completed, success_count, fail_count)
            
        except Exception as e:
            self._thread_safe_log(f"❌ 발송 스레드 오류: {e}", 'ERROR')
            import traceback
            self._thread_safe_log(f"🔍 상세 오류: {traceback.format_exc()}", 'ERROR')
            self.root.after(0, self._send_emails_error, str(e))
    
    def _send_emails_completed(self, success_count, fail_count):
        """이메일 발송 완료 후 UI 업데이트"""
        # 타이머 정리
        if hasattr(self, 'thread_check_timer'):
            self.root.after_cancel(self.thread_check_timer)
        
        # 연결 모니터링 재시작
        if self.get_connection_state():
            self.start_connection_monitor()
            total = success_count + fail_count
            if fail_count == 0:
                self.log(f"✅ 이메일 발송 완료: {total}건 모두 성공", 'INFO')
            else:
                self.log(f"⚠️ 이메일 발송 완료: 성공 {success_count}건, 실패 {fail_count}건 (총 {total}건)", 'WARNING')
        
        # 버튼 상태 복원
        self.send_button.config(state='normal', text="이메일 발송하기")
        self.scan_button.config(state='normal')
        
    
    def _send_emails_error(self, error_msg):
        """이메일 발송 오류 시 UI 업데이트"""
        self.log(f"🔧 UI 복원 시작: {error_msg}", 'INFO')
        
        # 타이머 정리
        if hasattr(self, 'thread_check_timer'):
            self.root.after_cancel(self.thread_check_timer)
        
        # 연결 모니터링 재시작
        if self.get_connection_state():
            self.start_connection_monitor()
            self.log("▶️ 이메일 발송 오류 후 연결 모니터링을 재시작합니다", 'INFO')
        
        # 버튼 상태 복원
        self.send_button.config(state='normal', text="이메일 발송하기")
        self.scan_button.config(state='normal')
        
        self.log("✅ 버튼 상태 복원 완료", 'INFO')
        
        # 오류 메시지 표시
        self._show_custom_message("발송 오류", f"이메일 발송 중 오류가 발생했습니다.\n\n{error_msg}", "error")
    
    def _check_thread_status(self):
        """스레드 상태 확인 (타임아웃 체크)"""
        if hasattr(self, 'send_thread') and self.send_thread.is_alive():
            timeout_seconds = self.config_manager.get('email_send_timeout', 180)
            self.log(f"⚠️ 스레드가 {timeout_seconds}초 이상 실행 중입니다. 응답이 없을 수 있습니다.", 'WARNING')
            # UI 복원
            self.send_button.config(state='normal', text="이메일 발송하기")
            self.scan_button.config(state='normal')
            self._show_custom_message("타임아웃", f"이메일 발송이 {timeout_seconds}초 이상 실행 중입니다.\n설정에서 대기 시간을 조정하거나 프로그램을 재시작해 주세요.", "error")
    
    def _start_time_display(self, start_time):
        """실시간 시간 표시 시작"""
        self.time_display_start = start_time
        self._update_time_display()
    
    def _update_time_display(self):
        """실시간 시간 표시 업데이트"""
        if self.time_display_start is None:
            return
        
        import time
        elapsed_time = time.time() - self.time_display_start
        
        # 초 단위로 표시 (소수점 1자리)
        elapsed_seconds = elapsed_time
        
        # 로그에 실시간 시간 표시 (기존 로그 마지막 줄 업데이트)
        self._thread_safe_log(f"   이메일 발송 중...⏱️ ({elapsed_seconds:.1f}초)", 'INFO', replace_last=True)
        
        # 0.5초마다 업데이트
        self.time_display_timer = self.root.after(500, self._update_time_display)
    
    def _stop_time_display(self):
        """실시간 시간 표시 정지"""
        if self.time_display_timer:
            self.root.after_cancel(self.time_display_timer)
            self.time_display_timer = None
        self.time_display_start = None
    
    def send_email_smtp(self, to_emails, subject, body, pdf_paths, smtp_server, smtp_port, sender_email, sender_password, retry_count=0):
        """SMTP를 통한 이메일 발송 (연결 재사용, 재시도 포함)"""
        self._thread_safe_log(f"   [DEBUG] send_email_smtp 시작", is_debug=True)
        self._thread_safe_log(f"   [DEBUG] 수신자: {to_emails}", is_debug=True)
        
        # 파일 크기 체크 (Gmail 25MB 제한)
        total_size = sum(pdf_path.stat().st_size for pdf_path in pdf_paths)
        max_size = 24 * 1024 * 1024  # 24MB (여유 있게)
        
        self._thread_safe_log(f"   [DEBUG] 첨부 파일 크기: {total_size / (1024*1024):.2f}MB", is_debug=True)
        
        if total_size > max_size:
            size_mb = total_size / (1024 * 1024)
            self._thread_safe_log(f"   ⚠ 첨부 파일 크기 초과: {size_mb:.1f}MB (제한: 25MB)", 'WARNING')
            return False
        
        # 이메일 메시지 생성
        self._thread_safe_log(f"   [DEBUG] 이메일 메시지 생성 중...", is_debug=True)
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = ', '.join(to_emails)
        msg['Subject'] = subject
        
        # 본문 첨부
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        self._thread_safe_log(f"   [DEBUG] 본문 첨부 완료", is_debug=True)
        
        # PDF 파일들 첨부
        for pdf_path in pdf_paths:
            with open(pdf_path, 'rb') as f:
                pdf = MIMEApplication(f.read(), _subtype='pdf')
                pdf.add_header('Content-Disposition', 'attachment', 
                             filename=('utf-8', '', pdf_path.name))
                msg.attach(pdf)
            self._thread_safe_log(f"   [DEBUG] PDF 첨부: {pdf_path.name}", is_debug=True)
        
        # 발송 정보 로그
        self._thread_safe_log(f"\n\n", is_debug=True)
        self._thread_safe_log(f"   📤 메일 발송 중...")
        self._thread_safe_log(f"   [DEBUG] ===== 발송 정보 =====", is_debug=True)
        self._thread_safe_log(f"   [DEBUG] 발신: {sender_email}", is_debug=True)
        self._thread_safe_log(f"   [DEBUG] 수신: {to_emails}", is_debug=True)
        self._thread_safe_log(f"   [DEBUG] 제목: {subject}", is_debug=True)
        self._thread_safe_log(f"   [DEBUG] 본문 미리보기: {body[:100]}...", is_debug=True)
        self._thread_safe_log(f"   [DEBUG] 첨부 파일: {[p.name for p in pdf_paths]}", is_debug=True)
        self._thread_safe_log(f"   [DEBUG] SMTP 서버: {smtp_server}:{smtp_port}", is_debug=True)
        self._thread_safe_log(f"   [DEBUG] 메시지 크기: {len(msg.as_string())} bytes", is_debug=True)
        self._thread_safe_log(f"   [DEBUG] ========================", is_debug=True)
        
        # 메일 전송 시간 측정 시작
        import time
        start_time = time.time()
        
        # 실시간 시간 표시를 위한 타이머 시작
        self._start_time_display(start_time)
        
        try:
            # 기존 연결 재사용 또는 새 연결 생성
            if self.get_connection_state():
                # 기존 연결 재사용
                self._thread_safe_log(f"   [DEBUG] 기존 SMTP 연결 재사용...", is_debug=True)
                server = self.connection_state['server_conn']
            else:
                # 새 연결 생성
                self._thread_safe_log(f"   [DEBUG] 새 SMTP 연결 생성...", is_debug=True)
                
                # 포트에 따라 SSL/TLS 선택
                if smtp_port == 465:
                    # SSL 연결
                    server = smtplib.SMTP_SSL(smtp_server, smtp_port, timeout=300)
                else:
                    # TLS 연결
                    server = smtplib.SMTP(smtp_server, smtp_port, timeout=300)
            server.starttls()
                
            server.login(sender_email, sender_password)
            self._thread_safe_log(f"   [DEBUG] SMTP 연결 성공", is_debug=True)
            
            # 연결 정보 저장
            self.connection_state['server_conn'] = server
            self.connection_state['connected'] = True
            self.connection_state['last_activity'] = time.time()
            
            # 메일 전송 실행
            self._thread_safe_log(f"   [DEBUG] 메일을 보내는 중...", is_debug=True)
            server.send_message(msg)
            
            # 전송 시간 계산 (초)
            end_time = time.time()
            send_duration_seconds = end_time - start_time
            self._thread_safe_log(f"   [DEBUG] 메일 전송 성공 ({send_duration_seconds:.1f}초)", is_debug=True)
            
            # 실시간 시간 표시 타이머 정지
            self._stop_time_display()
            
            # 마지막 활동 시간 업데이트
            self.connection_state['last_activity'] = time.time()
            
            self._thread_safe_log(f"   ✅ 발송 완료! (전송시간: {send_duration_seconds:.1f}초)")
            self._thread_safe_log(f"\n\n", is_debug=True)
            return True
            
        except smtplib.SMTPAuthenticationError as e:
            # 전송 시간 계산 (실패 시에도)
            end_time = time.time()
            send_duration_seconds = end_time - start_time
            self._thread_safe_log(f"   ✗ 인증 실패: {e} (실패시간: {send_duration_seconds:.1f}초)", 'ERROR')
            self._thread_safe_log(f"   💡 이메일 주소와 앱 비밀번호를 확인하세요.", 'ERROR')
            self._thread_safe_log(f"\n\n", is_debug=True)
            
            # 실시간 시간 표시 타이머 정지
            self._stop_time_display()
            
            # 연결 실패 시 상태 업데이트
            self.connection_state['connected'] = False
            self.connection_state['server_conn'] = None
            return False
            
        except smtplib.SMTPException as e:
            # SMTP 오류 시 재시도 (1번만)
            if retry_count < 1:
                retry_count += 1
                self._thread_safe_log(f"   ✗ SMTP 오류 (시도 {retry_count}/2): {e}", 'ERROR')
                self._thread_safe_log(f"   🔄 2초 후 재시도합니다...", 'WARNING')
                self._thread_safe_log(f"\n\n", is_debug=True)
                
                # 2초 대기 후 재시도
                import time
                time.sleep(2)
                return self.send_email_smtp(to_emails, subject, body, pdf_paths, smtp_server, smtp_port, sender_email, sender_password, retry_count)
            else:
                # 전송 시간 계산 (실패 시에도)
                end_time = time.time()
                send_duration_seconds = end_time - start_time
                self._thread_safe_log(f"   ✗ SMTP 최종 실패: {e} (실패시간: {send_duration_seconds:.1f}초)", 'ERROR')
                self._thread_safe_log(f"   [DEBUG] SMTPException 타입: {type(e).__name__}", is_debug=True)
                self._thread_safe_log(f"\n\n", is_debug=True)
                
                # 실시간 시간 표시 타이머 정지
                self._stop_time_display()
                
                # 연결 오류 시 상태 업데이트
                self.connection_state['connected'] = False
                self.connection_state['server_conn'] = None
            return False
            
        except Exception as e:
            # 일반 오류 시 재시도 (1번만)
            if retry_count < 1:
                retry_count += 1
                self._thread_safe_log(f"   ✗ 발송 실패 (시도 {retry_count}/2): {e}", 'ERROR')
                self._thread_safe_log(f"   🔄 2초 후 재시도합니다...", 'WARNING')
                self._thread_safe_log(f"\n\n", is_debug=True)
                
                # 2초 대기 후 재시도
                import time
                time.sleep(2)
                return self.send_email_smtp(to_emails, subject, body, pdf_paths, smtp_server, smtp_port, sender_email, sender_password, retry_count)
            else:
                # 전송 시간 계산 (실패 시에도)
                end_time = time.time()
                send_duration_seconds = end_time - start_time
                self._thread_safe_log(f"   ✗ 발송 최종 실패: {e} (실패시간: {send_duration_seconds:.1f}초)", 'ERROR')
                self._thread_safe_log(f"   [DEBUG] Exception 타입: {type(e).__name__}", is_debug=True)
            import traceback
            self._thread_safe_log(traceback.format_exc(), is_debug=True)
            self._thread_safe_log(f"\n\n", is_debug=True)
            
            # 실시간 시간 표시 타이머 정지
            self._stop_time_display()
            
            # 연결 오류 시 상태 업데이트
            self.connection_state['connected'] = False
            self.connection_state['server_conn'] = None
            return False
    
    def move_pdfs_to_completed(self, pdf_paths):
        """PDF 파일들을 전송완료 폴더로 이동"""
        try:
            pdf_folder = Path(self.config_manager.get('pdf_folder'))
            completed_folder = Path(self.config_manager.get('completed_folder'))
            
            for pdf_path in pdf_paths:
                # 원본 폴더 구조 유지
                rel_path = pdf_path.relative_to(pdf_folder)
                dest_path = completed_folder / rel_path
                
                # 대상 폴더 생성
                dest_path.parent.mkdir(parents=True, exist_ok=True)
                
                # 파일 이동
                import shutil
                shutil.move(str(pdf_path), str(dest_path))
                self._thread_safe_log(f"   → {dest_path.name} 이동 완료", is_debug=True)
                
        except Exception as e:
            self._thread_safe_log(f"   ⚠ 파일 이동 실패: {e}", 'WARNING')
        
    def log(self, message, level='INFO', is_debug=False):
        """로그 추가
        
        Args:
            message: 로그 메시지
            level: 로그 레벨 (INFO, WARNING, ERROR 등)
            is_debug: True이면 디버그 모드일 때만 표시
        """
        # 디버그 로그는 디버그 모드가 켜져있을 때만 표시
        if is_debug:
            debug_mode = self.config_manager.get('debug_mode', False)
            if not debug_mode:
                return
        
        timestamp = datetime.now().strftime('%H:%M:%S')
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update()
    
    def _thread_safe_log(self, message, level='INFO', is_debug=False, replace_last=False):
        """스레드 안전한 로그 추가 (별도 스레드에서 호출 가능)"""
        # 디버그 로그는 디버그 모드가 켜져있을 때만 표시
        if is_debug:
            debug_mode = self.config_manager.get('debug_mode', False)
            if not debug_mode:
                return
        
        # 메인 스레드에서 실행되도록 스케줄링
        self.root.after(0, self._add_log_to_gui, message, level, replace_last)
    
    def _add_log_to_gui(self, message, level, replace_last=False):
        """GUI에 로그 추가 (메인 스레드에서만 호출)"""
        timestamp = datetime.now().strftime('%H:%M:%S')
        
        if replace_last:
            # 마지막 줄을 교체 (더 안전한 방법)
            try:
                # 마지막 줄이 "이메일 발송 중"으로 시작하는지 확인
                last_line = self.log_text.get('end-2l', 'end-1l')
                if '이메일 발송 중' in last_line:
                    # 마지막 줄 삭제
                    self.log_text.delete('end-2l', 'end-1l')
                # 새 줄 추가
                self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
            except:
                # 오류 시 그냥 새 줄 추가
                self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        else:
            # 새 줄 추가
            self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        
        self.log_text.see(tk.END)
        self.root.update()
    
    def set_status(self, message, color='blue'):
        """상태 업데이트"""
        self.status_label.config(text=message, foreground=color)
    
    def set_email_status(self, message, color='blue'):
        """이메일 연결 상태 업데이트"""
        self.email_status_label.config(text=message, foreground=color)
    
    def get_connection_state(self):
        """연결 상태 확인"""
        if not self.connection_state['connected'] or not self.connection_state['server_conn']:
            return False
        
        try:
            # 연결 상태 테스트
            self.connection_state['server_conn'].noop()
            return True
        except:
            # 연결이 끊어진 경우 상태 초기화
            self.connection_state['connected'] = False
            self.connection_state['server_conn'] = None
            self.connection_state['last_activity'] = None
            return False
    
    def update_connection_status(self):
        """연결 상태 업데이트"""
        if self.get_connection_state():
            self.set_email_status("연결됨", 'green')
        else:
            self.set_email_status("연결 안됨", 'red')
    
    def disconnect_smtp(self):
        """SMTP 연결 종료"""
        if self.connection_state['server_conn']:
            try:
                self.connection_state['server_conn'].quit()
            except:
                pass
        
        # 상태 초기화
        self.connection_state['server_conn'] = None
        self.connection_state['connected'] = False
        self.connection_state['last_activity'] = None
        
        # 타이머 정지
        if self.connection_state['check_timer']:
            self.root.after_cancel(self.connection_state['check_timer'])
            self.connection_state['check_timer'] = None
        
        self.update_connection_status()
    
    def start_connection_monitor(self):
        """연결 상태 모니터링 시작 (1분마다 확인)"""
        if self.connection_state['check_timer']:
            self.root.after_cancel(self.connection_state['check_timer'])
        
        self.connection_state['check_timer'] = self.root.after(60000, self.check_connection_status)  # 60초 = 1분
    
    def stop_connection_monitor(self):
        """연결 상태 모니터링 중지"""
        if self.connection_state['check_timer']:
            self.root.after_cancel(self.connection_state['check_timer'])
            self.connection_state['check_timer'] = None
    
    def check_connection_status(self):
        """연결 상태 확인"""
        if self.get_connection_state():
            self.log("🔍 연결 상태 확인: 정상", is_debug=True)
            # 다음 확인 예약
            self.start_connection_monitor()
        else:
            # 연결이 끊어진 경우
            self.log("⚠️ 연결이 끊어짐을 감지했습니다. 재연결을 시도합니다.", 'WARNING')
            self.update_connection_status()
            self.set_status("연결 끊어짐 ⚠️", 'orange')
            # 모니터링 중지 (재연결 시도 중)
            self.stop_connection_monitor()
            # 재연결 시도
            self.check_and_connect_email()
    
    def reset_email_settings(self):
        """이메일 설정 초기화"""
        if not messagebox.askyesno("초기화 확인", "이메일 설정을 기본값으로 초기화하시겠습니까?", parent=self.root):
            return
        
        self.config_manager.set('email.smtp_server', 'smtp.gmail.com')
        self.config_manager.set('email.smtp_port', 587)
        self.config_manager.set('email.sender_email', '')
        self.config_manager.set('email.sender_password', '')
        
        # UI 업데이트
        self.smtp_server_var.set('smtp.gmail.com')
        self.smtp_port_var.set('587')
        self.sender_email_var.set('')
        self.sender_password_var.set('')
        
        self._show_custom_message("초기화 완료", "이메일 설정이 초기화되었습니다.", "success")
    
    def reset_company_info(self):
        """회사 정보 초기화"""
        if not messagebox.askyesno("초기화 확인", "모든 회사 정보를 삭제하시겠습니까?\n이 작업은 되돌릴 수 없습니다.", parent=self.root):
            return
        
        self.config_manager.set('companies', {})
        self.refresh_company_list()
        
        self._show_custom_message("초기화 완료", "회사 정보가 초기화되었습니다.", "success")
    
    def reset_templates(self):
        """이메일 양식 초기화"""
        if not self._show_confirm_dialog("초기화 확인", "모든 이메일 양식을 기본값으로 초기화하시겠습니까?"):
            return
        
        default_templates = {
            '이메일양식A': {
                'subject': '안녕하세요, {회사명}님',
                'body': '안녕하세요.\n\n{회사명}님께 보내드립니다.\n\n첨부 파일: {파일명}\n발송 일시: {날짜} {시간}\n\n감사합니다.'
            },
            '이메일양식B': {
                'subject': '[{회사명}] 문서 전송',
                'body': '{회사명}님께\n\n요청하신 문서를 첨부하여 보내드립니다.\n\n감사합니다.'
            },
            '이메일양식C': {
                'subject': '파일 전송 - {파일명}',
                'body': '첨부 파일을 확인해 주세요.\n\n전송 시간: {날짜} {시간}'
            }
        }
        
        self.config_manager.set('templates', default_templates)
        self.load_template()
        
        self._show_custom_message("초기화 완료", "이메일 양식이 초기화되었습니다.", "success")
    
    def reset_advanced_settings(self):
        """고급 설정 초기화"""
        if not self._show_confirm_dialog("초기화 확인", "고급 설정을 기본값으로 초기화하시겠습니까?"):
            return
        
        self.config_manager.set('pattern', '^([가-힣A-Za-z0-9\\s]+?)(?:___|\.pdf$)')
        self.config_manager.set('auto_select_timeout', 10)
        self.config_manager.set('auto_send_timeout', 10)
        self.config_manager.set('email_send_timeout', 180)
        
        # UI 업데이트
        self.pattern_var.set('^([가-힣A-Za-z0-9\\s]+?)(?:___|\.pdf$)')
        self.auto_select_var.set('10')
        self.auto_send_var.set('10')
        self.email_send_timeout_var.set('180')
        
        self._show_custom_message("초기화 완료", "고급 설정이 초기화되었습니다.", "success")


def resource_path(relative_path):
    """PyInstaller로 패키징된 리소스 경로를 반환하는 함수"""
    try:
        # PyInstaller가 임시 폴더에 압축 해제한 경로
        base_path = sys._MEIPASS
    except Exception:
        # 일반 Python 실행 시 현재 스크립트 경로
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

def set_app_icon(root):
    """
    모든 창(향후 생성 Toplevel 포함)에 공통 아이콘 지정.
    - 반드시 Toplevel 생성 전에 호출할 것.
    - PNG 권장(256px 등 충분한 해상도). Windows에서는 ICO도 보강 적용.
    """
    # 1) iconphoto: 전 플랫폼 공통 기본 아이콘 지정
    png = resource_path("favicon/favicon-96x96.png")
    if os.path.exists(png):
        img = tk.PhotoImage(file=png)
        root.iconphoto(True, img)      # True: 이후 생성되는 모든 Toplevel에 상속
        root._icon_ref = img           # GC 방지

    # 2) Windows 보강: 작업표시줄/제목표시줄에 ICO가 필요한 경우
    if sys.platform == "win32":
        ico = resource_path("favicon/favicon.ico")
        if os.path.exists(ico):
            try:
                root.iconbitmap(ico)   # 실패해도 무시 가능
            except Exception:
                pass


def main():
    if sys.platform == 'win32':
        try:
            myappid = f'watual.{MAIN_NAME}.gui.{VERSION}'
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        except:
            pass

    root = tk.Tk()
    set_app_icon(root)
    app = PDFEmailSenderGUI(root)
    root.mainloop()


if __name__ == "__main__":
    # 배치 파일에서 호출될 때만 실행
    if len(sys.argv) > 1 and sys.argv[1] == "--get-main-name":
        print(get_main_name())
        sys.exit(0)
    
    # 일반 실행
    main()
