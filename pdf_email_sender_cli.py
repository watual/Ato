#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF 이메일 자동 발송 프로그램 (CLI 버전)
- settings.json 파일에서 설정 로드
- PDF 파일명에서 회사명을 추출
- 회사별 이메일 주소와 양식을 사용하여 이메일 발송
"""

MAIN_NAME = ""  # GUI 버전과 동일하게 설정

import os
import re
import smtplib
import logging
import time
import socket
import sys
import json
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from datetime import datetime

# Windows에서 키 입력 감지
try:
    import msvcrt
except ImportError:
    msvcrt = None

# 로깅 설정
log_dir = Path(__file__).parent
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_dir / 'email_sender.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)

class PDFEmailSender:
    def __init__(self):
        # 기본 디렉토리 설정
        if getattr(sys, 'frozen', False):
            # exe로 실행 중
            self.base_dir = Path(sys.executable).parent
        else:
            # Python 스크립트로 실행 중
            self.base_dir = Path(__file__).parent
        
        # settings.json 경로 (MAIN_NAME이 있으면 접두사 추가)
        if MAIN_NAME:
            self.settings_file = self.base_dir / f'{MAIN_NAME}_settings.json'
        else:
            self.settings_file = self.base_dir / 'settings.json'
        
        # MAIN_NAME이 있으면 폴더명에 접두사 추가
        if MAIN_NAME:
            self.pdf_dir = self.base_dir / f'{MAIN_NAME}_전송할PDF'
            self.completed_dir = self.base_dir / f'{MAIN_NAME}_전송완료'
        else:
            self.pdf_dir = self.base_dir / '전송할PDF'
            self.completed_dir = self.base_dir / '전송완료'
        
        # 설정 초기화
        self.settings = {}
        self.pattern = None
        self.company_db = {}
        self.email_templates = {}
        self.smtp_server = None
        self.smtp_port = None
        self.sender_email = None
        self.sender_password = None
        self.auto_select_timeout = 10
        self.auto_send_timeout = 10
        
        # 필요한 디렉토리 생성
        self.pdf_dir.mkdir(exist_ok=True)
        self.completed_dir.mkdir(exist_ok=True)
    
    def load_settings(self):
        """settings.json에서 모든 설정 로드"""
        try:
            if not self.settings_file.exists():
                raise FileNotFoundError(f"설정 파일이 없습니다: {self.settings_file}")
            
            with open(self.settings_file, 'r', encoding='utf-8') as f:
                self.settings = json.load(f)
            
            # 이메일 설정
            email_config = self.settings.get('email', {})
            self.smtp_server = email_config.get('smtp_server')
            self.smtp_port = email_config.get('smtp_port')
            self.sender_email = email_config.get('sender_email')
            password = email_config.get('sender_password', '')
            # 공백/탭 제거
            self.sender_password = password.replace(' ', '').replace('\t', '')
            
            # 패턴
            self.pattern = self.settings.get('pattern', r'^([가-힣A-Za-z0-9\s]+?)(?:___|\.pdf$)')
            
            # 타임아웃
            self.auto_select_timeout = self.settings.get('auto_select_timeout', 10)
            self.auto_send_timeout = self.settings.get('auto_send_timeout', 10)
            
            # 회사 정보 로드
            companies = self.settings.get('companies', {})
            for company_name, info in companies.items():
                # _description 같은 메타 정보는 제외
                if company_name.startswith('_'):
                    continue
                # _description 키가 있는 경우도 제외
                if isinstance(info, dict):
                    self.company_db[company_name] = {
                        'emails': info.get('emails', []),
                        'template': info.get('template', '이메일양식A')
                    }
            
            # 이메일 양식 로드
            templates = self.settings.get('email_templates', {})
            for template_name, content in templates.items():
                # _description 같은 메타 정보는 제외
                if template_name.startswith('_'):
                    continue
                if isinstance(content, dict) and 'subject' in content and 'body' in content:
                    self.email_templates[template_name] = {
                        'subject': content.get('subject', ''),
                        'body': content.get('body', '')
                    }
            
            logging.info(f"설정 로드 완료: {self.settings_file}")
            logging.info(f"  - 이메일: {self.sender_email}")
            logging.info(f"  - 회사: {len(self.company_db)}개")
            logging.info(f"  - 양식: {len(self.email_templates)}개")
            
            # 필수 설정 확인
            if not all([self.smtp_server, self.smtp_port, self.sender_email, self.sender_password]):
                raise ValueError("이메일 설정이 완전하지 않습니다.")
            
            if not self.pattern:
                raise ValueError("정규식 패턴이 설정되지 않았습니다.")
            
        except Exception as e:
            logging.error(f"설정 로드 오류: {e}")
            raise
    
    def test_email_connection(self):
        """이메일 계정 연결 테스트"""
        print("\n" + "="*60)
        print("이메일 계정 연결 테스트 중...")
        print("="*60)
        
        try:
            # SMTP 서버 연결
            server = smtplib.SMTP(self.smtp_server, self.smtp_port, timeout=10)
            server.starttls()
            
            # 로그인 시도
            server.login(self.sender_email, self.sender_password)
            server.quit()
            
            print(f"✓ 연결 성공!")
            print(f"✓ 발신 이메일: {self.sender_email}")
            print("="*60 + "\n")
            return True
            
        except smtplib.SMTPAuthenticationError as e:
            error_msg = str(e)
            print(f"✗ 인증 실패!")
            print(f"✗ 발신 이메일: {self.sender_email}")
            print("\n가능한 원인:")
            
            if '535' in error_msg:
                if 'Invalid credentials' in error_msg or 'Username and Password not accepted' in error_msg:
                    print("1. 이메일 주소 또는 비밀번호가 잘못되었습니다.")
                    print("2. '앱 비밀번호'가 아닌 일반 비밀번호를 사용하고 있을 수 있습니다.")
                    print("\n해결 방법:")
                    print("- Gmail의 경우: 2단계 인증 활성화 후 앱 비밀번호 생성")
                    print("  https://myaccount.google.com/apppasswords")
                elif 'Application-specific password required' in error_msg:
                    print("1. 앱 비밀번호를 사용해야 합니다.")
                    print("2. 일반 비밀번호는 보안상 사용할 수 없습니다.")
                    print("\n해결 방법:")
                    print("- 2단계 인증 활성화 → 앱 비밀번호 생성")
                else:
                    print("1. 인증 정보가 올바르지 않습니다.")
                    print("2. 2단계 인증이 비활성화되어 있을 수 있습니다.")
            else:
                print(f"알 수 없는 인증 오류: {error_msg}")
            
            print("="*60 + "\n")
            
            if msvcrt:
                print("계속하려면 아무 키나 누르세요...")
                msvcrt.getch()
            
            return False
            
        except socket.timeout:
            print(f"✗ 연결 시간 초과")
            print(f"✗ SMTP 서버: {self.smtp_server}:{self.smtp_port}")
            print("\n네트워크 연결을 확인해주세요.")
            print("="*60 + "\n")
            return False
            
        except Exception as e:
            print(f"✗ 연결 실패: {e}")
            print("="*60 + "\n")
            return False
    
    def extract_company_name(self, filename):
        """파일명에서 회사명 추출"""
        match = re.match(self.pattern, filename)
        if match:
            return match.group(1).strip()
        return None
    
    def format_email_content(self, template_name, company_name, filename):
        """이메일 내용 포맷팅"""
        if template_name not in self.email_templates:
            raise ValueError(f"양식 '{template_name}'을 찾을 수 없습니다.")
        
        template = self.email_templates[template_name]
        subject = template.get('subject', '')
        body = template.get('body', '')
        
        # 변수 치환
        now = datetime.now()
        replacements = {
            '{회사명}': company_name,
            '{파일명}': filename,
            '{날짜}': now.strftime('%Y-%m-%d'),
            '{시간}': now.strftime('%H:%M:%S')
        }
        
        for key, value in replacements.items():
            subject = subject.replace(key, value)
            body = body.replace(key, value)
        
        return subject, body
    
    def send_email(self, to_emails, subject, body, pdf_paths):
        """이메일 발송 (여러 PDF 첨부 가능)"""
        try:
            # 파일 크기 체크 (Gmail 25MB 제한)
            total_size = sum(os.path.getsize(pdf_path) for pdf_path in pdf_paths)
            max_size = 24 * 1024 * 1024  # 24MB (여유 있게)
            
            if total_size > max_size:
                size_mb = total_size / (1024 * 1024)
                logging.warning(f"⚠ 첨부 파일 크기가 너무 큽니다: {size_mb:.1f}MB")
                logging.warning(f"   Gmail 제한: 25MB")
                logging.warning(f"   파일: {[str(p.name) for p in pdf_paths]}")
                logging.warning(f"   → 이메일 발송을 건너뜁니다.")
                return False
            
            # 이메일 메시지 생성
            msg = MIMEMultipart()
            msg['From'] = self.sender_email
            msg['To'] = ', '.join(to_emails)
            msg['Subject'] = subject
            
            # 본문에 첨부 파일 목록 추가 (여러 파일인 경우)
            if len(pdf_paths) > 1:
                file_list = '\n'.join([f"- {pdf.name}" for pdf in pdf_paths])
                body = body + f"\n\n[첨부 파일]\n{file_list}"
            
            msg.attach(MIMEText(body, 'plain', 'utf-8'))
            
            # PDF 파일들 첨부
            for pdf_path in pdf_paths:
                with open(pdf_path, 'rb') as f:
                    pdf = MIMEApplication(f.read(), _subtype='pdf')
                    pdf.add_header('Content-Disposition', 'attachment', 
                                 filename=('utf-8', '', pdf_path.name))
                    msg.attach(pdf)
            
            # SMTP 서버 연결 및 발송
            server = smtplib.SMTP(self.smtp_server, self.smtp_port, timeout=30)
            server.starttls()
            server.login(self.sender_email, self.sender_password)
            server.send_message(msg)
            server.quit()
            
            logging.info(f"✓ 이메일 발송 성공: {', '.join(to_emails)}")
            return True
            
        except Exception as e:
            logging.error(f"이메일 발송 오류: {e}")
            return False
    
    def analyze_pdfs(self):
        """전송할PDF 폴더의 모든 PDF 분석 (재귀적)"""
        company_pdfs = {}
        unrecognized_files = []
        no_company_info = []
        oversized_files = []
        
        # PDF 파일 재귀적으로 검색
        pdf_files = list(self.pdf_dir.rglob('*.pdf'))
        
        for pdf_path in pdf_files:
            # 상대 경로 계산
            rel_path = pdf_path.relative_to(self.pdf_dir)
            file_size = os.path.getsize(pdf_path)
            
            # 회사명 추출
            company_name = self.extract_company_name(pdf_path.name)
            
            if not company_name:
                unrecognized_files.append((rel_path, file_size))
                continue
            
            # 회사 정보 확인
            if company_name not in self.company_db:
                no_company_info.append((company_name, rel_path, file_size))
                continue
            
            # 파일 크기 체크
            if file_size > 24 * 1024 * 1024:
                oversized_files.append((company_name, rel_path, file_size))
            
            # 회사별로 PDF 그룹화
            if company_name not in company_pdfs:
                company_pdfs[company_name] = []
            company_pdfs[company_name].append((pdf_path, rel_path))
        
        return company_pdfs, unrecognized_files, no_company_info, oversized_files
    
    def print_analysis_report(self, company_pdfs, unrecognized_files, no_company_info, oversized_files):
        """분석 결과 출력"""
        print("\n" + "="*60)
        print("PDF 파일 분석 결과")
        print("="*60)
        
        # 발송 가능한 파일
        if company_pdfs:
            print(f"\n✓ 발송 가능: {sum(len(pdfs) for pdfs in company_pdfs.values())}개 파일")
            for company, pdfs in sorted(company_pdfs.items()):
                print(f"\n  [{company}] - {len(pdfs)}개 파일")
                for pdf_path, rel_path in pdfs:
                    size_mb = os.path.getsize(pdf_path) / (1024 * 1024)
                    print(f"    - {rel_path} ({size_mb:.1f}MB)")
        else:
            print("\n✓ 발송 가능: 0개")
        
        # 인식 실패
        if unrecognized_files:
            print(f"\n⚠ 파일명 패턴 불일치: {len(unrecognized_files)}개")
            for rel_path, file_size in unrecognized_files:
                size_mb = file_size / (1024 * 1024)
                print(f"    - {rel_path} ({size_mb:.1f}MB)")
        
        # 회사 정보 없음
        if no_company_info:
            print(f"\n⚠ 회사 정보 없음: {len(no_company_info)}개")
            for company, rel_path, file_size in no_company_info:
                size_mb = file_size / (1024 * 1024)
                print(f"    - [{company}] {rel_path} ({size_mb:.1f}MB)")
        
        # 크기 초과
        if oversized_files:
            print(f"\n⚠ 파일 크기 초과 (25MB): {len(oversized_files)}개")
            for company, rel_path, file_size in oversized_files:
                size_mb = file_size / (1024 * 1024)
                print(f"    - [{company}] {rel_path} ({size_mb:.1f}MB)")
        
        print("\n" + "="*60)
    
    def wait_for_input_with_timeout(self, timeout_seconds, auto_action='send'):
        """
        사용자 입력 대기 (타임아웃 포함)
        timeout_seconds > 0: 타임아웃 후 자동 실행
        timeout_seconds == 0: 즉시 실행
        timeout_seconds < 0: 무한 대기 (자동 실행 비활성화)
        """
        if timeout_seconds == 0:
            return auto_action
        
        if timeout_seconds < 0:
            # 무한 대기
            print("\n[Enter: 발송 | ESC: 취소]", end='', flush=True)
            if msvcrt:
                while True:
                    if msvcrt.kbhit():
                        key = msvcrt.getch()
                        if key == b'\r':  # Enter
                            print(" → 발송")
                            return 'send'
                        elif key == b'\x1b':  # ESC
                            print(" → 취소")
                            return 'cancel'
            else:
                # Windows 아닌 경우
                input()
                return 'send'
        
        # 타임아웃 카운트다운
        remaining = timeout_seconds
        while remaining > 0:
            print(f"\r{remaining}초 후 자동 {auto_action}... [Enter: 즉시 발송 | ESC: 취소]", 
                  end='', flush=True)
            
            if msvcrt and msvcrt.kbhit():
                key = msvcrt.getch()
                if key == b'\r':  # Enter
                    print(" → 발송")
                    return 'send'
                elif key == b'\x1b':  # ESC
                    print(" → 취소")
                    return 'cancel'
            
            time.sleep(1)
            remaining -= 1
        
        print(f"\r시간 초과 - 자동 {auto_action}                                        ")
        return auto_action
    
    def scan_and_process(self):
        """PDF 스캔 및 이메일 발송"""
        try:
            # 1. PDF 분석
            print("\nPDF 파일 검색 중...")
            company_pdfs, unrecognized, no_info, oversized = self.analyze_pdfs()
            
            # 2. 분석 결과 출력
            self.print_analysis_report(company_pdfs, unrecognized, no_info, oversized)
            
            if not company_pdfs:
                print("\n발송할 이메일이 없습니다.")
                return
            
            # 3. 발송 확인
            action = self.wait_for_input_with_timeout(self.auto_send_timeout, 'send')
            
            if action == 'cancel':
                print("\n이메일 발송이 취소되었습니다.")
                return
            
            # 4. 이메일 발송
            print("\n" + "="*60)
            print("이메일 발송 시작")
            print("="*60 + "\n")
            
            success_count = 0
            fail_count = 0
            
            for company_name, pdfs in company_pdfs.items():
                try:
                    company_info = self.company_db[company_name]
                    emails = company_info['emails']
                    template_name = company_info['template']
                    
                    # 이메일 내용 생성
                    pdf_paths = [pdf_path for pdf_path, _ in pdfs]
                    first_filename = pdf_paths[0].name
                    subject, body = self.format_email_content(template_name, company_name, first_filename)
                    
                    # 이메일 발송
                    print(f"[{company_name}] {len(pdf_paths)}개 파일 발송 중...")
                    if self.send_email(emails, subject, body, pdf_paths):
                        success_count += len(pdf_paths)
                        
                        # 파일 이동 (폴더 구조 유지)
                        for pdf_path, rel_path in pdfs:
                            dest_dir = self.completed_dir / rel_path.parent
                            dest_dir.mkdir(parents=True, exist_ok=True)
                            dest_path = dest_dir / pdf_path.name
                            
                            pdf_path.rename(dest_path)
                            logging.info(f"파일 이동: {rel_path} → 전송완료/{rel_path}")
                    else:
                        fail_count += len(pdf_paths)
                        
                except Exception as e:
                    logging.error(f"[{company_name}] 처리 오류: {e}")
                    fail_count += len(pdfs)
            
            # 5. 결과 요약
            print("\n" + "="*60)
            print("발송 완료")
            print("="*60)
            print(f"성공: {success_count}개")
            print(f"실패: {fail_count}개")
            print("="*60 + "\n")
            
        except Exception as e:
            logging.error(f"처리 오류: {e}")
            raise
    
    def monitor_folder(self):
        """PDF 폴더 실시간 모니터링"""
        print("\n" + "="*60)
        print("📁 PDF 폴더 실시간 모니터링 시작")
        print("="*60)
        print(f"감시 중인 폴더: {self.pdf_dir}")
        print("\nPDF 파일이 추가되면 자동으로 이메일을 발송합니다.")
        print("종료하려면 Ctrl+C를 누르세요.")
        print("="*60 + "\n")
        
        # 이미 처리된 파일 기록
        processed_files = set()
        
        # 기존 파일들을 처리 완료 목록에 추가
        for pdf_path in self.pdf_dir.rglob('*.pdf'):
            processed_files.add(pdf_path)
        
        try:
            while True:
                # 현재 PDF 파일 목록
                current_files = set(self.pdf_dir.rglob('*.pdf'))
                
                # 새로 추가된 파일 찾기
                new_files = current_files - processed_files
                
                if new_files:
                    for pdf_path in new_files:
                        print(f"\n🆕 새 파일 감지: {pdf_path.name}")
                        
                        # 회사명 추출
                        company_name = self.extract_company_name(pdf_path.name)
                        
                        if not company_name:
                            print(f"   ⚠️  파일명 패턴 불일치 - 건너뜀")
                            processed_files.add(pdf_path)
                            continue
                        
                        if company_name not in self.company_db:
                            print(f"   ⚠️  회사 정보 없음 ({company_name}) - 건너뜀")
                            processed_files.add(pdf_path)
                            continue
                        
                        # 파일 크기 체크
                        file_size = os.path.getsize(pdf_path)
                        if file_size > 24 * 1024 * 1024:
                            size_mb = file_size / (1024 * 1024)
                            print(f"   ⚠️  파일 크기 초과 ({size_mb:.1f}MB) - 건너뜀")
                            processed_files.add(pdf_path)
                            continue
                        
                        # 이메일 발송
                        try:
                            company_info = self.company_db[company_name]
                            emails = company_info['emails']
                            template_name = company_info['template']
                            
                            subject, body = self.format_email_content(template_name, company_name, pdf_path.name)
                            
                            print(f"   📧 [{company_name}]에게 발송 중...")
                            if self.send_email(emails, subject, body, [pdf_path]):
                                print(f"   ✅ 발송 성공!")
                                
                                # 파일 이동
                                rel_path = pdf_path.relative_to(self.pdf_dir)
                                dest_dir = self.completed_dir / rel_path.parent
                                dest_dir.mkdir(parents=True, exist_ok=True)
                                dest_path = dest_dir / pdf_path.name
                                
                                pdf_path.rename(dest_path)
                                print(f"   📁 파일 이동: 전송완료/{rel_path}")
                            else:
                                print(f"   ❌ 발송 실패")
                        
                        except Exception as e:
                            print(f"   ❌ 처리 오류: {e}")
                            logging.error(f"[실시간] {pdf_path.name} 처리 오류: {e}")
                        
                        # 처리 완료 목록에 추가 (성공/실패 상관없이)
                        processed_files.add(pdf_path)
                
                # 5초 대기
                time.sleep(5)
        
        except KeyboardInterrupt:
            print("\n\n" + "="*60)
            print("📁 폴더 모니터링 종료")
            print("="*60 + "\n")

def main():
    """메인 함수"""
    try:
        print("\n" + "="*60)
        print("PDF 이메일 자동 발송 프로그램")
        print("="*60 + "\n")
        
        # 초기화
        sender = PDFEmailSender()
        sender.load_settings()
        
        # 이메일 연결 테스트
        sender.test_email_connection()
        
        # 모드 선택
        print("\n" + "="*60)
        print("모드 선택")
        print("="*60)
        print("1. 전송할PDF 폴더의 모든 PDF 발송")
        print("2. PDF 폴더 실시간 모니터링")
        print("3. 종료")
        print("="*60)
        
        # 타임아웃 처리
        if sender.auto_select_timeout == 0:
            # 즉시 실행
            choice = '1'
            print("즉시 실행: 옵션 1 선택")
        elif sender.auto_select_timeout < 0:
            # 무한 대기
            print("\n[Enter: 옵션 1 | ESC: 종료] 선택: ", end='', flush=True)
            if msvcrt:
                while True:
                    if msvcrt.kbhit():
                        key = msvcrt.getch()
                        if key == b'\r':  # Enter
                            choice = '1'
                            print("1")
                            break
                        elif key == b'\x1b':  # ESC
                            choice = '3'
                            print("3")
                            break
            else:
                choice = input() or '1'
        else:
            # 카운트다운
            remaining = sender.auto_select_timeout
            while remaining > 0:
                print(f"\r{remaining}초 후 자동 선택... [Enter: 옵션 1 | ESC: 종료] 선택: ", 
                      end='', flush=True)
                
                if msvcrt and msvcrt.kbhit():
                    key = msvcrt.getch()
                    if key == b'\r':  # Enter
                        choice = '1'
                        print("1")
                        break
                    elif key == b'\x1b':  # ESC
                        choice = '3'
                        print("3")
                        break
                else:
                    time.sleep(1)
                    remaining -= 1
            else:
                choice = '1'
                print("\r시간 초과 - 옵션 1 자동 선택                                        ")
        
        if choice == '1':
            sender.scan_and_process()
        elif choice == '2':
            sender.monitor_folder()
        elif choice == '3':
            print("\n프로그램을 종료합니다.")
            return
        else:
            print("\n잘못된 선택입니다.")
        
        print("\n프로그램을 종료하려면 아무 키나 누르세요...")
        if msvcrt:
            msvcrt.getch()
        else:
            input()
        
    except Exception as e:
        logging.error(f"프로그램 오류: {e}")
        print(f"\n오류 발생: {e}")
        print("\n프로그램을 종료하려면 아무 키나 누르세요...")
        if msvcrt:
            msvcrt.getch()
        else:
            input()

if __name__ == '__main__':
    main()

