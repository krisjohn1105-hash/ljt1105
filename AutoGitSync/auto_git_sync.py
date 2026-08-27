"""
==============================================================================
프로그램명: Auto Git Sync (자동 Git 동기화 시스템)
작성 목적: 30분 주기로 Git 저장소의 변경사항을 확인하여 자동으로 Pull, Commit, Push 수행
==============================================================================
"""

import argparse
import datetime
import logging
import os
from pathlib import Path
import subprocess
import sys
import time
from typing import Optional, Tuple


class GitSyncManager:
    """
    Git 자동 동기화(Pull, Commit, Push) 및 스케줄러 관리를 담당하는 모듈
    """

    def __init__(self, repo_dir: Optional[Path] = None):
        """
        초기화 함수: 저장소 경로 및 로깅 설정
        """
        # 저장소 루트 경로 설정 (기본값: 현재 스크립트 위치의 상위 디렉터리)
        if repo_dir is None:
            self.repo_dir = Path(__file__).resolve().parent.parent
        else:
            self.repo_dir = Path(repo_dir).resolve()

        # 로그 디렉터리 설정
        self.log_dir = Path(__file__).resolve().parent / "logs"
        self.log_dir.mkdir(parents=True, exist_ok=True)

        # 로거 설정
        self.logger = self._setup_logger()

    def _setup_logger(self) -> logging.Logger:
        """
        한글 및 파일 출력을 지원하는 로거 설정 함수
        """
        logger = logging.getLogger("AutoGitSync")
        logger.setLevel(logging.INFO)

        # 핸들러 중복 방지
        if not logger.handlers:
            today_str = datetime.datetime.now().strftime("%Y-%m-%d")
            log_file = self.log_dir / f"git_sync_{today_str}.log"

            # 파일 핸들러 (UTF-8 인코딩)
            file_handler = logging.FileHandler(log_file, encoding="utf-8")
            file_handler.setLevel(logging.INFO)

            # 콘솔 스트림 핸들러
            stream_handler = logging.StreamHandler(sys.stdout)
            stream_handler.setLevel(logging.INFO)

            # 포맷터 설정
            formatter = logging.Formatter(
                "[%(asctime)s] [%(levelname)s] %(message)s",
                datefmt="%Y-%m-%d %H:%M:%S"
            )
            file_handler.setFormatter(formatter)
            stream_handler.setFormatter(formatter)

            logger.addHandler(file_handler)
            logger.addHandler(stream_handler)

        return logger

    def _run_git_command(self, args: list[str]) -> Tuple[bool, str, str]:
        """
        Git 명령어를 실행하고 결과를 반환하는 안전한 실행 함수 (예외 처리 포함)
        """
        cmd = ["git"] + args
        try:
            result = subprocess.run(
                cmd,
                cwd=str(self.repo_dir),
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                check=False
            )
            is_success = (result.returncode == 0)
            return is_success, result.stdout.strip(), result.stderr.strip()
        except Exception as e:
            self.logger.error(f"[Git 명령 실행 예외 발생] 명령: {' '.join(cmd)} | 오류: {str(e)}")
            return False, "", str(e)

    def get_current_branch(self) -> str:
        """
        현재 Git 브랜치 이름을 가져오는 함수
        """
        success, stdout, _ = self._run_git_command(["branch", "--show-current"])
        if success and stdout:
            return stdout
        return "main"

    def has_changes(self) -> bool:
        """
        로컬 워킹 디렉터리에 커밋되지 않은 변경사항(추적/미추적 파일 포함)이 있는지 확인하는 함수
        """
        success, stdout, _ = self._run_git_command(["status", "--porcelain"])
        if success and stdout:
            return True
        return False

    def pull_changes(self, branch: str) -> bool:
        """
        원격 저장소(origin)로부터 최신 커밋을 가져오는 함수 (git pull --rebase)
        """
        self.logger.info(f"[Pull 시작] 원격 저장소 origin/{branch} 로부터 최신 변경사항 수신 시도...")
        # 충돌 최소화를 위해 rebase 및 autostash 옵션 사용
        success, stdout, stderr = self._run_git_command(["pull", "--rebase", "--autostash", "origin", branch])
        if success:
            self.logger.info(f"[Pull 완료] 성공: {stdout if stdout else '최신 상태 유지 중'}")
            return True
        else:
            self.logger.warning(f"[Pull 경고/실패] stderr: {stderr} | stdout: {stdout}")
            # 만약 rebase 충돌 상태라면 안전하게 중단(abort)
            if "conflict" in stderr.lower() or "conflict" in stdout.lower():
                self.logger.error("[충돌 발생] Pull 도중 충돌이 발생하여 안전을 위해 rebase --abort 를 수행합니다. 수동 확인이 필요합니다.")
                self._run_git_command(["rebase", "--abort"])
            return False

    def commit_changes(self) -> Tuple[bool, int]:
        """
        모든 변경사항을 스테이징(git add -A)하고 자동 커밋 메시지로 커밋하는 함수
        """
        self.logger.info("[Add/Commit 시작] 변경사항 스테이징 및 커밋 생성 시도...")
        
        # 1. git add -A
        add_success, _, add_err = self._run_git_command(["add", "-A"])
        if not add_success:
            self.logger.error(f"[Add 실패] 파일 스테이징 실패: {add_err}")
            return False, 0

        # 변경 파일 수 확인
        status_ok, status_out, _ = self._run_git_command(["status", "--porcelain"])
        if not status_ok or not status_out:
            self.logger.info("[Commit 생략] 스테이징된 변경사항이 없습니다.")
            return True, 0

        changed_count = len(status_out.splitlines())
        current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        commit_msg = f"Auto commit: {current_time} ({changed_count} files changed)"

        # 2. git commit
        commit_success, commit_out, commit_err = self._run_git_command(["commit", "-m", commit_msg])
        if commit_success:
            self.logger.info(f"[Commit 완료] 커밋 생성 성공: {commit_msg}")
            return True, changed_count
        else:
            self.logger.error(f"[Commit 실패] 커밋 생성 실패: {commit_err} | {commit_out}")
            return False, 0

    def push_changes(self, branch: str) -> bool:
        """
        로컬 커밋들을 원격 저장소(origin)로 푸시하는 함수
        """
        self.logger.info(f"[Push 시작] 로컬 커밋을 원격 저장소 origin/{branch} 로 푸시 시도...")
        success, stdout, stderr = self._run_git_command(["push", "origin", branch])
        if success:
            self.logger.info(f"[Push 완료] 성공적으로 푸시되었습니다. {stdout}")
            return True
        else:
            self.logger.error(f"[Push 실패] 푸시 실패: {stderr} | {stdout}")
            return False

    def sync_repository(self) -> bool:
        """
        전체 Git 동기화 (Pull -> Commit -> Push) 워크플로 실행 함수
        """
        start_time = datetime.datetime.now()
        self.logger.info("=" * 60)
        self.logger.info(f"[자동 동기화 작업 시작] 경로: {self.repo_dir} | 시각: {start_time.strftime('%Y-%m-%d %H:%M:%S')}")

        try:
            # 1. 현재 브랜치 확인
            branch = self.get_current_branch()
            self.logger.info(f"[브랜치 확인] 현재 활성화된 브랜치: {branch}")

            # 2. 로컬 변경사항 여부 확인
            changes_detected = self.has_changes()
            committed_count = 0

            if changes_detected:
                self.logger.info("[로컬 변경 감지] 새로운 변경사항이 발견되어 커밋을 진행합니다.")
                commit_ok, committed_count = self.commit_changes()
                if not commit_ok:
                    self.logger.error("[동기화 중단] 로컬 커밋 생성에 실패하여 동기화를 중단합니다.")
                    return False
            else:
                self.logger.info("[로컬 변경 없음] 커밋할 새로운 파일 변경사항이 없습니다.")

            # 3. 원격 변경사항 Pull (rebase)
            pull_ok = self.pull_changes(branch)
            if not pull_ok:
                self.logger.warning("[Pull 실패/경고] 최신 변경사항 수신 중 문제가 발생했습니다.")

            # 4. 로컬에 푸시할 커밋이 있는지 확인 후 Push
            # (새로 커밋했거나, 이전에 푸시되지 않은 로컬 커밋이 있는 경우)
            success_ahead, stdout_ahead, _ = self._run_git_command(["status", "-sb"])
            need_push = committed_count > 0
            if success_ahead and "ahead" in stdout_ahead:
                need_push = True

            if need_push:
                push_ok = self.push_changes(branch)
                if not push_ok:
                    self.logger.error("[Push 실패] 원격 저장소 푸시에 실패했습니다.")
                    return False
            else:
                self.logger.info("[Push 생략] 원격으로 푸시할 새로운 커밋이 없습니다.")

            elapsed = (datetime.datetime.now() - start_time).total_seconds()
            self.logger.info(f"[자동 동기화 작업 완료] 소요 시간: {elapsed:.2f}초")
            self.logger.info("=" * 60)
            return True

        except Exception as e:
            self.logger.error(f"[동기화 중 예외 발생] 오류 메시지: {str(e)}")
            return False

    def register_windows_task(self, interval_minutes: int = 30) -> bool:
        """
        Windows 작업 스케줄러에 30분 주기 자동 동기화 작업을 등록하는 함수
        """
        task_name = "AutoGitSync_ljt1105"
        script_path = Path(__file__).resolve()
        # pythonw.exe를 사용하여 콘솔창 없이 조용히 백그라운드 실행
        python_exe = sys.executable
        pythonw_exe = Path(python_exe).parent / "pythonw.exe"
        if not pythonw_exe.exists():
            pythonw_exe = Path(python_exe)

        action_cmd = f'"{pythonw_exe}" "{script_path}" --once'

        self.logger.info(f"[작업 스케줄러 등록 시작] 작업 이름: {task_name}, 실행 주기: {interval_minutes}분")

        # schtasks 명령어로 등록 (30분마다 반복 실행, PC 부팅/로그온 시에도 동작)
        # /SC MINUTE /MO 30 /TN <task_name> /TR <action_cmd> /F
        cmd = [
            "schtasks", "/Create",
            "/TN", task_name,
            "/TR", action_cmd,
            "/SC", "MINUTE",
            "/MO", str(interval_minutes),
            "/F"
        ]

        try:
            result = subprocess.run(cmd, capture_output=True, text=True, check=False)
            if result.returncode == 0:
                self.logger.info(f"[작업 스케줄러 등록 성공] 작업 '{task_name}'이 {interval_minutes}분 주기로 성공적으로 등록되었습니다.")
                print(f"\n[성공] Windows 작업 스케줄러에 '{task_name}'이(가) 등록되었습니다. (매 {interval_minutes}분마다 자동 실행)")
                return True
            else:
                self.logger.error(f"[작업 스케줄러 등록 실패] stderr: {result.stderr}")
                print(f"\n[실패] 작업 스케줄러 등록 오류: {result.stderr}")
                return False
        except Exception as e:
            self.logger.error(f"[작업 스케줄러 등록 예외] 오류: {str(e)}")
            return False

    def unregister_windows_task(self) -> bool:
        """
        Windows 작업 스케줄러에서 자동 동기화 작업을 삭제하는 함수
        """
        task_name = "AutoGitSync_ljt1105"
        self.logger.info(f"[작업 스케줄러 삭제 시작] 작업 이름: {task_name}")

        cmd = ["schtasks", "/Delete", "/TN", task_name, "/F"]
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, check=False)
            if result.returncode == 0:
                self.logger.info(f"[작업 스케줄러 삭제 성공] 작업 '{task_name}'이 정상적으로 제거되었습니다.")
                print(f"\n[성공] 작업 '{task_name}'이(가) 스케줄러에서 삭제되었습니다.")
                return True
            else:
                self.logger.warning(f"[작업 스케줄러 삭제 안내] {result.stderr}")
                print(f"\n[안내] {result.stderr.strip()}")
                return False
        except Exception as e:
            self.logger.error(f"[작업 스케줄러 삭제 예외] 오류: {str(e)}")
            return False

    def query_windows_task(self) -> None:
        """
        Windows 작업 스케줄러 등록 상태를 조회하는 함수
        """
        task_name = "AutoGitSync_ljt1105"
        cmd = ["schtasks", "/Query", "/TN", task_name, "/FO", "LIST", "/V"]
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, encoding="cp949", errors="replace", check=False)
            if result.returncode == 0:
                print("\n=== Windows 작업 스케줄러 등록 상태 ===")
                print(result.stdout)
            else:
                print(f"\n[안내] 등록된 작업 '{task_name}'을(를) 찾을 수 없거나 아직 등록되지 않았습니다.")
        except Exception as e:
            print(f"[조회 오류] {str(e)}")

    def run_daemon(self, interval_seconds: int = 1800) -> None:
        """
        터미널에서 지속적으로 실행되는 상주형 데몬 모드 (30분 = 1800초)
        """
        self.logger.info(f"[데몬 모드 시작] 주기: {interval_seconds // 60}분 ({interval_seconds}초) 마다 자동 동기화 실행 중... (중단: Ctrl+C)")
        print(f"[*] Auto Git Sync 데몬이 실행되었습니다. (주기: {interval_seconds // 60}분)")
        print("[*] 종료하려면 Ctrl + C 를 누르세요.\n")

        try:
            while True:
                self.sync_repository()
                self.logger.info(f"[대기 중] 다음 실행까지 {interval_seconds // 60}분 대기합니다...")
                time.sleep(interval_seconds)
        except KeyboardInterrupt:
            self.logger.info("[데몬 모드 종료] 사용자에 의해 중단되었습니다.")
            print("\n[!] 자동 동기화 데몬이 종료되었습니다.")


def main():
    """
    메인 진입점 및 CLI 인자 파싱
    """
    parser = argparse.ArgumentParser(description="30분 주기 자동 Git 동기화 (Pull, Commit, Push) 도구")
    parser.add_argument("--once", action="store_true", help="1회 즉시 동기화 실행 (기본값)")
    parser.add_argument("--loop", action="store_true", help="상주형 데몬 모드로 반복 실행")
    parser.add_argument("--interval", type=int, default=1800, help="반복 실행 주기(초), 기본 1800초(30분)")
    parser.add_argument("--register", action="store_true", help="Windows 작업 스케줄러에 30분 주기 등록")
    parser.add_argument("--unregister", action="store_true", help="Windows 작업 스케줄러에서 작업 제거")
    parser.add_argument("--status", action="store_true", help="Windows 작업 스케줄러 등록 상태 확인")

    args = parser.parse_args()
    manager = GitSyncManager()

    if args.register:
        interval_min = args.interval // 60 if args.interval >= 60 else 30
        manager.register_windows_task(interval_minutes=interval_min)
    elif args.unregister:
        manager.unregister_windows_task()
    elif args.status:
        manager.query_windows_task()
    elif args.loop:
        manager.run_daemon(interval_seconds=args.interval)
    else:
        # 기본 동작: 1회 동기화 실행
        manager.sync_repository()


if __name__ == "__main__":
    main()
