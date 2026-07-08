import os
import shutil
from pathlib import Path
import time
import win32com.client
import pythoncom

# Z 드라이브 네트워크 경로 상수 정의 (사용자 규칙 7 준수)
Z_DRIVE = Path("Z:/")

def safe_mkdir(target_dir: Path):
    """
    네트워크 드라이브에서 exist_ok=True 임에도 WinError 183이 발생하는 버그를 방지하고,
    동일한 이름의 파일이 이미 존재하는 경우 해당 파일을 백업 폴더명으로 우회시켜 안전하게 디렉토리를 생성합니다.
    """
    if target_dir.exists():
        if target_dir.is_file():
            # 동일 이름의 '파일'이 존재하는 경우, 이 파일의 이름을 변경하여 우회합니다.
            backup_path = target_dir.with_name(f"{target_dir.name}_backup_file")
            print(f"[경고] '{target_dir}' 경로에 디렉토리가 아닌 파일이 존재합니다. 파일명을 '{backup_path.name}'으로 변경합니다.")
            try:
                # 덮어쓰기 방지를 위해 백업 경로가 이미 존재하면 삭제
                if backup_path.exists():
                    backup_path.unlink(missing_ok=True)
                os.rename(str(target_dir), str(backup_path))
            except Exception as rename_err:
                print(f"[오류] 기존 파일 '{target_dir.name}' 이름 변경 실패: {rename_err}")
                raise
        else:
            # 이미 디렉토리로 정상 존재하므로 생성 불필요
            return

    # 디렉토리 생성 시도
    try:
        target_dir.mkdir(parents=True, exist_ok=True)
    except FileExistsError:
        pass
    except OSError as e:
        if getattr(e, 'winerror', None) == 183 or target_dir.is_dir():
            pass
        else:
            raise

def safe_move_file(source_path: Path, target_dir: Path, retries: int = 3, delay: float = 0.5) -> bool:
    """
    파일을 대상 디렉토리로 안전하고 빠르게 이동합니다.
    대상 디렉토리가 없으면 생성하고, 이미 동일한 파일이 존재하면 덮어씁니다.
    네트워크 드라이브 및 Windows 지연 삭제 문제를 해결하기 위해 재시도 로직을 포함합니다.
    """
    # 작업의 시작 로그 명시 (사용자 규칙 2 준수)
    # print(f"[이동 시작] '{source_path.name}' -> {target_dir}")
    try:
        if not source_path.is_file():
            return False
        
        # 대상 디렉토리 생성
        safe_mkdir(target_dir)
        target_path = target_dir / source_path.name
        
        # 소스와 대상이 동일한 경로이면 아무 작업도 안 함
        if source_path.resolve() == target_path.resolve():
            return True
            
        for attempt in range(1, retries + 1):
            try:
                # 대상 파일이 이미 존재하면 명시적으로 먼저 삭제 시도 (덮어쓰기 강제)
                if target_path.exists():
                    try:
                        target_path.unlink(missing_ok=True)
                    except:
                        pass
                # os.replace는 동일 드라이브 내에서 대상 파일이 있어도 강제 원자적 덮어쓰기 수행
                os.replace(str(source_path), str(target_path))
                return True
            except OSError:
                # 다른 드라이브 간 이동이거나 네트워크 드라이브 특성에 따라 os.replace가 안 될 경우
                try:
                    if target_path.exists():
                        target_path.unlink(missing_ok=True)
                    shutil.copy2(str(source_path), str(target_path))
                    source_path.unlink(missing_ok=True)
                    return True
                except Exception as attempt_err:
                    if attempt == retries:
                        print(f"[오류] 파일 이동 최종 실패 ({source_path.name} -> {target_dir}): {attempt_err}")
                        return False
                    time.sleep(delay)
        return False
    except Exception as e:
        print(f"[오류] 파일 이동 중 예외 발생 ({source_path.name}): {e}")
        return False

def safe_copy_file(source_path: Path, target_dir: Path, retries: int = 3, delay: float = 0.5) -> bool:
    """
    파일을 대상 디렉토리로 안전하고 빠르게 복사합니다.
    대상 디렉토리가 없으면 생성하고, 이미 동일한 파일이 존재하면 덮어씁니다.
    """
    # 작업의 시작 로그 명시 (사용자 규칙 2 준수)
    # print(f"[복사 시작] '{source_path.name}' -> {target_dir}")
    try:
        if not source_path.is_file():
            return False
        
        # 대상 디렉토리 생성
        safe_mkdir(target_dir)
        target_path = target_dir / source_path.name
        
        # 소스와 대상이 동일한 경로이면 아무 작업도 안 함
        if source_path.resolve() == target_path.resolve():
            return True
            
        for attempt in range(1, retries + 1):
            try:
                if target_path.exists():
                    target_path.unlink(missing_ok=True)
                shutil.copy2(str(source_path), str(target_path))
                return True
            except Exception as attempt_err:
                if attempt == retries:
                    print(f"[오류] 파일 복사 최종 실패 ({source_path.name} -> {target_dir}): {attempt_err}")
                    return False
                time.sleep(delay)
        return False
    except Exception as e:
        print(f"[오류] 파일 복사 중 예외 발생 ({source_path.name}): {e}")
        return False

def process_kb_swap_files(kb_swap_dir: Path):
    """
    KB-SWAP 폴더를 1회 순회하여 파일명 키워드에 따라 opening, closing, position, collateral 폴더로 분류 이동합니다.
    """
    print(f"[작업 시작] KB-SWAP 파일 통합 분류: {kb_swap_dir}")
    if not kb_swap_dir.exists():
        print(f"[경고] KB-SWAP 디렉토리가 존재하지 않습니다: {kb_swap_dir}")
        return
        
    files_moved = 0
    try:
        opening_dir = kb_swap_dir / "opening"
        closing_dir = kb_swap_dir / "closing"
        position_dir = kb_swap_dir / "position"
        collateral_dir = kb_swap_dir / "collateral"
        
        for item in kb_swap_dir.iterdir():
            if not item.is_file():
                continue
            # 디렉토리는 제외하고 파일인 경우에만 검사하며, .pdf 파일은 제외 (기존 로직 준수)
            if item.name.endswith(".pdf"):
                continue
                
            target_dir = None
            if "New_Trade" in item.name:
                target_dir = opening_dir
            elif "Unwind_Trade" in item.name:
                target_dir = closing_dir
            elif "Daily_P&L_" in item.name:
                target_dir = position_dir
            elif "Collateral Summary" in item.name:
                target_dir = collateral_dir
                
            if target_dir:
                if safe_move_file(item, target_dir):
                    files_moved += 1
                    print(f"  └ '{item.name}' 파일 이동 완료 -> {target_dir.name}")
                    
        print(f"[작업 완료] KB-SWAP 총 {files_moved}개 파일 분류 완료")
    except Exception as e:
        print(f"[오류] KB-SWAP 파일 분류 중 오류 발생: {e}")

def process_kis_swap_files(kis_swap_dir: Path):
    """
    KIS-SWAP 폴더를 1회 순회하여 파일명 키워드에 따라 opening, closing, position 폴더로 분류 이동합니다.
    """
    print(f"[작업 시작] KIS-SWAP 파일 통합 분류: {kis_swap_dir}")
    if not kis_swap_dir.exists():
        print(f"[경고] KIS-SWAP 디렉토리가 존재하지 않습니다: {kis_swap_dir}")
        return
        
    files_moved = 0
    try:
        opening_dir = kis_swap_dir / "opening"
        closing_dir = kis_swap_dir / "closing"
        position_dir = kis_swap_dir / "position"
        
        for item in kis_swap_dir.iterdir():
            if not item.is_file():
                continue
            if item.name.endswith(".pdf"):
                continue
                
            target_dir = None
            if "New Trade" in item.name:
                target_dir = opening_dir
            elif "Termination" in item.name:
                target_dir = closing_dir
            elif "Open Position" in item.name:
                target_dir = position_dir
                
            if target_dir:
                if safe_move_file(item, target_dir):
                    files_moved += 1
                    print(f"  └ '{item.name}' 파일 이동 완료 -> {target_dir.name}")
                    
        print(f"[작업 완료] KIS-SWAP 총 {files_moved}개 파일 분류 완료")
    except Exception as e:
        print(f"[오류] KIS-SWAP 파일 분류 중 오류 발생: {e}")

def process_prelude_files(prelude_dir: Path):
    """
    PRELUDE_MTM 폴더를 1회 순회하여 특정 키워드의 CSV 파일을 ES16, ES24, ES40 폴더로 복사하고,
    오래된 보고서 파일(특정 키워드로 시작하는 파일)은 old 폴더로 백업 이동합니다.
    (복사 후 백업 이동이 순차적으로 처리되도록 보장합니다.)
    """
    print(f"[작업 시작] PRELUDE_MTM 파일 통합 복사 및 백업 이동: {prelude_dir}")
    if not prelude_dir.exists():
        print(f"[경고] PRELUDE_MTM 디렉토리가 존재하지 않습니다: {prelude_dir}")
        return
        
    files_copied = 0
    files_moved = 0
    backup_keywords = ["EQSWAP16", "EQSWAP16X", "EQSWAP24M", "EQSWAP24MX", "EQSWAP40", "EQSWAP40X"]
    
    try:
        es16_dir = prelude_dir / "ES16"
        es24_dir = prelude_dir / "ES24"
        es40_dir = prelude_dir / "ES40"
        old_dir = prelude_dir / "old"
        
        # 순회 도중 파일이 이동되면 iterdir 결과가 중간에 바뀔 수 있으므로 리스트로 미리 가져옴
        file_list = [item for item in prelude_dir.iterdir() if item.is_file()]
        
        for item in file_list:
            is_csv = item.name.endswith(".csv")
            
            # 1. CSV 보고서 복사 대상 검사
            if is_csv:
                target_copy_dir = None
                if "EQSWAP16X" in item.name:
                    target_copy_dir = es16_dir
                elif "EQSWAP24MX" in item.name:
                    target_copy_dir = es24_dir
                elif "EQSWAP40X" in item.name:
                    target_copy_dir = es40_dir
                    
                if target_copy_dir:
                    if safe_copy_file(item, target_copy_dir):
                        files_copied += 1
                        print(f"  └ '{item.name}' 파일 복사 완료 -> {target_copy_dir.name}")
                        
            # 2. 백업 이동(old) 대상 검사 (복사 완료 후 또는 복사 대상이 아니더라도 키워드 충족 시 이동)
            if any(item.name.startswith(kw) for kw in backup_keywords):
                if safe_move_file(item, old_dir):
                    files_moved += 1
                    print(f"  └ '{item.name}' 파일 백업 완료 -> old")
                    
        print(f"[작업 완료] PRELUDE_MTM 총 {files_copied}개 파일 복사, {files_moved}개 파일 백업 완료")
    except Exception as e:
        print(f"[오류] PRELUDE_MTM 파일 처리 중 오류 발생: {e}")

def refresh_excel_queries(file_path: Path):
    """
    Excel COM API를 호출하여 지정된 엑셀 파일 내의 모든 외부 데이터 쿼리 및 피벗 테이블을 새로고침합니다.
    """
    print(f"[작업 시작] 엑셀 쿼리 새로고침 시작: {file_path}")
    excel = None
    wb = None
    try:
        # COM 라이브러리 초기화
        pythoncom.CoInitialize()
        
        # 엑셀 애플리케이션 실행
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False  # 팝업 경고 메시지 방지
        
        # 파일 존재 여부 확인
        if not file_path.exists():
            print(f"[오류] 엑셀 파일이 존재하지 않습니다: {file_path}")
            return
            
        # 절대 경로로 파일 오픈
        wb = excel.Workbooks.Open(str(file_path.absolute()))

        # 모든 외부 데이터 연결 및 피벗 새로고침
        wb.RefreshAll()

        # 비동기 쿼리가 완료될 때까지 대기
        excel.CalculateUntilAsyncQueriesDone()

        # 변경사항 저장 후 닫기
        wb.Save()
        wb.Close(SaveChanges=True)
        print(f"[작업 성공] {file_path.name}의 쿼리가 성공적으로 새로고침되었습니다.")
    except Exception as e:
        print(f"[오류] {file_path.name} 쿼리 새로고침 중 오류 발생: {e}")
    finally:
        # Excel 자원 해제 및 프로세스 종료 보장 (try-finally 패턴)
        try:
            if wb is not None:
                try:
                    wb.Close(SaveChanges=False)
                except:
                    pass
            if excel is not None:
                excel.Quit()
        except Exception as close_err:
            print(f"[경고] 엑셀 프로세스 종료 시 에러 발생: {close_err}")
        finally:
            # COM 객체 완전 해제
            excel = None
            wb = None
            pythoncom.CoUninitialize()
            print("[작업 완료] 엑셀 쿼리 새로고침 프로세스가 종료되었습니다.")

def main():
    print("[시스템 시작] 파일 이동 및 엑셀 쿼리 새로고침 배치 작업을 시작합니다.")
    
    # 1. KB-SWAP 관련 파일 이동
    kb_swap_dir = Z_DRIVE / "02.펀드" / "003.매매보고서 대사" / "KB-SWAP"
    process_kb_swap_files(kb_swap_dir)

    # 2. KIS-SWAP 관련 파일 이동
    kis_swap_dir = Z_DRIVE / "02.펀드" / "003.매매보고서 대사" / "KIS-SWAP"
    process_kis_swap_files(kis_swap_dir)

    # 3 & 4. PRELUDE MTM 관련 파일 복사 및 백업 이동 (통합 처리)
    prelude_dir = Z_DRIVE / "02.펀드" / "003.매매보고서 대사" / "PRELUDE_MTM"
    process_prelude_files(prelude_dir)

    # 5. EQSWAP.xlsx 쿼리 자동 새로고침
    excel_report_path = prelude_dir / "EQSWAP.xlsx"
    refresh_excel_queries(excel_report_path)

    print("[시스템 종료] 모든 파일 이동 및 쿼리 새로고침 작업이 성공적으로 완료되었습니다.")

if __name__ == "__main__":
    main()