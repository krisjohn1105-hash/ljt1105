# Auto Git Sync (30분 주기 자동 Git 동기화)

이 모듈은 `d:\PythonProjects\ljt1105` 저장소의 변경사항을 30분마다 자동으로 감지하여 `git pull`, `git add`, `git commit`, `git push`를 안전하게 수행합니다.

---

## 1. 사용 방법

### 방법 1. Windows 작업 스케줄러 등록 (권장 - 백그라운드 자동 실행)
터미널을 열어두지 않아도 PC가 켜져 있는 동안 30분마다 백그라운드(`pythonw.exe`)에서 조용히 실행됩니다.

- **등록**: `register_scheduler.bat` 더블 클릭 또는
  ```powershell
  python AutoGitSync/auto_git_sync.py --register
  ```
- **해제**: `unregister_scheduler.bat` 더블 클릭 또는
  ```powershell
  python AutoGitSync/auto_git_sync.py --unregister
  ```
- **상태 확인**:
  ```powershell
  python AutoGitSync/auto_git_sync.py --status
  ```

---

### 방법 2. 1회 즉시 수동 동기화
지금 즉시 pull, commit, push를 1회 실행하고 싶을 때 사용합니다.

- `run_sync.bat` 더블 클릭 또는
  ```powershell
  python AutoGitSync/auto_git_sync.py --once
  ```

---

### 방법 3. 상주형 터미널 데몬 모드
터미널 창에서 30분 주기로 계속 돌아가게 하고 싶을 때 사용합니다.

```powershell
python AutoGitSync/auto_git_sync.py --loop --interval 1800
```

---

## 2. 동작 흐름
1. 현재 Git 브랜치 확인 (기본 `main`)
2. 로컬 변경 사항 감지 (`git status --porcelain`)
3. 변경 사항이 있을 경우 `git add -A` 및 `Auto commit: YYYY-MM-DD HH:MM:SS (N files changed)` 커밋 생성
4. 원격 저장소로부터 최신 상태 수신 (`git pull --rebase --autostash origin <branch>`)
   - 충돌 감지 시 자동으로 중단(`rebase --abort`) 후 에러 로그 기록
5. 원격 저장소로 로컬 커밋 푸시 (`git push origin <branch>`)
6. 실행 결과 및 타임스탬프를 `AutoGitSync/logs/git_sync_YYYY-MM-DD.log` 파일에 안전하게 누적 기록

---

## 3. 로그 파일 위치
- `d:\PythonProjects\ljt1105\AutoGitSync\logs\git_sync_YYYY-MM-DD.log`
- 매일 날짜별로 분할되어 저장됩니다.
