from apscheduler.schedulers.blocking import BlockingScheduler
from pathlib import Path
import os
import win32com.client
import datetime
import time
import schedule
import zipfile
import move_files

def extract_attachments(output_dir, attachments):
    for attachment in attachments:
        attachment_name = str(attachment)
        attachment_path = output_dir / attachment_name

        if attachment_name.lower().endswith('.zip'):
            # 첨부 파일이 .zip 형식인 경우 압축을 풉니다.
            with zipfile.ZipFile(attachment_path, 'r') as zip_ref:
                zip_ref.extractall(output_dir)
            # os.remove(attachment_path)  # 압축 파일은 삭제할 수도 있습니다.
        else:
            # .zip이 아닌 경우 무시
            continue


def kb_pbs_trade_report_download():
    output_dir = Path("Z:/02.펀드/003.매매보고서 대사/KB")
    output_dir.mkdir(parents=True, exist_ok=True)
    outlook = win32com.client.Dispatch("outlook.application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6).Folders("KB-PBS")
    messages = inbox.Items

    for message in messages:
        if message.Unread:

            attachments = message.Attachments

            target_folder = output_dir
            target_folder.mkdir(parents=True, exist_ok=True)

            for attachment in attachments:
                attachment.SaveAsFile(target_folder / str(attachment))
                
                if message.Unread:
                    message.Unread = False

    print("KB-PBS Trade report download complete")


def nh_trade_report_download():
    output_dir = Path("Z:/02.펀드/003.매매보고서 대사/NH")
    output_dir.mkdir(parents=True, exist_ok=True)
    outlook = win32com.client.Dispatch("outlook.application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6).Folders("NH")
    messages = inbox.Items

    for message in messages:
        if message.Unread:

            attachments = message.Attachments

            target_folder = output_dir
            target_folder.mkdir(parents=True, exist_ok=True)

            for attachment in attachments:
                attachment.SaveAsFile(target_folder / str(attachment))
                
                if message.Unread:
                    message.Unread = False

    print("NH Trade report download complete")


def kis_pbs_trade_report_download():
    output_dir = Path("Z:/02.펀드/003.매매보고서 대사/KIS-PBS")
    output_dir.mkdir(parents=True, exist_ok=True)
    outlook = win32com.client.Dispatch("outlook.application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6).Folders("KIS-PBS")
    messages = inbox.Items

    for message in messages:
        if message.Unread:

            attachments = message.Attachments

            target_folder = output_dir
            target_folder.mkdir(parents=True, exist_ok=True)

            for attachment in attachments:
                attachment.SaveAsFile(target_folder / str(attachment))
                
                if message.Unread:
                    message.Unread = False

    print("KIS-PBS Trade report download complete")

def kis_trade_report_download():
    output_dir = Path("Z:/02.펀드/003.매매보고서 대사/KIS")
    output_dir.mkdir(parents=True, exist_ok=True)
    outlook = win32com.client.Dispatch("outlook.application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6).Folders("KIS")
    messages = inbox.Items

    for message in messages:
        if message.Unread:

            attachments = message.Attachments

            target_folder = output_dir
            target_folder.mkdir(parents=True, exist_ok=True)

            for attachment in attachments:
                attachment.SaveAsFile(target_folder / str(attachment))
                
                if message.Unread:
                    message.Unread = False

    print("KIS Trade report download complete")


def yuanta_trade_report_download():
    output_dir = Path("Z:/02.펀드/003.매매보고서 대사/Yuanta")
    output_dir.mkdir(parents=True, exist_ok=True)
    outlook = win32com.client.Dispatch("outlook.application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6).Folders("Yuanta")
    messages = inbox.Items

    for message in messages:
        if message.Unread:

            attachments = message.Attachments

            target_folder = output_dir
            target_folder.mkdir(parents=True, exist_ok=True)

            for attachment in attachments:
                attachment.SaveAsFile(target_folder / str(attachment))
                
                if message.Unread:
                    message.Unread = False

    print("Yuanta Trade report download complete")


def miraeasset_trade_report_download():
    output_dir = Path("Z:/02.펀드/003.매매보고서 대사/MiraeAsset")
    output_dir.mkdir(parents=True, exist_ok=True)
    outlook = win32com.client.Dispatch("outlook.application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6).Folders("MiraeAsset")
    messages = inbox.Items

    for message in messages:
        if message.Unread:

            attachments = message.Attachments

            target_folder = output_dir
            target_folder.mkdir(parents=True, exist_ok=True)

            for attachment in attachments:
                attachment.SaveAsFile(target_folder / str(attachment))
                
                if message.Unread:
                    message.Unread = False

    print("MiraeAsset Trade report download complete")


def hsbc_trade_report_download():
    output_dir = Path("Z:/02.펀드/003.매매보고서 대사/HSBC")
    output_dir.mkdir(parents=True, exist_ok=True)
    outlook = win32com.client.Dispatch("outlook.application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6).Folders("HSBC")
    messages = inbox.Items

    for message in messages:
        if message.Unread:

            attachments = message.Attachments

            target_folder = output_dir
            target_folder.mkdir(parents=True, exist_ok=True)

            for attachment in attachments:
                attachment.SaveAsFile(target_folder / str(attachment))
                
                if message.Unread:
                    message.Unread = False

    print("HSBC Trade report download complete")


def eugene_trade_report_download():
    output_dir = Path("Z:/02.펀드/003.매매보고서 대사/Eugene")
    output_dir.mkdir(parents=True, exist_ok=True)
    outlook = win32com.client.Dispatch("outlook.application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6).Folders("Eugene")
    messages = inbox.Items

    for message in messages:
        if message.Unread:

            attachments = message.Attachments

            target_folder = output_dir
            target_folder.mkdir(parents=True, exist_ok=True)

            for attachment in attachments:
                attachment.SaveAsFile(target_folder / str(attachment))
                
                if message.Unread:
                    message.Unread = False

    print("Eugene Trade report download complete")


def jpm_trade_report_download():
    output_dir = Path("Z:/02.펀드/003.매매보고서 대사/J.P.Morgan")
    output_dir.mkdir(parents=True, exist_ok=True)
    outlook = win32com.client.Dispatch("outlook.application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6).Folders("J.P.Morgan")
    messages = inbox.Items

    for message in messages:
        if message.Unread:

            attachments = message.Attachments

            target_folder = output_dir
            target_folder.mkdir(parents=True, exist_ok=True)

            for attachment in attachments:
                attachment.SaveAsFile(target_folder / str(attachment))
                
                if message.Unread:
                    message.Unread = False

    print("JPM Trade report download complete")


def clsa_trade_report_download():
    output_dir = Path("Z:/02.펀드/003.매매보고서 대사/CLSA")
    output_dir.mkdir(parents=True, exist_ok=True)
    outlook = win32com.client.Dispatch("outlook.application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6).Folders("CLSA")
    messages = inbox.Items

    for message in messages:
        if message.Unread:

            attachments = message.Attachments

            target_folder = output_dir
            target_folder.mkdir(parents=True, exist_ok=True)

            for attachment in attachments:
                attachment.SaveAsFile(target_folder / str(attachment))
                
                if message.Unread:
                    message.Unread = False

    print("CLSA Trade report download complete")


def GoldmanSachs_trade_report_download():
    output_dir = Path("Z:/02.펀드/003.매매보고서 대사/GoldmanSachs")
    output_dir.mkdir(parents=True, exist_ok=True)
    outlook = win32com.client.Dispatch("outlook.application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6).Folders("GoldmanSachs")
    messages = inbox.Items

    for message in messages:
        if message.Unread:

            attachments = message.Attachments

            target_folder = output_dir
            target_folder.mkdir(parents=True, exist_ok=True)

            for attachment in attachments:
                attachment.SaveAsFile(target_folder / str(attachment))
                
                if message.Unread:
                    message.Unread = False

    print("GoldmanSachs Trade report download complete")


def gs_pbs_sma_trade_report_download():
    swap_dir = Path("D:/PythonProjects/ljt1105/Dunamis_Workflow/data/input/QRT/Swap")
    cash_dir = Path("D:/PythonProjects/ljt1105/Dunamis_Workflow/data/input/QRT/Cash")
    swap_dir.mkdir(parents=True, exist_ok=True)
    cash_dir.mkdir(parents=True, exist_ok=True)
    outlook = win32com.client.Dispatch("outlook.application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6).Folders("GS PBS (SMA)")
    messages = inbox.Items

    for message in messages:
        if message.Unread:

            attachments = message.Attachments

            # 발신자 정보(이름/이메일 주소)로 synthetics/cash 여부를 판단합니다.
            sender_info = ""
            for attr in ("SenderName", "SenderEmailAddress"):
                try:
                    sender_info += str(getattr(message, attr, "")) + " "
                except Exception:
                    pass
            sender_info = sender_info.lower()

            if "synthetics" in sender_info:
                target_folder = swap_dir
            elif "cash" in sender_info:
                target_folder = cash_dir
            else:
                print(f"발신자를 확인할 수 없어 건너뜁니다: {message.Subject}")
                continue

            for attachment in attachments:
                attachment.SaveAsFile(target_folder / str(attachment))

                if message.Unread:
                    message.Unread = False

            # 압축해제 함수 호출
            extract_attachments(target_folder, attachments)

    print("GS PBS (SMA) Trade report download complete")


def hmc_trade_report_download():
    output_dir = Path("Z:/02.펀드/003.매매보고서 대사/HMC")
    output_dir.mkdir(parents=True, exist_ok=True)
    outlook = win32com.client.Dispatch("outlook.application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6).Folders("HMC")
    messages = inbox.Items

    for message in messages:
        if message.Unread:

            attachments = message.Attachments

            target_folder = output_dir
            target_folder.mkdir(parents=True, exist_ok=True)

            for attachment in attachments:
                attachment.SaveAsFile(target_folder / str(attachment))
                
                if message.Unread:
                    message.Unread = False

    print("HMC Trade report download complete")


def kb_swap_trade_report_download():
    output_dir = Path("Z:/02.펀드/003.매매보고서 대사/KB-SWAP")
    output_dir.mkdir(parents=True, exist_ok=True)
    outlook = win32com.client.Dispatch("outlook.application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6).Folders("KB-SWAP")
    messages = inbox.Items

    for message in messages:
        if message.Unread:

            attachments = message.Attachments

            target_folder = output_dir
            target_folder.mkdir(parents=True, exist_ok=True)

            for attachment in attachments:
                attachment.SaveAsFile(target_folder / str(attachment))
                
                if message.Unread:
                    message.Unread = False

            # 압축해제 함수 호출
            extract_attachments(output_dir, attachments)

    print("KB-SWAP report download complete")


def kis_swap_trade_report_download():
    output_dir = Path("Z:/02.펀드/003.매매보고서 대사/KIS-SWAP")
    output_dir.mkdir(parents=True, exist_ok=True)
    outlook = win32com.client.Dispatch("outlook.application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6).Folders("KIS-SWAP")
    messages = inbox.Items

    for message in messages:
        if message.Unread:

            attachments = message.Attachments

            target_folder = output_dir
            target_folder.mkdir(parents=True, exist_ok=True)

            for attachment in attachments:
                attachment.SaveAsFile(target_folder / str(attachment))
                
                if message.Unread:
                    message.Unread = False

            extract_attachments(output_dir, attachments)

    print("KIS-SWAP report download complete")

def CIMB_trade_report_download():
    output_dir = Path("Z:/02.펀드/003.매매보고서 대사/CGS-CIMB")
    output_dir.mkdir(parents=True, exist_ok=True)
    outlook = win32com.client.Dispatch("outlook.application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6).Folders("CGS-CIMB")
    messages = inbox.Items

    for message in messages:
        if message.Unread:

            attachments = message.Attachments

            target_folder = output_dir
            target_folder.mkdir(parents=True, exist_ok=True)

            for attachment in attachments:
                attachment.SaveAsFile(target_folder / str(attachment))
                
                if message.Unread:
                    message.Unread = False

            extract_attachments(output_dir, attachments)

    print("CGS-CIMB report download complete")

def MacQuarie_trade_report_download():
    output_dir = Path("Z:/02.펀드/003.매매보고서 대사/MacQuarie")
    output_dir.mkdir(parents=True, exist_ok=True)
    outlook = win32com.client.Dispatch("outlook.application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6).Folders("MacQuarie")
    messages = inbox.Items

    for message in messages:
        if message.Unread:

            attachments = message.Attachments

            target_folder = output_dir
            target_folder.mkdir(parents=True, exist_ok=True)

            for attachment in attachments:
                attachment.SaveAsFile(target_folder / str(attachment))
                
                if message.Unread:
                    message.Unread = False

            extract_attachments(output_dir, attachments)

    print("MacQuarie report download complete")

def DAOL_trade_report_download():
    output_dir = Path("Z:/02.펀드/003.매매보고서 대사/DAOL")
    output_dir.mkdir(parents=True, exist_ok=True)
    outlook = win32com.client.Dispatch("outlook.application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6).Folders("DAOL")
    messages = inbox.Items

    for message in messages:
        if message.Unread:

            attachments = message.Attachments

            target_folder = output_dir
            target_folder.mkdir(parents=True, exist_ok=True)

            for attachment in attachments:
                attachment.SaveAsFile(target_folder / str(attachment))
                
                if message.Unread:
                    message.Unread = False

            extract_attachments(output_dir, attachments)

    print("DAOL report download complete")

def DS_trade_report_download():
    output_dir = Path("Z:/02.펀드/003.매매보고서 대사/DS")
    output_dir.mkdir(parents=True, exist_ok=True)
    outlook = win32com.client.Dispatch("outlook.application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6).Folders("DS")
    messages = inbox.Items

    for message in messages:
        if message.Unread:

            attachments = message.Attachments

            target_folder = output_dir
            target_folder.mkdir(parents=True, exist_ok=True)

            for attachment in attachments:
                attachment.SaveAsFile(target_folder / str(attachment))
                
                if message.Unread:
                    message.Unread = False

            extract_attachments(output_dir, attachments)

    print("DS report download complete")

def Daiwa_trade_report_download():
    output_dir = Path("Z:/02.펀드/003.매매보고서 대사/Daiwa")
    output_dir.mkdir(parents=True, exist_ok=True)
    outlook = win32com.client.Dispatch("outlook.application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6).Folders("Daiwa")
    messages = inbox.Items

    for message in messages:
        if message.Unread:

            attachments = message.Attachments

            target_folder = output_dir
            target_folder.mkdir(parents=True, exist_ok=True)

            for attachment in attachments:
                attachment.SaveAsFile(target_folder / str(attachment))
                
                if message.Unread:
                    message.Unread = False

            extract_attachments(output_dir, attachments)

    print("Daiwa report download complete")

def tradeTeam_report_download():
    output_dir = Path("Z:/02.펀드/019. 일간매매내역")
    output_dir.mkdir(parents=True, exist_ok=True)
    outlook = win32com.client.Dispatch("outlook.application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6).Folders("거래내역")
    messages = inbox.Items

    for message in messages:
        if message.Unread:

            attachments = message.Attachments

            target_folder = output_dir
            target_folder.mkdir(parents=True, exist_ok=True)

            for attachment in attachments:
                attachment.SaveAsFile(target_folder / str(attachment))
                
                if message.Unread:
                    message.Unread = False

            extract_attachments(output_dir, attachments)

    print("Trade report download complete")

def OESERVER_report_download():
    output_dir = Path("Z:/02.펀드/003.매매보고서 대사/OESERVER")
    output_dir.mkdir(parents=True, exist_ok=True)
    outlook = win32com.client.Dispatch("outlook.application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6).Folders("OESERVER")
    messages = inbox.Items

    for message in messages:
        if message.Unread:

            attachments = message.Attachments

            target_folder = output_dir
            target_folder.mkdir(parents=True, exist_ok=True)

            for attachment in attachments:
                attachment.SaveAsFile(target_folder / str(attachment))
                
                if message.Unread:
                    message.Unread = False

            extract_attachments(output_dir, attachments)

    print("OESERVER report download complete")

def PRELUDE_RECAP_report_download():
    # [작업 시작 명시] PRELUDE_RECAP 매매보고서 다운로드 시작
    print(f"[{datetime.datetime.now()}] PRELUDE_RECAP 매매보고서 다운로드 및 pre-trade 복사 작업 시작")
    
    output_dir = Path("Z:/02.펀드/003.매매보고서 대사/PRELUDE_RECAP")
    pre_trade_dir = Path("D:/PythonProjects/ljt1105/Dunamis_Workflow/data/input/pre-trade")
    
    # 디렉토리가 없을 경우 생성
    output_dir.mkdir(parents=True, exist_ok=True)
    pre_trade_dir.mkdir(parents=True, exist_ok=True)
    
    outlook = win32com.client.Dispatch("outlook.application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6).Folders("PRELUDE_RECAP")
    messages = inbox.Items

    # 읽지 않은 메일 수집
    unread_messages = []
    for message in messages:
        try:
            if message.Unread:
                unread_messages.append(message)
        except Exception:
            # MailItem이 아니거나 Unread 속성이 없는 인스턴스 예외 처리
            pass

    # 수신 시간(ReceivedTime) 기준 오름차순(과거에서 최신 순)으로 정렬
    # 이 순서대로 덮어씌워지면 최신 메일에서 받은 동일 파일명의 파일이 최종 저장됩니다.
    unread_messages.sort(key=lambda m: getattr(m, 'ReceivedTime', datetime.datetime.min))

    for message in unread_messages:
        attachments = message.Attachments

        for attachment in attachments:
            attachment_name = str(attachment)
            
            # 1. Z 드라이브 원본 대상 폴더에 저장
            attachment.SaveAsFile(output_dir / attachment_name)
            
            # 2. D 드라이브 pre-trade 폴더에 저장
            attachment.SaveAsFile(pre_trade_dir / attachment_name)

        # 메일 읽음 처리
        try:
            if message.Unread:
                message.Unread = False
        except Exception:
            pass

        # zip 파일인 경우 압축 해제 처리 (원본 폴더와 pre-trade 폴더 둘 다 수행)
        extract_attachments(output_dir, attachments)
        extract_attachments(pre_trade_dir, attachments)

    # [작업 종료 명시] PRELUDE_RECAP 매매보고서 다운로드 완료
    print(f"[{datetime.datetime.now()}] PRELUDE_RECAP 매매보고서 다운로드 및 pre-trade 복사 작업 완료")  

def trade_report_download():
    import pythoncom
    # 스케줄러의 별도 스레드에서 win32com 객체를 생성하려면 CoInitialize()가 필수적입니다.
    pythoncom.CoInitialize()
    
    try:
        print(f"[{datetime.datetime.now()}] 매매보고서 다운로드 시작")
        print("===================================================")
        kb_pbs_trade_report_download()
        print("===================================================")
        nh_trade_report_download()
        print("===================================================")
        kis_pbs_trade_report_download()
        print("===================================================")
        kis_trade_report_download()
        print("===================================================")
        yuanta_trade_report_download()
        print("===================================================")
        miraeasset_trade_report_download()
        print("===================================================")
        hsbc_trade_report_download()
        print("===================================================")
        eugene_trade_report_download()
        print("===================================================")
        jpm_trade_report_download()
        print("===================================================")
        clsa_trade_report_download()
        print("===================================================")
        GoldmanSachs_trade_report_download()
        print("===================================================")
        gs_pbs_sma_trade_report_download()
        print("===================================================")
        hmc_trade_report_download()
        print("===================================================")
        kb_swap_trade_report_download()
        print("===================================================")
        kis_swap_trade_report_download()
        print("===================================================")
        CIMB_trade_report_download()
        print("===================================================")
        MacQuarie_trade_report_download()
        print("===================================================")
        DAOL_trade_report_download()
        print("===================================================")
        DS_trade_report_download()
        print("===================================================")
        Daiwa_trade_report_download()
        print("===================================================")
        tradeTeam_report_download()
        print("===================================================")
        OESERVER_report_download()
        print("===================================================")
        PRELUDE_RECAP_report_download()
        print("===================================================")
        
        # 파일 다운로드 후 move_files 스크립트를 통해 파일 분류
        # print("파일 이동 및 정리 작업 시작")
        # move_files.main()
        # print("===================================================")

        print(f"[{datetime.datetime.now()}] 매매보고서 다운로드 및 분류 완료")
    except Exception as e:
        print(f"에러 발생: {e}")
    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    # 스케줄러 객체 생성
    # scheduler = BlockingScheduler()

    # 1. 특정 주기마다 실행 (예: 1시간마다) - 필요시 주석 해제
    # scheduler.add_job(trade_report_download, 'interval', hours=1)

    # 2. 매일 특정 시간에 실행 (예: 아침 8시 30분) - 필요시 조건 변경
    # scheduler.add_job(trade_report_download, 'cron', hour=8, minute=30)
    
    # [참고] 프로그램을 실행하자마자 한 번 작동하게 하려면 아래 주석을 해제하세요
    trade_report_download()

    # print("스케줄러가 시작되었습니다... (종료: Ctrl+C)")
    # try:
    #     scheduler.start()
    # except (KeyboardInterrupt, SystemExit):
    #     print("스케줄러가 종료되었습니다.")
