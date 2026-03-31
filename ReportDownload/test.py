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
    output_dir = Path("Z:/02.펀드/003.매매보고서 대사/PRELUDE_RECAP")
    output_dir.mkdir(parents=True, exist_ok=True)
    outlook = win32com.client.Dispatch("outlook.application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6).Folders("PRELUDE_RECAP")
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

    print("PRELUDE_RECAP report download complete")  

# def trade_report_download():
if __name__ == "__main__":
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