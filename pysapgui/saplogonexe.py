import subprocess
import os
import time
from pysapgui.sapgui import SAPGuiApplication


class SAPLogonEXE:
    path = None

    @staticmethod
    def __check_sap_logon():
        try:
            SAPGuiApplication.connect()
        except RuntimeError:
            return False
        else:
            return True

    @staticmethod
    def set_path(path):
        SAPLogonEXE.path = path

    @staticmethod
    def start(check=True):
        if not check or not SAPLogonEXE.__check_sap_logon():
            if SAPLogonEXE.path and os.path.exists(SAPLogonEXE.path):
                subprocess.Popen(SAPLogonEXE.path)
                time.sleep(7)
                return SAPLogonEXE.__check_sap_logon()
            else:
                raise ValueError('Path to saplogon.exe not found')
        return False
