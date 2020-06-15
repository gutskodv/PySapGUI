import pythoncom
import win32com.client

SAP_GUI_APPLICATION = "SAPGUI"
SAP_LOGON = "SAP Logon"

GUI_MAIN_WINDOW = "wnd[0]"
GUI_CHILD_WINDOW1 = "wnd[1]"
GUI_CHILD_WINDOW2 = "wnd[2]"

GUI_MAIN_USER_AREA = "wnd[0]/usr"
GUI_CHILD_USER_AREA1 = "wnd[1]/usr"
GUI_CHILD_USER_AREA2 = "wnd[2]/usr"


class SAPGuiApplication:
    @staticmethod
    def connect():
        appl = SAPGuiApplication.__get_object()
        if appl:
            return SAPGuiApplication.__get_script_engine(appl)

    @staticmethod
    def __get_script_engine(sap_gui):
        try:
            sap_application = sap_gui.GetScriptingEngine
            if isinstance(sap_application, win32com.client.CDispatch):
                return sap_application
        except pythoncom.com_error as error:
            hr, msg, exc, arg = error.args
            msg = "Application {0} not running.".format(SAP_LOGON)
            msg += "COM: {0} ({1})".format(msg, hr)
            raise RuntimeError(msg)

    @staticmethod
    def __get_object():
        pythoncom.CoInitialize()
        try:
            sap_gui = win32com.client.GetObject(SAP_GUI_APPLICATION)
            if isinstance(sap_gui, win32com.client.CDispatch):
                return sap_gui
        except pythoncom.com_error as error:
            hr, msg, exc, arg = error.args
            msg = "Application {0} not running.".format(SAP_LOGON)
            msg += "COM: {0} ({1})".format(msg, hr)
            raise RuntimeError(msg)
