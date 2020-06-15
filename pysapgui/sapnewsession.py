import pythoncom
import win32com.client
import pysapgui.sapgui
from pysapgui.sapguielements import SAPGuiElements
from pysapgui.saplogonpwd import SAPLogonPwd
from pysapgui.saplogonini import SAPLogonINI
from pysapgui.sapexistedsession import SAPExistedSession

USERNAME_FIELD = "wnd[0]/usr/txtRSYST-BNAME"
PASSWORD_FIELD = "wnd[0]/usr/pwdRSYST-BCODE"
CLIENT_FIELD = "wnd[0]/usr/txtRSYST-MANDT"
LANGUAGE_FIELD = "wnd[0]/usr/txtRSYST-LANGU"

MULTI_USER_DISCONNECT = "wnd[1]/usr/radMULTI_LOGON_OPT1"


class SAPNewSession:
    @staticmethod
    def get_new_session_by_name(name):
        sap_appl = pysapgui.sapgui.SAPGuiApplication.connect()
        if not sap_appl:
            return

        try:
            sap_appl.OpenConnection(name, True)
        except pythoncom.com_error as error:
            hr, msg, exc, arg = error.args
            sap_error = "COM: {0} ({1})".format(msg, hr)
            msg = "Connection '{0}' not found. {1}".format(name, sap_error)
            raise RuntimeError(msg)

        try:
            children_count = SAPExistedSession.get_session_count()
            sap_connection = sap_appl.Children(children_count - 1)
            if not isinstance(sap_connection, win32com.client.CDispatch):
                raise RuntimeError("Active connection(s) to SAP server not found in {0} application".format(
                    pysapgui.sapgui.SAP_LOGON))
        except pythoncom.com_error as error:
            hr, msg, exc, arg = error.args
            sap_error = "COM: {0} ({1})".format(msg, hr)
            raise RuntimeError(sap_error)
        else:
            return SAPExistedSession.get_session_by_connect(sap_connection)

    @staticmethod
    def is_already_logon_window(sap_session):
        try:
            select = SAPGuiElements.get_element(sap_session, MULTI_USER_DISCONNECT)
        except AttributeError:
            pass
        else:
            SAPGuiElements.select_element(sap_session, select.id)
            SAPGuiElements.press_keyboard_keys(sap_session, "Enter", pysapgui.sapgui.GUI_CHILD_WINDOW1)

    @staticmethod
    def close_new_windows_while_connect(sap_session):
        max_windows = 10
        try:
            for i in range(0, max_windows):
                SAPGuiElements.get_element(sap_session, pysapgui.sapgui.GUI_CHILD_WINDOW1)
                SAPGuiElements.press_keyboard_keys(sap_session, "Enter", pysapgui.sapgui.GUI_CHILD_WINDOW1)
        except AttributeError:
            pass

    @staticmethod
    def create_new_session(sid, user, pwd, client=None, langu=None, close_conn=False, change_pwd=True):
        conn_descr = SAPLogonINI.get_connect_name_by_sid(sid)
        if not conn_descr:
            raise ValueError("Connection to '{0}' not found".format(sid))

        sap_session = SAPNewSession.get_new_session_by_name(conn_descr)
        if sap_session:
            if user:
                SAPGuiElements.set_text(sap_session, USERNAME_FIELD, user)
            if pwd:
                SAPGuiElements.set_text(sap_session, PASSWORD_FIELD, pwd)
            if client:
                SAPGuiElements.set_text(sap_session, CLIENT_FIELD, client)
            if langu:
                SAPGuiElements.set_text(sap_session, LANGUAGE_FIELD, langu)

            SAPGuiElements.press_keyboard_keys(sap_session, "Enter")

            msg = SAPGuiElements.get_status_message(sap_session)
            if msg and msg[0] == "E":
                SAPGuiElements.press_keyboard_keys(sap_session, "Shift+F3")
                raise RuntimeError("Could not log to the SAP server. {0}".format(msg[2]))

            SAPNewSession.is_already_logon_window(sap_session)
            return_pwd = SAPLogonPwd.is_change_password_window_while_connect(sap_session, change_pwd=change_pwd)
            SAPNewSession.close_new_windows_while_connect(sap_session)
            if close_conn:
                SAPExistedSession.close_session(sap_session)

            return return_pwd
