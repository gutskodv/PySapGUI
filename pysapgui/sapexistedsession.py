import pythoncom
import win32com.client
import pysapgui.sapgui
from pysapgui.transaction import SAPTransaction, TRANSACTION_EXIT


class SAPExistedSession:
    @staticmethod
    def get_session_by_number(session_num=0):
        sap_appl = pysapgui.sapgui.SAPGuiApplication.connect()
        if not sap_appl:
            return

        try:
            sap_connection = sap_appl.Children(session_num)
            if not isinstance(sap_connection, win32com.client.CDispatch):
                raise RuntimeError("Active connection(s) to SAP server not found in {0} application".format(
                    pysapgui.sapgui.SAP_LOGON))
        except pythoncom.com_error as error:
            hr, msg, exc, arg = error.args
            sap_error = "Connection '{2}' not found. COM: {0} ({1})".format(msg, hr, session_num)
            raise RuntimeError(sap_error)

        return SAPExistedSession.get_session_by_connect(sap_connection)

    @staticmethod
    def get_session_by_connect(sap_connection):
        try:
            sap_session = sap_connection.Children(0)
            if not isinstance(sap_session, win32com.client.CDispatch):
                raise RuntimeError("SAP GUI scripting disabled. Please set parameter sapgui/user_scripting to TRUE")
        except pythoncom.com_error as error:
            hr, msg, exc, arg = error.args
            sap_error = "COM: {0} ({1})".format(msg, hr)
            raise RuntimeError("SAP GUI scripting disabled. {0}".format(sap_error))
        else:
            return sap_session

    @staticmethod
    def get_sap_session_with_info(session_num=0):
        sap_session = SAPExistedSession.get_session_by_number(session_num)
        outdict = dict()
        try:
            info = sap_session.Info
            outdict["sid"] = info.SystemName
            outdict["client"] = info.Client
            outdict["user"] = info.User
            outdict["app_server"] = info.applicationServer
        except pythoncom.com_error as error:
            hr, msg, exc, arg = error.args
            sap_error = "COM: {0} ({1})".format(msg, hr)
            msg_list = list()
            msg_list.append("Attached to not authorized session")
            if sap_error:
                msg_list.append(sap_error)
            raise RuntimeError(" ".join(msg_list))

        else:
            if outdict["user"]:
                return sap_session, outdict
            else:
                msg_list = list()
                msg_list.append("Attached to not authorized session. SID: {0}".format(outdict["sid"]))
                raise RuntimeError(" ".join(msg_list))

    @staticmethod
    def get_session_count():
        sap_appl = pysapgui.sapgui.SAPGuiApplication.connect()
        if not sap_appl:
            return

        try:
            sap_connection_count = sap_appl.Children.count
        except pythoncom.com_error as error:
            hr, msg, exc, arg = error.args
            msg = "Application {0} not running.".format(pysapgui.sapgui.SAP_LOGON)
            msg += "COM: {0} ({1})".format(msg, hr)
            raise RuntimeError(msg)
        else:
            return sap_connection_count

    @staticmethod
    def get_multi_sap_session_with_info(unique_sid=True):
        max_saps = 10
        try:
            max_saps = SAPExistedSession.get_session_count()
        except RuntimeError:
            pass
        out_list = list()
        for i in range(0, max_saps):
            try:
                sap_connect, sap_info = SAPExistedSession.get_sap_session_with_info(i)
            except RuntimeError:
                continue
            else:
                if unique_sid:
                    for session, info in out_list:
                        if info["sid"] == sap_info["sid"]:
                            continue
                out_list.append((sap_connect, sap_info))

        return out_list

    @staticmethod
    def close_session(sap_session):
        SAPTransaction.call(sap_session, TRANSACTION_EXIT, check_error=False)
