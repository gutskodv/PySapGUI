import pysapgui.sapgui
from pysapgui.sapguielements import SAPGuiElements

TRANSACTION_TEXT_FIELD = "{window}/tbar[0]/okcd"
TRANSACTION_PREFIX = "/N"
TRANSACTION_EXIT = "EX"


class SAPTransaction:
    @staticmethod
    def __call_transaction(sap_session, transaction, window_id=pysapgui.sapgui.GUI_MAIN_WINDOW, check_error=True):
        tcode = TRANSACTION_PREFIX + transaction
        SAPGuiElements.set_text(sap_session, TRANSACTION_TEXT_FIELD.format(window=window_id), tcode)
        SAPGuiElements.press_keyboard_keys(sap_session, "Enter")

        if not check_error:
            return
        gui_msg = SAPGuiElements.get_status_message(sap_session)

        if gui_msg:
            if gui_msg[1] == "343":
                msg = "Wrong transaction name '{0}'. GUI Message: {1}".format(transaction, gui_msg[2])
                raise ValueError(msg)

            elif gui_msg[1] == "077":
                msg = "Not authorized to execute '{0}' transaction. GUI Message: {1}".format(transaction, gui_msg[2])
                raise PermissionError(msg)

            elif gui_msg[1] == "410":
                msg = "Blocked action in '{0}' transaction. GUI Message: {1}".format(transaction, gui_msg[2])
                raise PermissionError(msg)

    @staticmethod
    def call(sap_session, tcode, check_error=True):
        SAPTransaction.__call_transaction(sap_session, tcode, check_error=check_error)

    @staticmethod
    def exit_transaction(sap_session):
        SAPTransaction.__call_transaction(sap_session, "", check_error=False)
