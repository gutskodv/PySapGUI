import string
import random
from pysapgui.sapguielements import SAPGuiElements
from pysapgui.transaction import SAPTransaction
import pysapgui.sapgui

MIN_PASSWORD_LNG = "min_password_lng"
MIN_PASSWORD_DIGITS = "min_password_digits"
MIN_PASSWORD_LETTERS = "min_password_letters"
MIN_PASSWORD_LOWERCASE = "min_password_lowercase"
MIN_PASSWORD_UPPERCASE = "min_password_uppercase"
MIN_PASSWORD_SPECIALS = "min_password_specials"

CHNAGE_PWD_BUTTON_SU3 = "wnd[0]/tbar[1]/btn[6]"
CHANGE_PWD_FIELD1 = "wnd[1]/usr/pwdRSYST-NCODE"
CHANGE_PWD_FIELD2 = "wnd[1]/usr/pwdRSYST-NCOD2"

OLD_PWD_FIELD = "wnd[1]/usr/subSUBSCREEN:SAPMSYST:0043/pwdRSYST-BCODE"
NEW_PWD_FIELD1 = "wnd[1]/usr/subSUBSCREEN:SAPMSYST:0043/pwdRSYST-NCODE"
NEW_PWD_FIELD2 = "wnd[1]/usr/subSUBSCREEN:SAPMSYST:0043/pwdRSYST-NCOD2"


CHANGE_PASSWORD_TRANSACTION = "SU3"


class SAPLogonPwd:
    password_policy = None

    @staticmethod
    def set_password_policy(policy):
        password_policy = {
            MIN_PASSWORD_LNG: 6,
            MIN_PASSWORD_DIGITS: 0,
            MIN_PASSWORD_LETTERS: 0,
            MIN_PASSWORD_LOWERCASE: 0,
            MIN_PASSWORD_UPPERCASE: 0,
            MIN_PASSWORD_SPECIALS: 0
        }

        for key, value in policy.items():
            try:
                password_policy[key] = int(value)
            except ValueError:
                continue
        SAPLogonPwd.password_policy = password_policy

    @staticmethod
    def gen_password():
        if not SAPLogonPwd.password_policy:
            raise AttributeError('Could not generate new password. Password policy not set')
        out_pwd = ""
        out_pwd += SAPLogonPwd.__gen_letters(SAPLogonPwd.password_policy[MIN_PASSWORD_LETTERS],
                                             SAPLogonPwd.password_policy[MIN_PASSWORD_UPPERCASE],
                                             SAPLogonPwd.password_policy[MIN_PASSWORD_LOWERCASE])
        out_pwd += SAPLogonPwd.__gen_digits(SAPLogonPwd.password_policy[MIN_PASSWORD_DIGITS])
        out_pwd += SAPLogonPwd.__gen_specials(SAPLogonPwd.password_policy[MIN_PASSWORD_SPECIALS])

        if len(out_pwd) < SAPLogonPwd.password_policy[MIN_PASSWORD_LNG]:
            out_pwd += SAPLogonPwd.__gen_lower_letters(SAPLogonPwd.password_policy[MIN_PASSWORD_LNG] - len(out_pwd))

        return SAPLogonPwd.__shuffle(out_pwd)

    @staticmethod
    def __gen_lower_letters(letters_num):
        letters = string.ascii_lowercase
        return ''.join(random.choice(letters) for i in range(letters_num))

    @staticmethod
    def __gen_upper_letters(letters_num):
        letters = string.ascii_uppercase
        return ''.join(random.choice(letters) for i in range(letters_num))

    @staticmethod
    def __gen_digits(digits):
        letters = string.digits
        return ''.join(random.choice(letters) for i in range(digits))

    @staticmethod
    def __gen_specials(specials):
        letters = "!@#$%^&*()_+-={}[]<>.,"
        return ''.join(random.choice(letters) for i in range(specials))

    @staticmethod
    def __gen_letters(letters, upper_letters, lower_letters):
        if letters > upper_letters + lower_letters:
            lower_letters = letters - upper_letters
        return SAPLogonPwd.__gen_upper_letters(upper_letters) + SAPLogonPwd.__gen_lower_letters(lower_letters)

    @staticmethod
    def __shuffle(in_str):
        temp_list = list(in_str)
        random.shuffle(temp_list)
        return ''.join(temp_list)

    @staticmethod
    def is_change_password_window_while_connect(sap_session, change_pwd=True):
        try:
            SAPGuiElements.get_element(sap_session, CHANGE_PWD_FIELD1)
            SAPGuiElements.get_element(sap_session, CHANGE_PWD_FIELD2)
        except AttributeError:
            # All is Ok. No need to change password
            pass
        else:
            if change_pwd:
                SAPLogonPwd.__change_password_login(sap_session)
            else:
                SAPGuiElements.press_keyboard_keys(sap_session, "F12", pysapgui.sapgui.GUI_CHILD_WINDOW1)
                raise RuntimeError('Change password to the SAP user')

    @staticmethod
    def __change_password_login(sap_session):
        new_pwd = SAPLogonPwd.gen_password()

        try:
            SAPGuiElements.set_text(sap_session, CHANGE_PWD_FIELD1, new_pwd)
            SAPGuiElements.set_text(sap_session,  CHANGE_PWD_FIELD2, new_pwd)
            SAPGuiElements.press_keyboard_keys(sap_session, "Enter", pysapgui.sapgui.GUI_CHILD_WINDOW1)
        except AttributeError as error:
            raise AttributeError("Could not find GUI elements to change password. {0}".format(str(error)))

        try:
            SAPGuiElements.get_element(sap_session, pysapgui.sapgui.GUI_CHILD_WINDOW2)
        except AttributeError:
            # All is Ok. Password changed
            pass
        else:
            raise AttributeError("Couldn't set new password. Reconfigure password policy")

        return new_pwd

    @staticmethod
    def change_password_su3(sap_session, old_pwd):
        SAPTransaction.call(sap_session, CHANGE_PASSWORD_TRANSACTION)
        new_pwd = SAPLogonPwd.gen_password()

        SAPGuiElements.press_button(sap_session, CHNAGE_PWD_BUTTON_SU3)
        msg = SAPGuiElements.get_status_message(sap_session)
        if msg:
            if msg[1] in ('190', '180'):
                raise RuntimeError("Password could not be changed. {0}".format(msg[2]))

        try:
            SAPGuiElements.set_text(sap_session, OLD_PWD_FIELD, old_pwd)
            SAPGuiElements.set_text(sap_session, NEW_PWD_FIELD1, new_pwd)
            SAPGuiElements.set_text(sap_session, NEW_PWD_FIELD2, new_pwd)
            SAPGuiElements.press_keyboard_keys(sap_session, "Enter", pysapgui.sapgui.GUI_CHILD_WINDOW1)
        except AttributeError as error:
            raise AttributeError("Could not find GUI elements to change password. {0}".format(str(error)))

        try:
            SAPGuiElements.get_element(sap_session, pysapgui.sapgui.GUI_CHILD_WINDOW2)
            SAPGuiElements.press_keyboard_keys(sap_session, "Enter", pysapgui.sapgui.GUI_CHILD_WINDOW2)
            SAPGuiElements.press_keyboard_keys(sap_session, "F12", pysapgui.sapgui.GUI_CHILD_WINDOW1)
        except AttributeError:
            # All is Ok. Password changed successfully
            return new_pwd
        else:
            raise AttributeError("Couldn't set new password. Reconfigure password policy")
