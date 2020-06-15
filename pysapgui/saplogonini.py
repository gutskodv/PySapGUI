import configparser
import os

DESCRIPTION_SECTION = "Description"
SID_SECTION = "MSSysName"


class SAPLogonINI:
    ini_files = list()

    @staticmethod
    def get_connect_name_by_sid(sid, first=True):
        return SAPLogonINI.__get_connection_description_by_sid(sid, first)

    @staticmethod
    def set_ini_files(*args):
        SAPLogonINI.ini_files = list()
        for file in args:
            if file and os.path.exists(file):
                SAPLogonINI.ini_files.append(file)

    @staticmethod
    def __get_item_by_sid(config, sid):
        items = dict(config[SID_SECTION].items())
        for item_key, item_value in items.items():
            if item_value == sid:
                return item_key

    @staticmethod
    def __get_descr_by_item_id(config, item_id):
        return config[DESCRIPTION_SECTION][item_id]

    @staticmethod
    def __get_connection_description_by_sid(sid, first=True):
        out_list = list()

        if not len(SAPLogonINI.ini_files):
            raise RuntimeError("Saplogon.ini files not found")

        for file in SAPLogonINI.ini_files:
            config = configparser.ConfigParser()
            config.read(file)
            item = SAPLogonINI.__get_item_by_sid(config, sid)
            if item:
                conn_name = SAPLogonINI.__get_descr_by_item_id(config, item)
                if first:
                    return conn_name
                else:
                    out_list.append(conn_name)

        if not len(out_list):
            raise ValueError("System '{0}' not found in saplogon.ini files".format(sid))
        else:
            return out_list
