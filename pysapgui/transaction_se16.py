# -*- coding: utf-8 -*-
from pysapgui.sapguielements import SAPGuiElements
from pysapgui.transaction import SAPTransaction
from pysapgui.sapgui import GUI_CHILD_WINDOW1, GUI_MAIN_WINDOW, GUI_CHILD_USER_AREA1

SE16_TCODE = 'SE16'
TABLENAME_FIELD = "wnd[0]/usr/ctxtDATABROWSE-TABLENAME"
MENU_FIELDS_FOR_SELECTION = ["Поля для выбора", "Fields for Selection", "Felder für Selektion"]
MENU_USER_PARAMETERS = ["ПользовПараметры...", "User Parameters...", "Benutzerparameter..."]
OK_BUTTON = 'wnd[1]/tbar[0]/btn[0]'
FIELD_NAME_SELECTION = "wnd[1]/usr/tabsG_TABSTRIP/tabp0400/ssubTOOLAREA:SAPLWB_CUSTOMIZING:0400/radSEUCUSTOM-FIELDNAME"
ALV_GRID_SELECTION = "wnd[1]/usr/tabsG_TABSTRIP/tabp0400/ssubTOOLAREA:SAPLWB_CUSTOMIZING:0400/radRSEUMOD-TBALV_GRID"
EXCLUDE_VALUES = "wnd[1]/usr/tabsTAB_STRIP/tabpNOSV"
INCLUDE_VALUES = "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA"
INCLUDE_VALUES_TEMPLATE = "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE" \
                          "/{type}RSCSEL_255-SLOW_I[1,{row}]"
INCLUDE_VALUES_BUTTON = "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE" \
                        "/btnRSCSEL_255-SOP_I[0,{0}]"
EXCLUDE_VALUES_TEMPLATE = "wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E" \
                          "/{type}RSCSEL_255-SLOW_E[1,{row}]"
EXCLUDE_VALUES_BUTTON = "wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E" \
                        "/btnRSCSEL_255-SOP_E[0,{0}]"
OK_BUTTON_FILTER = "wnd[2]/tbar[0]/btn[0]"
ROW_NUMBER_BUTTON = 'wnd[0]/tbar[1]/btn[31]'
NUMBER_ENTRIES_FIELD = 'wnd[1]/usr/txtG_DBCOUNT'
MAX_LINES = 'wnd[0]/usr/txtMAX_SEL'
FILEFORMAT = "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]"
FILEPATH = 'wnd[1]/usr/ctxtDY_PATH'
FILENAME = 'wnd[1]/usr/ctxtDY_FILENAME'


class FieldFilters:
    def __init__(self, do_log=False):
        self.filters = list()
        self.do_log = do_log

    def __bool__(self):
        return True if len(self.filters) else False

    def __str__(self):
        out_list = list()
        out_list.append("Filter")

        for item in self.filters:
            out_list.append(str(item))

        return "\n".join(out_list)

    def add_filter(self, field_filter):
        self.filters.append(field_filter)

    def init_filter_by_dict(self, filterlist):
        for item in filterlist:
            if "field_name" not in item:
                continue
            field_name = item['field_name']
            new_filter = FieldFilter(field_name)
            if "exclude_values" in item:
                new_filter.set_exclude_values(item['exclude_values'])
            if "equal_values" in item:
                new_filter.set_equal_values(item['equal_values'])
            self.add_filter(new_filter)

    def enable_columns_to_filter(self, session):
        column_list = [filter1.field_name for filter1 in self.filters]
        if not len(column_list):
            return

        SAPGuiElements.call_menu(session, MENU_FIELDS_FOR_SELECTION)
        max_scroll = int(SAPGuiElements.get_max_scroll_position(session, GUI_CHILD_USER_AREA1))
        pos_scroll = 0
        startpos = 5
        do_cycle = True
        while do_cycle:
            SAPGuiElements.set_scroll_position(session, pos_scroll, GUI_CHILD_USER_AREA1)
            max_i = 0
            for i, element in SAPGuiElements.iter_elements_by_template(session,
                                                                       GUI_CHILD_USER_AREA1,
                                                                       "wnd[1]/usr/lbl[4,{0}]",
                                                                       startpos):
                max_i = i
                if element.text in column_list:
                    SAPGuiElements.set_checkbox(session, "wnd[1]/usr/chk[2,{0}]".format(i))
            if pos_scroll < max_scroll:
                new_pos_scroll = min(pos_scroll + max_i, max_scroll)
                startpos = max_i - (new_pos_scroll - pos_scroll) + 1
                pos_scroll = new_pos_scroll
            else:
                do_cycle = False

        SAPGuiElements.press_keyboard_keys(session, "Enter", GUI_CHILD_WINDOW1)

    def get_filter_by_field_name(self, field_name):
        for field_filter in self.filters:
            if field_filter.field_name == field_name:
                return field_filter

    def set_filter_value(self, session):
        if not len(self.filters):
            return

        columnlist = [filter1.field_name for filter1 in self.filters]

        startpos = 1
        for i, element in SAPGuiElements.iter_elements_by_template(session,
                                                                   GUI_MAIN_WINDOW,
                                                                   "wnd[0]/usr/txt%_I{0}_%_APP_%-TEXT",
                                                                   startpos):
            element_text = element.text
            if element_text in columnlist:
                element_id = element.id
                button_id = element_id.replace("txt", "btn").replace("-TEXT", "-VALU_PUSH")
                SAPGuiElements.press_button(session, button_id)
                field_filter = self.get_filter_by_field_name(element_text)
                field_filter.set_filter(session)


class FieldFilter:
    def __init__(self, field_name):
        self.field_name = field_name

    def __str__(self):
        out_list = list()
        out_list.append("Field: %s" % (self.field_name,))
        if hasattr(self, "exclude_single_values") and self.exclude_single_values:
            out_list.append("Exclude: %s" % (", ".join(self.exclude_single_values)))
        if hasattr(self, "equal_single_values") and self.equal_single_values:
            out_list.append("Equal: %s" % (", ".join(self.equal_single_values)))

        return ", ".join(out_list)

    def set_equal_values(self, values):
        self.equal_single_values = list()
        if type(values) is list:
            if len(values) == 1 and values[0] is None:
                self.equal_single_values.append("")
            else:
                self.equal_single_values.extend(values)
        else:
            self.equal_single_values.append(values)

    def set_exclude_values(self, values):
        self.exclude_single_values = list()
        if type(values) is list:
            if len(values) == 1 and values[0] is None:
                self.exclude_single_values.append("")
            else:
                self.exclude_single_values.extend(values)
        else:
            self.exclude_single_values.append(values)

    def set_range_values(self, values):
        self.equal_range_values = list()
        if type(values) is list:
            self.equal_range_values.extend(values)
        else:
            self.equal_range_values.append(values)

    def set_exclude_range_values(self, values):
        self.exclude_range_values = list()
        if type(values) is list:
            self.exclude_range_values.extend(values)
        else:
            self.exclude_range_values.append(values)

    def set_filter(self, session):
        SAPGuiElements.press_keyboard_keys(session, "Shift+F4")
        if hasattr(self, "exclude_single_values"):
            if len(self.exclude_single_values):
                SAPGuiElements.select_element(session, EXCLUDE_VALUES)
                for i, item in enumerate(self.exclude_single_values):
                    if item == "":
                        SAPGuiElements.press_button(session, EXCLUDE_VALUES_BUTTON.format(i))
                        SAPGuiElements.press_button(session, OK_BUTTON_FILTER)
                    else:
                        SAPGuiElements.try_to_set_text(session, EXCLUDE_VALUES_TEMPLATE.format(type="{type}", row=i),
                                                       item)

        if hasattr(self, "equal_single_values"):
            if len(self.equal_single_values):
                SAPGuiElements.select_element(session, INCLUDE_VALUES)
                for i, item in enumerate(self.equal_single_values):
                    if item == "":
                        SAPGuiElements.press_button(session, INCLUDE_VALUES_BUTTON.format(i))
                        SAPGuiElements.press_button(session, OK_BUTTON_FILTER)
                    else:
                        SAPGuiElements.try_to_set_text(session, INCLUDE_VALUES_TEMPLATE.format(type="{type}", row=i),
                                                       item)

        SAPGuiElements.press_keyboard_keys(session, "F8")


class TCodeSE16:
    first_call = True

    def __init__(self, table_name, sap_session=None, do_log=False):
        self.tcode = SE16_TCODE
        self.table_name = table_name
        self.sap_session = sap_session
        self.do_log = do_log
        self.table_name = table_name
        self.problem = None

    def __get_entries_number(self, session):
        SAPGuiElements.press_keyboard_keys(session, "Ctrl+F7")
        entries_num = SAPGuiElements.get_text(session, NUMBER_ENTRIES_FIELD)
        SAPGuiElements.press_keyboard_keys(session, "Enter", GUI_CHILD_WINDOW1)
        return self.__parse_val(entries_num)

    @staticmethod
    def __parse_val(value):
        return value.replace('.', '').replace(',', '').replace(' ', '')

    def get_row_number_by_filter(self, sap_session, table_filter=None):
        if not sap_session:
            sap_session = self.sap_session
        SAPTransaction.call(sap_session, self.tcode)
        self.__set_table_name(sap_session)

        if TCodeSE16.first_call:
            TCodeSE16.__set_se16_parameters(sap_session)

        if table_filter:
            table_filter.enable_columns_to_filter(sap_session)
            table_filter.set_filter_value(sap_session)
        return self.__get_entries_number(sap_session)

    def save_table_content(self, filepath, filename, sap_session=None):
        if not sap_session:
            sap_session = self.sap_session
        if not sap_session:
            msg = "Bad session"
            raise ValueError(msg)
        SAPTransaction.call(sap_session, self.tcode)
        self.__set_table_name(sap_session)

        if TCodeSE16.first_call:
            TCodeSE16.__set_se16_parameters(sap_session)

        self.__set_max_lines(sap_session)
        SAPGuiElements.press_keyboard_keys(sap_session, "F8")

        gui_msg = SAPGuiElements.get_status_message(sap_session)
        if gui_msg:
            if gui_msg[1] == "429":
                msg = "No entries in the {0} table. GUI Message: {1}".format(self.table_name, gui_msg[2])
                raise ValueError(msg)

        SAPGuiElements.press_keyboard_keys(sap_session, "Ctrl+Shift+F9")

        SAPGuiElements.select_element(sap_session, FILEFORMAT)
        SAPGuiElements.press_keyboard_keys(sap_session, "Enter")

        if filepath:
            SAPGuiElements.set_text(sap_session, FILEPATH, filepath)

        if filename:
            SAPGuiElements.set_text(sap_session, FILENAME, filename)

        SAPGuiElements.press_keyboard_keys(sap_session, "Enter")

    @staticmethod
    def __set_max_lines(sap_session):
        limit = "2000000"
        SAPGuiElements.set_text(sap_session, MAX_LINES, limit)

    def __set_table_name(self, session):
        SAPGuiElements.set_text(session, TABLENAME_FIELD, self.table_name)
        SAPGuiElements.press_keyboard_keys(session, "Enter")
        gui_msg = SAPGuiElements.get_status_message(session)

        if gui_msg:
            if gui_msg[1] == "402":
                msg = "Table '{0}' not found. GUI Message: {1}".format(self.table_name, gui_msg[2])
                raise ValueError(msg)
            elif gui_msg[1] == "419":
                msg = "Not authorized to view the '{0}' table. GUI Message: {1}".format(self.table_name, gui_msg[2])
                raise PermissionError(msg)

    @staticmethod
    def __set_se16_parameters(session):
        SAPGuiElements.call_menu(session, MENU_USER_PARAMETERS)
        SAPGuiElements.select_element(session, FIELD_NAME_SELECTION)
        SAPGuiElements.select_element(session, ALV_GRID_SELECTION)
        SAPGuiElements.press_keyboard_keys(session, "Enter", GUI_CHILD_WINDOW1)
        TCodeSE16.first_call = False
