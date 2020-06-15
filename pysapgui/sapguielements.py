import pythoncom
import pysapgui.sapgui

STATUS_BAR = "{window}/sbar"
MENU = "{window}/mbar"

GUI_TYPES_WITH_TEXT_PROPERTY = ("GuiTextField",
                                "GuiOkCodeField",
                                "GuiCTextField",
                                "GuiLabel",
                                "GuiStatusbar",
                                "GuiTitlebar",
                                "GuiPasswordField")
GUI_TYPES_WITH_PRESS_METHOD = ("GuiButton", )
GUI_TYPES_WITH_SELECT_METHOD = ("GuiMenu",
                                "GuiTab",
                                "GuiRadioButton")
GUI_TYPES_WITH_SELECTED_PROPERTY = ("GuiCheckBox", )
GUI_TYPES_WITH_SENDVKEY_METHOD = ("GuiMainWindow",
                                  "GuiModalWindow")
GUI_GRID_TYPES = ("GuiShell", )
GUI_GRID_SUBTYPES = ("GridView", )
GUI_SCROLL_TYPES = ("GuiUserArea", )

GUI_VKEYS = {
    "Enter": 0,
    "F6": 6,
    "F8": 8,
    "F12": 12,
    "Shift+F3": 15,
    "Shift+F4": 16,
    "Ctrl+F7": 31,
    "Ctrl+Shift+F9": 45
}
TRY_TEXT_TYPES = ['ctxt', 'txt']


class SAPGuiElements:
    @staticmethod
    def set_text(sap_session, text_element_id, text):
        text_element = SAPGuiElements.__get_element(sap_session, text_element_id)

        if text_element.type in GUI_TYPES_WITH_TEXT_PROPERTY:
            text_element.text = str(text)
        else:
            msg = "Text not be set for GUI element '{0}' (wrong GUI element type '{1}')".format(text_element_id,
                                                                                                text_element.type)
            raise TypeError(msg)

    @staticmethod
    def get_text(sap_session, text_element_id):
        text_element = SAPGuiElements.__get_element(sap_session, text_element_id)

        if text_element.type in GUI_TYPES_WITH_TEXT_PROPERTY:
            return text_element.text
        else:
            msg = "GUI element '{0}' has no 'text' property(wrong GUI element type '{1}')".format(text_element_id,
                                                                                                  text_element.type)
            raise TypeError(msg)

    @staticmethod
    def press_button(sap_session, button_element_id):
        button_element = SAPGuiElements.__get_element(sap_session, button_element_id)

        if button_element.type in GUI_TYPES_WITH_PRESS_METHOD:
            button_element.press()
        else:
            msg = "GUI element '{0}' has no press() method (wrong GUI element type '{1}')".format(button_element_id,
                                                                                                  button_element.type)
            raise TypeError(msg)

    @staticmethod
    def get_element(sap_session, element_id):
        return SAPGuiElements.__get_element(sap_session, element_id)

    @staticmethod
    def __get_element(sap_session, element_id):
        try:
            element = sap_session.findById(element_id)

        except pythoncom.com_error:
            msg = "GUI element with id '{0}' not found".format(element_id)
            raise AttributeError(msg)

        except AttributeError:
            msg = "Wrong sap session object received"
            raise AttributeError(msg)
        else:
            return element

    @staticmethod
    def set_checkbox(sap_session, element_id):
        selected_element = SAPGuiElements.__get_element(sap_session, element_id)

        if selected_element.type in GUI_TYPES_WITH_SELECTED_PROPERTY:
            selected_element.selected = True
        else:
            msg = "GUI element '{0}' has no select() method (wrong GIU element type '{1}')".format(
                element_id,
                selected_element.type)
            raise TypeError(msg)

    @staticmethod
    def select_element(sap_session, element_id):
        select_element = SAPGuiElements.__get_element(sap_session, element_id)

        if select_element.type in GUI_TYPES_WITH_SELECT_METHOD:
            select_element.select()
        else:
            msg = "GUI element '{0}' has no select() method (wrong GUI element type '{1}')".format(element_id,
                                                                                                   select_element.type)
            raise TypeError(msg)

    @staticmethod
    def press_keyboard_keys(sap_session, keys, window_id=pysapgui.sapgui.GUI_MAIN_WINDOW):
        if keys in GUI_VKEYS.keys():
            SAPGuiElements.__send_vkey(sap_session, GUI_VKEYS[keys], window_id)
        else:
            msg = "Wrong vkey '{0}' received. Supported only {1}".format(keys, ", ".join(GUI_VKEYS.keys()))
            raise ValueError(msg)

    @staticmethod
    def __send_vkey(sap_session, vkey, window_id=pysapgui.sapgui.GUI_MAIN_WINDOW):
        window_element = SAPGuiElements.__get_element(sap_session, window_id)

        if window_element.type in GUI_TYPES_WITH_SENDVKEY_METHOD:
            if vkey in GUI_VKEYS.values():
                window_element.sendVKey(vkey)
        else:
            msg = "GUI Element '{0}' has no sendvkey() method (wrong GUI element type '{1}')".format(
                window_id,
                window_element.type)
            raise TypeError(msg)

    @staticmethod
    def get_status_message(sap_session, window_id=pysapgui.sapgui.GUI_MAIN_WINDOW):
        statusbar = SAPGuiElements.__get_element(sap_session, STATUS_BAR.format(window=window_id))

        if statusbar.text:
            return statusbar.MessageType, statusbar.MessageNumber, statusbar.text

    @staticmethod
    def try_to_set_text(sap_session, element_template, text):
        for element_type in TRY_TEXT_TYPES:
            try:
                SAPGuiElements.set_text(sap_session, element_template.format(type=element_type), text)
            except AttributeError:
                pass
            else:
                return

        msg = "GUI elements with mask '{0}' not found".format(element_template)
        raise AttributeError(msg)

    @staticmethod
    def get_grid_rows_number(sap_session, grid_id):
        grid_element = SAPGuiElements.__get_element(sap_session, grid_id)

        if grid_element.type in GUI_GRID_TYPES and \
                grid_element.subtype in GUI_GRID_SUBTYPES:
            return grid_element.RowCount
        else:
            msg = "GIU element '{0}' has no rowcount() method (wrong GUI element type '{1}')".format(
                grid_id,
                grid_element.type)
            raise TypeError(msg)

    @staticmethod
    def get_value_from_grid(sap_session, grid_id, row, column):
        grid_element = SAPGuiElements.__get_element(sap_session, grid_id)

        if grid_element.type in GUI_GRID_TYPES and \
                grid_element.subtype in GUI_GRID_SUBTYPES:
            return grid_element.GetCellValue(row, grid_element.ColumnOrder(column))
        else:
            msg = "GUI element '{0}' has no rowcount() method (wrong GUI element type '{1}')".format(grid_id,
                                                                                                     grid_element.type)
            raise TypeError(msg)

    @staticmethod
    def get_scroll_position(sap_session, area_id=pysapgui.sapgui.GUI_MAIN_USER_AREA):
        scroll_element = SAPGuiElements.__get_element(sap_session, area_id)
        if scroll_element.type in GUI_SCROLL_TYPES:
            return scroll_element.verticalScrollbar.position
        else:
            msg = "GUI element '{0}' has no verticalScrollbar object (wrong GUI element type '{1}')".format(
                area_id,
                scroll_element.type)
            raise TypeError(msg)

    @staticmethod
    def get_max_scroll_position(sap_session, area_id=pysapgui.sapgui.GUI_MAIN_USER_AREA):
        scroll_element = SAPGuiElements.__get_element(sap_session, area_id)
        if scroll_element.type in GUI_SCROLL_TYPES:
            return scroll_element.verticalScrollbar.Maximum
        else:
            msg = "GUI element '{0}' has no verticalScrollbar object (wrong GUI element type '{1}')".format(
                area_id,
                scroll_element.type)
            raise TypeError(msg)

    @staticmethod
    def set_scroll_position(sap_session, position, area_id=pysapgui.sapgui.GUI_MAIN_USER_AREA):
        scroll_element = SAPGuiElements.__get_element(sap_session, area_id)
        if scroll_element.type in GUI_SCROLL_TYPES:
            if not type(position) is int:
                msg = "Wrong position '{0}' to scroll received".format(position)
                raise TypeError(msg)
            scroll_element.verticalScrollbar.position = position
        else:
            msg = "GUI element '{0}' has no verticalScrollbar object (wrong GUI element type '{1}')".format(
                area_id,
                scroll_element.type)
            raise TypeError(msg)

    @staticmethod
    def __find_menu_element(menu, menu_names):
        if menu.text in menu_names:
            return menu.id

        for child_menu in menu.Children:
            found_element_id = SAPGuiElements.__find_menu_element(child_menu, menu_names)
            if found_element_id:
                return found_element_id

    @staticmethod
    def call_menu(sap_session, menu_names, window_id=pysapgui.sapgui.GUI_MAIN_WINDOW):
        if not len(menu_names):
            msg = "Empty Menu names list received"
            raise ValueError(msg)
        main_menu = SAPGuiElements.__get_element(sap_session, MENU.format(window=window_id))
        menu_element_id = SAPGuiElements.__find_menu_element(main_menu, menu_names)

        if not menu_element_id:
            msg = "Menu GUI element not found by text {text}.".format(text=", ".join(menu_names))
            raise ValueError(msg)
        SAPGuiElements.select_element(sap_session, menu_element_id)

    @staticmethod
    def iter_elements_by_template(sap_session, root_element_id, id_template, start_index, max_index=50):
        root_area = SAPGuiElements.__get_element(sap_session, root_element_id)

        for index in range(start_index, max_index):
            try:
                element = SAPGuiElements.__get_element(root_area, id_template.format(index))
                yield index, element
            except AttributeError:
                break

    @staticmethod
    def get_grid(sap_session, grid_id):
        grid_element = SAPGuiElements.__get_element(sap_session, grid_id)

        if grid_element.type in GUI_GRID_TYPES:
            return grid_element
        else:
            msg = "GUI element {0} not Grid. Wrong GUI element type '{1}'".format(
                grid_id,
                grid_element.type)
            raise TypeError(msg)
