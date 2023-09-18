from .sapguielements import SAPGuiElements
from .alv_grid import SAPAlvGrid

MENU_STATUS = ['Статус...', 'Status...']

COMPONENTS_BUTTON = "wnd[1]/usr/btnPRELINFO"
PAGE2 = "wnd[2]/usr/tabsVERSDETAILS/tabpPROD_VERS"
PAGE1 = "wnd[2]/usr/tabsVERSDETAILS/tabpCOMP_VERS"

GRID = "wnd[2]/usr/tabsVERSDETAILS/tabpCOMP_VERS/ssubDETAIL_SUBSCREEN:SAPLOCS_UI_CONTROLS:0301/cntlSCV_CU_CONTROL" \
       "/shellcont/shell"

OK_BUTTON2 = "wnd[2]/tbar[0]/btn[0]"
OK_BUTTON1 = "wnd[1]/tbar[0]/btn[0]"


class SAPSoftwareComponents:

    @staticmethod
    def load_software_components(sap_session):
        SAPGuiElements.call_menu(sap_session, MENU_STATUS)
        SAPGuiElements.press_button(sap_session, COMPONENTS_BUTTON)
        SAPGuiElements.select_element(sap_session, PAGE1)
        data = SAPSoftwareComponents.__read_grid(sap_session)
        SAPGuiElements.press_button(sap_session, OK_BUTTON2)
        SAPGuiElements.press_button(sap_session, OK_BUTTON1)
        return data

    @staticmethod
    def __read_grid(sap_session):
        out_list = list()
        grid = SAPAlvGrid(sap_session, GRID)
        for row in grid:
            out_list.append(row)
        return out_list

    @staticmethod
    def load_sap_core_info(sap_session):
        outdict = dict()
        SAPGuiElements.call_menu(sap_session, MENU_STATUS)
        SAPGuiElements.press_keyboard_keys(sap_session, "Shift+F5")
        outdict["krnl_patch_level"] = SAPGuiElements.get_text(sap_session, "wnd[2]/usr/txtKINFOSTRUC-KERNEL_PATCH_LEVEL")
        outdict["krnl_version"] = SAPGuiElements.get_text(sap_session, "wnd[2]/usr/txtKINFOSTRUC-KERNEL_RELEASE")
        SAPGuiElements.press_button(sap_session, OK_BUTTON2)
        SAPGuiElements.press_button(sap_session, OK_BUTTON1)
        return outdict



