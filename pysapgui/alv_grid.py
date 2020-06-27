from .sapguielements import SAPGuiElements


class SAPAlvGrid:
    def __init__(self, sap_session, grid_id):
        self.grid_id = grid_id
        self.sap_session = sap_session
        self.grid = self.__get_grid()
        self.columns = list()
        self.__set_columns()

    def __get_grid(self):
        return SAPGuiElements.get_grid(sap_session=self.sap_session, grid_id=self.grid_id)

    def __column_count(self):
        return self.grid.ColumnCount

    def __row_count(self):
        return self.grid.RowCount

    def get_row_count(self):
        return self.__row_count()

    def __get_column_by_index(self, index):
        return self.grid.ColumnOrder(index)

    def __scrool_grid(self, row):
        return self.grid.SetCurrentCell(row, self.columns[0])

    def __set_columns(self):
        if not self.grid:
            return
        for i in range(0, self.__column_count()):
            self.columns.append(self.__get_column_by_index(i))

    def __get_cell(self, row, column):
        return self.grid.GetCellValue(row, column)

    def get_cell(self, row, column):
        return self.__get_cell(row, column)

    def __iter__(self):
        for row in range(0, self.__row_count()):
            oneline = list()
            for column in self.columns:
                oneline.append(self.__get_cell(row, column))
            if row % 16 == 0:
                self.__scrool_grid(row)
            yield oneline
