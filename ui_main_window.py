from checklist_loader import start_check
from PyQt5 import QtWidgets, QtGui, QtCore

def handle_start_check(self):
    if self.df_checklist is None or self.extracted_texts is None:
        QtWidgets.QMessageBox.warning(self, "Missing Data", "Please load both checklist and PDF before checking.")
        return

    result_df = start_check(self.df_checklist, self.extracted_texts)

    if result_df.empty:
        QtWidgets.QMessageBox.information(self, "No Result", "No matching terms found.")
        return

    self.result_table.setRowCount(len(result_df))
    self.result_table.setColumnCount(len(result_df.columns))
    self.result_table.setHorizontalHeaderLabels(result_df.columns.tolist())

    for row_idx, (_, row) in enumerate(result_df.iterrows()):
        for col_idx, value in enumerate(row):
            item = QtWidgets.QTableWidgetItem(str(value))
            if str(row["Result"]).startswith("❌"):
                item.setForeground(QtGui.QBrush(QtGui.QColor("red")))
            self.result_table.setItem(row_idx, col_idx, item)
