from PyQt5 import QtWidgets
from ui.main_window import DSOApp
import sys

def run_app():
    app = QtWidgets.QApplication(sys.argv)
    window = DSOApp()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    run_app()
