from PyQt5 import QtWidgets, QtCore
from ui.main_window import DSOApp
from PyQt5.QtGui import QIcon, QPixmap
from PyQt5.QtWidgets import QSplashScreen
import os, sys
import logging

APP_ICON_PATH = os.path.join("assets", "app", "dso_icon.ico")
SPLASH_IMG_PATH = os.path.join("assets", "app", "splash_dso.png")

# [ADD] ให้รูปไอคอนไม่แตกบนจอ HiDPI
QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps, True)

if sys.platform.startswith("win"):
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("com.company.dso")  # เปลี่ยนเป็น ID ของคุณ
    except Exception:
        pass

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
    datefmt="%H:%M:%S",
)

def run_app():
    app = QtWidgets.QApplication(sys.argv)

    # Add ตั้งไอคอนให้ทั้งแอป (เผื่อ Title bar/Taskbar)
    try:
        app.setWindowIcon(QIcon(APP_ICON_PATH))
    except Exception:
        pass

    # Add Splash (ถ้ามีไฟล์)
    splash = None
    if os.path.exists(SPLASH_IMG_PATH):
        pix = QPixmap(SPLASH_IMG_PATH)
        if not pix.isNull():
            splash = QSplashScreen(pix)
            splash.show()
            app.processEvents()

    window = DSOApp()
    window.show()

    # Add ปิด splash เมื่อพร้อม
    if splash is not None:
        splash.finish(window)

    sys.exit(app.exec_())

if __name__ == "__main__":
    run_app()