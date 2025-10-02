import sys
from PySide6.QtWidgets import QApplication
from view import PPTView
from controller import PPTController

if __name__ == "__main__":
    app = QApplication(sys.argv)
    view = PPTView()
    controller = PPTController(view)
    view.show()
    sys.exit(app.exec())
