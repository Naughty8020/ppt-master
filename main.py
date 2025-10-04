from PySide6.QtWidgets import QApplication
from controller import PPTController, PPTView  # もし controller.py 内に PPTView を書いた場合

app = QApplication([])
view = PPTView()
controller = PPTController(view)
view.show()
app.exec()
