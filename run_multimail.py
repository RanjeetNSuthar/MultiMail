import sys
from PyQt5.QtWidgets import QApplication

from core_multimail import Mailer

app = QApplication(sys.argv)
window = Mailer()
window.show()
sys.exit(app.exec_())


