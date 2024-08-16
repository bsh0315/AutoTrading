from kiwoom.kiwoom import *

import sys
from PyQt5.QtWidgets import *
class UI_class():
    def __init__(self):
        print("UI 클래스 실행")
        
        self.app = QApplication(sys.argv)
        Kiwoom()
        
        self.app.exec_()
        