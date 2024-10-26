import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QTabWidget
from PyQt5.QtGui import QIcon
from Test import WellDataApp
from casing import DbCalculator
from Datainput import DataInputTab

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Well Data Analyzer Pro")
        self.setGeometry(100, 100, 1200, 800)

        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        self.setupTabs()
        self.connectTabs()
        self.setStyle()

    def setupTabs(self):
        self.equations_tab = WellDataApp()
        self.data_input_tab = DataInputTab()
        self.casing_tab = DbCalculator()

        self.tabs.addTab(self.equations_tab, QIcon("icons/equations.png"), "Equations")
        self.tabs.addTab(self.data_input_tab, QIcon("icons/datainput.png"), "Data Input")
        self.tabs.addTab(self.casing_tab, QIcon("icons/casing.png"), "Casing")

    def connectTabs(self):
        self.equations_tab.casing_tab = self.casing_tab
        self.equations_tab.data_input_tab = self.data_input_tab
        self.casing_tab.data_input_tab = self.data_input_tab

    def setStyle(self):
        self.setStyleSheet(WellDataApp.STYLE_SHEET)

def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
