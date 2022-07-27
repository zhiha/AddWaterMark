from PyQt5.QtWidgets import QMainWindow, QApplication
from main_window import *
from addmask_main import *
import sys


class Main_GUI(Ui_MainWindow):

    def init(self):
        self.pushButton.clicked.connect(self.openDir)
        self.pushButton_2.clicked.connect(self.openDir)
        self.pushButton_3.clicked.connect(self.addMask)
        self.pushButton_4.clicked.connect(self.close)
        self.setWindowTitle("WaterMask Add")

    def openDir(self):
        sender = self.sender()
        text = sender.text()
        fname = QFileDialog.getExistingDirectory(self,'open dir', '.')
        if text == self.pushButton.text():
            self.textEdit.setText(fname)
            self.input_file_path = fname
        if text == self.pushButton_2.text():
            self.textEdit_2.setText(fname)
            self.output_file_path = fname

    def close(self):
        self.close()

    def addMask(self):
        if self.input_file_path and self.output_file_path:
            self.addwaterMask = AddWaterMask(self.input_file_path,self.output_file_path)
            self.addwaterMask.run()



if __name__ == '__main__':
    app = QApplication(sys.argv)
    MainWindow = QMainWindow()
    main_ui = Main_GUI()
    main_ui.setupUi(MainWindow)
    main_ui.init()
    MainWindow.show()
    sys.exit(app.exec_())
    os.system('pause')
