from PyQt5.QtWidgets import QMainWindow, QApplication
from main_window import *
from addmask_main import *
import sys
import traceback

class Main_GUI(Ui_MainWindow):

    def init(self):
        self.pushButton.clicked.connect(self.openDir)
        self.pushButton_2.clicked.connect(self.openDir)
        self.pushButton_3.clicked.connect(self.addMask)
        self.pushButton_4.clicked.connect(self.close)
        self.setWindowTitle("WaterMask Add")
        self.input_file_path = None
        self.output_file_path = None

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
        TEMP_DIR = os.path.join(os.getcwd(), 'temp')
        if os.path.exists(TEMP_DIR):
            rmtree(TEMP_DIR)
        if self.input_file_path and self.output_file_path:
            self.addwaterMask = AddWaterMask(self.input_file_path,self.output_file_path)
            try:
                self.addwaterMask.run()
            except Exception as e:
                print("error: ")
                traceback.print_exc()

        else:
            QMessageBox.critical(None, "错误", "未设定输入/输出文件所在文件夹位置")



if __name__ == '__main__':
    app = QApplication(sys.argv)
    MainWindow = QMainWindow()
    main_ui = Main_GUI()
    main_ui.setupUi(MainWindow)
    main_ui.init()
    MainWindow.show()
    sys.exit(app.exec_())
    os.system('pause')
