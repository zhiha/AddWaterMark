# _*_coding: UTF-8_*_

from PyQt5.QtWidgets import QApplication, QDialog, QProgressBar
from PyQt5.QtCore import QRect

class ProgressBar(QDialog):
    def __init__(self, parent=None):
        super(ProgressBar, self).__init__(parent)

        # Qdialog窗体的设置
        self.resize(500, 32)  # QDialog窗的大小

        # 创建并设置 QProcessbar
        self.progressBar = QProgressBar(self)  # 创建
        self.progressBar.setMinimum(0)  # 设置进度条最小值
        self.progressBar.setMaximum(100)  # 设置进度条最大值
        self.progressBar.setValue(0)  # 进度条初始值为0
        self.progressBar.setGeometry(QRect(1, 3, 499, 28))  # 设置进度条在 QDialog 中的位置 [左，上，右，下]
        self.setWindowTitle("水印添加进度....")
        self.show()

    def setValue(self, task_number, total_task_number, value):  # 设置总任务进度和子任务进度
        if task_number == '0' and total_task_number == '0':
            self.setWindowTitle(self.tr('正在处理中'))
        else:
            label = "正在处理：" + "第" + str(task_number) + "/" + str(total_task_number) + '个任务'
            self.setWindowTitle(self.tr(label))  # 顶部的标题
        self.progressBar.setValue(value)


class pyqtbar():

    def __init__(self):
        self.progressbar = ProgressBar()

    def set_value(self, task_number, total_task_number, i):
        self.progressbar.setValue(str(task_number), str(total_task_number), i + 1)
        QApplication.processEvents()

    @property
    def close(self):
        self.progressbar.close()