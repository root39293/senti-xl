import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget, QFileDialog, QLabel, QProgressBar, QTextEdit, QLineEdit
from PyQt5.QtCore import Qt
import qdarkstyle
from analysis_thread import AnalysisThread
from dotenv import load_dotenv
import os

load_dotenv()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('감정 분석 프로그램')
        self.setGeometry(100, 100, 500, 400)
        layout = QVBoxLayout()

        self.label = QLabel("파일을 선택하여 감정 분석을 시작하세요.")
        self.label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.label)

        self.columnLabel = QLabel("분석할 열 이름을 입력하세요:")
        layout.addWidget(self.columnLabel)

        self.columnInput = QLineEdit()
        layout.addWidget(self.columnInput)

        self.logEdit = QTextEdit()
        self.logEdit.setReadOnly(True)
        layout.addWidget(self.logEdit)

        self.progressBar = QProgressBar(self)
        self.progressBar.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.progressBar)

        self.buttonSelect = QPushButton('파일 선택')
        self.buttonSelect.clicked.connect(self.openFileDialog)
        layout.addWidget(self.buttonSelect)

        self.buttonRun = QPushButton('실행')
        self.buttonRun.clicked.connect(self.runAnalysis)
        self.buttonRun.setEnabled(False)
        layout.addWidget(self.buttonRun)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)
        
        self.setStyleSheet(qdarkstyle.load_stylesheet_pyqt5())

    def openFileDialog(self):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(self, "엑셀 파일 선택", "", "Excel Files (*.xlsx);;All Files (*)", options=options)
        if fileName:
            self.fileName = fileName
            self.label.setText(f"선택된 파일: {fileName}")
            self.logEdit.append("파일이 선택되었습니다.")
            self.buttonRun.setEnabled(True)

    def runAnalysis(self):
        self.logEdit.append("분석을 시작합니다...")
        columnName = self.columnInput.text().strip()
        if columnName:
            self.analyzeFile(self.fileName, columnName)
        else:
            self.logEdit.append("분석할 열 이름을 입력하세요.")

    def analyzeFile(self, fileName, columnName):
        self.thread = AnalysisThread(fileName, columnName)
        self.thread.progress.connect(self.progressBar.setValue)
        self.thread.log.connect(self.logEdit.append)
        self.thread.finished.connect(self.onFinished)
        self.thread.start()

    def onFinished(self, outputFileName):
        self.label.setText(f"분석 완료! 결과 파일: {outputFileName}")
        self.logEdit.append(f"분석 완료! 결과 파일: {outputFileName}")
        self.progressBar.setValue(100)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    mainWindow = MainWindow()
    mainWindow.show()
    sys.exit(app.exec_())
