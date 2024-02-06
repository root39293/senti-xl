import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget, QFileDialog, QLabel, QProgressBar
from PyQt5.QtCore import QThread, pyqtSignal, Qt  
from openai import OpenAI
import qdarkstyle

OPENAI_API_KEY = "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"
client = OpenAI(api_key=OPENAI_API_KEY)

def analyze_sentiment(text):
    try:
        completion = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": """
                #지시문
                너는 지금부터 감정분석 기계이다. 감정분석기계의 역할로서 제약조건, 출력형태를 준수하여 사용자가 입력하는 내용에 대해 감정분석을 하는 것이 너의 역할이다. 예시를 참고하여 답하여라.
                
                #제약조건
                - 가급적 짧게 답한다.
                - 부연설명은 하지 않는다.
                
                #출력형태
                [ 사용자의 입력내용에 따른 감정 긍정, 부정, 평이 출력 ]
                
                #예시
                친절하게 잘 알려주셨습니다.
                긍정
                
                검사관이 늦게 오셨어요
                부정
                
                나쁘지 않았습니다
                평이
                """},
                {"role": "user", "content": f'#사용자의 입력 {text}'}
            ]
        )
        sentiment = completion.choices[0].message.content.strip()
        return sentiment
    except Exception as e:
        print(f"Error: {e}")
        return "분석 오류"

class AnalysisThread(QThread):
    finished = pyqtSignal(str)
    progress = pyqtSignal(int)

    def __init__(self, fileName):
        super().__init__()
        self.fileName = fileName

    def run(self):
        df = pd.read_excel(self.fileName)
        df['번호'] = df['번호'].astype(str)
        df.dropna(subset=['주관식응답내용'], inplace=True)
        
        total = len(df['주관식응답내용'])
        sentiment_results = []
        for i, text in enumerate(df['주관식응답내용'], 1):
            sentiment = analyze_sentiment(text)
            sentiment_results.append(sentiment)
            progress = int((i / total) * 100)
            self.progress.emit(progress)
        
        df['응답내용분석결과'] = sentiment_results
        outputFileName = self.fileName.replace('.xlsx', '_감정분석결과.xlsx')
        df.to_excel(outputFileName, index=False)
        
        self.finished.emit(outputFileName)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('감정 분석 프로그램')
        self.setGeometry(100, 100, 400, 300)
        layout = QVBoxLayout()

        self.label = QLabel("파일을 선택하여 감정 분석을 시작하세요.")
        self.label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.label)

        self.progressBar = QProgressBar(self)
        self.progressBar.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.progressBar)

        self.button = QPushButton('파일 선택')
        self.button.clicked.connect(self.openFileDialog)
        layout.addWidget(self.button)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)
        
        self.setStyleSheet(qdarkstyle.load_stylesheet_pyqt5())

    def openFileDialog(self):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(self, "엑셀 파일 선택", "", "Excel Files (*.xlsx);;All Files (*)", options=options)
        if fileName:
            self.label.setText(f"파일 분석 중입니다... : {fileName}")
            QApplication.processEvents()
            self.analyzeFile(fileName)

    def analyzeFile(self, fileName):
        self.thread = AnalysisThread(fileName)
        self.thread.progress.connect(self.progressBar.setValue)
        self.thread.finished.connect(self.onFinished)
        self.thread.start()

    def onFinished(self, outputFileName):
        self.label.setText(f"분석 완료! 결과 파일: {outputFileName}")
        self.progressBar.setValue(100)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyleSheet(qdarkstyle.load_stylesheet_pyqt5())  
    mainWindow = MainWindow()
    mainWindow.show()
    sys.exit(app.exec_())
