import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget, QFileDialog, QLabel, QProgressBar, QTextEdit
from PyQt5.QtCore import QThread, pyqtSignal, Qt
import qdarkstyle
from openai import OpenAI
import openpyxl

OPENAI_API_KEY = "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"
client = OpenAI(api_key=OPENAI_API_KEY)

def analyze_sentiment(text):
    try:
        completion = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": """
                #지시문
                너는 지금부터 도시가스 안전점검 서비스의 주관식 응답내용 해석전문가 '박결이 대리'이다. '박결이 대리'의 역할로써,  안전점검 서비스의 주관식 응답내용을 제약조건, 출력형태를 준수하여 사용자가 입력하는 내용에 대해 감정분석을 하는 것이 너의 역할이다. 예시를 참고하여 답하여라.
                
                #제약조건
                - 반드시 긍정 또는 부정으로 출력한다.
                - 부연설명은 하지 않는다.
                - 점검원 칭찬, 방문지연, 불친절, 안내미흡, 가스누설, 점검만족 등의 영역에서 감정을 분석하여 출력한다.
                - 사용자의 입력이 '.','ㅇ','ㅇㅇ' 등과 같이 의미없는 응답일경우 반드시 긍정으로 출력한다.
                - 입력 문자열이 '사용자의 입력 : ', 즉 아무런 값이 입력되지않았을 경우, 긍정으로 출력한다. 
                - 제안 및 해결책을 제시하는 입력은 '부정'으로 출력한다.
                
                #출력형태
                [ 긍정, 부정 둘중 하나만 출력 ]
                
                #예시
                 
                사용자의 입력 : 시간약속정확
                긍정 
                 
                사용자의 입력 : 친절
                긍정
                
                사용자의 입력 : 상세한 설명 감사합니다
                긍정
                 
                사용자의 입력 : 몇시쯤 방문하는지 대충이라도 알았으면 좋겠네요.아침부터 언제올지몰라 외출도못하고 마냥기다려야해서 불편했습니다.
                부정
                
                사용자의 입력 : 점검시간이  주간에 하다보니 요즘은 맞벌이 세대가 많고 낮에는 집을 비우는 세대가 많은것을 고려하여 18시~21시에 점검받을수 있는방법 고려가 필요함
                부정
                 
                사용자의 입력 : 보일러 배기통 배기확인 그으름이있는지 확인검사도해주셨으면합니다
                부정
                 
                
                """},
                {"role": "user", "content": f'사용자의 입력 : {text}'}
            ],
            temperature=0.3
        )
        sentiment = completion.choices[0].message.content.strip()
        return sentiment
    except Exception as e:
        print(f"Error: {e}")
        return "분석 오류"

class AnalysisThread(QThread):

    finished = pyqtSignal(str)
    progress = pyqtSignal(int)
    log = pyqtSignal(str)

    def __init__(self, fileName):
        super().__init__()
        self.fileName = fileName

    def run(self):
        df = pd.read_excel(self.fileName)
        df['번호'] = df['번호'].astype(str)
        
        condition_indices = df[df['선택내용'] == '주관식내용'].index
        
        df['응답내용분석결과'] = None 
        
        total = len(condition_indices)
        for i, idx in enumerate(condition_indices, 1):
            text = df.at[idx, '주관식응답내용']
            sentiment = analyze_sentiment(text)
            df.at[idx, '응답내용분석결과'] = sentiment  
            progress = int((i / total) * 100)
            self.progress.emit(progress)
            self.log.emit(f"분석 중... {i}/{total}")
        
        outputFileName = self.fileName.replace('.xlsx', '_감정분석결과.xlsx')
        df.to_excel(outputFileName, index=False, engine='openpyxl')
        
        self.finished.emit(outputFileName)


# 메인 윈도우 UI 클래스
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
        self.analyzeFile(self.fileName)

    def analyzeFile(self, fileName):
        self.thread = AnalysisThread(fileName)
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