import pandas as pd
from PyQt5.QtCore import QThread, pyqtSignal
from openai_api import analyze_sentiment

class AnalysisThread(QThread):
    finished = pyqtSignal(str)
    progress = pyqtSignal(int)
    log = pyqtSignal(str)

    def __init__(self, fileName, columnName):
        super().__init__()
        self.fileName = fileName
        self.columnName = columnName

    def run(self):
        try:
            df = pd.read_excel(self.fileName)
            if self.columnName not in df.columns:
                self.log.emit(f"열 '{self.columnName}'이(가) 파일에 없습니다.")
                self.finished.emit("")
                return

            df['응답내용분석결과'] = None 
            total = len(df)
            
            for i, row in df.iterrows():
                text = row[self.columnName]
                sentiment = analyze_sentiment(text)
                df.at[i, '응답내용분석결과'] = sentiment
                progress = int((i / total) * 100)
                self.progress.emit(progress)
                self.log.emit(f"분석 중... {i+1}/{total}")

            outputFileName = self.fileName.replace('.xlsx', '_감정분석결과.xlsx')
            df.to_excel(outputFileName, index=False, engine='openpyxl')
            self.finished.emit(outputFileName)
        except Exception as e:
            self.log.emit(f"오류 발생: {e}")
            self.finished.emit("")

