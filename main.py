import sys
import os
import win32com.client as win32
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QPushButton, QListWidget, QFileDialog, QMessageBox, 
                             QLabel, QAbstractItemView)
from PyQt5.QtCore import Qt

class HwpxMerger(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle('HWPX 파일 병합 도구')
        self.setGeometry(300, 300, 500, 400)
        
        layout = QVBoxLayout()
        
        # 파일 목록 레이블
        self.label = QLabel('병합할 HWPX 파일을 추가하세요 (순서대로 병합됩니다):')
        layout.addWidget(self.label)
        
        # 파일 목록 리스트 위젯
        self.listWidget = QListWidget()
        self.listWidget.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.listWidget.setDragDropMode(QAbstractItemView.InternalMove) # 드래그 앤 드롭으로 순서 변경 가능
        layout.addWidget(self.listWidget)
        
        # 버튼 레이아웃
        btnLayout = QHBoxLayout()
        
        self.btnAdd = QPushButton('파일 추가')
        self.btnAdd.clicked.connect(self.addFiles)
        btnLayout.addWidget(self.btnAdd)
        
        self.btnRemove = QPushButton('선택 삭제')
        self.btnRemove.clicked.connect(self.removeFiles)
        btnLayout.addWidget(self.btnRemove)
        
        self.btnClear = QPushButton('모두 비우기')
        self.btnClear.clicked.connect(self.listWidget.clear)
        btnLayout.addWidget(self.btnClear)
        
        layout.addLayout(btnLayout)
        
        # 병합 실행 버튼
        self.btnMerge = QPushButton('하나의 파일로 병합하기')
        self.btnMerge.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; height: 40px;")
        self.btnMerge.clicked.connect(self.mergeFiles)
        layout.addWidget(self.btnMerge)
        
        self.setLayout(layout)
        
    def addFiles(self):
        files, _ = QFileDialog.getOpenFileNames(self, 'HWPX 파일 선택', '', 'HWPX Files (*.hwpx)')
        if files:
            self.listWidget.addItems(files)
            
    def removeFiles(self):
        for item in self.listWidget.selectedItems():
            self.listWidget.takeItem(self.listWidget.row(item))
            
    def mergeFiles(self):
        count = self.listWidget.count()
        if count < 2:
            QMessageBox.warning(self, '경고', '최소 2개 이상의 파일을 선택해야 합니다.')
            return
            
        savePath, _ = QFileDialog.getSaveFileName(self, '병합 파일 저장', 'merged_file.hwpx', 'HWPX Files (*.hwpx)')
        if not savePath:
            return
            
        try:
            # 한컴오피스 객체 생성
            hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
            # 보안 승인 모듈 등록 (필요한 경우)
            # hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
            
            # 첫 번째 파일 열기
            first_file = self.listWidget.item(0).text()
            hwp.Open(first_file)
            
            # 문서 끝으로 이동
            hwp.Run("MoveDocEnd")
            
            # 나머지 파일들을 순서대로 끼워넣기
            for i in range(1, count):
                file_path = self.listWidget.item(i).text()
                
                # InsertFile 액션 설정
                hset = hwp.HParameterSet.HInsertFile
                hwp.HAction.GetDefault("InsertFile", hset.HSet)
                hset.FileName = file_path
                hset.KeepSection = 1 # 구역 나누기 유지 여부 (필요에 따라 0 또는 1)
                hwp.HAction.Execute("InsertFile", hset.HSet)
                
                # 다음 삽입을 위해 다시 끝으로 이동
                hwp.Run("MoveDocEnd")
            
            # 결과 저장
            hwp.SaveAs(savePath, "HWPX")
            hwp.Quit()
            
            QMessageBox.information(self, '성공', f'병합이 완료되었습니다!\n저장 위치: {savePath}')
            
        except Exception as e:
            QMessageBox.critical(self, '오류', f'병합 중 오류가 발생했습니다:\n{str(e)}')
            if 'hwp' in locals():
                hwp.Quit()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = HwpxMerger()
    ex.show()
    sys.exit(app.exec_())
