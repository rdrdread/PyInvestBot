# COM 방식이 아닌 OCX 방식이라 PyQt 패키지의 QAxContainer 모듈을 이용함
# PyQt는 Qt라는 GUI 프레임워크의 파이썬 바인딩임
from PyQt5.QAxContainer import *

# PyQt5라는 디렉터리의 QtWidgets 파일에 있는 모든 것 (*)을 import 하라는 의미이므로,
# 해당 모듈 (QtCore)에 정의된 변수, 함수, 클래스를 모듈 이름을 통해 접근할 필요 없이 바로 사용할 수 있음
# 위젯(widget)은 사용자 인터페이스를 구성하는 가장 기본적인 부품 역할임
# QGroupBox, QLabel, QTextEdit, QDateEdit, QTimeEdit, QLineEdit 등
from PyQt5.QtWidgets import *

# 다른 위젯에 포함되지 않은 최상위 위젯을 특별히 윈도우(window)라 칭함
# 윈도우를 생성하기 위한 클래스로 QMainWindow나 QDialog 클래스가 일반적으로 사용됨
# MyWindow는 QMainWindow를 상속받는 형태임
# MyWindow 클래스는 PyQt가 제공하는 QMainWindow 클래스를 단지 상속함으로써 QMainWindow와 동등한 능력을 갖추게 되는 것임
class MyWindow(QMainWindow):
  def __init__(self):
    # super 내장 함수를 이용해 부모 클래스(QMainWindow)의 생성자를 명시적으로 호출
    super ().__init__()
    # 윈도우의 타이틀을 변경하는 메소드
    self.setWindowTitle("PyInvestBot")
    # 윈도우의 좌표, 사이즈를 변경하는 메소드
    self.setGeometry(300, 300, 300, 400)
    
    # 키움증권에서 제공하는 클래스를 사용하기 위해 ProgID를 QAxWidget 클래스의 생성자로 전달하여 인스턴스를 생성
    self.kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
    # COM 방식에서 인스턴스를 통해 메서드를 호출했던 것과 달리 OCX 방식에서는 QAxBase 클래스의 dynamicCall 메서드를 사용해 원하는 메서드를 호출
    # QAxWidget 클래스는 QAxBase 클래스를 상속받았으므로 QAxWidget 클래스의 인스턴스는 dynamicCall 메서드를 호출할 수 있음
    self.kiwoom.dynamicCall("CommConnect()")
    
    # QTextEdit 객체는 최상위 윈도우 안으로 생성돼야 하므로 QTextEdit 객체를 생성할 때 인자로 self 매개변수를 전달
    # self를 사용하는 이유는 클래스의 다른 메서드에서도 해당 변수를 사용해 객체에 접근하기 위해서
    # self.text_edit는 생성자에서 객체를 바인딩할 때 사용될뿐더러 event_connect 메서드에서도 사용되기 때문에 text_edit가 아니라 self.text_edit를 사용
    self.text_edit = QTextEdit(self)
    # QTextEdit 클래스의 크기 및 출력 위치를 조절
    self.text_edit.setGeometry(10, 60, 280, 80)
    # 읽기/쓰기 모드를 변경
    self.text_edit.setEnabled(False)
    
    # Open API+는 통신 연결 상태가 바뀔 때 OnEventConnect라는 이벤트가 발생
    # 이벤트(OnEventConnect) 발생 시 자동으로 이벤트 처리 메서드(self.event_connect) 호출
    self.kiwoom.OnEventConnect.connect(self.event_connect)
    
    # 윈도우 버튼 추가를 위해 MyWindow 클래스의 생성자에서 QPushButton 클래스의 인스턴스를 생성
    btn1 = QPushButton("Login", self)
    # 버튼 출력 위치를 조정
    btn1.move(20, 20)
    # 'clicked' 이벤트와 btn1_clicked라는 이벤트 처리 메서드를 connect라는 메서드로 연결
    btn1.clicked.connect(self.btn1_clicked)
    
    btn2 = QPushButton("Check state", self)
    btn2.move(20, 70)
    btn2.clicked.connect(self.btn2_clicked)
    
  def btn1_clicked(self):
    ret = self.kiwoom.dynamicCall("CommConnect()")
      
  def btn2_clicked(self):
    if self.kiwoom.dynamicCall("GetConnectState()") == 0:
      self.statusBar().showMessage("Not connected")
    else:
      self.statusBar().showMessage("Connected")
      
  def event_connect(self, err_code):
    if err_code == 0:
      self.text_edit.append("로그인 성공")
      
if __name__ == "main":
  # QApplication 클래스의 인스턴스 생성
  # 이벤트 루프를 담당하는 부분
  app = QApplication (sys.argv)
  # 앞서 작성한 MyWindow 클래스의 인스턴스 생성
  myWindow =MyWindow()
  # 새롭게 작성한 윈도우를 화면에 출력
  myWindow.show()
  # 이벤트 루프를 실행
  app.exec_()
