import sys

# COM 방식이 아닌 OCX 방식이라 PyQt 패키지의 QAxContainer 모듈을 이용함
# PyQt는 Qt라는 GUI 프레임워크의 파이썬 바인딩임
from PyQt5.QAxContainer import *

from PyQt5.QtGui import *

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
    
    # 키움증권에서 제공하는 클래스를 사용하기 위해 ProgID를 QAxWidget 클래스의 생성자로 전달하여 인스턴스를 생성
    self.kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
    # COM 방식에서 인스턴스를 통해 메서드를 호출했던 것과 달리 OCX 방식에서는 QAxBase 클래스의 dynamicCall 메서드를 사용해 원하는 메서드를 호출
    # QAxWidget 클래스는 QAxBase 클래스를 상속받았으므로 QAxWidget 클래스의 인스턴스는 dynamicCall 메서드를 호출할 수 있음
    self.kiwoom.dynamicCall("CommConnect()")
    # Open API+는 통신 연결 상태가 바뀔 때 OnEventConnect라는 이벤트가 발생
    # 이벤트(OnEventConnect) 발생 시 자동으로 이벤트 처리 메서드(self.event_connect) 호출
    self.kiwoom.OnEventConnect.connect(self.event_connect)
    # 서버로부터 이벤트가 발생할 때까지 이벤트 루프를 사용해 대기
    # 서버와 통신한 후 서버로부터 데이터를 전달받은 시점에 발생
    self.kiwoom.OnReceiveTrData.connect(self.receive_trdata)
    
    # 윈도우의 타이틀을 변경하는 메소드
    self.setWindowTitle("PyInvestBot")
    # 윈도우의 좌표, 사이즈를 변경하는 메소드
    self.setGeometry(300, 300, 300, 400)
    
    # QLabel - 간단한 텍스트 출력 위젯
    # 첫 번째 인자 : 출력될 문자열, 두 번째 인자 : 부모 위젯
    # 생성자에서 텍스트를 출력하는 용도로만 사용되며 다른 메서드에서는 사용되지 않으므로 self.label로 바인딩 안함
    label = QLabel('종목코드: ', self)
    # 출력 위치 조정
    # 크기/위치 동시 조절하려면 setGeometry 사용
    label.move(20, 20)
    
    # QLineEdit - 간단한 사용자 입력 처리 위젯
    self.code_edit = QLineEdit(self)
    self.code_edit.move(80, 20)
    # 기본값으로 키움증권의 종목 코드를 QLineEdit에 출력하기 위해 setText 메서드를 사용
    self.code_edit.setText("039490")
    
    # QPushButton - 버튼 생성 위젯
    # 첫 번째 인자 : 출력될 텍스트, 두 번째 인자 : 부모 위젯
    btn1 = QPushButton("조회", self)
    btn1.move(190, 20)
    btn1.clicked.connect(self.btn1_clicked)
    
    # QTextEdit - 메시지 출력
    self.text_edit = QTextEdit(self)
    self.text_edit.setGeometry(10, 60, 280, 80)
    # 사용자가 QTextEdit 위젯을 통해 입력할 수 없고 오직 읽기 모드로만 사용하도록 setEnabled 메서드를 사용
    self.text_edit.setEnabled(False)
  
  def event_connect(self, err_code):
    if err_code == 0:
      self.text_edit.append("로그인 성공")
    
  def btn1_clicked(self):
    # text 메서드를 통해 QLineEdit에 사용자가 입력한 값을 가져옴
    code = self.code_edit.text()
    # 사용자로부터 입력받은 값을 QTextEdit 위젯에 출력하기 위해 self.text_edit라는 변수를 통해 append 메서드를 호출
    self.text_edit.append("종목코드: " + code)
    
    # SetInputValue 메서드를 사용해 TR 입력 값을 설정 (TR을 구성)
    # "종목코드", "입력값1"
    # QLineEdit 위젯을 통해 입력받은 종목 코드를 SetInputValue 메서드의 두 번째 인자로 전달
    self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)
    # CommRqData 메서드를 사용해 TR을 서버로 송신
    # "RQName", "opt10001", "0", "화면번호"
    self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "opt10001_req", "opt10001", 0, "0101")
    
  # OnReceiveTrData 메서드의 첫 번째 인자는 당연히 self이고 나머지 인자는 OnReceiveTrData 이벤트의 원형을 참조해서 인자를 순서대로 받을 수 있도록 구현 rn
  def receive_trdata(self, screen_no, rqname, trcode, recordname, prev_next, data_len, err_code, msg1, msg2):
    if rqname == "opt10001_req":
      # CommGetData 메서드를 사용해 수신 데이터를 가져옴
      # OnReceiveTrData 이벤트가 발생했다는 것은 서버로부터 데이터를 전달 받았음을 의미하므로 OnReceiveTrData 메서드에서 CommGetData 메서드를 호출해서 데이터를 가져오면ehla
      # CommGetData의 첫 번째 인자와 세 번째 인자에 TR 명과 Request 명을 입력해 어떤 TR에 대한 데이터를 얻고자 하는지 알려줘야 함
      # Tran 데이터 - Tran명, X, 레코드명, 반복인덱스, 아이템명
      # 실시간 데이터 - Key Code, Real Type, Item Index, X, X
      # 체결 데이터 - 체결구분, "-1", X, ItemIndex, X
      name = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "", rqname, 0, "종목명")
      volume = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "", rqname, 0, "거래량")
      # strip 메서드를 호출해 문자열의 공백을 제거
      # 그런 다음 QTexEdit 객체에 해당 문자열을 추가
      self.text_edit.append("종목명: " + name.strip())
      self.text_edit.append("거래량: " + volume.strip())
    
    ''''
    # QTextEdit 객체는 최상위 윈도우 안으로 생성돼야 하므로 QTextEdit 객체를 생성할 때 인자로 self 매개변수를 전달
    # self를 사용하는 이유는 클래스의 다른 메서드에서도 해당 변수를 사용해 객체에 접근하기 위해서
    # self.text_edit는 생성자에서 객체를 바인딩할 때 사용될뿐더러 event_connect 메서드에서도 사용되기 때문에 text_edit가 아니라 self.text_edit를 사용
    self.text_edit = QTextEdit(self)
    # QTextEdit 클래스의 크기 및 출력 위치를 조절
    self.text_edit.setGeometry(10, 60, 280, 80)
    # 읽기/쓰기 모드를 변경
    self.text_edit.setEnabled(False)
    
    ''''
    
    ''''
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
    # CommConnect() - 로그인 윈도우를 실행한다.
    # OCX 방식에서는 QAxBase 클래스의 dynamicCall 메서드를 사용해 원하는 메서드를 호출
    ret = self.kiwoom.dynamicCall("CommConnect()")
      
  def btn2_clicked(self):
    # GetConnectState() - 현재 접속상태를 반환한다.
    # 0:미연결, 1:연결완료
    if self.kiwoom.dynamicCall("GetConnectState()") == 0:
      self.statusBar().showMessage("Not connected")
    else:
      self.statusBar().showMessage("Connected")
    ''''
    
      
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
