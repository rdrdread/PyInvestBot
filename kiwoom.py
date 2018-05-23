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
    # 윈도우의 좌표, 사이즈 변굥하는 메소드
    self.setGeometry(300, 300, 300, 400)
    
if __name__ == "main":
  # QApplication 클래스의 인스턴스 생성
  # 이벤트 루프를 담당하는 부분
  app = QApplication (sys.argv)
  # 앞서 작성한 MyWindow 클래스의 인스턴스 생성
  myWindow =MyWindow()
  # 새롭게 작성한 윈도우를 화면에 출력
  myWindow.show()
  #
  app.exec_()
