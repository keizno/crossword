pip install customtkinter openpyxl pillow reportlab

pip install로 필요한 라이브러리 설치 후 더블클릭

#exe 로 만들때는..

PyInstaller 사용하시면 됩니다.
1. 설치
pip install pyinstaller

2. 단순 빌드 (테스트용)
pyinstaller --onefile --windowed crossword_puzzle.py

4. 제대로 된 빌드 (권장)
의존성이 많아서 spec 파일로 관리하는 게 낫습니다.
spec 파일 먼저 생성
pyinstaller --onefile --windowed --name "CrosswordMaker" crossword_puzzle.py

5. 생성된 CrosswordMaker.spec 파일을 열어서 hiddenimports 추가:

a = Analysis(
    ['crossword_puzzle.py'],
    ...
    hiddenimports=[
        'customtkinter',
        'PIL._tkinter_finder',
        'openpyxl',
        'reportlab',
        'reportlab.pdfbase.ttfonts',
        'reportlab.platypus',
    ],
    ...
)

6. 한글 폰트 같이 패키징할 경우
나눔고딕이 없는 PC에서도 돌아가게 하려면 폰트 파일을 포함시켜야 합니다.

spec 파일 안에
a = Analysis(
    ...
    datas=[
        ('NanumGothic.ttf', '.'),       # 폰트를 같은 폴더에 두고
        ('NanumGothicBold.ttf', '.'),
    ],
    ...
)

7. 그리고 코드에서 폰트 경로를 실행파일 기준으로 찾도록 수정:
   
import sys, os

def _get_base_path():
    # PyInstaller로 빌드된 경우 sys._MEIPASS 사용
    if hasattr(sys, '_MEIPASS'):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))

_register_korean_fonts() 안에 추가
base = _get_base_path()
candidates = [
    (os.path.join(base, 'NanumGothic.ttf'),
     os.path.join(base, 'NanumGothicBold.ttf')),
    # 기존 경로들...
]

8. 그 다음 spec으로 빌드:
pyinstaller CrosswordMaker.spec

9. dist/
└── CrosswordMaker.exe   ← 이걸 배포

