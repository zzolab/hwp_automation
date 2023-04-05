# 선생님들을 위한 한글 자동화 툴 개발
## 업로드 된 파일 설명
### checker.py
- 문서 내 통일되지 않은 폰트를 찾아 확인함
- 특정 폰트 외 다른 폰트를 모두 붉은 색으로 처리함
- [폰트의 종류](https://www.hancom.com/upload/HC/20161015/20161015191158328001.pdf)

### space_letter.py
- 글 작성중 단어가 줄을 넘어가는 경우, 글자의 자간을 자동으로 조정함.

## ascending.py
- 논문 등의 참고 문헌의 순서가 섞여있는 경우 오름차순으로 정렬
- 한글, 영어 순서로 정렬하도록 제작
- 중간에 separate_list 함수는 chatgpt를 이용하여 제작

## word_ascending.py
- ascending.py의 MS Word버전

### pyinstaller.md 
- pyinstaller를 이용하여 파이썬 설치 없이 사용할 수 있도록 배포하는 법을 설명(예정)

### hwp_to_python.md
- 한글의 매크로 기록 기능을 이용하여 파이썬 코드로 변환하는 법(예정)

#### Thanks for 일코(https://martinii.fun/)
- 일상의 코딩님 글과 강의를 참고하여 제작되었습니다.
