import os
from tkinter.filedialog import askopenfilenames
import win32com.client as win32

def start_hwp(visible=False, open_file=None):
    '''
    한글 파일을 시작하도록 하는 함수
    visible=False # 한글 창을 보이지 않도록 설정 (True로 변경시 한글 창이 보임)
    '''
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    if visible:
        hwp.XHwpWindows.Item(0).Visible = True
    if open_file:
        hwp.Open(open_file)
    else:
        pass
    return hwp

def choose_file():
    """
    파일선택 함수
    """
    filelist = askopenfilenames(title="수정할 한/글문서를 모두 선택해주세요.",
                     initialdir=os.getcwd(),
                     filetypes=[("한/글파일", "*.hwp *.hwpx")])
    return filelist

def get_name(path):
    '''
    파일 이름이 --(이름) 꼴일때
    (이름)을 뽑아내는 함수
    '''
    start = path.find("(")
    end = path.find(")") + 1
    result = path[start:end]
    return result

def get_text(hwp):
    '''
    한글에서 적혀있는 글자를 추출하는 함수
    '''
    hwp.InitScan(Range=0xff)
    total_text = ""
    state = 2
    while state not in [0, 1]:
        state, text = hwp.GetText()
        total_text += text
    hwp.ReleaseScan()
    return total_text

def InsertText(text):
    '''
    한글에서 글을 입력하는 함수
    '''
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = text
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

def move_to_begin(hwp):
    hwp.HAction.Run("MoveDocBegin")  # 문서 시작으로 이동

if __name__ == '__main__':              # 실행
    contents = []                       # 내용을 담을 리스트
    names=[]                            # 이름을 담을 리스트
    filelist=choose_file()              # 파일선택
    for file in filelist:               # 선택한 파일들 중에 하나씩 선택해서
        hwp = start_hwp(open_file=file) # 파일을 열고
        move_to_begin(hwp)              # 시작점으로 돌아간 후
        path = hwp.Path                 # 경로를 추출하고(최종 저장할 파일의 경로를 같에 만들어주기 위해)
        name=get_name(hwp.Path)         # 해당 파일의 이름을 추출하고
        names.append(name)              # names 리스트에 담는다.
        ctrl = hwp.HeadCtrl             # 컨트롤을 활성화하고
        hwp.FindCtrl()                  # 컨트롤을 찾고
        hwp.Run("ShapeObjTableSelCell") # 첫 번째 셀로 진입
        for i in range(8):              # range(n) : n은 첫번째 셀에서 F5를 누르고 목표 셀까지 오른쪽 키로 이동한 횟수
            hwp.HAction.Run("TableRightCell") # 오른쪽 키보드를 누른 것과 같은 효과
        hwp.Run("ShapeObjTableSelCell") # 도착한 셀에 진입해서
        contents.append(get_text(hwp))  # 글자를 가져오고
        hwp.Clear(option=1)             # 파일을 닫음(새로 파일을 열기 위함)
    hwp.Quit()                          # 한글 일단 종료


    # 파일 이름만 추출
    directory, filename = os.path.split(path)

    hwp = start_hwp()                   # 한글 새로 시작
    for i in range(len(names)):         # names 리스트에 담긴 이름만큼 (제출한만큼)
        text_name=names[i]              # 이름을 선택하고
        text=contents[i]                # 내용을 선택하고
        InsertText(text_name)           # 이름을 적고
        InsertText("\r\n")              # 한줄 띄고
        InsertText(text)                # 내용을 적고
        InsertText("\r\n")              # 한줄 띄고
        InsertText("\r\n")              # 한줄 띄고
    
    hwp.SaveAs(Path=directory+'\\'+'통합.hwp', Format=hwp.XHwpDocuments.Item(0).Format)  # 저장하고 
    hwp.Clear() # 문서 닫기
    hwp.Quit() # 한글 종료
