import win32com.client as win32 # 한글 실행을 위한 모듈
from tkinter import Tk # 파일 선택을 위한 모듈
from tkinter.filedialog import askopenfilename # 파일 선택창과 관련된 모듈
import os



def start_hwp(visible=False, open_file=None):
    '''
    한글 파일을 실행하는 함수
    visible 기본 값은 False, True로 하면 한글 창의 띄워져 보이도록 if 문을 이용하여 설정
    open_file 기본 값은 None,  선택한 파일 경로를 넣을 수 있도록 설정
    '''
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")

    if visible:
        hwp.XHwpWindows.Item(0).Visible = True
 
    if open_file:
        hwp.Open(open_file)
    else:
        pass
    return hwp

def select_file():
    win = Tk()  # GUI 실행하고
    win.withdraw()
    hwpx = askopenfilename(title="한글 파일을 선택해주세요. by 우혁쌤",   #파일 선택 창 맨위에 보이는 문구
                             initialdir=os.getcwd(),                   # 기본적으로 현재 폴더를 먼저 띄우게 설정
                             filetypes=[("한/글파일", "*.hwp *.hwpx")]) # 선택하는 파일의 종류를 제한
    win.quit()  # GUI 종료
    return hwpx

def change_letter_color(face, r, g, b):
    '''
    글꼴을 하나 선택 후 
    해당 글꼴의 색을 모두 rgb에 맞춰서 바꿈
    한양신명조는 한글 전용폰트(HTF) 이므로 hwp.FontType("HTF")을 사용
    대부분 사실 hwp.FontType("TTF")
    '''
    hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
    hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("AllDoc")
#     hwp.HParameterSet.HFindReplace.FindCharShape.FontTypeHangul = hwp.FontType("TTF")
    hwp.HParameterSet.HFindReplace.FindCharShape.FontTypeHangul = hwp.FontType("HTF")
    hwp.HParameterSet.HFindReplace.FindCharShape.FaceNameHangul = face
#     hwp.HParameterSet.HFindReplace.ReplaceCharShape.FontTypeHangul = hwp.FontType("TTF")
    hwp.HParameterSet.HFindReplace.ReplaceCharShape.FontTypeHangul = hwp.FontType("HTF")
    hwp.HParameterSet.HFindReplace.ReplaceCharShape.TextColor = hwp.RGBColor(r, g, b)
    hwp.HParameterSet.HFindReplace.ReplaceMode = 1
    hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
    hwp.HParameterSet.HFindReplace.FindType = 1
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

# def change_letter_color_all(r, g, b):
#     '''
#     한글의 버그를 이용(찾는 글꼴이 없으면 모든 글꼴을 선택함)
#     모든 글자의 색을 모두 rgb에 맞춰서  
#     '''
#     hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
#     hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("AllDoc")
#     hwp.HParameterSet.HFindReplace.FindCharShape.FontTypeHangul = hwp.FontType("TTF")
#     hwp.HParameterSet.HFindReplace.FindCharShape.FaceNameHangul = "모두"
#     hwp.HParameterSet.HFindReplace.ReplaceCharShape.FontTypeHangul = hwp.FontType("TTF")
#     hwp.HParameterSet.HFindReplace.ReplaceCharShape.TextColor = hwp.RGBColor(r, g, b)
#     hwp.HParameterSet.HFindReplace.ReplaceMode = 1
#     hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
#     hwp.HParameterSet.HFindReplace.FindType = 1
#     hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

def check_letter(face, r, g, b):
    hwp.SetPos(0, 0, 0)
    hwp.Run("MoveSelDocBegin")
    hwp.InitScan()
    hwp.Run("MoveSelDocEnd")
    hwp.Run("CharShapeTextColorRed")
    change_letter_color(face, r, g, b)
    area = 2
    while True:
        hwp.SetPos(area, 0, 0)
        if hwp.GetPos()[0] == 0:
            break
        while True:
            hwp.Run("MoveSelDocBegin")
            hwp.InitScan()
            hwp.Run("MoveSelDocEnd")
            hwp.Run("CharShapeTextColorRed")
            change_letter_color(face, 0, 0, 0)
            area += 1
            hwp.SetPos(area, 0, 0)
            if hwp.GetPos()[0] == 0:
                break

if __name__ == '__main__':

    hwp = start_hwp(visible=True, open_file=select_file())

    # 모든 글자 빨갛게 일괄변경
    change_letter_color_all(255, 0, 0)
        
    # "한양신명조" 서체만 검게
    check_letter("한양신명조", 0, 0, 0)
