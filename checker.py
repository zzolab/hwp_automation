import win32com.client as win32
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import os



def start_hwp(visible=False, open_file=None):
    
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
    hwpx = askopenfilename(title="한글 파일을 선택해주세요. by 우혁쌤",
                             initialdir=os.getcwd(),
                             filetypes=[("한/글파일", "*.hwp *.hwpx")])
    win.quit()  # GUI 종료
    return hwpx

def change_letter_color(face, r, g, b):
    hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
    hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("AllDoc")
    hwp.HParameterSet.HFindReplace.FindCharShape.FontTypeHangul = hwp.FontType("HFT")
    hwp.HParameterSet.HFindReplace.FindCharShape.FaceNameHangul = face
    hwp.HParameterSet.HFindReplace.ReplaceCharShape.FontTypeHangul = hwp.FontType("HFT")
    hwp.HParameterSet.HFindReplace.ReplaceCharShape.TextColor = hwp.RGBColor(r, g, b)
    hwp.HParameterSet.HFindReplace.ReplaceMode = 1
    hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
    hwp.HParameterSet.HFindReplace.FindType = 1
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

def change_letter_color_all(r, g, b):
    hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
    hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("AllDoc")
    hwp.HParameterSet.HFindReplace.FindCharShape.FontTypeHangul = hwp.FontType("TTF")
    hwp.HParameterSet.HFindReplace.FindCharShape.FaceNameHangul = "모두"
    hwp.HParameterSet.HFindReplace.ReplaceCharShape.FontTypeHangul = hwp.FontType("TTF")
    hwp.HParameterSet.HFindReplace.ReplaceCharShape.TextColor = hwp.RGBColor(r, g, b)
    hwp.HParameterSet.HFindReplace.ReplaceMode = 1
    hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
    hwp.HParameterSet.HFindReplace.FindType = 1
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

if __name__ == '__main__':

    hwp = start_hwp(visible=True, open_file=select_file())

    # 모든 글자 빨갛게 일괄변경
    change_letter_color_all(255, 0, 0)
        
    # "한양신명조" 서체만 검게
    change_letter_color("한양신명조", 0, 0, 0)
