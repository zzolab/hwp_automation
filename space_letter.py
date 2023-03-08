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


def num_letter():
    """
    자간자동조정 함수에서
    라인 끝에 걸쳐진 단어의
    앞뒤길이를 각각 계산하기 위함.
    """
    hwp.InitScan(Range=0xff)
    _, text = hwp.GetText()
    hwp.ReleaseScan()
    return len(text)


def adjustment():
    """
    모든 라인을 순회하면서
    끝에 걸쳐친 단어를 탐색함.

    잘린 단어의 앞이 길면
    라인 전체의 자간을 줄이고,

    잘린 단어의 뒤가 길면
    라인 전체의 자간을 늘임.

    한 줄 문단이 되거나
    걸쳐진 단어가 없으면 종료.
    """
    while True:
        hwp.Run("MoveLineEnd")
        hwp.Run("MoveSelWordBegin")
        front_length = num_letter()
        if front_length == 0:  # 단어가 잘려있지 않으면 자간조정 중지
            break
        hwp.Run("MoveSelWordEnd")
        back_length = num_letter()
        if not (front_length and back_length):  # 한 줄 문단이면 자간조정 중지
            hwp.Run("Cancel")
            break
        hwp.Run("MoveWordBegin")
        hwp.Run("MoveLineEnd")
        hwp.Run("MoveSelLineBegin")
        if front_length >= back_length:  # 앞이 길면
            hwp.Run("CharShapeSpacingDecrease")  # 라인 자간 -1%
        else:  # 뒤가 길면
            hwp.Run("CharShapeSpacingIncrease")  # 라인 자간 +1%
        hwp.Run("Cancel")


def ctrl_adjustment():
    """
    표나 글상자 등 텍스트가 들어가는
    모든 영역의 자간을 조정하기 위함
    """
    area = 2
    while True:
        hwp.SetPos(area, 0, 0)
        if hwp.GetPos()[0] == 0:
            break
        while True:
            adjustment()
            hwp.Run("MoveLineEnd")
            hwp.Run("MoveNextPosEx")
            if hwp.GetPos()[0] == 0:
                break
        area += 1


def end_position():
    """
    본문 탐색 while문의 종료 조건으로
    "문서 끝에 도착하면 반복종료"를 구현하기 위해
    문서 끝 위치를 미리 추출해 둠
    """
    hwp.Run("MoveDocEnd")
    end_pos = hwp.GetPos()  # 종료위치 저장
    hwp.Run("MoveDocBegin")
    return end_pos

def select_file():
    win = Tk()  # GUI 실행하고
    win.withdraw()
    hwpx = askopenfilename(title="한글 파일을 선택해주세요. by 우혁쌤",
                             initialdir=os.getcwd(),
                             filetypes=[("한/글파일", "*.hwp *.hwpx")])
    win.quit()  # GUI 종료
    return hwpx

if __name__ == '__main__':
    hwp = start_hwp(visible=True, open_file=select_file())
    end = end_position()

    # 본문 자간조정
    while hwp.GetPos() != end:
        adjustment()
        hwp.Run("MoveLineEnd")
        hwp.Run("MoveNextPosEx")

    # 표 및 글상자 자간조정
    ctrl_adjustment()
    # print("자간조정 작업 끝!")
    
    hwp.SaveAs(Path=hwp.Path.replace(".hwp", "(자간조정).hwp"), Format=hwp.XHwpDocuments.Item(0).Format)
