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

def InsertText(text):
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = text
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

def separate_list(lst):
    eng_lst = []
    kor_lst = []
    for item in lst:
        if isinstance(item, str) and item[0].isalpha():
            if ord(item[0]) < 256: # 영어로 시작하는 경우
                eng_lst.append(item)
            else: # 한글로 시작하는 경우
                kor_lst.append(item)
    return eng_lst, kor_lst

if __name__ == '__main__':
    hwp = start_hwp(visible=True, open_file=select_file())

    content = []
    hwp.InitScan()
    while True:
        state, text = hwp.GetText()
        content.append(text)
        if state <= 1:
            break
    hwp.ReleaseScan()
 
    # 전처리(중복제거)
    content=set(content)
    content=list(content)
    # ''제거
    content.remove('')

    eng_lst, kor_lst=separate_list(content)

    # 정렬
    kor_lst.sort()
    eng_lst.sort()


    hwp.HAction.Run("MoveDocEnd")
    hwp.Run("BreakPage")
    InsertText("정렬된 문헌\r\n")

    for i in range(0, len(kor_lst)):
        text=kor_lst[i]
        InsertText(text)

    for i in range(0, len(eng_lst)):
        text=eng_lst[i]
        InsertText(text)

    hwp.SaveAs(Path=hwp.Path.replace(".hwp", "(정렬).hwp"), Format=hwp.XHwpDocuments.Item(0).Format)
