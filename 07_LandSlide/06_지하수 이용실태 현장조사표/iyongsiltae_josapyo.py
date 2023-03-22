

import os, shutil
import win32com.client as win32
import pandas as pd

XL_INPUT = "iyong_template.xlsx"
HWP_INPUT = "iyong(field).hwp"
HWP_OUTPUT = "iyong(result).hwp"

def get_desktop():
    desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    return desktop

desktop = get_desktop()
excel = pd.read_excel(f"{desktop}\\{XL_INPUT}")
hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")  # 한/글 열기

# import win32gui  # 한/글 창을 백그라운드로 숨기기 위한 모듈
# hwnd = win32gui.FindWindow(None, '빈 문서 1 - 한글')  # 한/글 창의 윈도우핸들값을 알아내서
# win32gui.ShowWindow(hwnd, 0)  # 한/글 창을 백그라운드로 숨김


hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
shutil.copyfile(f"{desktop}\\{HWP_INPUT}", f"{desktop}\\{HWP_OUTPUT}")

hwp.Open(f"{desktop}\\{HWP_OUTPUT}")
field_list = [i for i in hwp.GetFieldList().split("\x02")]

print(len(field_list), field_list)

hwp.Run('SelectAll')
hwp.Run('Copy')
hwp.MovePos(3)

print('페이지 복사를 시작합니다.')

print(len(excel) - 1)

for i in range(len(excel) - 1):
    hwp.Run('Paste')
    hwp.MovePos(3)

print(f'{len(excel)}페이지 복사를 완료하였습니다.')

for page in range(len(excel)):
    for field in field_list:
        data = excel[field].iloc[page]
        # if type(data) == str:
        if pd.isna(data):
            write_data = " "
        else:
            write_data = data

        hwp.MoveToField(f'{field}{{{{{page}}}}}')
        hwp.PutFieldText(f'{field}{{{{{page}}}}}', write_data)

    print(f'{page + 1}:{excel.ad1[page]}')


hwp.Save()
hwp.Quit()




