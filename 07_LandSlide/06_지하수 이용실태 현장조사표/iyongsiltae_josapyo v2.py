
import os
import shutil
import win32com.client as win32
import pandas as pd

XL_INPUT = "iyong_template.xlsx"
HWP_INPUT = "iyong(field).hwp"
HWP_OUTPUT = "iyong(result).hwp"

def get_desktop():
    desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    return desktop

def initial_work():
    desktop = get_desktop()
    excel = pd.read_excel(f"{desktop}\\{XL_INPUT}")
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")  # 한/글 열기

    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
    return hwp, excel

def initial_opencopy(hwp, excel):
    desktop = get_desktop()
    shutil.copyfile(f"{desktop}\\{HWP_INPUT}", f"{desktop}\\{HWP_OUTPUT}")

    hwp.Open(f"{desktop}\\{HWP_OUTPUT}")
    field_list = [i for i in hwp.GetFieldList().split("\x02")]

    print(len(field_list), field_list)

    hwp.Run('SelectAll')
    hwp.Run('Copy')
    hwp.MovePos(3)

    print('------------------------------------------------------')
    print('page copy started ...')
    print(len(excel))

    for i in range(len(excel) - 1):
        hwp.Run('Paste')
        hwp.MovePos(3)

    print(f'{len(excel)} page copy completed ...')
    print('------------------------------------------------------')
    return field_list



def copy_work(hwp, excel, field_list):
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

        print(f'{page + 1}:{excel.address[page]}')


def end_work(hwp, excel):
    hwp.Save()
    hwp.Quit()

def main():
    hwp, excel = initial_work()
    field_list = initial_opencopy(hwp, excel)
    copy_work(hwp, excel, field_list)
    end_work(hwp, excel)
    print('------------------------------------------------------')


if __name__ == "__main__":
    main()
