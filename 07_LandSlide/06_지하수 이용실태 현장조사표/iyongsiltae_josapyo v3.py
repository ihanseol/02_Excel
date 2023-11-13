import os
import shutil
import pandas as pd
from hwpapi.core import App

XL_INPUT = "iyong_template.xlsx"
HWP_INPUT = "iyong(field).hwp"
HWP_OUTPUT = "iyong(result).hwp"


def get_desktop():
    desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    return desktop


def initial_work():
    desktop = get_desktop()
    excel = pd.read_excel(f"{desktop}\\{XL_INPUT}")
    app = App(None, False)
    return app, excel


def initial_opencopy(app, excel):
    desktop = get_desktop()
    shutil.copyfile(f"{desktop}\\{HWP_INPUT}", f"{desktop}\\{HWP_OUTPUT}")
    app.open(f"{desktop}\\{HWP_OUTPUT}")
    hwp = app.api

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

"""
def copy_work(app, excel, field_list):
    hwp = app.api
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

"""


def copy_work(app, excel, field_list):
    hwp = app.api

    for page, address in enumerate(excel.address):
        for field in field_list:
            data = excel[field].iloc[page]
            write_data = " " if pd.isna(data) else data

            field_tag = f'{field}{{{{{page}}}}}'
            hwp.MoveToField(field_tag)
            hwp.PutFieldText(field_tag, write_data)

        print(f'{page + 1}:{address}')


def end_work(app, excel):
    app.api.Save()
    app.quit()


def main():
    app, excel = initial_work()
    field_list = initial_opencopy(app, excel)
    copy_work(app, excel, field_list)
    end_work(app, excel)
    print('------------------------------------------------------')


if __name__ == "__main__":
    main()
