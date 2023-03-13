# 该程序用于将我NAS 网盘中的视频和图片文件移动到video目录下去，并且根据视频文件的拍摄时间进行分类存放
#1. 遍历整个Z盘目录中的所有文件
#2. 判断这个文件是视频文件（mov. mp4. ) 还是 图片文件
#3. 读取文件前面4个数字存放到year中，56两个数字存放到month中
#   src = 当前目录   dst = dst\year\month (如果dst不存在，则创建dst先），接下来执行move操作

import os
import shutil
#from openpyxl import Workbook
from openpyxl import load_workbook

# Press the green button in the gutter to run the script.
if __name__ == '__main__':

    srcpathroot = 'Y:\\'
    dstpathroot = 'Z:\\family'
    videofileext = ('mov', 'mp4', 'avi', 'mkv','flv','mpg','3gp')

    failedfile = []
    movedfilesrc = []
    movedfiledst = []
    movednumber = 0

    print(f" search src is {srcpathroot}, destination director is {dstpathroot}")
    # input("请检查源目录和目标目录是否正确, 敲回车则开始遍历整个src目录，并将找到的视频文件移动到dst目录下")
    for dirpath, dirnames, files in os.walk(srcpathroot, topdown=False):
        print(f'Found directory: {dirpath}')

        for files_name in files:
            suffix = files_name.lower().endswith(tuple(videofileext))
            srcpath = os.path.join(dirpath, files_name)
            # 只有是视频文件才处理
            if suffix:
                year = files_name[0:4]
                month = files_name[4:6]

                if year.isdigit() is False or month.isdigit() is False:
                    failedfile.append(srcpath)
                    continue

                if int(year) not in range(2000,2030) or int(month) not in range(1,13):
                    failedfile.append(srcpath)
                    continue

                # 得到目的地路径
                dstpath = os.path.join(dstpathroot, year, month)
                # 如果该路径不存在，则创建该路径下对应的全部目录
                if not os.path.exists(dstpath):
                    os.makedirs(dstpath)
                # 得到需要移动的文件的全路径
                dstpath = os.path.join(dstpath, files_name)
                #print(f"move file {srcpath} to {dstpath}")
                # 保存源路径和目的地全路径
                movedfilesrc.append(srcpath)
                movedfiledst.append(dstpath)
                # 移动文件
                shutil.move(srcpath, dstpath)
                movednumber = movednumber + 1
            else:
                failedfile.append(srcpath)


    print("start to store the result into the result.xlsx")
    print('%d files were moved',movednumber)
    wb = load_workbook('result.xlsx')
    ws = wb.active

    for idx in range(1, len(failedfile)+1):
        ws.cell(idx, 1).value = failedfile[idx-1]
    for idx in range(1, len(movedfilesrc)+1):
        ws.cell(idx, 4).value = movedfilesrc[idx-1]
        ws.cell(idx, 6).value = movedfiledst[idx-1]
    wb.save("result.xlsx")


