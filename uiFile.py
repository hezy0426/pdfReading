import main2
import os

# Making sure these two folders exist
inputFolder = os.path.join(os.getcwd(), '输入')
if not os.path.exists(inputFolder):
    os.makedirs(inputFolder)
print(inputFolder)

outputPath = os.path.join(os.getcwd(), '输出')
if not os.path.exists(outputPath):
    os.makedirs(outputPath)
print(outputPath)

obj1 = main2.ohNo(outputPath, inputFolder)

while True:
    ans1 = input(
        '\n输入1, 2, 3: 选1做月报。选2做在途 选3改PDF文件名\n输入exit退出\n')
    if ans1 == '1':
        yearNum = input('输入最终报价表格的年份：比如 “2022最终报价”,就输入2022\n')
        if obj1.validateNumber(yearNum):
            try:
                obj1.weeklyReport(yearNum)
            except Exception as err:
                print(err)
        else:
            print(yearNum + ' 不是4位数,请重试\n')
    elif ans1 == '2':
        try:
            obj1.orderFollowup()
        except Exception as err:
            print(err)
    elif ans1 == '3':
        try:
            obj1.nameChange()
        except Exception as err:
            print(err)
    elif ans1 == 'test':
        obj1.testing()
    elif ans1.upper() == 'EXIT':
        break
