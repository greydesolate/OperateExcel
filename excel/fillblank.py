import pandas as pd
import numpy as np
import openpyxl
from xlsxwriter.workbook import Workbook
import datetime

# 参数名	参数类型	参数默认值
# 题干起始行	数字	1
# 题干起始列	数字	2
questionBeginColIndex = 2
# 题干被替换符号	字符串	____
questionToBeReplacedStr = "___"
# 作答数据起始行	数字	2
userSet_answerBeginRowIndex = 2
# 读取整个Excel时，Excel的第二行等于 dataframe 里的第0行。所以下面第一个2 是可变参数，后面 ajustRowIndex 是固定值。
ajustRowIndex = 2
answerBeginRowIndex = userSet_answerBeginRowIndex - ajustRowIndex
# 作答数据起始列	数字	2
userSet_answerBeginColIndex = 2
ajustColIndex = 1
answerBeginColIndex = userSet_answerBeginColIndex - ajustColIndex
# 作答内容起始符号	字符串多个字符逗号分隔	(
answerBeginStr = "("
# 作答内容结束符号	字符串多个字符逗号分隔	)
answerBeginStr = ")"

# 是否保留作答内容起始结束符号
isKeepAnswerContainChar = True

# 对于空白作答要替换为的字内容，默认空字符串不进行替换
replaceEmptyAnswerStr = ""

# 作答数据填充的颜色	字符串	红色；如果内容为 "" 空字符串或者空值，则不进行颜色替换。
replacedFontColor = "red"
# 颜色对象
style_format = ({'color': replacedFontColor})

# 文件名	字符串	空
fileName = ""
# 文件路径（如果脚本跟Excel文件在同一个目录下就不用填写）	字符串	空
filePath = ""

re = pd.DataFrame(pd.read_excel("G:\\work\\OperateExcel\\FillBlank.xlsx", sheet_name=0))

filetype_xlsx = ".xlsx"

def not_empty(s):
  return s and s.strip()
# print(list(filter(not_empty, ['A', '', 'B', None,'C', ' '])))
def mergeQuesionAndAnswer(isAddColor, question_list, answer_list, red):
    fill_list = []
    answer_index = 0
    for question in question_list:
        if len(question) > 0:
            fill_list.append(question)

        if answer_index < len(answer_list):
            answer_str = answer_list[answer_index]
            if isKeepAnswerContainChar:
                if bool(isAddColor):
                    fill_list.append(red)
                fill_list.append(answer_str)
            else:
                answer_str = answer_str[1:len(answer_list[answer_index])-1]
                if len(answer_str) > 0:
                    if bool(isAddColor):
                        fill_list.append(red)
                    fill_list.append(answer_str)

        answer_index += 1
    print(fill_list)
    return fill_list


# def fillExcelBlank():

# print(re)
# print(re.head())
#  获取行号
# print(re._stat_axis.values.tolist())
# 获取第一行表头List
headerList = re.columns.values.tolist()
# headerList_splitedList 保存根据 questionToBeReplacedStr 拆分的结果。
headerList_splitedList = []
# print(headerList)


createfilename = str(datetime.datetime.now())
workbook = Workbook("createfilename" + filetype_xlsx)  # 创建xlsx

worksheet = workbook.add_worksheet('结果sheet')  # 添加sheet

red = workbook.add_format(style_format)  # 颜色对象

# 将 headerList 处理为 List（str.split(questionToBeReplacedStr))
for headerListIndex in range(questionBeginColIndex - 1, len(headerList)):
    # 根据配置的分隔符 questionToBeReplacedStr 拆成新的list
    headerList_splitedList.append(str(headerList[headerListIndex]).split(questionToBeReplacedStr))
    # 往目标Excel的sheet 填充表头
    worksheet.write(0, headerListIndex, headerList[headerListIndex])
# print(headerList_splitedList)


# 修改re 的表头 从[0, len(headerList))

# re.columns = [str(i) for i in range(0, len(headerList))]
# re.columns = map(str,range(0, len(headerList))
re.columns = range(0, len(headerList))
# print("re修改表头后")
# print(re)
# print(re.columns.values.tolist())


# 填写题目题干内容作为结果 Excel 表头

# worksheet.write(0, 0, 'sentences')  # 0，0表示row，column，sentences表示要写入的字符串

# test_list = ["我爱", "中国", "天安门"]

# test_list.insert(1, red)  # 将颜色对象放入需要设置颜色的词语前面
# print(test_list)
# worksheet.write_rich_string(1, 0, *test_list)  # 写入工作簿
# workbook.close()  # 记得关闭

# 从配置信息 对数据 re 开始行遍历每列数据
# print(len(re))
for answerRow in re.itertuples(index=True):
    # 读取考生作答数据局，按顺序替换掉 题干里的空白，并设置字色 replacedFontColor
    for answerColIndex in range(0, len(headerList)):
        # print(len(answerRow))
        print("answerColIndex" + str(answerColIndex))
        # 重置表头后 第一列是 1.所以需要answerColIndex 进行+1
        cell_content = str(getattr(answerRow, "_" + (str(answerColIndex + 1))))
        cell_content_split_list = cell_content.split("\n")
        cell_content_split_list = list(filter(not_empty, cell_content_split_list))
        print(cell_content_split_list)
        print("---")

        current_row_index = userSet_answerBeginRowIndex - 1 + int(getattr(answerRow, "Index"))
        # 如果没有到作答数据列,则进行复制内容到对应单元格
        if answerColIndex <answerBeginColIndex:
            worksheet.write(current_row_index, answerColIndex, cell_content)
        # 如果到达作答数据列,进行拼接题目和作答数据
        else:
            fill_content = mergeQuesionAndAnswer(True, headerList_splitedList[answerColIndex - answerBeginColIndex],
                                                 cell_content_split_list, red)
            worksheet.write_rich_string(current_row_index,
                                        answerColIndex, *fill_content)
workbook.close()


# print(re[answerIndex])

# data = {"one": np.random.randn(4), "two": np.linspace(1, 4, 4), "three": ['zhangsan', '李四', 999, 0.1]}
# df = pd.DataFrame(data, index=[1, 2, 3, 4])
#
# print(df)
# 获取题干数组

# def out_data(re):
#     print(re)

def createXlsxWorkBook(createPath, createfilename):
    createPath = str(createPath).strip()
    createfilename = str(createfilename).strip()
    filetype_xlsx = ".xlsx"
    if len(createfilename) == 0:
        createfilename = str(datetime.datetime.now())
    workbook = Workbook(createfilename + filetype_xlsx)  # 创建xlsx

    worksheet = workbook.add_worksheet('结果sheet')  # 添加sheet

    red = workbook.add_format({'color': 'red'})  # 颜色对象

    worksheet.write(0, 0, 'sentences')  # 0，0表示row，column，sentences表示要写入的字符串

    test_list = ["我爱", "中国", "天安门"]

    test_list.insert(1, red)  # 将颜色对象放入需要设置颜色的词语前面
    print(test_list)
    worksheet.write_rich_string(1, 0, *test_list)  # 写入工作簿
    workbook.close()  # 记得关闭
