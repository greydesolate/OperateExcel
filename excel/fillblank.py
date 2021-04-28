import pandas as pd
import numpy as np
import openpyxl
from xlsxwriter.workbook import Workbook

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
# 文件名	字符串	空
fileName = ""
# 文件路径（如果脚本跟Excel文件在同一个目录下就不用填写）	字符串	空
filePath = ""

re = pd.DataFrame(pd.read_excel("G:\\work\\OperateExcel\\FillBlank.xlsx", sheet_name=0))

# print(re)
# print(re.head())
#  获取行号
# print(re._stat_axis.values.tolist())
# 获取第一行表头List
headerList = re.columns.values.tolist()
# headerList_splitedList 保存根据 questionToBeReplacedStr 拆分的结果。
headerList_splitedList = []
# print(headerList)
# 将 headerList 处理为 List（str.split(questionToBeReplacedStr))
for headerListIndex in range(1, len(headerList)):
    # 根据配置的分隔符 questionToBeReplacedStr 拆成新的list
    headerList_splitedList.append(str(headerList[headerListIndex]).split(questionToBeReplacedStr))
# print(headerList_splitedList)


# 修改re 的表头 从[0, len(headerList))

# re.columns = [str(i) for i in range(0, len(headerList))]
# re.columns = map(str,range(0, len(headerList))
re.columns = range(0, len(headerList))
# print("re修改表头后")
# print(re)
# print(re.columns.values.tolist())

# 从配置信息 对数据 re 开始行遍历每列数据
# print(len(re))
for answerRow in re.itertuples(index=False):
    # 读取考生作答数据局，按顺序替换掉 题干里的空白，并设置字色 replacedFontColor
    for answerColIndex in range(answerBeginColIndex,len(headerList)):
        # print(len(answerRow))
        cellContent = getattr(answerRow,"_"+(str(answerColIndex)))
        # print(cellContent)
        print(cellContent.split("\n"))
        print("---")
# print(re[answerIndex])

# data = {"one": np.random.randn(4), "two": np.linspace(1, 4, 4), "three": ['zhangsan', '李四', 999, 0.1]}
# df = pd.DataFrame(data, index=[1, 2, 3, 4])
#
# print(df)
# 获取题干数组

# def out_data(re):
#     print(re)
