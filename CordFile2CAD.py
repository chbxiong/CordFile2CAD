#!/usr/bin/env python
# -*- coding: utf-8 -*-
# 作者：chbxiong
# 单位：常德市国土资源规划测绘院
# 开发时间：2020-6-7 14:18
# 文件名称：CordFile2CAD.py
# 开发工具：PyCharm


from dxfwrite import DXFEngine as dxf
import xlrd

xlsFilePath = "坐标.xlsx"

data = xlrd.open_workbook(xlsFilePath)
table = data.sheets()[0]
row_len = table.nrows  # 6行
col_len = table.ncols  # 4列

# 获取value值，传入行/列索引
readed_Data = []
for i in range(1, row_len):  # 首行为标题，不要
    temp = []
    temp.append(table.row(i)[0].value)
    temp.append(table.row(i)[1].value)
    temp.append(table.row(i)[2].value)
    temp.append(table.row(i)[3].value)

    readed_Data.append(temp)

# print(readed_Data)  # [[A1,1,x1,y1],[A2,1,x2,y2],[A3,1,x3,y3],[B1,2,x4,y4],……]

group = []
groupNum = 1
totalGroups = []
for j in range(0, len(readed_Data)):  # 0至4
    if readed_Data[j][1] == groupNum:
        group.append(readed_Data[j])
        #print(group)
    else:
        m1 = group[:]
        totalGroups.append(m1)
        #print("test1=", totalGroups)
        groupNum = groupNum + 1
        group.clear()
        group.append(readed_Data[j])

totalGroups.append(group)

print(totalGroups)
print(len(totalGroups))

for i in range(0, len(totalGroups)):   # 循环的次数为分组的个数
    if len(totalGroups[i]) >= 3:   # 如果某个分组只有两个点，则不形成闭合图形
        totalGroups[i].append(totalGroups[i][0])   # 如果每个分组的点数不少于3个，则最后增加首个节点，以形成闭合图形


print(totalGroups)



drawing = dxf.drawing('myDwg.dxf')
drawing.add_layer('红线', color=1)  # 3：绿色，2：黄色，1：红色，7：白色。

xyData = ()
for i in range(0, len(totalGroups)):
    points = []
    for j in range(0, len(totalGroups[i])):
        xyData = (totalGroups[i][j][3], totalGroups[i][j][2])
        points.append(xyData)
    drawing.add(dxf.polyline(points, color=1, layer='红线'))


print(points)



groupCord = []
for i in range(0, len(totalGroups)):
    if len(totalGroups[i]) < 3:  # 如果只有两个点，则求中点
        yyy = (totalGroups[i][0][3] + totalGroups[i][1][3]) / 2
        xxx = (totalGroups[i][0][2] + totalGroups[i][1][2]) / 2
        tempAA = (yyy, xxx)
        groupCord.append(tempAA)
    else:   # 如果点数不小于3个，则可构成闭合图形，循环求几何中心
        x_tatal = 0
        y_tatal = 0
        for j in range(0,len(totalGroups[i])-1):
            x_tatal += totalGroups[i][j][2]
            y_tatal += totalGroups[i][j][3]

        x_ave = x_tatal / (len(totalGroups[i]) - 1)
        y_ave = y_tatal / (len(totalGroups[i]) - 1)
        tempBB = (y_ave, x_ave)
        groupCord.append(tempBB)


drawing.add_layer('分组编号', color=1)  # 3：绿色，2：黄色，1：红色，7：白色
for i in range(0,len(totalGroups)):
    drawing.add(dxf.text("第 " + str(i+1) + " 分组", insert=groupCord[i], layer="分组编号"))



# drawing.add(dxf.polyline(
#     points=[(496942.225, 3214446.719), (497021.432, 3214447.495), (497021.434, 3214370.677), (496943.968, 3214369.230),
#             (496942.225, 3214446.719)], color=1, layer='红线'))




drawing.save()

