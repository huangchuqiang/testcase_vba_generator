#!usr/bin/python
#FileName : delete_case.py

def writeToFile(fileOut, stringList):
	for string in stringList:
		fileOut.write(string)
		
def canDo(actionStr, fileOut):
	fileOut.write("\t\'操作：%s\n"%(actionStr))

def canNotDo(actionStr, fileOut):
	fileOut.write("\t\'备注：%s vba实现不了\n"%(actionStr))
	
def getOperate(actionStr):
	actionStr = actionStr.strip()
	if (actionStr == ""):
		return ""
	
	if (actionStr.find("、") != -1):
		myList = actionStr.split ('、')
		return myList[1]
	else:
		return actionStr[0:4]
		
		
def beginFunc(funcName, fileOut):
	writeToFile(fileOut, "Sub %s()\n"%(funcName))
	writeToFile(fileOut, "\twindowsIndex = 1\n\tApplication.Windows(windowsIndex).Activate\n")
	
def endFunc(fileOut):
	writeToFile(fileOut, "End Sub\n")
	
def openAction(actionStr, fileOut):
	canDo(actionStr, fileOut)

def closeAction(actionStr, fileOut):
	canDo(actionStr, fileOut)
	writeToFile(fileOut, "\tPresentations.Item(windowsIndex).Saved = ksoTrue\n\
\tActivePresentation.Close\n")

def noAction(fileOut, actionStr):
	fileOut.write("\t还没有实现\n")
	print ("\t还没有实现\n")
	
def operateMaster(file, fileOut, actionStr):

	return getLine(file)
	
def inputAction(actionStr, fileOut):
	if (actionStr[0:3] == "占位符" or actionStr[0:2] == "内容"):
		list = actionStr.split("“")
		str_value = list[1][0:-1]
		print (str_value)
		writeToFile(fileOut, "\tIf ActivePresentation.Slides.Count > 0 Then\n\
\t\tIf ActivePresentation.Slides(1).Shapes.Count > 0 Then\n\
\t\t\tActivePresentation.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange.Text = \"%s\"\n\
\t\tEnd If\n\
\tEnd If\n"%(str_value))

def insertAction(actionStr, fileOut):
	print (actionStr)
	dict = {
	"矩形" : "\tActiveWindow.Selection.SlideRange.Shapes.AddShape(msoShapeRectangle, 133.25, 94.25, 232.38, 130.38).Select\n",
	"线条" : "\tActiveWindow.Selection.SlideRange.Shapes.AddLine(70.88, 338, 428, 394.75).Select\n",
	"表格" : "\tActiveWindow.Selection.SlideRange.Shapes.AddTable(3, 3).Select\n",
	"图表" : """\tActiveWindow.Selection.SlideRange.Shapes.AddOLEObject(Left:=120, Top:=110, Width:=480, Height:=320, ClassName:="Et.chart.6", Link:=msoFalse).Select
    With ActiveWindow.Selection.ShapeRange
        .Left = 120
        .Top = 109.875
        .Width = 480
        .Height = 320.25
    End With\n""",
	"符号" : "\t\'没有找到相关的API\n",
	"艺术字" : "\tActiveWindow.Selection.SlideRange.Shapes.AddTextEffect msoTextEffect15, \"WPS Office\", \"Arial\", 30, ksoFalse, ksoFalse, 0, 0\n",
	"组织结构图" : "\tapi调用有问题\n",
	"公式" : "\tActiveWindow.Selection.SlideRange.Shapes.AddOLEObject(Left:=120, Top:=110, Width:=480, Height:=320, ClassName:=\"Equation.KSEE3\", Link:=msoFalse).Select\n\
\t\'上面只能调出公式编辑窗口，以下流程应该用测试工具保障\n"}
	
	writeToFile(fileOut, dict.get(actionStr.strip(), "\tTodo\n"))
	
	
def deleteAction(actionStr, fileOut):
	if(actionStr == "输入内容"):
		writeToFile(fileOut, "\tActivePresentation.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange.Text = \"\"\n")
	
def operateNormal(file, fileOut, actionStr):
	canDo(actionStr, fileOut)
	dict = {
	"插入" : insertAction,
	"输入" : inputAction,
	"删除" : deleteAction
	}
	print(actionStr[2:])
	dict.get(actionStr[0:2], noAction)(actionStr[2:], fileOut)	
	return getLine(file)
	
def getLine(file):
	return file.readline()	
	
def main():
	import sys
	fileName = sys.argv[0]
	fileDir = fileName.split("delete_case.py")[0]

	file = open(fileDir + "\delete1.txt")
	fileOut = open(fileDir + "\delete_case.txt", 'w')
	index = 0
	line = getLine(file)
	while(line):
		line = getOperate(line)
		if (line == "Case"):
			if (index != 0):
				endFunc(fileOut)
			index += 1
			beginFunc(line + "_%d"%(index), fileOut)
			#print (line + "_%d"%(index))
			line = getLine(file)			
		else:
			print (line)
			if (line[0:2] == "切换"):
				line = operateMaster(file, fileOut, line)
			elif (line[0:4] == "打开文档"):
				openAction(line, fileOut)
				line = getLine(file)
			elif (line[0:4] == "关闭文档"):
				closeAction(line, fileOut)
				line = getLine(file)
			else:
				print ("\tnoramal:\n")
				line = operateNormal(file, fileOut, line)

	endFunc(fileOut)
	index += 1
	file.close()
	fileOut.write("Sub main()\n")
	for i in range(1, index):
		fileOut.write("\tCase%d\n"%(i))
	endFunc(fileOut)
        
	fileOut.close()
if (__name__ == "__main__"):
	main()