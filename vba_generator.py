#!usr/bin/python
#FileName: vba_generator.py

#辅助方法
def write_to_file(file_out, list):
	for str in list:
		file_out.write(str)
		
def delete_useless_chart(str):
	str = str.strip()
	if str == "":
		return ""
	print ("\tfunc delete_useless_chart: %s \n"%(str))
	if (str.find("、") != -1):
		list = str.split("、")
		return list[1]
	else:
		return str
		
def init(file_out):
	print ("\tinit\n")
	list = ["""\tindex_slide =  ActivePresentation.Slides.Count
	If index_slide < 1 then Exit Sub\n""",
	"\tset shapes = ActiveWindow.Selection.SlideRange.Shapes\n"]
	write_to_file(file_out, list)
	
def end_case(file_out):
	write_to_file(file_out, "End Sub\n")	

def begin_case(func_name, file_out):
	list = ["Sub %s()\n"%(func_name), 
			"\twindowsIndex = 1\n",
			"\tApplication.Windows(windowsIndex).Activate\n",
			"\tDim shapes as shapes\n",
			"\tDim index_slide as Integer\n"]
	write_to_file(file_out, list)
	
def open_action(line, file_out):
	print ("\topen_action: %s\n"%(line))
	write_to_file(file_out, "\t\'操作： %s\n"%(line))

def close_action(line, file_out):
	write_to_file(file_out, "\t\'操作： %s\n"%(line))
	write_to_file(file_out, "\tPresentations.Item(windowsIndex).Saved = ksoTrue\n\
\tActivePresentation.Close\n")

def change_action(line, file_out):
	dict = { "母版视图" : """\tActiveWindow.ViewType = ppViewSlideMaster
	index_slide = ActivePresentation.Designs.Count
	If index_slide < 1 then Exit Sub
	set shapes = ActivePresentation.Designs(1).SlideMaster.Shapes\n"""
	}
	
	temp_str = line.split("到")[1]
	str = dict.get(temp_str, "Todo")
	if (str != "Todo"):
		write_to_file(file_out, "\t\'操作：%s\n"%(line))
		write_to_file(file_out, str)
	else:
		write_to_file(file_out, "\t\'备注: %s vba实现不了\n"%(line))
	
def insert_action(line, file_out):
	dict = { "矩形" : """\tshapes.AddShape(msoShapeRectangle, 133.25, 94.25, 232.38, 130.38).Select\n""",
			"线条" : """\tshapes.AddLine(70.88, 338, 428, 394.75).Select\n""",
			"表格" : """\tshapes.AddTable(3, 3).Select\n""",
			"图表" : """\tshapes.AddOLEObject(Left:=120, Top:=110, Width:=480, Height:=320, ClassName:="Et.chart.6", Link:=msoFalse).Select
    With ActiveWindow.Selection.ShapeRange
        .Left = 120
        .Top = 109.875
        .Width = 480
        .Height = 320.25
    End With\n""",
			"艺术字" : """\tshapes.AddTextEffect msoTextEffect15, \"WPS Office\", \"Arial\", 30, ksoFalse, ksoFalse, 0, 0\n""",
			"公式" : """\tActiveWindow.Selection.SlideRange.Shapes.AddOLEObject(Left:=120, Top:=110, Width:=480, Height:=320, ClassName:=\"Equation.KSEE3\", Link:=msoFalse).Select
			\t\'上面只能调出公式编辑窗口，以下流程应该用测试工具保障\n""",
			"符号" : """\tTodo\n"""}
	str = dict.get(line[2:], "Todo")
	if (str != "Todo"):
		write_to_file(file_out, "\t\'操作：%s\n"%(line))
		write_to_file(file_out, str)
	else:
		write_to_file(file_out, "\t\'备注: %s vba实现不了\n"%(line))

def delete_action(line, file_out):
	#print (line[2:])
	if (line[2:] == "输入内容"):
		write_to_file(file_out, "\t\'操作：%s\n"%(line))
		write_to_file(file_out, "\tshapes.Item(1).TextFrame.TextRange.Text = \"\"\n")
	elif (line[2:] != "组织结构图"):
		write_to_file(file_out, "\t\'操作：%s\n"%(line))
		write_to_file(file_out, "\tcount = shapes.count\n\
\tif count > 0 Then shapes(count).Delete\n")
	else:
		write_to_file(file_out,  "\t\'备注: %s vba实现不了\n"%(line))
		
def input_action(line, file_out):
	if (line[2:5] == "占位符" or line[2:4] == "内容"):
		list = line.split("“")
		str_value = list[1][0:-1]
		print (str_value)
		write_to_file(file_out, "\t\'操作：%s\n"%(line))
		write_to_file(file_out, "\tshapes.Item(1).TextFrame.TextRange.Text = \"%s\"\n"%(str_value))
	
def main():
	import sys
	file_name = sys.argv[0]
	file_dir = file_name.split("vba_generator.py")[0]
	
	file_in = open(file_dir + "\delete.txt")
	file_out = open(file_dir + "\delete_case_vba.txt", 'w')
	
	case_index = 0
	first_step = 1
	for str in file_in.readlines():
		str = delete_useless_chart(str)
		print (str)
		if (str.strip() == ""):
			continue
		if (str[0:4] == "Case"):
			if (case_index != 0):
				end_case(file_out)
			first_step = 1
			case_index += 1
			begin_case(str[0:4] + "_%d"%(case_index), file_out)
		else:
			if (first_step == 1 and str[0:2] != "切换"):
				init(file_out)
				first_step = 0
			print (str)
			if (str[0:4] == "打开文档"):
				open_action(str, file_out)
			elif (str[0:4] == "关闭文档"):
				close_action(str, file_out)
			elif (str[0:2] == "切换"):
				change_action(str, file_out)
			elif (str[0:2] == "插入"):
				insert_action(str, file_out)
			elif (str[0:2] == "删除"):
				delete_action(str, file_out)
			elif (str[0:2] == "输入"):
				input_action(str, file_out)
			else:
				print ("\tmain function: to %s\n"%(str))
		
	end_case(file_out)
	case_index += 1
	file_in.close()
	file_out.write("Sub main()\n")
	for i in range(1, case_index):
		file_out.write("\tCase_%d\n"%(i))
	end_case(file_out)
        
	file_out.close()
if (__name__ == "__main__"):
	main()
				
			
			
			
			
			