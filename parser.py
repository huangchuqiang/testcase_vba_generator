#!us/bin/python
#FileName: parser.py

def delete_useless_chart(str):
	str = str.strip()
	if str == "":
		return ""
	#print ("\tfunc delete_useless_chart: %s \n"%(str))
	if (str.find("、") != -1):
		list = str.split("、")
		return list[1]
	else:
		return str

def open_action(str, file_out):
	return ""
	
def no_action(str,file_out):
	return ""
	
def main():
	import os
	import sys
	homedir = os.getcwd()
	print (sys.argv[1])
	os.chdir(homedir)
	file_in = open(sys.argv[1])
	file_out = open("parser_" + sys.argv[1], 'w')	
	
	case_index = 1
	for str in file_in.readlines():
		str = delete_useless_chart(str)
		if (str.strip() == ""):
			continue
		print (str)
		if (str[0:4] == "Case"):
			if (case_index == 0):
				file_out.write("end_case\n")
			file_out.write("#%s\n"%(str))
			file_out.write("begin_case\n")
			case_index = 0
		else:
			dict = {
			 "打开" : open_action,
			 #"输入" : input_action,
			 #"插入" : insert_action,
			 #"删除" : delete_action,
			 #"切换" : change_action,
			 #"关闭" : close_action,
			 #"双击" : double_hit_action,
			 #"右键" : right_hit_action,
			 #"缩放" : mini_action,
			 #"复制" : copy_action,
			 #"粘贴" : paste_action
			}
			file_out.write("#%s\n"%(str))
			func = dict.get(str[2:], no_action)
			if (func != no_action):
				func(str[0:2], file_out)
	file_out.write("end_case\n")
	file_in.close()
	file_out.close()
if (__name__ == "__main__"):
	main()
						