"""
HUXCODE REVIEW RECORD TOOL
Copyright (c) 2018 Melon, Lingfo Pty Ltd
This module is designed by meloneat github: and released under a
MIT licence.
"""

import sublime,sublime_plugin
import os,sys

sys.path.append(os.path.dirname(__file__))
import xlwt 
import xlrd
import datetime,platform
from xlutils.copy import copy;

# Time and more:
today_date = datetime.date.today()

# Global const
the_view_ident = 0 #
the_input_panel_title = "HUX Code Review:"
the_input_panel_defaulttext = ""
point_line_text = "Nothing to write"
main_dir = "huCode"
main_file = "Recording_default"
sheet_name = "CodeReview"
row0 = [u'Number',u'Selection',u'ReviewContext',u'FileDirection',u'Row']
# tool function: 
# 1. 
def mkDir(file_dir):
	saveProDir = os.path.join(findRootPath(),main_dir)
	print(saveProDir)
	if not os.path.exists(saveProDir):
		os.mkdir(saveProDir)
		print('目录创造成功！')
	else:
		print('目录已存在--')
	pass

# 2.
def getSeparator():
    if 'Windows' in platform.system():
        separator = '\\'
    else:
        separator = '/'
    return separator

# 3.
def findRootPath(filedir = os.getcwd()):
    o_path = filedir
    separator = getSeparator()
    str_path = o_path
    str_path = str(str_path.split(separator)[0])
    return str_path+separator
    pass

# 4.
def currentDir():
	return os.path.abspath(__file__)
	pass
  
class CoderReviewCommand(sublime_plugin.TextCommand):
	"""TextCommands在每个视图中实例化一次。View对象可以通过self.view获取"""
	"""docstring for HuCodeCommand"""
	
	def run(self, edit):

		self.lineGet = self.selectLine(self.view)
		self.selectGet =  self.selectContext(self.view)
		self.openInputPanel()
		pass

	def confirmPoint(self):
		for region in self.view.sel():
			print(region)
			if region.empty():
				line = self.view.line(region)
				line_contents = self.view.substr(line)
				self.pointLine = line_contents
				print('confirmPoint:',line_contents)
		pass

	def openInputPanel(self,title=the_input_panel_title):
		self.view.window().show_input_panel(
			caption=title,
      		initial_text=the_input_panel_defaulttext,
      		on_done=self.onDone,
      		on_change=self.onChange,
      		on_cancel=self.onCancel
      	)
		pass

	def selectLine(self,viewSelf):
		currentPointBlock = viewSelf.sel()[0]
		currentLine = 0
		allBlock = viewSelf.find_all(".*")
		print(currentPointBlock.a)
		for allB in allBlock:
			currentLine += 1
			if allB.a <= currentPointBlock.a and allB.b >= currentPointBlock.b:
				break
		return currentLine
		pass

	def selectContext(self,viewSelf):
		currSel = viewSelf.sel()[0]
		return viewSelf.substr(currSel) 
		pass

	def onDone(self,val):
		self.inputvalue = val
		# infoArr : structor [value,line,select]
		infoArr = [self.inputvalue,self.lineGet,self.selectGet]
		exo = execlAbout(sublime_plugin.TextCommand)
		realBook = exo.createBook(infoArr)
		print('done!',val)
		pass

	def onChange(self,a):
		print('change!',a)
		pass

	def onCancel(self):
		pass

class execlAbout(object):
	"""docstring for execlAbout"""
	def __init__(self, arg):
		self.sheetname = sheet_name
		mkDir(main_dir)
		self.arg = arg

	def createBook(self,getArr,filename=main_file):
		self.fileAndDir = os.path.join(findRootPath(),main_dir)+getSeparator()+str(today_date)+filename+'.et'
		self.revText = getArr[0]
		self.selLine  = getArr[1]
		self.selText  = getArr[2]
		
		line = '1000'
		if not os.path.exists(self.fileAndDir):
			print("建立新的！")
			wbook = xlwt.Workbook()
			# cell_overwrite_ok =True是为了能对同一个单元格重复操作。False反之
			sheet = wbook.add_sheet(self.sheetname,cell_overwrite_ok=True)
			row1 = [1,self.selText,self.revText,currentDir(),self.selLine]
			# init the execl style
			for i in range(0,len(row0)):
				# from 00 row
				sheet.write(0,i,row0[i],self.setStyle('Times New Roman',320,True))
				# from 01 row
				sheet.write(1,i,row1[i],self.setStyle('Times New Roman',250,False))
			wbook.save(self.fileAndDir)
		else:
			print("文件已经存在！")
			#sheet = wbook.get_sheet("Sheet One")
			self.loadSheet(self.fileAndDir)
		pass

	def loadSheet(self,filen):
		
		wbook = xlrd.open_workbook(filen)
		# get the nrows
		sheet = wbook.sheet_by_name(sheet_name);
		print('name',sheet.nrows)
		rows = sheet.nrows

		# write mode 
		newWb = copy(wbook);
		sheetWriten = newWb.get_sheet(sheet_name);
		rownew = [rows,self.selText,self.revText,currentDir(),self.selLine]
		# init the execl style
		for i in range(0,len(rownew)):
			sheetWriten.write(0,i,row0[i],self.setStyle('Times New Roman',320,True))
			sheetWriten.write(rows,i,rownew[i])

		newWb.save(self.fileAndDir)
		
		pass

	def setSepareStyle(self,str):
		return xlwt.easyxf(str)
		pass

	def setStyle(self,name,height,bold=False,refresh=True):  

		style = xlwt.XFStyle() 	# 初始化样式  
		font = xlwt.Font() 		# 创建字体  
		font.name = name  
		font.bold = bold  
		font.color_index = 4  
		font.height = height  

		borders= xlwt.Borders()  
		borders.left= 6  
		borders.right= 6  
		borders.top= 6  
		borders.bottom= 6  

		style.font = font
		style.borders = borders  

		return style 
		