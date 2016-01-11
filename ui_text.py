# coding: utf-8


'''
File:   ui_text.py
Author: Luliying
Date:   2016.01.08
'''


import re, os, xlwt


jsPath = "J:\\hipap\\smb_priv\\webpages_eap\\locale\\en_US\\lan.js"
excelPath = "J:\\lan.xls"
fileWrite = "J:\\text.txt"
htmlPath = "J:\\hipap\\smb_priv\\webpages_eap\\pages\\userrpm\\"

jsBegin = re.compile(r'(\S*):\s*\{')
jsEnd = re.compile(r'\s*\},')
jsName = re.compile(r'\"?([\S^\"]*?)\"?\s*:\s*[\"\'].*[\"\'],?')
jsValue = re.compile(r'.*:\s*[\"\'](.*)[\"\'],?')
jsNotWord = re.compile(r'^\W*$')

jsNotText = re.compile(r'^\s*//.*')
jsNotText2 = re.compile(r'^\s*/\*.*\*/\s*$')
jsNotTextBegin = re.compile(r'^\s*/\*.*')
jsNotTextEnd = re.compile(r'.*\*/\s*$')

htmlNotText = re.compile(r'<!--[^-]*-->')
htmlNormalText = re.compile(r'<[^<]*>([^<]*)<[^<]*>')
htmlNormalBlank = re.compile(r'^[\s]*$')
htmlNotWord = re.compile(r'^\W*$')

scriptBegin = re.compile(r'<script')
scriptEnd = re.compile(r'</script')
scriptTextbox = re.compile(r'textbox\("[^"]*"\s*,\s*"([^"]*)"\s*\)')
scriptCombobox = re.compile(r'combobox\("[^"]*"\s*,\s*"([^"]*)"\s*\)')


class writeToExcel:

    def __init__(self):
        self.lanToExcel = {}
        self.htmlToExcel = {}
        self.currentKey = ""
        self.excelLineNum = 0
        self.fileLineNum = 0
        self.isText = True
        self.isScript = False
        self.htmlFileList = []
        
    def pickFromJs(self, jsPath):
        self.currentKey = ""
        self.fileLineNum = 0
        self.isText = True
        print "pick from js: " + jsPath
        fRead = open(jsPath, 'r')
        #fWrite = open(fileWrite, 'w')
        while True:
            lineText = fRead.readline()
            if lineText:
                self.fileLineNum += 1
                if re.findall(jsBegin, lineText):   # 左括号
                    self.currentKey = re.findall(jsBegin, lineText)[0]
                    #print self.currentKey
                    currentDict = {}
                    continue
                if re.findall(jsEnd, lineText): # 右括号
                    #print lineText
                    if "" == self.currentKey:
                        continue
                    else:
                        self.lanToExcel[self.currentKey] = currentDict
                        self.currentKey = ""
                        continue
                if "" == self.currentKey:   # 不在括号对之间
                    continue
                if re.findall(jsNotText, lineText) or re.findall(jsNotText2, lineText):   # 单行注释
                    continue
                if re.findall(jsNotTextBegin, lineText):    # 多行注释起始
                    self.isText = False
                    continue
                if re.findall(jsNotTextEnd, lineText):  # 多行注释结尾
                    self.isText = True
                    continue
                if re.findall(jsName, lineText):  # 内容
                    if False == self.isText:    # 忽略注释
                        continue
                    else:
                        textName = re.findall(jsName, lineText)[0]
                        textValue = re.findall(jsValue, lineText)[0]
                        if re.findall(jsNotWord, textValue):
                            continue    # 忽略空白或纯字符
                        currentDict[textName] = textValue
                        #fWrite.write(lineText)
                        continue
            else:
                break
        print self.fileLineNum, "lines read"

    def searchHtmlFiles(self, htmlPath):
        self.htmlFileList = []
        fileCount = 0
        for root, subdir, files in os.walk(htmlPath):
            for htmlFile in files:
                self.htmlFileList.append(htmlFile)
                fileCount += 1
        print fileCount, "html files found"

    def pickFromHtml(self, htmlPath):
        fWrite = open(fileWrite, 'w')
        print "pick from html files: " + htmlPath
        self.searchHtmlFiles(htmlPath)
        for htmlFile in self.htmlFileList:  # 下面逐个文件处理
            fWrite.write("\n" + htmlFile + "\n")
            self.currentKey = htmlFile
            self.fileLineNum = 0
            self.isText = True
            self.isScript = False
            textCount = 0
            currentDict = {}
            #print htmlFile
            fRead = open(htmlPath + htmlFile, 'r')
            while True:
                lineText = fRead.readline()
                if lineText:
                    self.fileLineNum += 1
                    if re.findall(scriptBegin, lineText):   # 脚本开头
                        self.isScript = True
                        continue
                    if re.findall(scriptEnd, lineText): # 脚本结尾
                        self.isScript = False
                        continue
                    if re.findall(jsNotText, lineText) or re.findall(jsNotText2, lineText) or re.findall(htmlNotText, lineText):
                        continue    # 忽略单行注释
                    if re.findall(jsNotTextBegin, lineText):    # 多行注释起始
                        self.isText = False
                        continue
                    if re.findall(jsNotTextEnd, lineText):  # 多行注释结尾
                        self.isText = True
                        continue
                    if False == self.isScript:  # 处理html
                        if re.findall(htmlNormalText, lineText):    # 普通文本（两对<>之间）
                            if False == self.isText:
                                continue    # 忽略注释
                            else:
                                textValue = re.findall(htmlNormalText, lineText)[0]
                                if re.findall(htmlNotWord, textValue):
                                    continue    # 忽略空白或纯字符
                                values = currentDict.values()
                                if textValue in values:
                                    continue    # 忽略重复
                                textName = "string%d"%textCount   # 自己命名
                                textCount += 1
                                currentDict[textName] = textValue
                                fWrite.write(lineText)
                                continue
                    else:   # 处理脚本
                        if re.findall(scriptTextbox, lineText):
                            if False == self.isText:
                                continue    # 忽略注释
                            else:
                                textValue = re.findall(scriptTextbox, lineText)[0]
                                if re.findall(htmlNotWord, textValue):
                                    continue    # 忽略空白或纯字符
                                values = currentDict.values()
                                if textValue in values:
                                    continue    # 忽略重复
                                textName = "string%d"%textCount   # 自己命名
                                textCount += 1
                                currentDict[textName] = textValue
                                fWrite.write(lineText)
                                continue
                        if re.findall(scriptCombobox, lineText):
                            if False == self.isText:
                                continue    # 忽略注释
                            else:
                                textValue = re.findall(scriptCombobox, lineText)[0]
                                if re.findall(htmlNotWord, textValue):
                                    continue    # 忽略空白或纯字符
                                values = currentDict.values()
                                if textValue in values:
                                    continue    # 忽略重复
                                textName = "string%d"%textCount   # 自己命名
                                textCount += 1
                                currentDict[textName] = textValue
                                fWrite.write(lineText)
                                continue
                else:
                    break
            if 0 != len(currentDict):
                self.htmlToExcel[self.currentKey] = currentDict

    def writeDataToExcel(self):
        self.excelLineNum = 0
        font = xlwt.Font()
        font.name = 'Calibri'
        font.height = 0x00DC
        pattern = xlwt.Pattern()
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = 23 # 背景色
        styleCommon = xlwt.XFStyle()
        styleCommon.font = font
        styleBackground = xlwt.XFStyle()
        styleBackground.font = font
        styleBackground.pattern = pattern
        workBook = xlwt.Workbook(encoding='utf-8')
        sheet = workBook.add_sheet('lan')
        sheet.write(self.excelLineNum, 0, 'name', styleCommon)
        sheet.write(self.excelLineNum, 1, 'value', styleCommon)
        sheet.write(self.excelLineNum, 2, 'translate', styleCommon)
        if 0 == len(self.lanToExcel):
            raise Exception("data is empty !")
        else:
            print "write lan data to excel"
            keys = self.lanToExcel.keys()
            self.excelLineNum += 1
            for key in keys:
                sheet.write_merge(self.excelLineNum, self.excelLineNum, 0, 2, key, styleBackground) # 标题
                self.excelLineNum += 1
                insideKeys = self.lanToExcel[key].keys()
                for insideKey in insideKeys:
                    sheet.write(self.excelLineNum, 0, insideKey, styleCommon)   # name
                    sheet.write(self.excelLineNum, 1, self.lanToExcel[key][insideKey], styleCommon)    # value
                    self.excelLineNum += 1
        if 0 == len(self.htmlToExcel):
            raise Exception("data is empty !")
        else:
            print "write html data to excel"
            sorted(self.htmlToExcel)
            keys = self.htmlToExcel.keys()
            for key in keys:
                sheet.write_merge(self.excelLineNum, self.excelLineNum, 0, 2, key, styleBackground) # 标题
                self.excelLineNum += 1
                insideKeys = self.htmlToExcel[key].keys()
                for insideKey in insideKeys:
                    sheet.write(self.excelLineNum, 0, insideKey, styleCommon)   # name
                    sheet.write(self.excelLineNum, 1, self.htmlToExcel[key][insideKey], styleCommon)    # value
                    self.excelLineNum += 1
        sheet.col(0).width = 10000  # 单元格宽度
        sheet.col(1).width = 30000
        print self.excelLineNum, "lines written"
        workBook.save(excelPath)


def main():
    dataToExcel = writeToExcel()
    dataToExcel.pickFromJs(jsPath)
    dataToExcel.pickFromHtml(htmlPath)
    dataToExcel.writeDataToExcel()


def test():
    testLine = r' | '
    print testLine
    if re.findall(htmlNotWord, testLine):
        resultLine = re.findall(htmlNotWord, testLine)[0]
        print resultLine
    else:
        print "not found !"


#test()
main()
