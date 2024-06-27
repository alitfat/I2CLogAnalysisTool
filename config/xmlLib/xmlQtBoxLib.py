import xml.etree.ElementTree as ET
from xml.etree.ElementTree import  ElementTree,Element
from config.xmlLib.xmlLib import xmlLib
from config.xmlLib.xmlQtLib import xmlQLog
from Lib.FileSysProcess import FileSysProcess

class logList(xmlLib):
    """ log Setting"""
    #ディレクトリ選択ダイアログボックス
    logDirSelectDialogBox:xmlQLog
    #内容無し
    logNoComment:xmlQLog
    #ファイル無し
    logFileNoExsit:xmlQLog
    #パス無し
    logPathNoExsit:xmlQLog
    #文字列に規定対象外
    logIllegalChar:xmlQLog
    #ファイル閉じる
    logFileClose:xmlQLog
    #ファイル上書き
    logFileOverWrite:xmlQLog
    #ファイル拡張子判定
    logFileExtension:xmlQLog
    
    def __init__(self)-> None:
 
        super(logList, self).__init__()

        #ディレクトリ選択ダイアログボックス
        self.logDirSelectDialogBoxName = "logDirSelectDialogBox" 
        self.logDirSelectDialogBoxText = "ディレクトリ選択ダイアログボックス"
        #内容無し
        self.logNoCommentName = "logNoComment" 
        self.logNoCommentText = "__titleName__内容取得失敗\n" + "取得失敗理由:\n" +  "内容無し\n"
        #ファイル無し
        self.logFileNoExsitName = "logFileNoExsit" 
        self.logFileNoExsitText = "__titleName__内容取得失敗\n" + "取得失敗理由:\n" +  "    下記ファイルが存在していない\n    __fileAddress__"
        #パス無し
        self.logPathNoExsitName = "logPathNoExsit" 
        self.logPathNoExsitText =  "__titleName__内容取得失敗\n" + "取得失敗理由:\n" +  "    下記パスが存在していない\n    __PathAddress__"
        #文字列に規定対象外
        self.logIllegalCharName = "logIllegalChar" 
        self.logIllegalCharText =  "__titleName__内容取得失敗\n" + "取得失敗理由:\n" +  "下記文字列に規定対象外の文字が存在します\n        __IllegalChar__"
        #ファイル閉じる
        self.logFileCloseName = "logFileClose" 
        self.logFileCloseText =  "下記のファイルを閉じてください：\n" + "__fileAddress__"
        #ファイル上書き
        self.logFileOverWriteName = "logFileOverWrite" 
        self.logFileOverWriteText =  "下記のファイルが存在しています、上書きを実施しますか：\n" + "    __fileAddress__"

        #ファイル拡張子判定
        self.logFileExtensionName = "logFileExtension" 
        self.logFileExtensionText =  "ファイル内容取得失敗\n" + "内容取得失敗理由:\n" + "下記ファイルの拡張子が間違いました\n" + "    __fileAddress__"

        self.createConfigSetting()
        return

    def createConfigSetting(self) -> None:
        #pdf変換失敗エラー
        self.logDirSelectDialogBox = xmlQLog(self.logDirSelectDialogBoxName, self.logDirSelectDialogBoxText)
        #内容無し
        self.logNoComment = xmlQLog(self.logNoCommentName, self.logNoCommentText)
        #ファイル無し
        self.logFileNoExsit = xmlQLog(self.logFileNoExsitName, self.logFileNoExsitText)
        #パス無し
        self.logPathNoExsit = xmlQLog(self.logPathNoExsitName, self.logPathNoExsitText)
        #文字列に規定対象外
        self.logIllegalChar = xmlQLog(self.logIllegalCharName, self.logIllegalCharText)
        #ファイル閉じる
        self.logFileClose = xmlQLog(self.logFileCloseName, self.logFileCloseText)
        #ファイル上書き
        self.logFileOverWrite = xmlQLog(self.logFileOverWriteName, self.logFileOverWriteText)
        #ファイル拡張子判定
        self.logFileExtension = xmlQLog(self.logFileExtensionName, self.logFileExtensionText)

        return

    def getConfigSetting(self, eleTool:Element|None)-> None:

        if eleTool is None :
            return
        
        #pdf変換失敗エラー
        self.logDirSelectDialogBox.getConfigSetting(eleTool, self.logDirSelectDialogBoxName)
        self.logDirSelectDialogBoxText = self.logDirSelectDialogBox.getLabelValue('Text')
        #内容無し
        self.logNoComment.getConfigSetting(eleTool, self.logNoCommentName)
        self.logNoCommentText = self.logNoComment.getLabelValue('Text')
        #ファイル無し
        self.logFileNoExsit.getConfigSetting(eleTool, self.logFileNoExsitName)
        self.logFileNoExsitText = self.logFileNoExsit.getLabelValue('Text')
        #パス無し
        self.logPathNoExsit.getConfigSetting(eleTool, self.logPathNoExsitName)
        self.logPathNoExsitText = self.logPathNoExsit.getLabelValue('Text')
        #文字列に規定対象外
        self.logIllegalChar.getConfigSetting(eleTool, self.logIllegalCharName)
        self.logIllegalCharText = self.logIllegalChar.getLabelValue('Text')
        #ファイル閉じる
        self.logFileClose.getConfigSetting(eleTool, self.logFileCloseName)
        self.logFileCloseText = self.logFileClose.getLabelValue('Text')
        #ファイル上書き
        self.logFileOverWrite.getConfigSetting(eleTool, self.logFileOverWriteName)
        self.logFileOverWriteText = self.logFileOverWrite.getLabelValue('Text')
        #ファイル拡張子判定
        self.logFileExtension.getConfigSetting(eleTool, self.logFileExtensionName)
        self.logFileExtensionText = self.logFileExtension.getLabelValue('Text')
        return
    
    def outputConfigSetting(self, eleLogList:Element)-> None:
        comment = ET.Comment('ログ内容設定')
        eleLogList.append(comment)
        self.logDirSelectDialogBox.outputConfigSetting(eleLogList)
        self.logNoComment.outputConfigSetting(eleLogList)
        self.logFileNoExsit.outputConfigSetting(eleLogList)
        self.logPathNoExsit.outputConfigSetting(eleLogList)
        self.logIllegalChar.outputConfigSetting(eleLogList)
        self.logFileClose.outputConfigSetting(eleLogList)
        self.logFileOverWrite.outputConfigSetting(eleLogList)
        self.logFileExtension.outputConfigSetting(eleLogList)

        return
       
    def updateConfigSetting(self) -> None:
        return
    
    def updateGUISetting(self) -> None:
        return


class xmlQtBoxLibConfig():
    
    def __init__(self, configFileAddress:str)-> None:
        super(xmlQtBoxLibConfig, self).__init__()
        self.FSP = FileSysProcess()
        self.CreateFileAddress =configFileAddress + "\\QtBoxLib_Jp.xml"
        if self.FSP.judgeFileExsit(self.CreateFileAddress) :
            self.configFileAddress = self.CreateFileAddress
        else :
            self.configFileAddress = configFileAddress + "\\Jp\\QtBoxLib_Jp.xml"

        self.logList = logList()
        return
        
    def getConfigSetting(self) -> bool:
        bResult = True
        try:
            tree = ET.parse(self.configFileAddress)
            elePdfEditerGUI =  tree.getroot()
            eleLogList = elePdfEditerGUI.find("LogList")
            self.logList.getConfigSetting(eleLogList)

        except Exception as e:
            bResult = False  
        return bResult
    
    def outputConfigSetting(self) -> None:
        #ツール名称
        elePdfEditerGUI = Element("xmlQtBoxLib")

        #【LogList】設定
        comment = ET.Comment('【LogList】設定')
        elePdfEditerGUI.append(comment)
        eleLogList = Element("LogList")
        self.logList.outputConfigSetting(eleLogList)
        elePdfEditerGUI.append(eleLogList)

        #ElementTreeを追加
        ET.indent(elePdfEditerGUI)
        tree = ElementTree(elePdfEditerGUI)
    
        #ファイル出力
        with open (self.CreateFileAddress, "wb") as xmlFileAddr :
            tree.write(xmlFileAddr, encoding='utf-8', xml_declaration=True)
        return
        
    def updateConfigSetting(self) -> None:
        self.logList.updateConfigSetting()
        return

    def updateGUISetting(self) -> None:
        self.logList.updateGUISetting()
        return
        
    def print(self)-> None:
        #self.menubarList.print()
        return
        
