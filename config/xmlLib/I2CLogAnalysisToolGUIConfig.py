import xml.etree.ElementTree as ET
from xml.etree.ElementTree import  ElementTree,Element
from PyQt5.QtWidgets import QRadioButton
from config.xmlLib.xmlLib import xmlLib
from config.xmlLib.xmlQtLib import xmlQWidget, xmlQAction, xmlQMenuAction, xmlQRadioButton, xmlQTextEdit,xmlQLabel,xmlQPushButton,xmlQCheckBox
from GUI.I2CLogAnalysisToolGUI import Ui_I2CLogAnalysisToolGUI as ToolUI

from Lib.FileSysProcess import FileSysProcess

class toolNameLib(xmlLib):
    
    #ツール名称
    widTool:xmlQWidget
    """ toolName Setting"""
    def __init__(self)-> None:
 
        super(toolNameLib, self).__init__()

        #ツール名称
        self.toolName = "widTool" 
        self.toolTitle = "I2ClogTool"
        self.toolWindowTitle = ""
        self.createConfigSetting()
        return

    def createConfigSetting(self) -> None:
        #ツール名称
        self.widTool = xmlQWidget(self.toolName, self.toolTitle)
        return

    def getConfigSetting(self, eleTool:Element|None)-> None:

        if eleTool is None :
            return
        
        #ツール名称
        self.widTool.getConfigSetting(eleTool, self.toolName)
        return
    
    def outputConfigSetting(self, eleTool:Element)-> None:
        comment = ET.Comment('ツール名称設定')
        eleTool.append(comment)
        self.widTool.outputConfigSetting(eleTool)
        return
       
    def updateConfigSetting(self, toolUi:ToolUI) -> None:
        return
    
    def updateGUISetting(self, toolUi:ToolUI) -> None:
        toolNameText = self.widTool.getLabelValue('WindowTitle')
        self.toolWindowTitle = (toolNameText)
        return


class menubarList(xmlLib):
    xmlMenuActions :list[xmlQMenuAction] = []
    """ Tool Menu Bar Setting"""
    def __init__(self)-> None:
 
        super(menubarList, self).__init__()

        #【設定】
        self.menuSetting = "menuSetting"
        #【その他】
        self.actionOther = "Other"
        #【ヘルプ】
        self.menuHelp = "help"
        #【I2CLogAnalysisToolについて】
        self.actionAboutDstTool = "About I2CLogAnalysisTool"
        self.clearConfigSetting()
        self.createConfigSetting()
        return

    def clearConfigSetting(self) -> None:
        self.xmlMenuActions.clear()
        return

    def createConfigSetting(self) -> None:
        #【menubar0】【設定】
        xmlMenuAction = xmlQMenuAction(self.menuSetting, self.menuSetting)
        #【menubar0|menu1】【その他】
        xmlAction = xmlQAction("menu0", self.actionOther)
        xmlMenuAction.addQMenuAction(xmlAction)
        self.xmlMenuActions.append(xmlMenuAction)

        #【menubar1】 【ヘルプ】
        xmlMenuAction = xmlQMenuAction(self.menuHelp, self.menuHelp)
        #【menubar1|menu0】 【I2CLogAnalysisToolについて】
        xmlAction = xmlQAction("menu0", self.actionAboutDstTool)
        xmlMenuAction.addQMenuAction(xmlAction)
        self.xmlMenuActions.append(xmlMenuAction)
        return

    def getConfigSetting(self, eleMenubarList:Element|None)-> None:

        if eleMenubarList is None :
            return
        
        #【menubar0】【設定】
        eleMenu = self.xmlMenuActions[0].getConfigSetting(eleMenubarList, "menubar0")
        #【menubar1】【ヘルプ】
        eleMenu = self.xmlMenuActions[1].getConfigSetting(eleMenubarList, "menubar1")
        
        return
    
    def outputConfigSetting(self, eleMenubar:Element)-> Element:
        #【menubar0】 #【設定】
        eleSubMenubar = self.__outputLabelSetting(eleMenubar, 0)
        #【menubar1】 【ヘルプ】
        eleSubMenubar = self.__outputLabelSetting(eleMenubar, 1)
        return eleSubMenubar
       
    def __outputLabelSetting(self, eleMenubar:Element, qtIndex:int = 0) -> Element:
        eleSubMenubar = self.xmlMenuActions[qtIndex].outputLabelProperty(eleMenubar, qtIndex)
        return eleSubMenubar
       
    def updateConfigSetting(self, toolUi:ToolUI ) -> None:

        return
    
    def updateGUISetting(self, toolUi:ToolUI) -> None:
        #【設定】
        self.xmlMenuActions[0].updateGUISetting(toolUi.menuSetting)
        #【設定】【その他】
        qMenuActionList = self.xmlMenuActions[0].qMenuActionList
        qMenuActionList[0].updateGUISetting(toolUi.actionOther)

        #【ヘルプ】
        self.xmlMenuActions[1].updateGUISetting(toolUi.menuHelp)
        #【ヘルプ】【I2Cログ解析Toolについて】
        qMenuActionList = self.xmlMenuActions[1].qMenuActionList
        qMenuActionList[0].updateGUISetting(toolUi.actionAboutDstDecodeTool)
        return
        
    def print(self) -> None:
        return

class xmlQEncryptButton(xmlQRadioButton):
    label:dict[str, str]= {}  
    def __init__(self, labelName = "", text="", bEnabled = True, bHidden = False, strEncrypt = "", strEncrypt2 = "", qtObj:QRadioButton|None = None)-> None:
        super(xmlQEncryptButton, self).__init__(labelName, text, bEnabled, bHidden,qtObj=qtObj)
        self.label["Encrypt"] = strEncrypt
        self.label["Encrypt2"] = strEncrypt2
        return
        
class actionToolList(xmlLib):
    #lbObjFile
    lbObjFile:xmlQLabel
    #txObjFile
    txObjFile:xmlQTextEdit
    #btObjFile
    btObjFile:xmlQPushButton
    
    #lbDstFilePath
    lbDstFilePath:xmlQLabel
    #txDstFilePath
    txDstFilePath:xmlQTextEdit
    #cbDstFilePath
    cbDstFilePath:xmlQCheckBox
    #btDstFilePath
    btDstFilePath:xmlQPushButton
    
    #lbDstFile
    lbDstFile:xmlQLabel
    #txDstFile
    txDstFile:xmlQTextEdit
    #cbDstFile
    cbDstFile:xmlQCheckBox
    #lbDstFileNote
    lbDstFileNote:xmlQLabel
    #cbDstFileOverWrite
    cbDstFileOverWrite:xmlQCheckBox
    #cb2ExcelFile
    cb2ExcelFile:xmlQCheckBox
    
    #btDstFileProcess
    btDstFileProcess:xmlQPushButton

    
    def __init__(self)-> None:
        super(actionToolList, self).__init__()
        #lbObjFile
        self.lbObjFileName = "lbObjFileName" 
        self.lbObjFileText = "ObjectFileAddress"
        self.lbObjFileHidden = False
        
        #txObjFile
        self.txObjFileName = "txObjFileName"
        self.txObjFileText = ""
        self.txObjFileEnabled = True
        self.txObjFileHidden = False

        #btObjFile
        self.btObjFileName = "btObjFileName"
        self.btObjFileText = "select"
        self.btObjFileEnabled = True
        self.btObjFileHidden = False
        
        #lbDstFilePath
        self.lbDstFilePathName = "lbDstFilePath" 
        self.lbDstFilePathText = "PdfFilePath"
        self.lbDstFilePathHidden = False

        #txDstFilePath
        self.txDstFilePathName = "txDstFilePath"
        self.txDstFilePathText = ""
        self.txDstFilePathEnabled = True
        self.txDstFilePathHidden = False

        #cbDstFilePath
        self.cbDstFilePathName = "cbDstFilePath"
        self.cbDstFilePathText = "Handwritting"
        self.cbDstFilePathChecked = False
        self.cbDstFilePathEnabled = True
        self.cbDstFilePathHidden = False

        #btDstFilePath
        self.btDstFilePathName = "btDstFilePath"
        self.btDstFilePathText = "select"
        self.btDstFilePathEnabled = True
        self.btDstFilePathHidden = False
        
        #lbDstFile
        self.lbDstFileName = "lbDstFileName"
        self.lbDstFileText = "PdfFileName"
        self.lbDstFileHidden = False

        #txDstFile
        self.txDstFileName = "txDstFileName"
        self.txDstFileText = ""
        self.txDstFileEnabled = True
        self.txDstFileHidden = False
        
        #cbDstFile
        self.cbDstFileName = "cbDstFileName"
        self.cbDstFileText = "Handwritting"
        self.cbDstFileChecked = False
        self.cbDstFileEnabled = True
        self.cbDstFileHidden = False
        
        #lbDstFileNote
        self.lbDstFileNoteName = "lbDstFileNote" 
        self.lbDstFileNoteText = "※ 👆 Do not input the extension please."
        self.lbDstFileNoteHidden = True
        
        #cbDstFileOverWrite
        self.cbDstFileOverWriteName = "cbDstFileOverWrite"
        self.cbDstFileOverWriteText = "overwrite saving"
        self.cbDstFileOverWriteChecked = False
        self.cbDstFileOverWriteEnabled = True
        self.cbDstFileOverWriteHidden = True
        
        #cb2ExcelFile
        self.cb2ExcelFileName = "cb2ExcelFile"
        self.cb2ExcelFileText = "outputExcelFile"
        self.cb2ExcelFileChecked = False
        self.cb2ExcelFileEnabled = True
        self.cb2ExcelFileHidden = True
        
        #btDstFileProcess
        self.btDstFileProcessName = "btDstFileProcess"
        self.btDstFileProcessText = "ObjFile2pdfFileExecute"
        self.btDstFileProcessEnabled = True
        self.btDstFileProcessHidden = False
        
        self.clearConfigSetting()
        self.createConfigSetting()
        return
        
    def clearConfigSetting(self) -> None:
        return

    def createConfigSetting(self) -> None:
        #lbObjFile
        self.lbObjFile = xmlQLabel(self.lbObjFileName,
                             self.lbObjFileText,
                             self.lbObjFileHidden)
        
        #txObjFile
        self.txObjFile = xmlQTextEdit(self.txObjFileName,
                                   self.txObjFileText,
                                   self.txObjFileEnabled,
                                   self.txObjFileHidden
                                   )

        #btObjFile
        self.btObjFile = xmlQPushButton(self.btObjFileName, 
                                       self.btObjFileText, 
                                       self.btObjFileEnabled,  
                                       self.btObjFileHidden)

        #lbDstFilePath
        self.lbDstFilePath = xmlQLabel(self.lbDstFilePathName, 
                             self.lbDstFilePathText, 
                             self.lbDstFilePathHidden)

        #txDstFilePath
        self.txDstFilePath = xmlQTextEdit(self.txDstFilePathName, 
                                   self.txDstFilePathText,
                                   self.txDstFilePathEnabled,
                                   self.txDstFilePathHidden
                                   )
        
        #cbDstFilePath
        self.cbDstFilePath = xmlQCheckBox(self.cbDstFilePathName, 
                                   self.cbDstFilePathText,
                                   self.cbDstFilePathChecked,
                                   self.cbDstFilePathEnabled,
                                   self.cbDstFilePathHidden)
        
        #btDstFilePath
        self.btDstFilePath = xmlQPushButton(self.btDstFilePathName, 
                                       self.btDstFilePathText, 
                                       self.btDstFilePathEnabled,  
                                       self.btDstFilePathHidden)
        
        #lbDstFile
        self.lbDstFile = xmlQLabel(self.lbDstFileName, 
                             self.lbDstFileText, 
                             self.lbDstFileHidden)

        #txDstFile
        self.txDstFile = xmlQTextEdit(self.txDstFileName, 
                                   self.txDstFileText,
                                   self.txDstFileEnabled,
                                   self.txDstFileHidden)
        
        #cbDstFile
        self.cbDstFile = xmlQCheckBox(self.cbDstFileName, 
                                   self.cbDstFileText,
                                   self.cbDstFileChecked,
                                   self.cbDstFileEnabled,
                                   self.cbDstFileHidden)
        
        
        #lbDstFileNote
        self.lbDstFileNote = xmlQLabel(self.lbDstFileNoteName, 
                             self.lbDstFileNoteText, 
                             self.lbDstFileNoteHidden)
        
        
        #cbDstFileOverWrite
        self.cbDstFileOverWrite = xmlQCheckBox(self.cbDstFileOverWriteName,
                                   self.cbDstFileOverWriteText,
                                   self.cbDstFileOverWriteChecked,
                                   self.cbDstFileOverWriteEnabled,
                                   self.cbDstFileOverWriteHidden)
        
        
        #cb2ExcelFile
        self.cb2ExcelFile = xmlQCheckBox(self.cb2ExcelFileName,
                                         self.cb2ExcelFileText,
                                         self.cb2ExcelFileChecked,
                                         self.cb2ExcelFileEnabled,
                                         self.cb2ExcelFileHidden)
        
        #btDstFileProcess
        self.btDstFileProcess = xmlQPushButton(self.btDstFileProcessName,
                                       self.btDstFileProcessText, 
                                       self.btDstFileProcessEnabled,  
                                       self.btDstFileProcessHidden)
        
        return
        
        
    def getConfigSetting(self, eleActionTool:Element|None)-> None:
        if eleActionTool is None :
            return
        
        #lbObjFile
        self.lbObjFile.getConfigSetting(eleActionTool,self.lbObjFileName)
        #txObjFile
        self.txObjFile.getConfigSetting(eleActionTool,self.txObjFileName)
        #btObjFile
        self.btObjFile.getConfigSetting(eleActionTool,self.btObjFileName)
        
        #lbDstFilePath
        self.lbDstFilePath.getConfigSetting(eleActionTool,self.lbDstFilePathName)
        #txDstFilePath
        self.txDstFilePath.getConfigSetting(eleActionTool,self.txDstFilePathName)
        #cbDstFilePath
        self.cbDstFilePath.getConfigSetting(eleActionTool,self.cbDstFilePathName)
        #btDstFilePath
        self.btDstFilePath.getConfigSetting(eleActionTool,self.btDstFilePathName)
        
        #lbDstFile
        self.lbDstFile.getConfigSetting(eleActionTool,self.lbDstFileName)
        #txDstFile
        self.txDstFile.getConfigSetting(eleActionTool,self.txDstFileName)
        #cbDstFile
        self.cbDstFile.getConfigSetting(eleActionTool,self.cbDstFileName)
        #lbDstFileNote
        self.lbDstFileNote.getConfigSetting(eleActionTool,self.lbDstFileNoteName)
        #cbDstFileOverWrite
        self.cbDstFileOverWrite.getConfigSetting(eleActionTool,self.cbDstFileOverWriteName)
        #cb2ExcelFile
        self.cb2ExcelFile.getConfigSetting(eleActionTool,self.cb2ExcelFileName)
        
        #btDstFileProcess
        self.btDstFileProcess.getConfigSetting(eleActionTool,self.btDstFileProcessName)

        return
        
    def outputConfigSetting(self, eleActionTool:Element) -> None:

        #lbObjFile
        self.lbObjFile.outputConfigSetting(eleActionTool)
        #txObjFile
        self.txObjFile.outputConfigSetting(eleActionTool)
        #btObjFile
        self.btObjFile.outputConfigSetting(eleActionTool)

        #lbDstFilePath
        self.lbDstFilePath.outputConfigSetting(eleActionTool)
        #txDstFilePath
        self.txDstFilePath.outputConfigSetting(eleActionTool)
        #cbDstFilePath
        self.cbDstFilePath.outputConfigSetting(eleActionTool)
        #btDstFilePath
        self.btDstFilePath.outputConfigSetting(eleActionTool)

        #lbDstFile
        self.lbDstFile.outputConfigSetting(eleActionTool)
        #txDstFile
        self.txDstFile.outputConfigSetting(eleActionTool)
        #cbDstFile
        self.cbDstFile.outputConfigSetting(eleActionTool)
        #lbDstFileNote
        self.lbDstFileNote.outputConfigSetting(eleActionTool)
        #cbDstFileOverWrite
        self.cbDstFileOverWrite.outputConfigSetting(eleActionTool)
        #cb2ExcelFile
        self.cb2ExcelFile.outputConfigSetting(eleActionTool)
        
        #btDstFileProcess
        self.btDstFileProcess.outputConfigSetting(eleActionTool)
        
        return
    
    def updateConfigSetting(self, toolUi:ToolUI) -> None:
        
        #lbObjFile
        self.lbObjFile.updateConfigSetting(toolUi.lbObjFileName)
        #txObjFile
        self.txObjFile.updateConfigSetting(toolUi.txObjFileName)
        #btObjFile
        self.btObjFile.updateConfigSetting(toolUi.btObjFileName)
        
        #lbDstFilePath
        self.lbDstFilePath.updateConfigSetting(toolUi.lbDstFilePath)
        #txDstFilePath
        self.txDstFilePath.updateConfigSetting(toolUi.txDstFilePath)
        #btDstFilePath
        self.btDstFilePath.updateConfigSetting(toolUi.btDstFilePath)
        #cbDstFilePath
        self.cbDstFilePath.updateConfigSetting(toolUi.cbDstFilePath)
        
        #lbDstFile
        self.lbDstFile.updateConfigSetting(toolUi.lbDstFileName)
        #txDstFile
        self.txDstFile.updateConfigSetting(toolUi.txDstFileName)
        #lbDstFileNote
        self.lbDstFileNote.updateConfigSetting(toolUi.lbDstFileNote)
        #cbDstFile
        self.cbDstFile.updateConfigSetting(toolUi.cbDstFileName)
        #cbDstFileOverWrite
        self.cbDstFileOverWrite.updateConfigSetting(toolUi.cbDstFileOverWrite)
        #cb2ExcelFile
        self.cb2ExcelFile.updateConfigSetting(toolUi.cb2ExcelFile)

        #btDstFileProcess
        self.btDstFileProcess.updateConfigSetting(toolUi.btDstFileProcess)
        
        return
    
    def updateGUISetting(self, toolUi:ToolUI) -> None:
        #lbObjFile
        self.lbObjFile.updateGUISetting(toolUi.lbObjFileName)
        #txObjFile
        self.txObjFile.updateGUISetting(toolUi.txObjFileName)
        #btObjFile
        self.btObjFile.updateGUISetting(toolUi.btObjFileName)

        #lbDstFilePath
        self.lbDstFilePath.updateGUISetting(toolUi.lbDstFilePath)
        #txDstFilePath
        self.txDstFilePath.updateGUISetting(toolUi.txDstFilePath)
        #cbDstFilePath
        self.cbDstFilePath.updateGUISetting(toolUi.cbDstFilePath)
        #btDstFilePath
        self.btDstFilePath.updateGUISetting(toolUi.btDstFilePath)
        
        #lbDstFile
        self.lbDstFile.updateGUISetting(toolUi.lbDstFileName)
        #txDstFile
        self.txDstFile.updateGUISetting(toolUi.txDstFileName)
        #cbDstFile
        self.cbDstFile.updateGUISetting(toolUi.cbDstFileName)
        #lbDstFileNote
        self.lbDstFileNote.updateGUISetting(toolUi.lbDstFileNote)
        #cbDstFileOverWrite
        self.cbDstFileOverWrite.updateGUISetting(toolUi.cbDstFileOverWrite)
        #cb2ExcelFile
        self.cb2ExcelFile.updateGUISetting(toolUi.cb2ExcelFile)
        
        #btDstFileProcess
        self.btDstFileProcess.updateGUISetting(toolUi.btDstFileProcess)

        return
        
    def print(self) :
        return
        
class I2CLogAnalysisToolGUIConfig():
    
    def __init__(self, configFileAddress:str, toolUi:ToolUI)-> None:
        super(I2CLogAnalysisToolGUIConfig, self).__init__()
        self.FSP = FileSysProcess()
        self.CreateFileAddress = configFileAddress + "\\I2CLogAnalysisToolGUI_Jp.xml"
        if self.FSP.judgeFileExsit(self.CreateFileAddress) :
            self.configFileAddress = self.CreateFileAddress
        else :
            self.configFileAddress = configFileAddress + "\\Jp\\I2CLogAnalysisToolGUI_Jp.xml"
            
        self.toolName = toolNameLib()
        self.ui = toolUi
        self.menubarList = menubarList()
        self.actionToolList = actionToolList()
        return
        
    def getConfigSetting(self) -> bool:
        bResult = True
        try:
            tree = ET.parse(self.configFileAddress)
            eleToolGUI =  tree.getroot()
            eleTool = eleToolGUI.find("ToolName")
            self.toolName.getConfigSetting(eleTool)
            eleMenubarList = eleToolGUI.find("menubarList")
            self.menubarList.getConfigSetting(eleMenubarList)
            eleActionToolList = eleToolGUI.find("actionToolList")
            self.actionToolList.getConfigSetting(eleActionToolList)

        except Exception as e:
            bResult = False  
        return bResult
    
    def outputConfigSetting(self) -> None:
        #ツール名称
        eleToolGUI = Element("I2CLogAnalysisTool")

        #ツール名称GUI作成
        eleTool = Element("ToolName")
        self.toolName.outputConfigSetting(eleTool)
        eleToolGUI.append(eleTool)

        #【メニューバー】GUI作成
        comment = ET.Comment('【メニューバー】GUI')
        eleToolGUI.append(comment)
        eleMenubar = Element("menubarList")
        self.menubarList.outputConfigSetting(eleMenubar)
        eleToolGUI.append(eleMenubar)

        #【各pdfツール設定】GUI作成
        comment = ET.Comment('【各I2Cツール設定】GUI')
        eleToolGUI.append(comment)
        eleActionToolList = Element("actionToolList")
        self.actionToolList.outputConfigSetting(eleActionToolList)
        eleToolGUI.append(eleActionToolList)

        #ElementTreeを追加
        ET.indent(eleToolGUI)
        tree = ElementTree(eleToolGUI)
    
        #ファイル出力
        with open (self.CreateFileAddress, "wb") as xmlFileAddr :
            tree.write(xmlFileAddr, encoding='utf-8', xml_declaration=True)
        return
        
    def updateConfigSetting(self) -> None:
        self.toolName.updateConfigSetting(self.ui)
        self.menubarList.updateConfigSetting(self.ui)
        self.actionToolList.updateConfigSetting(self.ui)
        return

    def updateGUISetting(self) -> None:
        self.toolName.updateGUISetting(self.ui)
        self.menubarList.updateGUISetting(self.ui)
        self.actionToolList.updateGUISetting(self.ui)
        return
        
    def print(self)-> None:
        #self.menubarList.print()
        return
        
