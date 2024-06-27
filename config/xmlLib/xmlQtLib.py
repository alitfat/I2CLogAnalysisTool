import copy
import xml.etree.ElementTree as ET
from xml.etree.ElementTree import  ElementTree,Element
from PyQt5.QtWidgets import QWidget, QMenu, QAction, QTextEdit, QLabel, QCheckBox, QComboBox, QPushButton, QGroupBox, QRadioButton, QListWidget,QLineEdit
from config.xmlLib.xmlLib import xmlLib, convertStr

class xmlQtLabelValue(xmlLib):
    labelValueList:dict[str, str]= {}
    def __init__(self)-> None:
        self.labelValueList.clear()
        return
    
    def getLabelValueList(self, eleElement:ElementTree|Element, eleFindTagName:str)-> ElementTree|Element|None:
       eleSubElement = self.getEleLabelList(eleElement, eleFindTagName, self.labelValueList)
       return eleSubElement
    
    class getLabelValue(convertStr):
        def __init__(self, labelName:str = "", defaultLabelValue:bool|int|str = "")-> None:
            self.strValue = str(xmlQtLabelValue.labelValueList.get(labelName, defaultLabelValue))
            return

class xmlQt(xmlLib):
    label:dict[str, str]= {}
    qtObj:QWidget|QMenu|QAction|QTextEdit|QLineEdit|QLabel|QPushButton|QCheckBox|QGroupBox|QComboBox|QRadioButton|QListWidget|None

    def __init__(self, labelName = "", WidgetName = "xmlQt", qtObj:QWidget|QMenu|QAction|QTextEdit|QLineEdit|QLabel|QPushButton|QCheckBox|QGroupBox|QComboBox|QRadioButton|QListWidget|None = None)-> None:
        self.label.clear()
        self.label = copy.deepcopy(self.label)
        self.label["Name"] = labelName
        self.label["QtWidget"] = WidgetName
        self.qtObj = qtObj
        return

    def clear(self)-> None:
        self.label.clear()
        return
    
    def addLabelProperty(self, PropertyName:str, PropertyValue:bool|int|str)-> None:
        self.label[PropertyName] = str(PropertyValue)
        return
    
    def delLabelProperty(self, PropertyName:str)-> None:
        del self.label[PropertyName]
        return
      
    def outputLabelProperty(self, eleQt:Element)-> Element:
        eleSubQt = self.addEleLabelList(eleQt, self.label)
        return eleSubQt
    
    def getLabelValue(self, labelProperty:str = "", defaultLabelProperty:bool|int|str = "")->str:
        strValue = str(self.label.get(labelProperty, defaultLabelProperty))
        return strValue
    
    def modifyLabelProperty(self, PropertyName:str, PropertyValue:bool|int|str)-> bool:
        bResult = False
        if PropertyName in self.label:
            self.label[PropertyName] = str(PropertyValue)
            bResult = True
        else:
            print(f'{PropertyName} is not in labelList')
        return bResult

    def outputConfigSetting(self, eleQt:Element,tagName = "")-> Element:
        if tagName == "":
            tagName = self.label["Name"]
        eleQt = self.addSubElement(eleQt,tagName)
        eleSubQt = self.outputLabelProperty(eleQt)
        return eleSubQt
        
    def getConfigSetting(self, eleElement:ElementTree|Element, eleFindTagName:str = "")-> ElementTree|Element|None:
        labellist:dict[str, str]= {}
        if eleFindTagName == "":
            eleFindTagName = self.getLabelValue("Name")
        eleSubElement = self.getEleLabelList(eleElement, eleFindTagName, labellist)
        for labelName, labelValue in labellist.items():
            if labelName in self.label:
                self.label[labelName] = labelValue.replace("\\n","\n")
        return eleSubElement
     
    def updateConfigSetting(self, qtObj:QWidget|QMenu|QAction|QTextEdit|QLineEdit|QLabel|QPushButton|QCheckBox|QGroupBox|QComboBox|QRadioButton|QListWidget|None = None)-> None:
        if qtObj is None:
            qtObj = self.qtObj
        if qtObj is None:
            return
        
        for labelName, labelValue in self.label.items():
            match labelName:
                case "Text":
                    if isinstance(qtObj, QTextEdit) :
                        self.label[labelName] = qtObj.toPlainText().replace("\n","\\n")
                    elif isinstance(qtObj, (QComboBox)) :
                        self.label[labelName] = qtObj.currentText().replace("\n","\\n")
                    else:
                        self.label[labelName] = qtObj.text().replace("\n","\\n")

                case "Title":
                    if isinstance(qtObj, QMenu|QGroupBox) :
                        self.label[labelName] = qtObj.title().replace("\n","\\n")

                case "WindowTitle":
                    if isinstance(qtObj, QWidget) :
                        self.label[labelName] = qtObj.windowTitle().replace("\n","\\n")
                    
                case "Enabled": self.label[labelName] = str(qtObj.isEnabled())
                case "Hidden":
                    if isinstance(qtObj, (QMenu|QTextEdit|QLineEdit|QLabel|QPushButton|QCheckBox|QGroupBox|QComboBox|QRadioButton|QListWidget|QGroupBox)) :
                        self.label[labelName] = str(qtObj.isHidden())
                case "Checked":
                    if isinstance(qtObj, (QCheckBox,QRadioButton)) :
                        self.label[labelName] = str(qtObj.isChecked())
                case _: pass
        return
    
    def updateGUISetting(self, qtObj:QWidget|QMenu|QAction|QTextEdit|QLineEdit|QLabel|QPushButton|QCheckBox|QGroupBox|QComboBox|QRadioButton|QListWidget|None = None)-> None:
        if qtObj is None:
            qtObj = self.qtObj
        if qtObj is None:
            return

        for labelName, labelValue in self.label.items():
            match labelName:
                case "Text":
                    if isinstance(qtObj, QComboBox) :
                        qtObj.setCurrentText(labelValue.replace("\\n","\n"))
                    else:
                        qtObj.setText(labelValue.replace("\\n","\n"))

                case "Title":
                    if isinstance(qtObj, QMenu|QGroupBox):
                        qtObj.setTitle(labelValue)

                case "WindowTitle":
                    if isinstance(qtObj, QWidget) :
                        qtObj.setWindowTitle(labelValue)
                       
                case "Enabled": qtObj.setEnabled(convertStr(labelValue).bool())
                case "Hidden":
                    if isinstance(qtObj, (QMenu|QTextEdit|QLineEdit|QLabel|QPushButton|QCheckBox|QGroupBox|QComboBox|QRadioButton|QListWidget|QGroupBox)) : 
                        qtObj.setHidden(convertStr(labelValue).bool())
                case "Checked":
                    if isinstance(qtObj, QCheckBox|QRadioButton) :
                        qtObj.setChecked(convertStr(labelValue).bool())
                case _: pass
        return

class xmlQLog(xmlQt):
    label:dict[str, str]= {}
    
    def __init__(self,labelName:str = "", text:str = "")-> None:
        super(xmlQLog, self).__init__(labelName, "QLog") 
        self.label["Text"] = text.replace("\n","\\n")
        return

    def outputConfigSetting(self, eleQt:Element,tagName = "")-> Element:
        if tagName == "":
            tagName = self.label["Name"]
        text = self.getLabelValue("Text")
        self.modifyLabelProperty("Text", text.replace("\n","\\n"))
        eleQt = self.addSubElement(eleQt,tagName)
        eleSubQt = self.outputLabelProperty(eleQt)
        return eleSubQt

class xmlQWidget(xmlQt):
    label:dict[str, str]= {}

    def __init__(self, labelName = "", windowTitle = "", qtObj:QWidget|None = None)-> None:
        super(xmlQWidget, self).__init__(labelName, "QWidget", qtObj=qtObj)
        self.label["WindowTitle"] = windowTitle
        return

class xmlQAction(xmlQt):
    label:dict[str, str]= {}

    def __init__(self, labelName = "", text = "", bEnabled = True, qtObj:QAction|None = None)-> None:
        super(xmlQAction, self).__init__(labelName, "QAction", qtObj=qtObj)
        self.label["Text"] = text
        self.label["Enabled"] = str(bEnabled)
        return
    

class xmlQMenu(xmlQt):
    label:dict[str, str]= {}
    qActionList:list[xmlQAction] = []


    def __init__(self,labelName = "", labeltitle = "", bEnabled = True, qtObj:QMenu|None = None)-> None:
        super(xmlQMenu, self).__init__(labelName, "QMenu", qtObj=qtObj)
        self.label["Title"] = labeltitle
        self.label["Enabled"] = str(bEnabled)
        self.qActionList.clear()
        self.qActionList =  copy.deepcopy(self.qActionList)
        return
    
    def clear(self)-> None:
        super(xmlQMenu, self).clear()
        self.qActionList.clear()

    def addQAction(self, qActionName:str, qActionText:str = "")-> None :
        if qActionText == "":
            qActionText = qActionName
        self.qActionList.append(xmlQAction(qActionName, qActionText))
        return
        
    def addQActions(self, qActionTextList:list[str])-> None :
        for index, qActionText in enumerate(qActionTextList):
            self.addQAction("action" + str(index), qActionText)
        return
          
    def outputLabelProperty(self, eleMenu:Element)-> Element:
       super(xmlQMenu, self).outputLabelProperty(eleMenu)
       eleAction = self.__outputQActionList(eleMenu)
       return eleAction
    
    def __outputQActionList(self, eleMenu:Element)-> Element:
        if len(self.qActionList) == 0:
            return eleMenu
        
        for index, qAction in enumerate(self.qActionList):
            eleActionComment = "action" + str(index)
            eleAction = self.addSubElement(eleMenu, eleActionComment)
            qAction.outputLabelProperty(eleAction)
        return eleAction
    
    def getConfigSetting(self, eleElement:ElementTree|Element, eleFindTagName:str = "")-> ElementTree|Element|None:
        if eleFindTagName == "":
            eleFindTagName = self.getLabelValue("Name")
        eleSubElement = super(xmlQMenu, self).getConfigSetting(eleElement, eleFindTagName)
        if eleSubElement is not None :
            for qAction in self.qActionList:
                qAction.getConfigSetting(eleSubElement)
        return eleSubElement

class xmlQMenuAction(xmlQt):
    label:dict[str, str]= {}
    qMenuActionList:list[xmlQMenu|xmlQAction]= []
        
    def __init__(self,labelName = "", labeltitle = "", bEnabled = True, qtObj:QMenu|QAction|None = None)-> None:
        super(xmlQMenuAction, self).__init__(labelName, "QMenu|QAction", qtObj=qtObj)      
        self.label["Title"] = labeltitle
        self.label["Enabled"] = str(bEnabled)
        self.qMenuActionList.clear()
        self.qMenuActionList = copy.deepcopy(self.qMenuActionList)
        return
    
    def clear(self)-> None:
        super(xmlQMenuAction, self).clear()
        self.qMenuActionList.clear()

    def addQMenuAction(self,eMenuAction:xmlQMenu|xmlQAction)-> None :
        self.qMenuActionList.append(copy.deepcopy(eMenuAction))
        return

    def outputLabelProperty(self, eleMenubar:Element, qtIndex:int = 0)-> Element:
        labelName = self.label["Title"]
        comment =  ET.Comment( '【' +  labelName + '】'  + 'メニュー')
        eleMenubar.append(comment)
        eleMenu = self.addSubElement(eleMenubar,"menubar" + str(qtIndex))
        super(xmlQMenuAction, self).outputLabelProperty(eleMenu)
        self.__outputQMenuActionList(eleMenu)
        return eleMenu
    
    def __outputQMenuActionList(self, eleMenu:Element)-> None :

        for index, qMenuAction in enumerate(self.qMenuActionList):
            eleMenuComment = "menu" + str(index)
            eleMenuAction = self.addSubElement(eleMenu, eleMenuComment)
            qMenuAction.outputLabelProperty(eleMenuAction)
        return
    
    def getConfigSetting(self, eleElement:ElementTree|Element, eleFindTagName:str = "")-> ElementTree|Element|None:
        if eleFindTagName == "":
            eleFindTagName = self.getLabelValue("Name")
        eleSubElement = super(xmlQMenuAction, self).getConfigSetting(eleElement, eleFindTagName)
        if eleSubElement is not None :
            for qMenuAction in self.qMenuActionList:
                qMenuAction.getConfigSetting(eleSubElement)

        return eleSubElement

class xmlQMenuBar(xmlQt):
    label:dict[str, str]= {}

    def __init__(self, labelName = "")-> None:
        super(xmlQMenuBar, self).__init__(labelName, "QMenuBar") 
        return
    
    def clear(self)-> None:
        super(xmlQMenuBar, self).clear()


class xmlQRadioButton(xmlQt):
    label:dict[str, str]= {}
    def __init__(self, labelName = "",text = "", bChecked = False, bEnabled = True, bHidden = False, qtObj:QRadioButton|None = None)-> None:
        super(xmlQRadioButton, self).__init__(labelName, "QRadioButton", qtObj=qtObj)
        self.label["Text"] = text
        self.label["Checked"] = str(bChecked)
        self.label["Enabled"] = str(bEnabled)
        self.label["Hidden"] = str(bHidden)
        return

class xmlQPushButton(xmlQt):
    label:dict[str, str]= {}
    
    def __init__(self, labelName = "",text = "", bEnabled = True, bHidden = False, qtObj:QPushButton|None = None)-> None:
        super(xmlQPushButton, self).__init__(labelName, "QPushButton", qtObj=qtObj) 
        self.label["Text"] = text
        self.label["Enabled"] = str(bEnabled)
        self.label["Hidden"] = str(bHidden)
        return


class xmlQTextEdit(xmlQt):
    label:dict[str, str]= {}
    
    def __init__(self,labelName:str = "", text:str = "", bEnabled = True, bHidden = False, qtObj:QTextEdit|None = None)-> None:
        super(xmlQTextEdit, self).__init__(labelName,"QTextEdit", qtObj=qtObj) 
        self.label["Text"] = text
        self.label["Enabled"] = str(bEnabled)
        self.label["Hidden"] = str(bHidden)
        return

class xmlQLineEdit(xmlQt):
    label:dict[str, str]= {}
    
    def __init__(self,labelName:str = "", text:str = "", bEnabled = True, bHidden = False, qtObj:QLineEdit|None = None)-> None:
        super(xmlQLineEdit, self).__init__(labelName,"QLineEdit", qtObj=qtObj) 
        self.label["Text"] = text
        self.label["Enabled"] = str(bEnabled)
        self.label["Hidden"] = str(bHidden)
        return


class xmlQGroupBox(xmlQt):
    label:dict[str, str]= {}
    
    def __init__(self,labelName = "",labeltitle:str = "", bHidden = False, qtObj:QGroupBox|None = None)-> None:
        super(xmlQGroupBox, self).__init__(labelName, "QGroupBox", qtObj=qtObj)
        self.label["Title"] = labeltitle
        self.label["Hidden"] = str(bHidden)
        return
    

class xmlQLabel(xmlQt):
    label:dict[str, str]= {}
    
    def __init__(self,labelName = "", text:str = "", bHidden:bool = False, qtObj:QLabel|None = None)-> None:
        super(xmlQLabel, self).__init__(labelName, "QLabel", qtObj=qtObj) 
        self.label["Text"] = text
        self.label["Hidden"] = str(bHidden)   
        return


class xmlQCheckBox(xmlQt):
    label:dict[str, str]= {}
    
    def __init__(self, labelName = "",text = "", bChecked = False, bEnabled = True, bHidden = False, qtObj:QCheckBox|None = None)-> None:
        super(xmlQCheckBox, self).__init__(labelName, "QCheckBox", qtObj=qtObj) 
        self.label["Text"] = text
        self.label["Checked"] = str(bChecked)
        self.label["Enabled"] = str(bEnabled)
        self.label["Hidden"] = str(bHidden)
        return


class xmlQItemList(xmlLib):
    labelName:str = ""
    ItemList:dict[str, str] = {}

    def __init__(self,labelName = "" )-> None:
        self.labelName = labelName
        self.ItemList.clear()
        self.ItemList = copy.deepcopy(self.ItemList)
        return
    
    def addItem(self, ItemName:str, ItemValue:str = "")-> None:
        self.ItemList[ItemName] = ItemValue
        return
        
    def addItems(self, ItemNameList:dict[str, str])-> None:
        for ItemName, ItemValue in ItemNameList.items():
            self.ItemList[ItemName] = ItemValue
        return
    
    def getItemValueDict(self)->dict[str, str]:
        ItemDataList:dict[str, str] = {}
        for ItemName, ItemValue in self.ItemList.items():
            ItemDataList[ItemValue] = ItemName
        return ItemDataList
    
    def getItemNameDict(self)->dict[str, str]:
        return self.ItemList
        
    def outputEleItems(self, eleItemList:Element)-> None:
        if len(self.ItemList) == 0:
            return
        eleItems =self.addEleLabel(eleItemList, "ItemList")
        for index,  (ItemName, ItemValue) in enumerate(self.ItemList.items()):
            eleItem =self.addEleLabel(eleItems, "Item" + str(index))
            self.addEleLabel(eleItem, "ItemName" , ItemName)
            self.addEleLabel(eleItem, "ItemValue" , ItemValue)
        return
    
    def updateEleItems(self, eleItemList:ElementTree|Element, clearFlg:bool = False)-> None:
        eleItems = self.findElement(eleItemList, "ItemList")
        if eleItems is None:
            return
        
        if clearFlg == False:
            for index, (ItemName, ItemValue) in enumerate(self.ItemList.items()):
                eleItem = self.findElement(eleItems, "Item" + str(index))
                if eleItem is not None:
                    ItemName = self.getEleText(eleItem, "ItemName",  ItemName).str()
                    ItemValue = self.getEleText(eleItem, "ItemValue",  ItemValue).str()
                    self.ItemList[ItemName] = ItemValue
            return
        
        index = 0
        self.ItemList.clear()
        while 1:
            eleItem = self.findElement(eleItems, "Item" + str(index))
            if eleItem is None:
                break
            ItemName = self.getEleText(eleItem, "ItemName").str()
            ItemValue = self.getEleText(eleItem, "ItemValue").str()
            self.ItemList[ItemName] = ItemValue
            index += 1
        return
    
    def updateItems(self, qtObj:QComboBox|QListWidget)-> None:

        self.ItemList.clear()
        count: int = qtObj.count()
        for itemIndex in range(count) :
            itemData =  qtObj.item(itemIndex)
            if itemData is not None:
                self.ItemList[str(itemIndex)] = itemData.text()
        return

    def updateGUIItems(self, qtObj:QComboBox|QListWidget)-> None:
        if len(self.ItemList) == 0:
            return
        
        qtObj.clear()
        for ItemName, ItemData in self.ItemList.items():
            qtObj.addItem(ItemData)
        return
        
       
class xmlQComboBox(xmlQt, xmlQItemList):
    label:dict[str, str]= {}
    ItemList:dict[str, str] = {}

    def __init__(self,labelName = "", text:str = "", bHidden:bool = True, ItemList:dict[str, str] = {}, qtObj:QComboBox|None = None)-> None:
        if labelName == "":
            return
        super(xmlQComboBox, self).__init__(labelName, "QComboBox", qtObj=qtObj) 
        self.label["Text"] = text
        self.label["Hidden"] = str(bHidden)
        self.ItemList.clear()
        self.ItemList = copy.deepcopy(self.ItemList)
        if len(ItemList) == 0:
            return
        self.addItems(ItemList)
        return
    
    def clear(self)-> None:
        super(xmlQComboBox, self).clear()
        self.ItemList.clear()

    def getConfigSetting(self, eleElement:ElementTree|Element, eleFindTagName:str = "")-> ElementTree|Element|None:
        if eleFindTagName == "":
            eleFindTagName = self.getLabelValue("Name")
        eleSubElement = super(xmlQComboBox, self).getConfigSetting(eleElement, eleFindTagName)
        if eleSubElement is None:
            return eleSubElement
        self.updateEleItems(eleSubElement)
        return eleSubElement

    def outputConfigSetting(self, eleElement:Element,tagName = "")-> Element:
        eleSubQt = super(xmlQComboBox, self).outputConfigSetting(eleElement, tagName)
        eleQt = self.findElement(eleElement,self.label["Name"],eleElement)
        if isinstance(eleQt, Element) :
            self.outputEleItems(eleQt)
        return eleSubQt
    
    def updateGUISetting(self, qtObj:QMenu|QAction|QTextEdit|QLineEdit|QLabel|QPushButton|QCheckBox|QGroupBox|QComboBox|QRadioButton|QListWidget|None = None)-> None:
        if qtObj is None:
            qtObj = self.qtObj
        if qtObj is None:
            return
        if isinstance(qtObj, QComboBox) :
            super(xmlQComboBox, self).updateGUISetting(qtObj)
            self.updateGUIItems(qtObj)
        return


class xmlQListWidget(xmlQt, xmlQItemList):
    label:dict[str, str]= {}
    ItemList:dict[str, str] = {}

    def __init__(self,labelName = "", bHidden:bool = True, ItemList:dict[str, str] = {}, qtObj:QListWidget|None = None)-> None:
        if labelName == "":
            return
        super(xmlQListWidget, self).__init__(labelName, "QListWidget", qtObj=qtObj) 
        self.label["Hidden"] = str(bHidden)
        self.ItemList.clear()
        self.ItemList = copy.deepcopy(self.ItemList)
        if len(ItemList) == 0:
            return
        self.addItems(ItemList)
        return
    
    def clear(self)-> None:
        super(xmlQListWidget, self).clear()
        self.ItemList.clear()

    def getConfigSetting(self, eleElement:ElementTree|Element, eleFindTagName:str = "")-> ElementTree|Element|None:
        if eleFindTagName == "":
            eleFindTagName = self.getLabelValue("Name")
        eleSubElement = super(xmlQListWidget, self).getConfigSetting(eleElement, eleFindTagName)
        if eleSubElement is None:
            return eleSubElement
        self.updateEleItems(eleSubElement, True)
        return eleSubElement

    def outputConfigSetting(self, eleElement:Element,tagName = "")-> Element:
        eleSubQt = super(xmlQListWidget, self).outputConfigSetting(eleElement, tagName)
        eleItemList = self.findElement(eleElement,self.label["Name"],eleElement)
        if isinstance(eleItemList, Element) :
            self.outputEleItems(eleItemList)
        return eleSubQt
    
    def updateConfigSetting(self, qtObj:QWidget|QMenu|QAction|QTextEdit|QLineEdit|QLabel|QPushButton|QCheckBox|QGroupBox|QComboBox|QRadioButton|QListWidget|None = None)-> None:
        super(xmlQListWidget, self).updateConfigSetting(qtObj)
        if isinstance(qtObj, QListWidget) :
            self.updateItems(qtObj)
        return

    def updateGUISetting(self, qtObj:QMenu|QAction|QTextEdit|QLineEdit|QLabel|QPushButton|QCheckBox|QGroupBox|QComboBox|QRadioButton|QListWidget|None = None)-> None:
        if qtObj is None:
            qtObj = self.qtObj
        if qtObj is None:
            return
        if isinstance(qtObj, QListWidget) :
            super(xmlQListWidget, self).updateGUISetting(qtObj)
            self.updateGUIItems(qtObj)
        return



