import xml.etree.ElementTree as ET
from xml.etree.ElementTree import  ElementTree,Element

class convertStr():
    strValue:str
    def __init__(self, strValue:bool|int|str = "")-> None:
        self.strValue = str(strValue)
        return
        
    def str(self)-> str:
        return self.strValue
        
    def int(self)-> int:
        return int(self.strValue)
        
    def bool(self)-> bool:
        boolValue:bool = False
        if self.strValue.lower() == "true":
            boolValue = True
        return boolValue

class xmlLib(object):
    def __init__(self)-> None:
        super(xmlLib, self).__init__()
        return
    
    def findElement(self, eleTree:ElementTree|Element, eleFindtr:str, eleDefaultTree:ElementTree|Element|None = None) -> ElementTree|Element|None:
        element = eleTree.find(eleFindtr)
        if element is None:
            return eleDefaultTree
        else :
            return element
    
    def judgeEleTagDict(self, element:Element, eleTagDict:dict[str, str])->bool:
        bResult = False
        childData:dict[str, str] = {}
        if len(element.items()) == 0:
            return bResult
        for eleName,eleValue in element.items():
             childData[eleName] = eleValue
        for childName, childValue in eleTagDict.items():
            if childName not in childData:
                return bResult
            if childValue != childData[childName]:
                return bResult
        bResult = True
        return bResult

    class getEleText(convertStr):
        strValue:str
        def __init__(self, eleTree:ElementTree|Element, eleLabelName:str, defalutEleLabelName:bool|int|str = "")-> None:
            element = eleTree.find(eleLabelName)
            if element is None:
                print(f'defalutEleLabelName:{defalutEleLabelName}')
                self.strValue = str(defalutEleLabelName)
            elif element.text is None:
                self.strValue = str(defalutEleLabelName)
            else:
                self.strValue = element.text
            return
    
    def getEleLabelList(self, eleElement:ElementTree|Element, eleFindTagName:str, label:dict[str, str])-> ElementTree|Element|None:
        label.clear()
        eleSubElement = self.findElement(eleElement, eleFindTagName)
        if eleSubElement is None:
            return eleSubElement
        for  iterData in eleSubElement.iter():
            if iterData.tag is None or iterData.text is None:
                continue
            if iterData.text.strip(" ") == "\n":
                continue
            if iterData.tag in label:
                continue
            label[iterData.tag] = iterData.text
        return eleSubElement

    def getEleLabelName(self, eleTree:ElementTree|Element, defalutEleLabelName:str = "")-> str:
        return self.getEleText(eleTree, "Name", defalutEleLabelName).str()
    
    def addEleLabel(self, eleElement:Element, eleLabelName:str = "Name", eleLabelText:str = "", eleTagDict: dict[str, str] = {})-> Element:
        eleSubElement = ET.SubElement(eleElement, eleLabelName, eleTagDict )
        eleSubElement.text = eleLabelText
        return eleSubElement

    def addEleLabelList(self, eleElement:Element, labelList:dict[str, str])-> Element:
        for labelKey, labelData in labelList.items():
            eleSubElement = self.addEleLabel(eleElement, labelKey, labelData)
        return eleSubElement

    def addSubElement(self, eleElement:Element, eleSubElementName:str = "", eleTagDict: dict[str, str] = {})-> Element:
        eleSubElement = ET.SubElement(eleElement, eleSubElementName, eleTagDict)
        return eleSubElement

    def judgeElement(self, eleElement:Element, eleTagNme:str, eleTagDict:dict[str, str]|None = None)->bool:
        if eleElement.tag != eleTagNme :
            return False
        if eleTagDict is None :
            return True
        for eleTagName, eleTagValue in eleTagDict.items():
            eleElementValue = eleElement.get(eleTagName, None)
            if eleElementValue == None:
                return False
            if eleElementValue != eleTagValue:
                return False
        return True

    def removeElement(self, eleElement:Element, eleTagNme:str)-> None:
        for eleSubElement in eleElement:
            if eleSubElement.tag == eleTagNme:
                eleElement.remove(eleSubElement)
                continue
            self.removeElement(eleSubElement, eleTagNme)
        return