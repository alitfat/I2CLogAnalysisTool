from Lib.FileSysProcess import FileSysProcess

class StrLib(object):
    
    def __init__(self)->None:
        super(StrLib, self).__init__()
        self.FileSysProcess = FileSysProcess()
        return
    
    def splitFileComment(self, srcfileComment:list[str], dstSplitData:list[list[str]], SplitStr=" ") -> bool:
        """
        --------------------------------------------------
        ファイル内容分割処理\n
        【引数 】\n
            srcfileComment:分割対象ファイル内容\n
            dstSplitData:分割対象結果\n
            SplitStr:分割文字列\n
        【戻り値】\n
            bool:処理結果\n
        --------------------------------------------------
        """
        bResult:bool =  True
        try:
            dstSplitData.clear()
            for oneLineComment in srcfileComment:
                dstSplitData.append(self.splitString(oneLineComment, SplitStr))
        except Exception as e:
            bResult = False
        
        return bResult

    def splitString(self, strComment:str, SplitStr=" ") -> list[str]:
        """
        --------------------------------------------------
        文字列内容分割処理\n
        【引数 】\n
            strComment:分割対象文字列\n
            SplitStr:分割文字列\n
        【戻り値】\n
            st:分割後結果\n
        --------------------------------------------------
        """
        tempStr = strComment
        tempStr = tempStr.replace("\n","")
        tempStr = tempStr.replace("\t"," ")
        tempStr = tempStr.strip(" ")
        splitDataList =  tempStr.split(SplitStr)
        strDataList:list[str] = []
        for splitData in splitDataList:
            strDataList.append(splitData.strip(" "))
        return strDataList
    
    def sortDict(self, dictList:dict, resultList:dict)-> None:
        """
        --------------------------------------------------
        dict列順処理(小⇒大)\n
        【引数 】\n
            dictList:対象列\n
            SplitStr:結果列\n
        【戻り値】無し\n
        --------------------------------------------------
        """
        tempList =  dict(sorted(dictList.items(), key=lambda item:item[0]))
        resultList.clear()
        for keyValue, itemValue in tempList.items():
            resultList[keyValue] = itemValue
        return

    