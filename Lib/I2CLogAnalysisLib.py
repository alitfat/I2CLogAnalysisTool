from threading import Thread
from Lib.FileSysProcess import FileSysProcess
from Lib.ExcelLib import ExcelLib
from Lib.StrLib import StrLib

from openpyxl.styles.borders import Border, Side,BORDER_THIN, BORDER_DOTTED
from openpyxl.styles.colors import BLACK

memoryBorderL1 = Border(top=Side(style=BORDER_THIN, color=BLACK), 
                        bottom=Side(style=BORDER_THIN, color=BLACK), 
                        left=Side(style=BORDER_THIN, color=BLACK),
                        right=Side(style=BORDER_DOTTED, color=BLACK))
memoryBorderL2 = Border(top=Side(style=BORDER_THIN, color=BLACK), 
                        bottom=Side(style=BORDER_THIN, color=BLACK), 
                        left=Side(style=BORDER_DOTTED, color=BLACK),
                        right=Side(style=BORDER_DOTTED, color=BLACK))
memoryBorderL3 = Border(top=Side(style=BORDER_THIN, color=BLACK), 
                        bottom=Side(style=BORDER_THIN, color=BLACK), 
                        left=Side(style=BORDER_DOTTED, color=BLACK),
                        right=Side(style=BORDER_DOTTED, color=BLACK))
memoryBorderL4 = Border(top=Side(style=BORDER_THIN, color=BLACK), 
                        bottom=Side(style=BORDER_THIN, color=BLACK), 
                        left=Side(style=BORDER_DOTTED, color=BLACK),
                        right=Side(style=BORDER_THIN, color=BLACK))
    
class I2CLogAnalysisLib(object):
    
    def __init__(self)->None:
        super(I2CLogAnalysisLib, self).__init__()
        self.FileSysProcess = FileSysProcess()
        self.StrLib = StrLib()
        self.ExcelLib = ExcelLib() 
        return
    
    def I2CLogAnalysisLib_Process(self, srcFileAddr:str, dstFileAddrList:list[str]) -> bool:
        """
        --------------------------------------------------
        I2CLogAnalysisLib処理\n
        【引数 】\n
            srtFileAddr:読取対象ファイル\n
            dstFileAddrList:生成結果対象ファイル\n
        【戻り値】\n
            bool:処理結果\n
        --------------------------------------------------
        """
        srcfileComment:list[str] = []
        dstAddrData:dict[int, list[str]] ={}
        dstMemoryDataList:dict[str, list[str]] ={}
        dstMemoryFileComment:list[str] = []
        dstExcelSheetComment:list[list[str]] = []


        #対象ファイル内容取得処理
        bResult = self.FileSysProcess.getFileComment(srcFileAddr, srcfileComment)
        if bResult == False:
            return bResult
        
        #対象ファイルからメモリ内容取得
        I2CType = self.__analyseSrcData(srcfileComment, dstAddrData)
        if I2CType == -1:
            return False
        
        #出力メモリ内容フォーマット作成
        self.__createMemoryList(dstAddrData, dstMemoryDataList)

        #メモリテキスト出力フォーマット作成
        self.__createMemoryTextFile(dstMemoryDataList, dstMemoryFileComment)
        #メモリテキストファイル出力
        bResult = self.FileSysProcess.writeFileComment(dstFileAddrList[0], dstMemoryFileComment)
        #タイムテキスト出力フォーマット作成
        dstExcelSheetComment.clear()
        dstExcelSheetComment.append(dstMemoryFileComment)
        if len(dstFileAddrList) > 1:
            #Excelファイル出力
            bResult = self.__CreateExcelFile(dstFileAddrList[1], dstExcelSheetComment)
        return bResult
    
    def __analyseSrcData(self,srcfileComment:list[str], dstAddrData:dict[int, list[str]]) -> int:
        analyseStrList:list[list[str]] = []
        I2CType = -1
        #ファイル内容分割処理
        bResult = self.StrLib.splitFileComment(srcfileComment, analyseStrList, ",")
        if bResult == False:
            return I2CType
        while analyseStrList[0][0] == '':
            analyseStrList.pop(0)
        
        if analyseStrList[0][0] == '[ISP CONFIG DUMP]':
            analyseStrList.pop(0)
            I2CType = 1
            bResult = self.__analyseSrcData_ConfigDump(analyseStrList, dstAddrData)
        elif analyseStrList[0][0] == '[Onsemi AR0147 regdump]':
            I2CType = 2
            bResult = self.__analyseSrcData_Regdump(analyseStrList, dstAddrData)
        else:
            pass
        return I2CType

    def __analyseSrcData_ConfigDump(self, analyseStrList:list[list[str]], dstAddrData:dict[int, list[str]]) -> bool:
        
        memroryDataBase:dict[int, list[str]] = {}

        for analyseStr in analyseStrList:
            if len(analyseStr) == 2 :
                hexAddr = analyseStr[0].replace('CONF=','').strip().replace('0x', '').strip()
                strvalue = analyseStr[1].replace('0x', '').strip()
                decAddr = int(hexAddr, 16)
                memroryDataBase[decAddr] = [strvalue]
        self.__createDstAddrData(memroryDataBase, dstAddrData)

        return True

    def __analyseSrcData_Regdump(self,analyseStrList:list[list[str]], dstAddrData:dict[int, list[str]]) -> bool:
        memroryDataBase:dict[int, list[str]] = {}

        for analyseStr in analyseStrList:
            if len(analyseStr) == 2 :
                hexAddr = analyseStr[0].replace('REG=','').strip().replace('0x', '').strip()
                strvalue = analyseStr[1].replace('0x', '').strip()
                decAddr = int(hexAddr, 16)
                memroryDataBase[decAddr] = [strvalue[0:2],strvalue[2:4]]
        self.__createDstAddrData(memroryDataBase, dstAddrData)
        return True

    def __createDstAddrData(self, memroryDataBase:dict[int, list[str]], dstAddrData:dict[int, list[str]]) ->None:
        memroryDataSort:dict[int, list[str]] = {} 
        self.StrLib.sortDict(memroryDataBase, memroryDataSort)

        dstAddrData.clear()
        curBaseAddr = -1
        curEndAddr = -1
        for address, valueList in memroryDataSort.items():
            if address != curEndAddr :
                curBaseAddr = address
                dstAddrData[curBaseAddr] = []
                curEndAddr = address + len(valueList)
            else:
                curEndAddr += len(valueList)
            for value in valueList:
                dstAddrData[curBaseAddr].append(value)  
        return
    
    def __createMemoryList(self,dstResult:dict[int, list[str]],dstMemoryDataList:dict[str, list[str]] ) -> None:
        dstMemoryDataList.clear()
        hexAddress_old = ""
        for decStartAddress, valueList in dstResult.items():
            startIndex = decStartAddress % 16
            curAddress = decStartAddress - startIndex
            hexAddress = format(curAddress, "04X")

            if hexAddress != hexAddress_old:
                dstMemoryDataList[hexAddress] = []
                self.__createXXDataList(dstMemoryDataList[hexAddress])
                hexAddress_old = hexAddress

            valueIndex = 0
            curIndex = startIndex
            valueLength = len(valueList)
            while(valueIndex < valueLength):

                if curIndex < 16:
                    dstMemoryDataList[hexAddress][curIndex] = valueList[valueIndex]
                else:
                    curIndex = 0
                    curAddress += 16
                    hexAddress = format(curAddress, "04X")
                    dstMemoryDataList[hexAddress] = []
                    hexAddress_old = hexAddress
                    self.__createXXDataList(dstMemoryDataList[hexAddress])
                    dstMemoryDataList[hexAddress][curIndex] = valueList[valueIndex]
                curIndex += 1
                valueIndex += 1

        return
    
    def __createXXDataList(self, dataList:list[str])->None:
        dataList.clear()
        for index in range(16):
            dataList.append('xx')
        return
    
    def __createMemoryTextFile(self, dstMemoryDataList:dict[str, list[str]], textFileComment:list[str]) -> None:
        textFileComment.clear()
        #タイトル作成
        lineStr = 'Address\t'
        lineStr += '00\t'+'01\t'+'02\t'+'03\t'+'04\t'+'05\t'+'06\t'+'07\t'
        lineStr += '08\t'+'09\t'+'0A\t'+'0B\t'+'0C\t'+'0D\t'+'0E\t'+'0F\n'
        textFileComment.append(lineStr)
        #メモリデータ内容取得
        for address, valueList in dstMemoryDataList.items():
            lineStr = address + "\t"
            for value in valueList:
                lineStr += value + "\t"
            lineStr =lineStr[:-1]
            lineStr +="\n"
            textFileComment.append(lineStr)
        return
    
    def __CreateExcelFile(self,dstFileAddr:str, dstExcelSheetComment:list[list[str]]) ->bool:
        dstMemoryFileComment = dstExcelSheetComment[0]

        #ExcelFile作成
        bResult = self.ExcelLib.createExcelFile(dstFileAddr)
        bResult = self.ExcelLib.setWorkSheet()
        fileName:list[str] = []
        self.FileSysProcess.getFileNameInfoByFileFullAddr(dstFileAddr,fileName)

        #Memoryシート作成
        self.ExcelLib.modifySheetName(dstSheetName=fileName[1] + '_Memory')
        self.__CreateExcelFile_MemoryDataSheet(dstMemoryFileComment)

        #ExcelFile保存
        self.ExcelLib.save()
        return True
    
    def __CreateExcelFile_MemoryDataSheet(self,dstMemoryFileComment:list[str]) ->None:
        
        rowIndex = 2

        #列幅設定
        self.ExcelLib.setColumnsWidth(2,2, 9)
        self.ExcelLib.setColumnsWidth(3,18, 3.5)

        #タイトル背景色設定
        colorValue = 'B4E6A0'
        self.ExcelLib.setBackGroundColor(2, 2, 18, colorValue)

        #セルデータ入力
        for rowStrComment in dstMemoryFileComment:
            colIndex = 2
            rowStrComment = rowStrComment.replace('\n', '')
            rowStrList = rowStrComment.split('\t')
            for rowStr in rowStrList:
                self.ExcelLib.addCellValue(rowIndex,colIndex, rowStr.strip(" "))
                colIndex += 1
            rowIndex += 1

        #枠線設定
        rowIndex -= 1
        self.ExcelLib.setBorder(2,rowIndex, 2, 2)
        self._SetRowAddrDataBorder(2, rowIndex, 3)
        colorValue = 'C0C0C0'
        self.ExcelLib.setBackGroundColorByCellValue(3, rowIndex, 3, 18,'xx', colorValue)
        return

    def _SetRowAddrDataBorder(self, startRowIndex:int, endRowIndex:int, columnIndex:int) ->None:
        '''
        Excelメモリデータ線描画処理
        '''

        '''
        #マルチスレッド実施
        thread1 = Thread(target=self.__SetAddrDataBorder,args =(startRowIndex, endRowIndex, columnIndex))
        columnIndex += 4
        thread2 = Thread(target=self.__SetAddrDataBorder,args =(startRowIndex, endRowIndex, columnIndex))
        columnIndex += 4
        thread3 = Thread(target=self.__SetAddrDataBorder,args =(startRowIndex, endRowIndex, columnIndex))
        columnIndex += 4
        thread4 = Thread(target=self.__SetAddrDataBorder,args =(startRowIndex, endRowIndex, columnIndex))

        thread1.start()
        thread2.start()
        thread3.start()
        thread4.start()

        thread1.join()
        thread2.join()
        thread3.join()
        thread4.join() 
        '''
        self.__SetAddrDataBorder(startRowIndex, endRowIndex, columnIndex)
        columnIndex += 4
        self.__SetAddrDataBorder(startRowIndex, endRowIndex, columnIndex)
        columnIndex += 4
        self.__SetAddrDataBorder(startRowIndex, endRowIndex, columnIndex)
        columnIndex += 4
        self.__SetAddrDataBorder(startRowIndex, endRowIndex, columnIndex)

        return

    def __SetAddrDataBorder(self, startRowIndex:int, endRowIndex:int, columnIndex:int) ->None:
        rowIndex = endRowIndex
        self.ExcelLib.setBorder(startRowIndex, rowIndex, columnIndex, columnIndex, memoryBorderL1)
        columnIndex += 1
        self.ExcelLib.setBorder(startRowIndex, rowIndex, columnIndex, columnIndex, memoryBorderL2)
        columnIndex += 1
        self.ExcelLib.setBorder(startRowIndex, rowIndex, columnIndex, columnIndex, memoryBorderL3)
        columnIndex += 1
        self.ExcelLib.setBorder(startRowIndex, rowIndex, columnIndex, columnIndex, memoryBorderL4)
        return
