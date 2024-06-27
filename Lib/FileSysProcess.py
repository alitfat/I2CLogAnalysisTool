import os
import re
import shutil

from tkinter import filedialog
from tkinter import messagebox
from datetime import datetime as dt

class FileSysProcess():

    def __init__(self):
        super(FileSysProcess, self).__init__()
    
    def getDirByFileFullAddr(self, fileFullAddr: str, errLogOutputFlg : bool = False) -> tuple[bool, str]:
        """
        -----------------------------------------------------------------
        文字列から、フォルダ名称取得処理\n
        【引数 】\n
            fileFullAddr :ファイルフルアドレス\n
            errLogOutputFlg:取得失敗の場合、エラーログ出力フラグ\n
        【戻り値】
            bool:実施結果\n
            str:パス文字列\n
        -----------------------------------------------------------------
        """
        bResult:bool = os.path.isdir(fileFullAddr)
        if ( bResult  == True ) :
            return (True, fileFullAddr)
        
        fileFullAddr = os.path.split(fileFullAddr)[0]
        try:
            bResult = os.path.isdir(fileFullAddr)
        except Exception as e:
           if ( errLogOutputFlg  == True ) :
                messagebox.showinfo("文字列処理", "文字列からパス内容取得失敗")
            
        return (bResult, fileFullAddr)



    def getFileNameInfoByFileFullAddr(self, fileFullAddr: str, fileName : list[str]) -> bool:
        """
        -----------------------------------------------------------------
        文字列から、ファイル名称、拡張子取得処理\n
        【引数 】\n
            fileFullAddr :ファイルフルアドレス\n
            fileName[0]  :ファイル名称(拡張子あり)\n
            fileName[1]  :ファイル名称(拡張子なし)\n
            fileName[2]  :拡張子(.をつける)\n
            fileName[3]  :拡張子(.をつけない)\n
        【戻り値】実施結果\n
        -----------------------------------------------------------------
        """
        bResult:bool =  True
        fileName.clear()
        try:
            fileName.append(os.path.basename(fileFullAddr))
            
            splitStr =  os.path.splitext(fileName[0])
            fileName.append(splitStr[0])
            fileName.append(splitStr[1])
            fileName.append(fileName[2].replace('.',''))           
        except Exception as e:
            bResult = False
        return bResult


    def getFileFullAddrByFileDialog(self, filetype:list[tuple], filePath: str) -> str:
        """
        -----------------------------------------------------------------
        文字列から、ファイルダイアグからファイルフルアドレス取得処理\n
        【引数 】\n
            filetypes    :ファイル種類\n
            filePath     :開始指定パス\n
        【戻り値】ファイルフルアドレス\n
        -----------------------------------------------------------------
        """
        
        #ファイルパス指定（絶対パスをPathメソッドでWindowsPath型に変換要）
        fileAddr:str = filedialog.askopenfilename(filetypes = filetype, initialdir = filePath)
        
        if ( fileAddr  == "" ) :
           return fileAddr
        
        fileAddr = fileAddr.replace('/','\\')
        fileAddr = fileAddr.replace('\\\\\\\\','\\')
        return fileAddr


    def getFilePathByPathDialog(self, fullFilePath: str) -> str:
        """
        -----------------------------------------------------------------
        文字列から、ファイルダイアグにて、ファイルフルパス取得処理\n
        【引数 】\n
            fullFilePath:ファイルアドレス文字列\n
        【戻り値】ファイルフルパス\n
        -----------------------------------------------------------------
        """
        
        #ファイルパス指定（絶対パスをPathメソッドでWindowsPath型に変換要）
        filePath:str  = filedialog.askdirectory(initialdir=fullFilePath)
        
        if ( filePath  == "" ) :
           return filePath
        
        filePath = filePath.replace('/','\\')
        filePath = filePath.replace('\\\\\\\\','\\')
        
        return filePath


    def judgeFileExsit(self, fullFileAddr: str) -> bool:
        """
        -----------------------------------------------------------------
        ファイルが存在するかどうかを判断\n
        【引数 】\n
            fullFileAddr:ファイルフルアドレス\n
        【戻り値】実施結果\n
        -----------------------------------------------------------------
        """
        bResult:bool = os.path.isfile(fullFileAddr)  
        return bResult
    
    def judgeFilesExsit(self, fullFilesAddr: list[str], flag:bool|None = None) -> list[tuple[bool, str]]:
        """
        -----------------------------------------------------------------
        ファイルが存在するかどうかを判断\n
        【引数 】\n
            fullFilesAddr:ファイルフルアドレスリスト\n
            flag:フィルター識別子\n
                  None:全ファイルリストの判断結果を戻す\n
                  True:存在しているファイル名称を戻す\n
                  False:存在してないファイル名称を戻す\n
        【戻り値】実施結果\n
        -----------------------------------------------------------------
        """
        bResultList:list[tuple[bool, str]] = []
        for fullFileAddr in fullFilesAddr:
            bResult = self.judgeFileExsit(fullFileAddr)
            if flag == None:
                bResultList.append((bResult, fullFileAddr))
                continue
            if flag == bResult:
                bResultList.append((bResult, fullFileAddr))
                continue

        return bResultList

    def judgeDirExsit(self, fullDirAddr: str) -> bool:
        """
        -----------------------------------------------------------------
        フォルダが存在するかどうかを判断\n
        【引数 】\n
            fullDirAddr:フォルダフルアドレス\n
        【戻り値】実施結果\n
        -----------------------------------------------------------------
        """
        bResult:bool = os.path.isdir(fullDirAddr)
        return bResult

    def JudgeFileExtensionByFileAddr(self, objFullFileAddr : str, extensionName:str= "txt" ) -> bool:
        """
        -----------------------------------------------------------------
        ファイル拡張子判定処理\n
        【引数 】\n
            objFullFileAddr:対象ファイルフルアドレス\n
            extensionName:拡張子\n
        【戻り値】判定結果\n
        -----------------------------------------------------------------
        """

        bResult = True
        objFileName: list[str] = []
        bResult = self.getFileNameInfoByFileFullAddr(objFullFileAddr, objFileName)
        
        if ( bResult  == False ) :
            return bResult
            
        pdfFileType = objFileName[3]
        if ( pdfFileType  != extensionName ) :
            bResult  = False

        return bResult

    def copyfile(self, srtFileAddr: str, dstFileAddr: str) -> None:
        """
        -----------------------------------------------------------------
        ファイルコピー\n
        【引数 】\n
            srtFileAddr:コピー元\n
            dstFileAddr:コピー先\n
        【戻り値】無し\n
        -----------------------------------------------------------------
        """
        shutil.copyfile(srtFileAddr, dstFileAddr)


    def cleanIllegalCharacter(self, strCharacter: str, illegalCharacter: str = r'[\\/:*?"<>|]+')-> str:
        """
        -----------------------------------------------------------------
        違法アルファベット抜き処理\n
        【引数 】\n
            strCharacter:対象アルファベット文字列\n
            illegalCharacter:削除対象アルファベット\n
        【戻り値】削除対象アルファベット抜き後のアルファベット\n
        -----------------------------------------------------------------
        """
        return re.sub(illegalCharacter,'',strCharacter)


    def delFile(self, fileFullAddr: str) -> bool:
        """
        -----------------------------------------------------------------
        ファイル削除処理\n
        【引数 】\n
            fileFullAddr:削除ファイルのフルアドレス\n
        【戻り値】実施結果\n
        -----------------------------------------------------------------
        """
        bResult:bool =  True
        try:
            os.remove(fileFullAddr)
        except Exception as e:
            bResult = False
        
        return bResult


    def getRandomValue(self, valueFormat:str = '%Y%M%D%H%M%S%f') -> str:
        """
        ------------------------------------------
        ランダム値取得処理\n
        【引数 】\n
            valueFormat:出力フォーマット\n
        【戻り値】ランダム値\n
        ------------------------------------------
        """
        tmpStr:str = dt.now().strftime(valueFormat)
        tmpStr = tmpStr.replace('/','')
        return tmpStr
    
    def getFileComment(self, fileFullAddr: str, fileDataList:list[str]) -> bool:
        """
        ------------------------------------------
        ファイル内容取得処理\n
        【引数 】\n
            fileFullAddr:ファイルアドレス\n
            fileDataList:ファイル内容の格納場所\n
        【戻り値】処理結果\n
        ------------------------------------------
        """
        bResult:bool =  True
        try:
            file = open(fileFullAddr, 'r')
            fileDataList.clear()
            datalist = file.readlines()
            for fileLineData in datalist:
                fileDataList.append(fileLineData)
            file.close()
        except Exception as e:
            bResult = False
        return bResult

    def getBinFileComment(self, fileFullAddr: str, startAddr:int = -1, MemorySize:int = -1) -> bytes|None:
        """
        ------------------------------------------
        バイナルファイル内容取得処理\n
        【引数 】\n
            fileFullAddr:ファイルアドレス\n
            fileDataList:ファイル内容の格納場所\n
            startAddr:開始アドレス\n
            MemorySize:メモリサイズ\n
        【戻り値】処理結果\n
        ------------------------------------------
        """
        try:
            file = open(fileFullAddr, 'rb')
            if startAddr == -1 and MemorySize == -1 :
                fileData = file.read()
            else:
                fileData = file.read(startAddr + MemorySize + 1)
                fileData = fileData[startAddr:startAddr + MemorySize]
            file.close()
        except Exception as e:
            fileData = None
        return fileData

    def writeFileComment(self, fileFullAddr: str, fileDataList:list[str]) -> bool:
        """
        ------------------------------------------
        ファイル内容埋込処理\n
        【引数 】\n
            fileFullAddr:ファイルアドレス\n
            dataList:ファイル内容の格納場所\n
        【戻り値】処理結果\n
        ------------------------------------------
        """
        bResult:bool =  True
        try:
            file = open(fileFullAddr, 'w')
            file.writelines(fileDataList)
            file.close()
        except Exception as e:
            bResult = False
        return bResult

