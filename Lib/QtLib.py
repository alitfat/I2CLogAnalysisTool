import os

from PyQt5.QtCore import QMimeData, QObject
from PyQt5.QtWidgets import QDialog, QTextEdit, QListWidget, QMessageBox, QWidget, QFileDialog, QAbstractButton
from PyQt5.QtGui import QDropEvent, QFocusEvent,QDragEnterEvent

from Lib.FileSysProcess import FileSysProcess
from config.xmlLib.xmlQtBoxLib import xmlQtBoxLibConfig
        
class qtBoxLib(QDialog):
    
    qMsgButton:QMessageBox.StandardButton
    
    def __init__(self, parent=None):
        super(qtBoxLib, self).__init__(parent)
        self.qMsgButton = QMessageBox.StandardButton.Ok
        self.fsps = FileSysProcess()
        self.msg = QMessageBox()
        xmlQtBoxLibConfigDir = os.path.join(os.getcwd(), "config")
        self.xmlQtBoxLib = xmlQtBoxLibConfig(xmlQtBoxLibConfigDir)
        self.xmlQtBoxLib.getConfigSetting()
        self.xmlQtBoxLib.outputConfigSetting()
        return


    def getFilePathByPathDialog(self, fullFilePath : str, objClass: QWidget|None = None) -> str:
        """
        -----------------------------------------------------------------
        文字列から、ファイルダイアグにて、ファイルフルパス取得処理\n
        【引数 】\n
            fullFilePath:ファイルアドレス文字列\n
            objClass:QWidget対象\n
        【戻り値】ファイルフルパス\n
        -----------------------------------------------------------------
        """
        # ディレクトリ選択ダイアログを表示
        dirSelectDialogBox = self.xmlQtBoxLib.logList.logDirSelectDialogBoxText
        filePath:str = QFileDialog.getExistingDirectory(objClass, dirSelectDialogBox, fullFilePath)
        
        if ( filePath  == "" ) :
           return fullFilePath
        
        filePath = filePath.replace('/','\\')
        filePath = filePath.replace('\\\\\\\\','\\')
        return filePath


    def GetFileNameByFileDialog(self, objFullFileAddr : str,
                                fileDialogTitle : str = "File Dialog Box",
                                fileFilter : str = "All Files( *.*)",
                                objClass: QWidget|None = None) -> str:
        """
        ------------------------------------------
        対象ファイル指定処理\n
        【引数 】\n
            objFullFileAddr:デフォルトファイルアドレス\n
            fileDialogTitle:ファイルダイアグボックスタイトル設定内容\n
            fileFilter:拡張子フィルター\n
            objClass:QWidget対象\n
        【戻り値】\n
            str:ファイルアドレス文字列\n
        ------------------------------------------
        """
        objFilePath = os.getcwd()
        resultList = self.fsps.getDirByFileFullAddr(objFullFileAddr)
        
        if resultList[0] :
            objFilePath = resultList[1]
        
        dialog = QFileDialog()
        dialog.setFileMode(QFileDialog.FileMode.AnyFile)
        fileName = dialog.getOpenFileName(objClass, fileDialogTitle, objFilePath, fileFilter)
        
        fileAddr = fileName[0]
        if fileAddr  == "" :
           return objFullFileAddr
        
        fileAddr = fileAddr.replace('/','\\')
        fileAddr = fileAddr.replace('\\\\\\\\','\\')
        return  fileAddr
    
    
    def GetOfficeFileName(self, objFullFileAddr : str, objClass: QWidget|None = None) -> str:
        """
        ------------------------------------------------------------
        対象Officeファイルアドレス取得処理\n
        【引数 】\n
            objFullFileAddr:デフォルトファイルアドレス\n
            objClass:QWidget対象\n
        【戻り値】\n
            str:ファイルアドレス文字列\n
        ------------------------------------------------------------
        """
        fileDialogTitle = "Select Office File"
        fileFilter = "All Office File(*.ppt *.pptx *.pptm *.doc *.docx  *.docm *.xls *.xlsx *.xlsm);;" 
        fileFilter += "Powerpoint File(*.ppt *.pptx *.pptm);;"
        fileFilter += "Word File(*.doc *.docx *.docm);;"
        fileFilter += "Excel File(*.xls *.xlsx *.xlsm);;"
        fileFilter += "All Files( *.*)"

        return self.GetFileNameByFileDialog( objFullFileAddr, fileDialogTitle, fileFilter, objClass )

    def GeExcelFileName(self, objFullFileAddr : str, objClass: QWidget|None = None) -> str:
        """
        ------------------------------------------------------------
        対象Excelファイルアドレス取得処理\n
        【引数 】\n
            objFullFileAddr:デフォルトファイルアドレス\n
            objClass:QWidget対象\n
        【戻り値】\n
            str:ファイルアドレス文字列\n
        ------------------------------------------------------------
        """
        fileDialogTitle = "Select Excel File"
        fileFilter = "Excel File(*.xls *.xlsx *.xlsm);;"
        fileFilter += "All Files( *.*)"

        return self.GetFileNameByFileDialog( objFullFileAddr, fileDialogTitle, fileFilter, objClass )

    def GetPdfFileName(self, objFullFileAddr : str, objClass: QWidget|None = None) -> str:
        """
        --------------------------------------------------
        対象PDFファイルアドレス取得処理\n
        【引数 】\n
            objFullFileAddr:デフォルトファイルアドレス\n
            objClass:QWidget対象\n
        【戻り値】\n
            str:ファイルアドレス文字列\n
        --------------------------------------------------
        """
        fileDialogTitle = "Select PDF File"
        fileFilter = "PDF File(*.pdf);;"
        fileFilter += "All Files( *.*)"

        return self.GetFileNameByFileDialog( objFullFileAddr, fileDialogTitle, fileFilter, objClass )


    def GetBinFileName(self, objFullFileAddr : str, objClass: QWidget|None = None) -> str:
        """
        --------------------------------------------------
        対象CSVファイルアドレス取得処理\n
        【引数 】\n
            objFullFileAddr:デフォルトファイルアドレス\n
            objClass:QWidget対象\n
        【戻り値】\n
            str:ファイルアドレス文字列\n
        --------------------------------------------------
        """
        fileDialogTitle = "Select Bin File"
        fileFilter = "Bin File(*.bin *.rom);;"
        fileFilter += "All Files( *.*)"

        return self.GetFileNameByFileDialog( objFullFileAddr, fileDialogTitle, fileFilter, objClass )

    def GetCsvFileName(self, objFullFileAddr : str, objClass: QWidget|None = None) -> str:
        """
        --------------------------------------------------
        対象CSVファイルアドレス取得処理\n
        【引数 】\n
            objFullFileAddr:デフォルトファイルアドレス\n
            objClass:QWidget対象\n
        【戻り値】\n
            str:ファイルアドレス文字列\n
        --------------------------------------------------
        """
        fileDialogTitle = "Select CSV File"
        fileFilter = "CSV File(*.csv);;"
        fileFilter += "All Files( *.*)"

        return self.GetFileNameByFileDialog( objFullFileAddr, fileDialogTitle, fileFilter, objClass )

    def GetTextFileName(self, objFullFileAddr : str, objClass: QWidget|None = None) -> str:
        """
        --------------------------------------------------
        対象テキストファイルアドレス取得処理\n
        【引数 】\n
            objFullFileAddr:デフォルトファイルアドレス\n
            objClass:QWidget対象\n
        【戻り値】\n
            str:ファイルアドレス文字列\n
        --------------------------------------------------
        """
        fileDialogTitle = "Select Text File"
        fileFilter = "Text File(*.txt);;"
        fileFilter += "All Files( *.*)"

        return self.GetFileNameByFileDialog( objFullFileAddr, fileDialogTitle, fileFilter, objClass )

    def showMsgBoxInfo(self, msgTitle : str = "MessageBox Info",
                       msgStr : str = "This is a message box",
                       qMsgBoxType : QMessageBox.Icon = QMessageBox.Icon.Information,
                       qMsgButtons : QMessageBox.StandardButton = QMessageBox.StandardButton.Ok) -> None:
        """
        ------------------------------------------------------------------------------
        メッセージボックス出力処理\n
        【引数 】\n
            msgTitle:メッセージボックスタイトル
            msgStr:メッセージボックス出力内容
            qMsgBoxType:メッセージボックスタイプ\n
            qMsgButtons:メッセージボックス表示ボタン\n
        【戻り値】無し\n
            注:メッセージボックス出力結果取得方法:\n
                関数setMsgBoxClickResultの戻り値を使ってください\n
        ------------------------------------------------------------------------------
        """
        self.msg.setIcon(qMsgBoxType)
        self.msg.setWindowTitle(msgTitle)
        self.msg.setText(msgStr)
        
        if qMsgBoxType != QMessageBox.Icon.Question:
            self.msg.setStandardButtons(qMsgButtons)
        else:
            self.msg.setStandardButtons(QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.No)
        self.msg.buttonClicked.connect(self.__setMsgBoxClickResult)
        retval = self.msg.exec_()
        return


    def __setMsgBoxClickResult(self, qMsgButtons:QAbstractButton) -> None:
        """
        ------------------------------------------------------------------------------
        メッセージボックス出力内容保存処理\n
        【引数 】\n
            qMsgButtons:メッセージボックスボタンイベントクラス
        【戻り値】無し\n
            注:メッセージボックス出力結果取得方法:\n
                関数setMsgBoxClickResultの戻り値を使ってください\n
        ------------------------------------------------------------------------------
        """
        buttontext = qMsgButtons.text()
        
        if ( buttontext  == "OK" ) :
            #print("buttontext： ", buttontext, " QMessageBox.Ok：", format(QMessageBox.StandardButton.Ok, '#x'))
            self.qMsgButton = QMessageBox.StandardButton.Ok
            return
        
        if ( buttontext  == "&Yes" ) :
            #print("buttontext： ", buttontext, " QMessageBox.Yes：", format(QMessageBox.StandardButton.Yes, '#x'))
            self.qMsgButton = QMessageBox.StandardButton.Yes
            return
        
        if ( buttontext  == "&No" ) :
            #print("buttontext： ", buttontext, " QMessageBox.No：", format(QMessageBox.StandardButton.No, '#x'))
            self.qMsgButton = QMessageBox.StandardButton.No
            return
        return


    def getMsgBoxClickResult(self) -> bool:
        """
        ------------------------------------------------------------------------------
        メッセージボックス出力処理\n
        【引数 】無し\n
        【戻り値】メッセージボックス出力結果\n
            注:本関数は関数showMsgBoxInfoの実施結果をリターンするので、\n
                関数showMsgBoxInfoをコールする前提です\n
        ------------------------------------------------------------------------------
        """
        bResult = True
        #print( "self.qMsgButton：", format(self.qMsgButton, '#x'))
        if ( self.qMsgButton  == QMessageBox.StandardButton.No ) :
           bResult =  False
        return bResult


    def showQuestionMsgBoxInfo(self, msgTitle : str = "MessageBox Info", msgStr : str = "This is a question box" ) -> None:
        """
        ------------------------------------------------------------------------------
        questionメッセージボックス出力処理\n
        【引数 】\n
            msgTitle:メッセージボックスタイトル
            msgStr:メッセージボックス出力内容
        【戻り値】無し\n
            注:メッセージボックス出力結果取得方法:\n
                関数setMsgBoxClickResultの戻り値を使ってください\n
        ------------------------------------------------------------------------------ 
        """
        self.showMsgBoxInfo(msgTitle, msgStr, QMessageBox.Icon.Question)
        return


    def showErrMsgBoxInfo(self, msgTitle : str = "MessageBox Info", msgStr : str = "This is a error box", 
                          qMsgButtons : QMessageBox.StandardButton = QMessageBox.StandardButton.Ok) -> None:
        """
        ------------------------------------------------------------------------------
        criticalメッセージボックス出力処理\n
        【引数 】\n
            msgTitle:メッセージボックスタイトル
            msgStr:メッセージボックス出力内容
        【戻り値】無し\n
            注:メッセージボックス出力結果取得方法:\n
                関数setMsgBoxClickResultの戻り値を使ってください\n
        ------------------------------------------------------------------------------ 
        """
        self.showMsgBoxInfo(msgTitle, msgStr, QMessageBox.Icon.Critical, qMsgButtons)
        return


    def showStdMsgBoxInfo(self, msgTitle : str = "MessageBox Info", msgStr : str = "This is a message box", 
                          qMsgButtons : QMessageBox.StandardButton = QMessageBox.StandardButton.Ok) -> None:
        """
        ------------------------------------------------------------------------------
        informationメッセージボックス出力処理\n
        【引数 】\n
            msgTitle:メッセージボックスタイトル
            msgStr:メッセージボックス出力内容
        【戻り値】無し\n
            注:メッセージボックス出力結果取得方法:\n
                関数setMsgBoxClickResultの戻り値を使ってください\n
        ------------------------------------------------------------------------------ 
        """
        self.showMsgBoxInfo(msgTitle, msgStr, QMessageBox.Icon.Information, qMsgButtons)
        return


    def judgeFileExsit(self, fileAddressName : str, msgTitle : str = "judge File Exsit", lbName: str = "") -> bool:
        """
        ------------------------------------------------------------------------------
        ファイルアドレス存在判定処理\n
        【引数 】\n
            fileAddressName:ファイルアドレス文字列
            msgTitle:メッセージボックスタイトル
            lbName:ラベル名称
        【戻り値】存在判定結果\n
            注:対象ファイル(fileAddressName)が存在しない場合、エラーメッセージボックス出力\n
        ------------------------------------------------------------------------------ 
        """
        fileAddressName = fileAddressName.strip(" ")
    
        bResult = self.fsps.judgeFileExsit(fileAddressName)
        if ( bResult  == False ) :
            if len(fileAddressName) == 0:
                msgStr = self.xmlQtBoxLib.logList.logNoCommentText
                msgStr = msgStr.replace('__titleName__',lbName)
            else :
                msgStr = self.xmlQtBoxLib.logList.logFileNoExsitText
                msgStr = msgStr.replace('__titleName__',lbName)
                msgStr = msgStr.replace('__fileAddress__',fileAddressName)

            self.showErrMsgBoxInfo(msgTitle, msgStr)
            
        return bResult


    def judgeDirExsit(self, dirAddr: str, msgTitle : str = "judge Directory Exsit", lbName: str = "") -> bool:
        """
        ------------------------------------------------------------------------------
        ファイルパス存在判定処理\n
        【引数 】\n
            dirAddr:ファイルパス文字列
            msgTitle:メッセージボックスタイトル
            lbName:ラベル名称
        【戻り値】存在判定結果\n
            注:対象ファイルパス(dirAddr)が存在しない場合、エラーメッセージボックス出力\n
        ------------------------------------------------------------------------------ 
        """
        dirAddr = dirAddr.strip(" ")

        bResult = self.fsps.judgeDirExsit(dirAddr)
        if ( bResult  == False ) :
            if ( len(dirAddr)  == 0 ) :
                msgStr = self.xmlQtBoxLib.logList.logNoCommentText
                msgStr = msgStr.replace('__titleName__',lbName)
            else :
                msgStr = self.xmlQtBoxLib.logList.logPathNoExsitText
                msgStr = msgStr.replace('__titleName__',lbName)
                msgStr = msgStr.replace('__PathAddress__',dirAddr)
            self.showErrMsgBoxInfo(msgTitle, msgStr)
        return bResult


    def judgeIllegalCharacter(self, strCharacter : str, 
                              msgTitle : str = "judge Illegal Character", 
                              lbName: str = "",
                              illegalCharacter: str = r'[\\/:*?"<>|]+') -> bool:
        """
        ------------------------------------------------------------------------------
         違法アルファベット存在判定処理\n
        【引数 】\n
            strCharacter:対象アルファベット文字列\n
            msgTitle:メッセージボックスタイトル
            lbName:ラベル名称
            illegalCharacter:違法アルファベット\n
        【戻り値】存在判定結果\n
            注: 違法アルファベット(illegalCharacter)が存在する場合、\n
                エラーメッセージボックス出力\n
        ------------------------------------------------------------------------------ 
        """

        strCharacter = strCharacter.strip(" ")
        if ( len(strCharacter)  == 0 ) :
            msgStr = self.xmlQtBoxLib.logList.logNoCommentText
            msgStr = msgStr.replace('__titleName__',lbName)
            self.showErrMsgBoxInfo(msgTitle, msgStr)
            return False

        FileName2 = self.fsps.cleanIllegalCharacter(strCharacter, illegalCharacter)
        if ( len(strCharacter)  != len(FileName2) ) :
            msgStr = self.xmlQtBoxLib.logList.logIllegalCharText
            msgStr = msgStr.replace('__titleName__',lbName)
            msgStr = msgStr.replace('__IllegalChar__',strCharacter)
            self.showErrMsgBoxInfo(msgTitle, msgStr)
            return False
        
        return True

    def judgeHexCharacter(self, strCharacter : str, 
                              msgTitle : str = "judge Illegal Character", 
                              lbName: str = "") -> bool:
        """
        ------------------------------------------------------------------------------
         十六進数アルファベット存在判定処理\n
        【引数 】\n
            strCharacter:対象アルファベット文字列\n
            msgTitle:メッセージボックスタイトル
            lbName:ラベル名称
        【戻り値】存在判定結果\n
            注: 十六進数以外の場合、\n
                エラーメッセージボックス出力\n
        ------------------------------------------------------------------------------ 
        """

        strCharacter = strCharacter.strip(" ")
        if ( len(strCharacter)  == 0 ) :
            msgStr = self.xmlQtBoxLib.logList.logNoCommentText
            msgStr = msgStr.replace('__titleName__',lbName)
            self.showErrMsgBoxInfo(msgTitle, msgStr)
            return False
        
        bResult:bool =  True
        try:
             int(strCharacter, 16)
        except Exception as e:
            bResult = False
        
        if bResult == False :
            msgStr = self.xmlQtBoxLib.logList.logIllegalCharText
            msgStr = msgStr.replace('__titleName__',lbName)
            msgStr = msgStr.replace('__IllegalChar__',strCharacter)
            self.showErrMsgBoxInfo(msgTitle, msgStr)
            return False
        
        return True

    def delFile(self, fileAddr : str, msgTitle : str = "del File") -> bool:
        """
        -----------------------------------------------------------------
        ファイル削除処理\n
        【引数 】\n
            fileAddr:削除ファイルのフルアドレス\n
            msgTitle:メッセージボックスタイトル
        【戻り値】実施結果\n
        -----------------------------------------------------------------
        """
        bResult = self.fsps.delFile(fileAddr)
        if ( bResult  == False ) :
            msgStr = self.xmlQtBoxLib.logList.logFileCloseText
            msgStr = msgStr.replace('__fileAddress__',fileAddr)
            self.showErrMsgBoxInfo(msgTitle, msgStr)
        return bResult

    def JudgeOverWriteFile(self, fileAddress: str, msgTitle : str = "JudgeOverWriteFile") -> bool:
        """
        -----------------------------------------------------------------
        ファイル上書き処理判定\n
        【引数 】\n
            fileAddress:ファイルのフルアドレス\n
            msgTitle:メッセージボックスタイトル
        【戻り値】実施結果\n
        ------------------------------------------------------------------
        """
        msgStr = self.xmlQtBoxLib.logList.logFileOverWriteText
        msgStr = msgStr.replace('__fileAddress__',fileAddress)
        self.showQuestionMsgBoxInfo(msgTitle, msgStr)
        bResult = self.getMsgBoxClickResult()
        return bResult

    def JudgeOverWriteFiles(self, fileAddressList: list[str], msgTitle : str = "JudgeOverWriteFileList") -> bool:
        """
        -----------------------------------------------------------------
        ファイル上書き処理判定\n
        【引数 】\n
            fileAddressList:ファイルのフルアドレスリスト\n
            msgTitle:メッセージボックスタイトル
        【戻り値】実施結果\n
        ------------------------------------------------------------------
        """
        fileList = ""
        for fileAddress in fileAddressList:
            fileList +=fileAddress + "\n    "
        msgStr = self.xmlQtBoxLib.logList.logFileOverWriteText
        msgStr = msgStr.replace('__fileAddress__',fileList)
        self.showQuestionMsgBoxInfo(msgTitle, msgStr)

        bResult = self.getMsgBoxClickResult()
        return bResult

    def JudgeFileExtension(self, fileAddress: str, extensionName: str = "txt", msgTitle : str = "JudgeFileExtension") -> bool:
        """
        -----------------------------------------------------------------
        ファイル拡張子判定処理\n
        【引数 】\n
            fileAddress:ファイルのフルアドレス\n
            extensionName:拡張子\n
            msgTitle:メッセージボックスタイトル
        【戻り値】実施結果\n
        ------------------------------------------------------------------
        """
        bResult = self.fsps.JudgeFileExtensionByFileAddr(fileAddress, extensionName)

        if bResult  == False :
            msgStr = self.xmlQtBoxLib.logList.logFileExtensionText
            msgStr = msgStr.replace('__fileAddress__',fileAddress)
            self.showErrMsgBoxInfo(msgTitle, msgStr)
        return bResult

class qtTextEdit(QTextEdit):
    fsps:FileSysProcess
    
    def __init__(self, parent=None):
        super(qtTextEdit, self).__init__(parent)
        self.fsps = FileSysProcess()
        return


    def DragEnterEvent(self, e: QDragEnterEvent) -> None:
        """
        ------------------------------------------------------------------------------
         ドラッグ処理：ドラッグインフォ内容存在判断\n
        ------------------------------------------------------------------------------ 
        """
        strText:str =self.GetDragEventText(e)
        
        if len(strText) != 0 :
            e.accept()
        else :
            e.ignore()
        return


    def DragEvent_FileAddr(self, textobj: QTextEdit, e: QDropEvent) -> None:
        """
        ------------------------------------------------------------------------------
         ドラッグ処理：ファイルアドレス取得\n
        ------------------------------------------------------------------------------ 
        """
        strText:str =self.GetDragEventText(e)
        textobj.setText(strText)
        return


    def DragEvent_FilePath(self, textobj: QTextEdit, e: QDropEvent) -> None:
        """
        ------------------------------------------------------------------------------
         ドラッグ処理：ファイルパス取得\n
        ------------------------------------------------------------------------------ 
        """
        strText:str =self.GetDragEventText(e)
        
        resultList = self.fsps.getDirByFileFullAddr(strText)
        if resultList[0]  == False :
            textobj.setText(resultList[1])
        return
        

    def DragEvent_FileName(self, textobj: QTextEdit, e: QDropEvent) -> None:
        """
        ------------------------------------------------------------------------------
         ドラッグ処理：ファイル名称取得\n
        ------------------------------------------------------------------------------ 
        """
        strText:str =self.GetDragEventText(e)
        
        strFileInfo: list[str] = []
        bResult:bool = self.fsps.getFileNameInfoByFileFullAddr(strText, strFileInfo)
        if ( bResult  == True ) :
            textobj.setText(strFileInfo[1])
        return

    
    def FocusOutEvent(self, textobj: QTextEdit, e: QFocusEvent) -> None:
        pass
        return

    def GetDragEventText(self, e: QDropEvent) -> str:
        """
        ------------------------------------------------------------------------------
         ドラッグ処理：ドラッグ内容取得\n
        ------------------------------------------------------------------------------ 
        """
        strText:str = ""
        mimeData:QMimeData|None
        mimeData = e.mimeData()
        if mimeData is None:
            return strText
        
        strText = self.DragText(mimeData.text())
        return strText
    

    def DragText(self, strText: str) -> str:
        """
        ------------------------------------------------------------------------------
         ドラッグ処理：ドラッグインフォ内容整理\n
        ------------------------------------------------------------------------------ 
        """
        strText = strText.replace('file:','')
        strText = strText.replace('///','')
        strText = strText.replace('/','\\')
        return strText


class qtListWidget(QListWidget):

    def __init__(self, parent=None):
        super(qtListWidget, self).__init__(parent)
        return


    def DragEnterEvent(self, event: QDragEnterEvent| None ) -> None:
        
        """
        ------------------------------------------------------------------------------
         ドラッグ処理：ドラッグインフォ選択内容存在判断\n
        ------------------------------------------------------------------------------ 
        """
        if event is None:
            return
        bResult:bool = False
        sourceObject:QObject | None
        
        try:
            sourceObject = event.source()
            if sourceObject is not None:
                if sourceObject.isWidgetType() :
                        bResult = True
             
        except Exception as ex:
            pass
            
        if bResult :
            event.accept()
        else :
            event.ignore()
        return




