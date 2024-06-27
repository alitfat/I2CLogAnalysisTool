from PyQt5.QtWidgets import QWidget, QFormLayout, QLabel

class I2CLogAnalysisTool_VerGUI(QWidget):
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("I2CLogAnalysisToolについて")
        self.resize(300, 90)
        
        layout = QFormLayout()
        self.lbVer = QLabel("Version   \t：1.0.0\n" + 
                            "Created by \t：GaoHongyu\n" + 
                            "MailAddress\t：alitfat142@gmail.com\n" + 
                            "Created at \t：自宅")
        layout.addWidget(self.lbVer)
        
        self.setLayout(layout)
        
