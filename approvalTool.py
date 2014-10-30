import sys, pyodbc, ctypes
from PyQt5.QtWidgets import QMainWindow, QApplication, QWidget, QFormLayout, QLabel, QButtonGroup, QRadioButton, QHBoxLayout, QTextEdit, QPushButton, QLineEdit, QMessageBox
from PyQt5.QtGui import QIcon

#testing.
class Approval(QMainWindow):
    def __init__(self):
        super(Approval, self).__init__()
        
        self.setWindowTitle("Expedited Approval Tool")
        self.setWindowIcon(QIcon('icon/PS_Icon.png'))
        
        self.frmMain = QFormLayout()
        #test
        mainLayout = QWidget()
        mainLayout.setLayout(self.frmMain)

        self.setCentralWidget(mainLayout)
        
        self.formMain()
        
    def formMain(self):
        
        leProbSheet = QLineEdit()
        leProbSheet.setMaximumWidth(50)
        self.frmMain.addRow(QLabel("Enter Problem Sheet ID:"), leProbSheet)
        
        hbShip = self.createRadioGroup("Ship")
        rbtnShipYes = self.rb.get("ShipYes")
        rbtnShipNo = self.rb.get("ShipNo")
        rbtnShipNo.setChecked(True)
        self.frmMain.addRow(QLabel("Did you confirm the correct shipping was \n selected and the correct pricing was quoted?"), hbShip)

        hbYY = self.createRadioGroup("YY")
        rbtnYYYes = self.rb.get("YYYes")
        rbtnYYNo = self.rb.get("YYNo")
        rbtnYYNo.setChecked(True)
        self.frmMain.addRow(QLabel("Is the order a YY or a non-suspicious NY/YN?"), hbYY)
        
        hbOver = self.createRadioGroup("Over")
        rbtnOverYes = self.rb.get("OverYes")
        rbtnOverNo = self.rb.get("OverNo")
        rbtnOverNo.setChecked(True)
        self.frmMain.addRow(QLabel("Did you override and upgrade the shipping \n in Stone Edge?"), hbOver)
        
        teNotes = QTextEdit()
        teNotes.setMaximumHeight(125)
        self.frmMain.addRow(QLabel("Additional Notes:"))
        self.frmMain.addRow(teNotes)
        
        btnApprove = QPushButton("Approved")
        btnApprove.clicked.connect(lambda: self.btnApprove_Click(leProbSheet.text(), rbtnShipYes.isChecked(), rbtnYYYes.isChecked(), rbtnOverYes.isChecked(), teNotes.toPlainText()))
        
        btnDeny = QPushButton("Not Approved")
        btnDeny.clicked.connect(lambda: self.btnDeny_Click(leProbSheet.text(), rbtnShipYes.isChecked(), rbtnYYYes.isChecked(), rbtnOverYes.isChecked(), teNotes.toPlainText()))
        
        btnQuit = QPushButton("Quit")
        btnQuit.clicked.connect(self.btnCancel_Click)
        
        hbBtnBox = QHBoxLayout()
        hbBtnBox.addWidget(btnApprove)
        hbBtnBox.addWidget(btnDeny)
        hbBtnBox.addWidget(btnQuit)
        
        self.frmMain.addRow(hbBtnBox)
        
    def btnApprove_Click(self, probSheetID, correctShip, YYorder, Override, notes):
        ad = ApprovalData()
        eid = ad.checkProbSheet(probSheetID)
        if eid:
            cnt = ad.checkApproved(eid)
        else: cnt = 0
        
        if eid > 0 and cnt == 0:
            try:
                ad.insApproval(probSheetID, eid, 1, correctShip, YYorder, Override, notes)
                ad.updExpedited(eid, 1)
                self.removeWidgets()
                self.formMain()
            except BaseException as e:
                QMessageBox.warning(self, "Database", "There was an issue inserting into the database. \n" + str(e), QMessageBox.Ok)
        else:
            if cnt:
                QMessageBox.warning(self, "Approved", "This problem sheet has already been approved.", QMessageBox.Ok)
            else:
                QMessageBox.warning(self, "Missing", "This problem sheet has not been entered into the system as expedited.", QMessageBox.Ok)
        
    def btnDeny_Click(self, probSheetID, correctShip, YYorder, Override, notes):
        ad = ApprovalData()
        eid = ad.checkProbSheet(probSheetID)
        if eid:
            cnt = ad.checkApproved(eid)
        else: cnt = 0
        
        if eid > 0 and cnt == 0:
            try:
                ad.insApproval(probSheetID, eid, 0, correctShip, YYorder, Override, notes)
                ad.updExpedited(eid, 0)
                self.removeWidgets()
                self.formMain()
            except BaseException as e:
                QMessageBox.warning(self, "Database", "There was an issue inserting into the database. \n" + str(e), QMessageBox.Ok)
        else:
            if cnt:
                QMessageBox.warning(self, "Approved", "This problem sheet has already been approved.", QMessageBox.Ok)
            else:
                QMessageBox.warning(self, "Missing", "This problem sheet has not been entered into the system as expedited.", QMessageBox.Ok)
       
    def btnCancel_Click(self):
        sys.exit()
        
    def createRadioGroup(self, name):
        #create dict items for all of the widgets, and to pull from after creation.
        bg = {}
        self.rb = {}
        hb = {}
        
        #create the group to hold the button.
        bg[(name)] = QButtonGroup(self)
        
        
        #create yes no buttons.
        self.rb[name+"Yes"] = QRadioButton("Yes")
        self.rb[name+"No"] = QRadioButton("No")
        
        #add buttons to group
        bg[(name)].addButton(self.rb[name+"Yes"])
        bg[(name)].addButton(self.rb[name+"No"])
    
        #add the whole thing to a Horz. Box layout.
        hb[(name)] = QHBoxLayout()
        hb[(name)].addWidget(self.rb[name+"Yes"])
        hb[(name)].addWidget(self.rb[name+"No"])
        hb[(name)].addStretch()      
        
        #print(self.rb[name+"Yes"].isChecked())

        return hb[(name)]  
    
    def removeWidgets(self):
        for cnt in reversed(range(self.frmMain.count())):
            # takeAt does both the jobs of itemAt and removeWidget
            # namely it removes an item and returns it
            widget = self.frmMain.takeAt(cnt).widget()
    
            if widget is not None: 
                # widget will be None if the item is a layout
                widget.deleteLater()

class ApprovalData():
    
    def __init__(self, *args):
        super(ApprovalData, self).__init__()
    
    def connect(self):
        try:
            conn = pyodbc.connect('DRIVER={SQL Server}; SERVER=SQLSERVER; DATABASE=ProblemSheets; Trusted_Connection=yes')
            db = conn.cursor()
        except BaseException as e:
            QMessageBox.critical(apr, 'Database Error', "Can not connect to database: \n" + str(e), QMessageBox.Ok)
        return db        
    
    def checkProbSheet(self, problemSheetID):
        db = self.connect()
        db.execute("SELECT ID FROM dbo.tblExpedited WHERE problemSheetID = ?",(problemSheetID))
        ds = db.fetchone()
        
        if ds:
            eid = ds[0]
        else:
            eid = 0
        
        return eid
    
    def checkApproved(self, eid):
        db = self.connect()
        db.execute("SELECT COUNT(ID) FROM dbo.tblExpeditedApproval WHERE expeditedSheetID = ? AND isApproved = 1", eid)
        ds = db.fetchone()
        
        cnt = ds[0]
        
        return cnt
    
    def insApproval(self, probID, expID, isApproved, shipping, billMatch, upgrade, notes):
        db = self.connect()
        db.execute("""IF EXISTS (SELECT * FROM dbo.tblExpeditedApproval WHERE expeditedSheetID = ?)
                        UPDATE dbo.tblExpeditedApproval 
                        SET isApproved = ?, correctShippingPrice = ?, isAddressMatch = ?, overrideUpgrade = ?, notes = ?
                        WHERE expeditedSheetID = ?
                     ELSE
                         INSERT INTO dbo.tblExpeditedApproval
                        (problemSheetID, expeditedSheetID, isApproved, correctShippingPrice, isAddressMatch, overrideUpgrade, notes)
                        VALUES
                        (?,?,?,?,?,?,?)""",(expID, isApproved, shipping, billMatch, upgrade, notes, expID, probID, expID, isApproved, shipping, billMatch, upgrade, notes))
        db.commit()
        
    
    def updExpedited(self, eid, approved):
        db = self.connect()
        if approved == 1:
            db.execute("UPDATE dbo.tblExpedited SET approved = ?, dateApproved = GETDATE() WHERE ID = ?", (approved, eid))
        elif approved == 0:
            db.execute("UPDATE dbo.tblExpedited SET approved = ?, dateApproved = NULL WHERE ID = ?", (approved, eid))
        db.commit()
        
if __name__ == "__main__":
    
    myappid = 'Approval Tool' # arbitrary string
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid) 
    
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    apr = Approval()
    apr.show()
    
    sys.exit(app.exec_())