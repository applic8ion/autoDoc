import sys, win32com.client
from UiGraphicsView import Ui_MainWindow
from PyQt5 import QtGui, QtWidgets, QtCore

class MyWindow(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

class WordAuto:
    def __init__(self, templeteFile=None):
        """
        class for word automation
        :param templeteFile:
        """
        self.wordApp = win32com.client.gencache.EnsureDispatch('Word.Application')
        if templeteFile == None:
            self.wordDoc = self.wordApp.Documents.Add()
        else:
            self.wordDoc = self.wordApp.Documents.Add(Templete=templeteFile)

        self.wordDoc.PageSetup.LeftMargin = 30
        self.wordDoc.PageSetup.TopMargin = 30
        self.wordDoc.PageSetup.BottomMargin = 30
        self.wordDoc.PageSetup.RightMargin = 30

        # Setup the Selection
        self.wordApp.Visible = True
        self.wordDoc.Range(0, 0).Select()
        self.wordSel = self.wordApp.Selection

    def Quit(self):
        self.wordDoc.Close(SaveChanges=1)

def main():
    print('Test Start!')
    wa = WordAuto()
    wa.wordDoc.Content.Text = "Hello"

if __name__ == "__main__":
        app = QtWidgets.QApplication(sys.argv)
        # view = MyView()
        ui = MyWindow()
        ui.show()
        sys.exit(app.exec_())
