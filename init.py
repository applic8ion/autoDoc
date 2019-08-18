import win32com.client

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

main()
