Attribute VB_Name = "Module3"
Sub PrintLetters()

Dim file As String
Dim path As String
Dim wrdApp As Object
Dim wrdDoc As Object
Dim printr As String
Dim defaultPrintr As String

defaultPrintr = Application.ActivePrinter 'Finds the default printer

path = ActiveWorkbook.path & "\Distribution Letters\" 'Requires a distribution letters folder to exist

Application.Dialogs(xlDialogPrinterSetup).Show 'Select printer to use
printr = Application.ActivePrinter

'Start word application and use the selected printer
Set wrdApp = CreateObject("Word.Application")
wrdApp.ActivePrinter = printr
wrdApp.Visible = False

'Cycle through the letters already saved
file = Dir(path)

Do While file <> ""
Set wrdDoc = wrdApp.Documents.Open(path & file, ReadOnly:=True)
wrdDoc.PrintOut
wrdDoc.Close
file = Dir()
Loop

wrdApp.ActivePrinter = defaultPrintr 'Set back to previous default printer
wrdApp.Quit
Set wrdApp = Nothing
Set wrdDoc = Nothing
End Sub
