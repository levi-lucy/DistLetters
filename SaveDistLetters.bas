Attribute VB_Name = "Module2"
'Author: Levi Lucy
'Date: 3/8/2018
'SaveDistLetters is a macro to save distribution letters using the information found within the excel file

Sub SaveDistLetters()

Dim wrdApp As Object
Dim wrdDoc As Object
Dim i As Integer
Dim j As Integer
Dim rowEnd As Integer
Dim propRowEnd As Integer
Dim property As String
Dim distAmt As Double
Dim wrkSht As Worksheet
Dim result As Variant

Sheets("Home").Activate
rowEnd = Cells(Rows.Count, 1).End(xlUp).Row 'Last element of Property list

For i = 2 To rowEnd

'Property Information
distAmt = Cells(i, 2).Value 'Amount of distribution

If distAmt > 0 Then
    property = Trim(Cells(i, 1).Value) 'Property name sending distribution
    Set wrkSht = Sheets(property) 'Worksheet containing partner information for given property

    'Partner information
    propRowEnd = wrkSht.Cells(Rows.Count, 1).End(xlUp).Row 'Last partner in list
    For j = 2 To propRowEnd
        result = PrintDistLetter(property, distAmt, wrkSht.Cells(j, 1).Value, wrkSht.Cells(j, 2).Value) 'Print letter for each partner
    Next
End If

'Determine if the procedure worked
Next
End Sub

'PrintDistLetter:
'Property should be the name of the property sending out distribution, amt is the distribution amount,
'Partner is the name of the partner receiving the letter (should have a sheet with the same name),
'prcnt is the percent ownership the partner has. The end result should be to produce a letter as a word doc with specific
'formatting. Will return true if successful, false if an error occurs.

Function PrintDistLetter(property As String, amt As Double, partner As String, prcnt As Double) As Boolean
Dim wrdApp As Object
Dim wrdDoc As Object
Dim partSheet As Worksheet
Dim partEnd As Double
Dim k As Integer
Dim distDate As Date

On Error GoTo errOutput
Set partSheet = Sheets(partner) 'Use name of partner for worksheet name with address
partEnd = WorksheetFunction.CountA(partSheet.Range("A:A")) 'Last row with data
distDate = Sheets("Home").Cells(2, 4)

'To create the word document
Set wrdApp = CreateObject("Word.Application")
wrdApp.Visible = True
Set wrdDoc = wrdApp.Documents.Add

With wrdDoc
'Set initial format
.PageSetup.TopMargin = 1.75 * 72
.Range.ParagraphFormat.SpaceAfter = 8
.Range.Font.name = "Times New Roman"
.Range.Font.Size = 11

'Format for letterhead
.Content.InsertParagraphAfter
.Content.InsertParagraphAfter
.Content.InsertParagraphAfter
.Content.InsertAfter Format(Date, "mmmm d, yyyy") 'Date of letters
.Content.InsertParagraphAfter
.Content.InsertParagraphAfter

'Change format for address
.Content.InsertAfter partSheet.Cells(2, 1).Value
.Paragraphs(.Paragraphs.Count).Range.Font.Size = 12
.Paragraphs(.Paragraphs.Count).Range.ParagraphFormat.SpaceAfter = 0
.Content.InsertParagraphAfter
For k = 3 To partEnd
.Content.InsertAfter partSheet.Cells(k, 1).Value
.Content.InsertParagraphAfter
Next
.Content.InsertParagraphAfter

'Bold subject line and change format again
.Content.InsertAfter UCase("Re: " & property & Format(distDate, " mmmm yyyy ") & "Distribution")
.Paragraphs(.Paragraphs.Count).Range.Font.Bold = True
.Paragraphs(.Paragraphs.Count).Range.Font.Size = 11
.Paragraphs(.Paragraphs.Count).Range.ParagraphFormat.SpaceAfter = 8
.Content.InsertParagraphAfter
.Content.InsertParagraphAfter

'Remove bold font
.Content.InsertAfter partSheet.Cells(2, 2).Value
.Paragraphs(.Paragraphs.Count).Range.Font.Bold = False
.Content.InsertParagraphAfter
.Content.InsertAfter "Enclosed, please find a check in the amount of " _
                     & Format(amt * prcnt, "$#,###.#0 ") & "representing your proportionate share of the" _
                     & Format(distDate, " mmmm, yyyy ") _
                     & "distribution totaling " & Format(amt, "$#,###.#0.")
.Content.InsertParagraphAfter
.Content.InsertAfter "Should you have any questions, please do not hesitate to contact me at mark@ackerberg.com " _
                      & "or          612-924-6402."
.Content.InsertParagraphAfter
.Content.InsertAfter "Yours sincerely,"
.Content.InsertParagraphAfter
.Content.InsertParagraphAfter
.Content.InsertParagraphAfter
.Content.InsertAfter "Mark Schlitter"
.Content.InsertParagraphAfter
.Content.InsertAfter "Chief Financial Officer"
.Content.InsertParagraphAfter
.Content.InsertAfter "The Ackerberg Group"

'Save document with meaningful name
.SaveAs (ActiveWorkbook.path & "\Distribution Letters\Cash Distribution Letter - " & property & " to " & partner & ".docx")
.Close
End With
wrdApp.Quit
Set wrdDoc = Nothing
Set wrdApp = Nothing
PrintDistLetter = True
Exit Function

errOutput:
PrintDistLetter = False
MsgBox ("Failure: " & property & " distribution to " & partner & " not saved.")

End Function


