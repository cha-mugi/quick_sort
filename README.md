# quick_sort



Sub makeText()
Dim ws As Worksheet
Set ws = ThisWorkbook.Worksheets(1)

Dim datFile As String
datFile = ActiveWorkbook.Path & "\data.txt"

Open datFile For Output As #1

Print #1, "aiueo"


Dim i As Long
i = 1
Do While ws.Cells(i, 1).Value <> ""
    Print #1, ws.Cells(i, 1).Value
    i = i + 1
Loop

Close #1

MsgBox "data.txtに書き出しました"

End Sub
''http://www.tipsfound.com/vba/07001

