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

Private Sub CommandButton4_Click()
 
  Dim Stream As Object
  
  ' VB標準のADODB.Streamオブジェクトを作成する
  Set Stream = CreateObject("ADODB.Stream")
  
  ' ストリームの文字コードをUTF8に設定する
  Stream.Charset = "UTF-8"
  ' ファイルのタイプ(1:バイナリ 2:テキスト)
  Stream.Type = 2
  ' ストリームを開く
  Stream.Open
  ' ストリームの保存形式をテキスト形式にする
  Stream.WriteText "エクセル講座" & vbCrLf & "http://www.petitmonte.com/excel/excel.html"
  ' ストリームに名前を付けて保存する(1は新規作成 2は上書き保存)
  Stream.SaveToFile ("utf8の書き込みテスト.txt"), 2
  ' ストリームを閉じる
  Stream.Close
  
  Set Stream = Nothing
  
End Sub
