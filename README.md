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
http://officetanaka.net/excel/vba/variable/08.htm


Sub CreateArgStatement()

Dim sheet1 As Worksheet
Set s = Worksheets(1)

'現在位置から親を探す
Dim arg_statement As String
arg_statement = parent_name(current_y_pos, current_x_pos, s)

End Sub


'親のnameを返す
Function parent_name(current_y_pos As Long, current_x_pos, sheet As Worksheet) As String


    '現在位置が親の位置ならそのまま親を返す
    If current_x_pos = 5 Then
    '''arg_child_x_pos_argnameは引数名
        parente_name = s.Cells(current_y_pos, current_x_pos_argName).Value
        
        GoTo Continue
    End If
    '
    
    '直接の親を探す
    arg_child_y_pos = current_y_pos
      '自分の直接の親のx座標を指す
    arg_child_x_pos = current_x_pos - 1
    Do While s.Cells(arg_child_y_pos, arg_child_x_pos).Value <> ""
        arg_child_y_pos = arg_child_y_pos - 1
    Loop
    '''arg_child_x_pos_argnameは引数名
    parente_name = s.Cells(arg_child_y_pos, arg_child_x_pos_argname)
    
    parent_name = parent_name(arg_child_y_pos, arg_child_x_pos) & parent_name
Continue:
End Function

Sub CreateNextMessage()
    '複数のMessageを作成する
    
    ''受信メッセージの隣にIF名がなければ、Messageを作成する
End Sub

Sub Create()
    '複数の引数を一つのEnumに変換
    '
    
    ''受信メッセージの隣にIF名がなければ、Messageを作成する
End Sub


NString string;
string.GetString(char[],size)

#open ($fh, '<:encoding(cp932)', 'text01.txt')
my $header = 'Message.h'
my $cpp    = 'Message.cpp'

my $CPP
my $HEADER
open(CPP,cpp)
open(HEADER,header)

while(<$CPP>)
{
	if ($_ =~ /¥d+yen/)
	{
  		print "$str¥n";
	}		
}

close(CPP)
close(HEADER)

strSamp = "123456789"
strSamp = Replace(strSamp, "123", "000")
'000456789を返す

https://allabout.co.jp/gm/gc/420438/
