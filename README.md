# Sample_ASP
vbscript readFile
' 読み込みファイルの指定 (相対パスなのでこのスクリプトと同じフォルダに置いておくこと)
Dim input
Set input = CreateObject("ADODB.Stream")
input.Type = 2    ' 1：バイナリ・2：テキスト
input.Charset = "UTF-8"    ' 文字コード指定
input.Open    ' Stream オブジェクトを開く
input.LoadFromFile "sample.tsv"    ' ファイルを読み込む

' 書き出しファイルの指定 (今回は新規作成する)
Dim output
Set output = CreateObject("ADODB.Stream")
output.Type = 2
output.Charset = "UTF-8"
output.Open

' 変数の定義
Dim strLine
Dim arrFields
Dim strMessage

Do Until input.EOS
        strLine = input.ReadText(-2)
        arrFields = Split(strLine,vbTab)
        strMessage = arrFields(0) & vbTab & arrFields(1) & vbTab & arrFields(2) & vbTab & arrFields(3) & vbTab & arrFields(4) & vbTab & arrFields(5) _
                     & vbTab & arrFields(6) & vbTab & arrFields(7) & vbTab & arrFields(8) & vbTab & arrFields(9) & vbCrLf
        output.WriteText strMessage,0
Loop

' MsgBox strMessage, vbInformation + vbOkOnly, "kakunin"
output.SaveToFile "output.tsv",2

' Stream を閉じる
input.Close
output.Close
