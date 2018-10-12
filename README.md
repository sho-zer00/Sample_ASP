' 開始ログの出力※後で記述

' 読み込みファイルの指定 
Dim input
Set input = CreateObject("ADODB.Stream")
input.Type = 2    ' 1：バイナリ・2：テキスト
input.Charset = "UTF-8"    ' 文字コード指定
input.Open    ' Stream オブジェクトを開く
input.LoadFromFile "C:\Users\A0832\desktop\sample.tsv"    ' ファイルを読み込む

' 書き出しファイルの指定 (今回は新規作成する)
Dim output
Set output = CreateObject("ADODB.Stream")
output.Type = 2
output.Charset = "UTF-8"
output.Open

' 変数の定義
Dim strLine
Dim arrFields

' 表示用のメッセージ変数
Dim strMessage

' 固定文字格納用の変数
Dim strFix

' 0行目のカラムを格納する配列
Dim columName(18)

' 何行目を呼び込んでいるかのカウンタ
Dim lineCount
lineCount = 0

' ループカウンタ
Dim intCounter
intCounter = 0

' --縦横変換
Do Until input.EOS
    strLine = input.ReadText(-2) '-2という数字は一行ずつ呼び込むことを表しています。ちなみに-1は全読み込み
    arrFields = Split(strLine,vbTab)
    strFix = "" 
    For intCounter = 0 To 6 Step 1
        strFix = strFix & arrFields(intCounter) & vbTab
    Next
    If lineCount = 0 Then

        For intCounter = 0 To 18 Step 1
            columName(intCounter) = arrFields(intCounter)
        Next
        strMessage = strFix & "Month" & vbTab & "Value" & vbCrLf
        output.WriteText strMessage,0
    Else
        For intCounter = 7 To 18 Step 1
            strMessage = strFix & columName(intCounter) & vbTab & arrFields(intCounter) & vbCrLf
            output.WriteText strMessage,0
        Next 
        
    End If
    
    lineCount = lineCount + 1

Loop

output.SaveToFile "C:\Users\A0832\desktop\output.tsv",2

' Stream を閉じる
input.Close
output.Close
