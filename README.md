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

' 読み込まれたデータを代入する変数
Dim strLine

' 読み込んだデータを配列に入れるための宣言
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

' --縦横変換--
Do Until input.EOS
    ' ファイルを1行ずつ読み込む。-2という数字は一行ずつ呼び込むことを表している。
    strLine = input.ReadText(-2) 

    ' 読み込んだデータを一次配列に入れる
    arrFields = Split(strLine,vbTab) 

    ' 縦横変換の際の固定文字を格納する
    strFix = "" 
    For intCounter = 0 To 6 Step 1
        strFix = strFix & arrFields(intCounter) & vbTab
    Next

    ' 最初の行を読み込むときの処理
    If lineCount = 0 Then

        ' 項目名は配列に格納（縦横変換の値で使用するため）
        For intCounter = 0 To 18 Step 1
            columName(intCounter) = arrFields(intCounter)
        Next

        ' 項目名の表示
        strMessage = strFix & "Month" & vbTab & "Value" & vbCrLf
        output.WriteText strMessage,0

    ' 最初の行以外を読み込むときの処理
    Else

        ' 固定文字＋4月～3月の出荷台数の表示
        For intCounter = 7 To 18 Step 1
            strMessage = strFix & columName(intCounter) & vbTab & arrFields(intCounter) & vbCrLf
            output.WriteText strMessage,0
        Next 
        
    End If
    
    ' ループカウンタを増やす
    lineCount = lineCount + 1

Loop

output.SaveToFile "C:\Users\A0832\desktop\output.tsv",2

' Stream を閉じる
input.Close
output.Close
