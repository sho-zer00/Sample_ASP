' --メイン処理--

Option Explicit

' 定数宣言
Const ColumNum = 18 ' 0行目のカラムを格納する配列の大きさ、固定文字＋4月〜3月の出荷台数の出力をするループカウンタの最後の数
Const OutputStrFix = 6 ' 固定文字を格納するためのループカウンタの最後の数字
Const OutputStartMonth = 7 ' 固定文字＋4月〜3月の出荷台数の出力をするループカウンタの最初の数
Const LogFile = "log.txt" ' ログファイルパス
Const InputFile = "sample.txt" ' 読み込むファイルパス
Const OutputFile = "output.tsv" ' 出力するファイルパス

' 変数宣言
Dim strLine ' 読み込まれたデータを代入する変数
Dim arrFields ' 読み込んだデータを配列に入れるための配列宣言
Dim strMessage ' 表示用のメッセージ変数
Dim strFix ' 固定文字格納用の変数

Dim columName() ' 0行目のカラムを格納する配列宣言
ReDim columName(ColumNum) ' ※VBSでは定数を配列の大きさと指定したい場合再定義の必要あり

Dim lineCount ' 0行目のカラムを格納する配列宣言
lineCount = 0

Dim intCounter ' ループカウンタ
intCounter = 0


' ログファイルの指定（ない場合、新規作成）
Dim objFso
Dim log
Set objFso = CreateObject("Scripting.FileSystemObject")
Set log = objFso.OpenTextFile(LogFile, 8, True)


' 処理開始ログの出力
log.WriteLine FormatDateTime(Now, 0) & "出荷台数データファイル加工処理開始============================="

' エラー発生時にも処理を続行するよう設定
On Error Resume Next

' 出力ファイルの指定 （新規作成）
Dim output
Set output = CreateObject("ADODB.Stream")
output.Type = 2
output.Charset = "UTF-8"
output.Open

' 読み込みファイルの指定 
Dim input
Set input = CreateObject("ADODB.Stream")
input.Type = 2    ' 1：バイナリ・2：テキスト
input.Charset = "UTF-8"    ' 文字コード指定
input.Open    ' Stream オブジェクトを開く
input.LoadFromFile InputFile    ' ファイルを読み込む

' ファイルの存在チェック
If Err.Number <> 0 Then
	log.WriteLine FormatDateTime(Now, 0)& ":" & "エラー : " & Err.Description  
Else
	' 縦横変換処理を行う
	Unpivot
End If


' 出力したファイルを保存する
output.SaveToFile OutputFile,2

' Stream を閉じる
input.Close
output.Close

' 処理終了ログの出力
log.WriteLine FormatDateTime(Now, 0) & "出荷台数データファイル加工処理終了============================="
log.Close

Err.Clear
On Error Goto 0



' --縦横変換を行うメソッド--
Sub Unpivot()

    On Error Resume Next
	Do Until input.EOS

	    ' ファイルを1行ずつ読み込む。-2という数字は一行ずつ呼び込むことを表す
	    strLine = input.ReadText(-2) 

	    ' ファイル情報を読み込めるかどうかのチェック
	    If Err.Number <> 0 Then
	    	log.WriteLine FormatDateTime(Now, 0)& ":" & "エラー" & Err.Description 
	    Else
	    	log.WriteLine FormatDateTime(Now, 0)& ":" & "出荷台数データ [" & InputFile & "]の読み込み完了"  
	    End If

	    ' 読み込んだデータを一次配列に入れる
	    arrFields = Split(strLine,vbTab) 

	    ' 縦横変換の際の固定文字を格納する
	    strFix = "" 
	    For intCounter = 0 To OutputStrFix Step 1
	        strFix = strFix & arrFields(intCounter) & vbTab
	    Next

	    ' 最初の行を出力するときの処理
	    If lineCount = 0 Then

	        ' 項目名は配列に格納（縦横変換の値で使用するため）
	        For intCounter = 0 To ColumNum Step 1
	            columName(intCounter) = arrFields(intCounter)
	        Next

	        ' 項目名の出力
	        strMessage = strFix & "Month" & vbTab & "Value" & vbCrLf
	        output.WriteText strMessage,0

	    ' 最初の行以外を出力するときの処理
	    Else

	        ' 固定文字＋4月〜3月の出荷台数の出力
	        For intCounter = OutputStartMonth To ColumNum Step 1
	            strMessage = strFix & columName(intCounter) & vbTab & arrFields(intCounter) & vbCrLf
	            output.WriteText strMessage,0
	        Next 
	        
	    End If
	    
	    ' ループカウンタを増やす
	    lineCount = lineCount + 1
	    log.WriteLine FormatDateTime(Now, 0) & lineCount & "の読み込み"
	Loop
	If Err.Number <> 0 Then
	    	log.WriteLine FormatDateTime(Now, 0)& ":" & "エラー" & Err.Description 
	    Else
	    	log.WriteLine FormatDateTime(Now, 0)& ":" & "出荷台数データ加工完了"  
	End If
	Err.Clear
	On Error Goto 0
End Sub
