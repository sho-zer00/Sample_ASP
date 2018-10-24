' -----------------------------------------------------------------------
' 処理名    ：Unpivot.vbs
' 処理概要  ：出荷台数データの縦横変換を行う
' パラメータ：InputFile ：変換前tsvファイル
'             OutputFile：変換後のtsvファイル
' 作成者    ：Isogai Sho
' 作成日    ：2018/10/18
' -----------------------------------------------------------------------

Option Explicit

' エラー発生時にも処理を続行するよう設定
On Error Resume Next

' -----------------------------------
' 定数宣言
' -----------------------------------

' 0行目のカラムを格納する配列の大きさ、固定文字＋4月から3月の出荷台数の出力をするループカウンタの最後の数
Const ColumNum = 18 

' 固定文字を格納するためのループカウンタの最後の数字
Const OutputStrFix = 6 

' 固定文字＋4月から3月の出荷台数の出力をするループカウンタの最初の数
Const OutputStartMonth = 7 

' Unpivot.logはコマンドライン引数では渡さない（コマンドライン引数として渡してしまうと、エラーによってはエラー情報をログに残せないため）
Const LogFile = "Log\Unpivot.log" 

'------------------------------------
' 変数宣言
'------------------------------------

' 読み込まれたデータを代入する変数
Dim strLine 

' 読み込んだデータを配列に入れるための配列宣言
Dim arrFields 

' 表示用のメッセージ変数
Dim strMessage 

' 固定文字格納用の変数
Dim strFix 

' 0行目のカラムを格納する配列宣言
Dim columName() 
ReDim columName(ColumNum) 

' 0行目のカラムを格納する配列宣言
Dim lineCount 
lineCount = 0

' ループカウンタ
Dim intCounter 
intCounter = 0

' 出荷台数データファイルパスをコマンドラインから受け取る
Dim InputFile 
InputFile = WScript.Arguments(0)

' Unpivot変換後のファイルパスをコマンドラインから受け取る
Dim OutputFile
OutputFile = WScript.Arguments(1)

' ログファイルの指定（ない場合、新規作成）
Dim objFso
Dim log
Set objFso = CreateObject("Scripting.FileSystemObject")

' 読み込みファイルの指定 
Dim input
Set input = CreateObject("ADODB.Stream")
input.Type = 2
input.Charset = "UTF-8"
input.Open
input.LoadFromFile InputFile

' 出力ファイルの指定 （新規作成）
Dim output
Set output = CreateObject("ADODB.Stream")
output.Type = 1

' 一時書き込みファイルの指定（BOMなしにするため）
Dim preout
Set preout = CreateObject("ADODB.Stream")
preout.Type = 2
preout.Charset = "UTF-8"

' 最終項目に年度を追記するための変数（取得するコマンドライン引数から数字の部分を抜き出す）
Dim year
year = Mid(InputFile,9,4)

'------------------------------------
' ファイルの有無チェック
'------------------------------------

If Err.Number = 0 Then

	' ログを上書きモードに
	Set log = objFso.OpenTextFile(LogFile, 2, True)

	' 処理開始ログの出力
	log.WriteLine FormatDateTime(Now, 0) & " "& "出荷台数データファイル加工処理開始================================"

	' 縦横変換処理を行う
	Unpivot

Else 
	' ログを上書きモードに
	Set log = objFso.OpenTextFile(LogFile, 2, True)

	' エラーログの出力
	log.WriteLine FormatDateTime(Now, 0) & " " & "エラー :" & Err.Description 

End If

' Stream を閉じる
input.Close
log.Close

Err.Clear
On Error Goto 0


'------------------------------------
' メイン処理（縦横変換）
'------------------------------------

Sub Unpivot()

	' 縦横変換開始時に出力ファイル、一次書き込みファイルのStreamオブジェクトを開く
	output.Open
	preout.Open

    ' エラー発生時にも処理を続行するよう設定
    On Error Resume Next

	' 縦横変換開始
	Do Until input.EOS

	    ' ファイルを1行ずつ読み込む。-2という数字は一行ずつ呼び込むことを表す
	    strLine = input.ReadText(-2) 

	    '------------------------------------
		' 読み込みチェック
		'------------------------------------
		If Err.Number <> 0 Then
			' エラーの場合、ループを抜けてエラーメッセージの表示をして異常終了をする
	    	log.WriteLine FormatDateTime(Now, 0)& " " & "エラー :" & Err.Description 
            Exit Do

	    Else
		
			' 読み込み成功の場合、読み込み成功の旨のメッセージを表示する
	    	log.WriteLine FormatDateTime(Now, 0)& " " & "出荷台数データ [" & InputFile & "]の" & lineCount + 1 &"行目の読み込み成功"  
			
			' 読み込んだデータを一次配列に入れる
	    	arrFields = Split(strLine,vbTab) 

	    	' 縦横変換の際の固定文字を格納する
	    	strFix = "" 
			For intCounter = 0 To OutputStrFix Step 1
	        	strFix = strFix & arrFields(intCounter) & vbTab
	    	Next

	    End If

	    '------------------------------------
		' １行目を一時ファイルへ出力する処理
		'------------------------------------
	    If lineCount = 0 Then

	        ' 項目名は配列に格納（縦横変換の値で使用するため）
	        For intCounter = 0 To ColumNum Step 1
	            columName(intCounter) = arrFields(intCounter)
	        Next

	        ' 項目名の出力
	        strMessage = strFix & "Month" & vbTab & "Value" & vbTab & "Fiscal_Year_ID" & vbCrLf
	        preout.WriteText strMessage,0

	    '------------------------------------
		' ２行目以降を一時ファイルへ出力する処理
		'------------------------------------
	    Else

	        ' 固定文字＋4月から3月の出荷台数の出力（縦横変換）
	        For intCounter = OutputStartMonth To ColumNum Step 1

                ' 出力内容
	            strMessage = strFix & columName(intCounter) & vbTab & arrFields(intCounter) & vbTab & year & vbCrLf

                '------------------------------------
				' 縦横変換時のエラーチェック
				'------------------------------------
                If Err.Number <> 0 Then
	    	        log.WriteLine FormatDateTime(Now, 0)& " " & "エラー :" & Err.Description 
                    Exit Do
	            ElseIf Err.Number = 0 AND intCounter = ColumNum Then
	    	        log.WriteLine FormatDateTime(Now, 0)& " " & "出荷台数データ [" & InputFile & "]の" & lineCount + 1 &"行目の縦横変換成功"  
	            End If

                ' 一時書き込み
	            preout.WriteText strMessage,0
	        Next
	    End If

        '------------------------------------
		' 一時書き込み時のエラーチェック
		'------------------------------------
	    If Err.Number <> 0 Then
	    	log.WriteLine FormatDateTime(Now, 0)& " " & "エラー :" & Err.Description 
            Exit Do
	    Else
	    	log.WriteLine FormatDateTime(Now, 0)& " " & "出荷台数データ [" & InputFile & "]の" & lineCount + 1 &"行目の一時書き込み成功"  
	    End If

	    ' ループカウンタを増やす
	    lineCount = lineCount + 1
	Loop

	' バイナリモードにする（BOMをスキップする）
    preout.Position = 0
	preout.Type = 1
	preout.Position = 3
    
	' 一時書き込みファイルデータの内容を読み込む
	Dim bin
	bin = preout.Read

	' 出力ファイルへデータを渡す
	output.Write(bin)
	
    '------------------------------------------------------
	' これまでの処理の中で一つでもエラーがあるかどうかのチェック
	'------------------------------------------------------
	If Err.Number <> 0 Then

		' 異常終了ログの出力
		log.WriteLine FormatDateTime(Now, 0) & " "& "出荷台数データファイル加工処理異常終了============================="

	Else

		' 出力ファイルへ書き込む
    	output.SaveToFile OutputFile,2

		'-----------------------------------
		' 出力ファイル書き込み時のエラーチェック
		'-----------------------------------
		If Err.Number <> 0 Then

		' 異常終了ログの出力
		log.WriteLine FormatDateTime(Now, 0) & " "& "エラー :" & Err.Description
		log.WriteLine FormatDateTime(Now, 0) & " "& "出荷台数データファイル加工処理異常終了============================="

		Else

		' 正常終了ログの出力
		log.WriteLine FormatDateTime(Now, 0) & " "& "出力ファイルへの書き込み成功"
		log.WriteLine FormatDateTime(Now, 0) & " "& "出荷台数データファイル加工処理正常終了============================="

		End If

	End If


	' 縦横変換処理終了時に出力ファイルのStreamオブジェクトを閉じる
	output.Close
	preout.Close
	Err.Clear
	On Error Goto 0
	
End Sub
