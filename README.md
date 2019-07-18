■資料・動画ダウンロード方法
1．マイページログインID・PWで下記ページへログインしてください。
https://jpn01.safelinks.protection.outlook.com/?url=https%3A%2F%2Faws.summitregist.jp%2Fpublic%2Flogin%3Fpage%3Dauth%26return_path%3D%2Fpublic%2Fapplication%2Fadd%2F79&amp;data=02%7C01%7C%7Cf8c5a95191e84038fcde08d6fea3d382%7Ca4dd529424e441028420cb86d0baae1e%7C1%7C0%7C636976376927644641&amp;sdata=c6VtzyEFVt%2FnppjsBnaJFz4glpk6HOVKAO9o%2F1yjfAQ%3D&amp;reserved=0
ログインID：cyamada
ログインPW：ご登録いただいたパスワード
2．ご希望の資料・動画を選択し、完了画面まで遷移してください。
3．その後、マイページへ遷移いただき、選択された資料・動画名下のリンクをクリックいただくことでご覧いただけます。


ログインID：cyamada
ログインPW：@10kages
	
rem コマンドの実行結果を変数に入れる  /F "usebackq"がその役目を持っている
for /F "usebackq" %%a IN (`powershell [DateTime]::Today.AddMonths"("-3")".ToString"("'yyyyMMdd'")"`) do set delete_day=%%a
echo %delete_day%
forfiles /s /m %delete_day%.log* /c "cmd /c echo del @file >> log.txt"
if %errorlevel% equ 1 echo no such %delete_day%.log file >> log.txt
if %errorlevel% equ 0 echo complete >> log.txt
