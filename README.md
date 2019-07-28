
	

rem コマンドの実行結果を変数に入れる  /F "usebackq"がその役目を持っている
for /F "usebackq" %%a IN (`powershell [DateTime]::Today.AddMonths"("-3")".ToString"("'yyyyMMdd'")"`) do set delete_day=%%a
echo %delete_day%
forfiles /s /m %delete_day%.log* /c "cmd /c echo delete target @path">>"log.txt
forfiles /s /m %delete_day%.log* /c "cmd /c del @path">>"log.txt
if %errorlevel% equ 1 echo no such %delete_day%.log file >> log.txt
if %errorlevel% equ 0 echo complete >> log.txt
