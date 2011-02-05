@ECHO OFF

REM httpget example

SET MSG=Done
CScript httpget.vbs http://cowares.nobody.jp/favicon.ico C:\tmp\favicon.ico
IF NOT ERRORLEVEL 0 SET MSG=Error %ERRORLEVEL%
ECHO %MSG%

SET MSG=Done
CScript httpget.vbs http://cowares.nobody.jp/favicon.ic C:\tmp\favicon.err
IF NOT ERRORLEVEL 0 SET MSG=Error %ERRORLEVEL%
ECHO %MSG%

REM httpgets example

SET MSG=Done
CScript httpgets.vbs < urls.txt
IF NOT ERRORLEVEL 0 SET MSG=Error %ERRORLEVEL%
ECHO %MSG%

