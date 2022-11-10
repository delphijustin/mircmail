set from=
set smtp_user=%from%
set smtp_pass=
set smtp_server=
set sendto=
set smtp_port=587
set smtp_options=/SSL
timeout /T 5
c:\SwithMail.exe /s /from "%from%" /u "%smtp_user%" /pass "%smtp_pass%" /server "%smtp_server%" /p "%smtp_port%" %smtp_options% /to "%sendto%" /sub %2 /b %1 /Log "%tmp%\mircmail.log"
