library mircmail;

uses
  SysUtils,
  windows,shellapi,
  Classes,
  mIRCc in '..\TmIRCControl\mIRCc.pas';
function GetConsoleWindow:HWND;stdcall;external kernel32;
procedure mIRCExecuteA(hw:hwnd;inst:hinst;command:pansichar;nShow:integer);stdcall;
var mIRC:TmIRCControl;
conf:TStringlist;
target:string;
BEGIN
mirc:=tmirccontrol.Create(nil);
mirc.Active:=true;
target:=mirc.ActiveChan;
conf:=tstringlist.Create;
conf.LoadFromFile('mircmail.conf');
if conf.IndexOfName('rem lastTarget')>-1then
target:=conf.Values['rem lastTarget'];
case command[0] of
'/':mirc.Command(strpas(command),1);
'=':target:=strpas(@command[1]);
else mirc.Say(target,strpas(command));
end;
conf.Values['rem lastTarget']:=target;
conf.SaveToFile('mircmail.conf');
mirc.Free;
END;
function vbsError(errorlevel:dword):string;
begin
result:=format('Unknown error %u',[errorlevel]);
case errorlevel of
0:Result:=syserrormessage(error_success);
1:result:='From address is wrong';
2:result:='Subject is non blank';
end;
end;
procedure InboxCheckerA(hw:hwnd;inst:hinst;mailbox:pansichar;nShow:integer);stdcall;
var emlSR:tsearchrec;
emlname:array[0..max_path]of ansichar;
el:dword;
exec:tshellexecuteinfoa;
b:boolean;
begin
allocconsole;setconsoletitle('mIRCMail Messages');showwindow(getconsolewindow,
nshow);
while FindWindow('mIRC',nil)<>0do
begin
zeromemory(@emlsr,sizeof(emlsr));
b:=(findfirst(format('%s*.eml',[mailbox]),faanyfile,emlSR)=0);
if b then begin
el:=0;
sysutils.FindClose(emlsr);
zeromemory(@exec,sizeof(exec));
exec.cbSize:=sizeof(exec);
exec.fMask:=SEE_MASK_NOCLOSEPROCESS or SEE_MASK_FLAG_NO_UI or SEE_MASK_NO_CONSOLE;
exec.Wnd:=hw;
exec.lpFile:='cscript.exe';
exec.lpParameters:=strlfmt(stralloc(2049),2048,
'recvmail.vbs //Nologo //B /eml:%s%s',[mailbox,emlsr.FindData.cFileName]);
if not shellexecuteexa(@exec)then writeln(datetimetostr(now),' - ShellExecuteEx:',
syserrormessage(getlasterror))else begin write(datetimetostr(now),' - ',
emlsr.finddata.cfilename,':');
 waitforsingleobject(exec.hprocess,infinite);getexitcodeprocess(exec.hprocess,el);
 writeln(vbserror(el));
closehandle(exec.hprocess);if not deletefile(strfmt(emlname,'%s%s',[mailbox,
emlsr.finddata.cfilename]))
then begin writeln(emlname,': delete error');sleep(15000);
exitprocess(1);
 end;
end;
end;
end;
end;
exports mIRCExecuteA,InboxCheckerA;
begin
end.
