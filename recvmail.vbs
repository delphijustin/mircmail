Dim std_internal_LibFiles
Set std_internal_LibFiles = CreateObject("Scripting.Dictionary")
' *************************************************************************************************
'! Includes a VbScript file
'! @param p_Path    The path of the file to include
' *************************************************************************************************
Sub IncludeFile(p_Path)
    ' only loads the file once
    If std_internal_LibFiles.Exists(p_Path) Then
        Exit Sub
    End If
    
    ' registers the file as loaded to avoid to load it multiple times
    std_internal_LibFiles.Add p_Path, p_Path

    Dim objFso, objFile, strFileContent, strErrorMessage
    Set objFso = CreateObject("Scripting.FileSystemObject")
    
    ' opens the file for reading
    On Error Resume Next
    Set objFile = objFso.OpenTextFile(p_Path)
    If Err.Number <> 0 Then
        ' saves the error before reseting it
        strErrorMessage = Err.Description & " (" &  Err.Source & " " & Err.Number & ")"
        On Error Goto 0
        Err.Raise -1, "ERR_OpenFile", "Cannot read '" & p_Path & "' : " & strErrorMessage
    End If
    
    ' reads all the content of the file
    strFileContent = objFile.ReadAll
    If Err.Number <> 0 Then
        ' saves the error before reseting it
        strErrorMessage = Err.Description & " (" &  Err.Source & " " & Err.Number & ")"
        On Error Goto 0
        Err.Raise -1, "ERR_ReadFile", "Cannot read '" & p_Path & "' : " & strErrorMessage
    End If
    
    ' this allows to run vbscript contained in a string
    ExecuteGlobal strFileContent
    If Err.Number <> 0 Then
        ' saves the error before reseting it
        strErrorMessage = Err.Description & " (" &  Err.Source & " " & Err.Number & ")"
        On Error Goto 0
        Err.Raise -1, "ERR_Include", "An error occurred while including '" & p_Path & "' : " & vbCrlf & strErrorMessage
    End If
End Sub
dim adoStream
dim cdoMsg
dim namedArgs
dim objShell
IncludeFile "mIRCMail.conf"
set namedArgs = Wscript.Arguments.Named
set adoStream = CreateObject("ADODB.Stream")
adoStream.Open
adoStream.LoadFromFile(namedArgs.Item("eml"))
set cdoMsg = CreateObject("CDO.Message")
cdoMsg.DataSource.OpenObject adoStream,"_Stream"
if InStr(Lcase(cdoMsg.From),Lcase(myemail))=0 then
Wscript.Quit 1
end if
if Len(cdoMsg.Subject) > 0 then
if blankSubjectOnly then
wscript.quit 2
end if
end if
Set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "rundll32.exe", "mircmail.dll,mIRCExecute "&cdoMsg.TextBody,"","open",10
Wscript.Quit 0