' Версия 0.2
' Получение имени ini-файла
Set objArgs=WScript.arguments
If objArgs.Count=0 Then
  strConfig="gpinst.ini"
Else
  strConfig=objArgs.item(0)
End If

' Получение переменных из ini-файла
Dim dIni
Dim objFS: Set objFS=CreateObject("Scripting.FileSystemObject")
Dim objWshShell: Set objWshShell=CreateObject("WScript.Shell")
strIniFile=GetCurrentFolder&strConfig
GetIniFile "",strIniFile
HKEY_LOCAL_MACHINE=&H80000002

arrDists=split(Ini("::main::install",""),",")
strDistsFolderPath=Ini("::main::dists","")
strRegistrySection=Ini("::main::regsection","")
QUOT=chr(34)
Set objRegistry=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

For i=0 To UBound(arrDists)
  name=Ini("::"&arrDists(i)&"::name","") 
  version=Ini("::"&arrDists(i)&"::version","")
  folder=Ini("::"&arrDists(i)&"::folder","")
  exe=Ini("::"&arrDists(i)&"::exe","")
  keys=Ini("::"&arrDists(i)&"::keys","")

  strKeyPath=strRegistrySection&"\"&name
  objRegistry.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,"version",strValue
  If Len(strValue)>0 Then 
    If IsNewVersion(version,strValue) Then
      strCommand=QUOT&strDistsFolderPath&"\"&folder&"\"&version&"\"&exe&QUOT&" "&keys
      If Instr(1,exe,".vbs",1)>0 Then strCommand="wscript.exe "&strCommand 
      intResult=ExecuteCommand(strCommand)
      If intResult=0 Then
         objRegistry.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath,"version",version
      End If
    End If
  Else 
    strCommand=QUOT&strDistsFolderPath&"\"&folder&"\"&version&"\"&exe&QUOT&" "&keys
    If Instr(1,exe,".vbs",1)>0 Then strCommand="wscript.exe "&strCommand
    intResult=ExecuteCommand(strCommand)
    If intResult=0 Then
      objRegistry.CreateKey HKEY_LOCAL_MACHINE,strKeyPath
      objRegistry.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath,"version",version
    End If
  End If
Next

Function IsNewVersion(strVersionInFile,strVersionInRegistry)
  Dim i
  arrVerFile=split(strVersionInFile,".")
  arrVerReg=split(strVersionInRegistry,".")
  If UBound(arrVerFile)>UBound(arrVerReg) Then
    intVerMaxDigit=UBound(arrVerFile) 
  Else
    intVerMaxDigit=UBound(arrVerReg)
  End If
  For i=0 To intVerMaxDigit
    If CInt(arrVerReg(i))<CInt(arrVerFile(i)) Then
      IsNewVersion=True
      Exit For
    Else
      IsNewVersion=False
    End If
  Next 
End Function

Function ExecuteCommand(strCommand)
  Set objProcess = objWshShell.Exec(strCommand)
  Do While objProcess.Status=0
    s=s&objProcess.StdOut.ReadAll()
  Loop
  executeCommand = objProcess.ExitCode
End Function

' Подпрограммы и функции общего назначения (для работы с ini и т.д.)
Sub GetIniFile(ByVal strPreffix,ByVal strIniFile)
	Set dIni = CreateObject("Scripting.Dictionary")
  Dim objFS: Set objFS = CreateObject("Scripting.FileSystemObject")
  If objFS.FileExists(strIniFile) = False Then
    If InStr(strIniFile, ":\") = 0 And Left (strIniFile,2)<>"\\" Then
     'Искать в папке Windows, если нет в текущей
      strIniFile = objFS.GetSpecialFolder(0) & "\" & strIniFile
    End If
  End If
  Dim arrIni: arrIni = Split(objFS.OpenTextFile(strIniFile).ReadAll, vbNewLine)
  For Each strIniLine In arrIni
	  If strIniLine<>"" Then
	    If InStr(1,strIniLine,";",1) <> 1 Then
	      If InStr(1,strIniLine,"[",1) = 1 Then
	        strSection=Mid(strIniLine,2,InStr(1,strIniLine,"]",1)-2)
		    Else
		      intEqPos=InStr(1,strIniLine,"=",1)
		      If intEqPos > 0 Then
		        strParameter=Left(strIniLine,intEqPos-1)
            dIni.Add strPreffix&"::"&strSection&"::"&strParameter,Mid(strIniLine,intEqPos+1,len(strIniLine)) 
		      End If
		    End If
	    End If
	  End If
	Next   
End Sub
Function Ini(strSecParam,strDefault)
  If dIni.Exists(strSecParam) Then
    Ini = dIni.Item(strSecParam)
  Else
    Ini = strDefault
  End If
End Function
Function getCurrentFolder()
    Dim strCurrentFolder
    strCurrentFolder = WScript.ScriptFullName
    getCurrentFolder = Left(strCurrentFolder, InstrRev(strCurrentFolder, "\"))
End Function
