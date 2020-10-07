Const ForReading = 1
Const ForWriting = 2

Dim path1
dim drv1
dim mconfig32
dim mconfig64

Set objFSO = CreateObject("Scripting.FileSystemObject")

Set wshShell = CreateObject("WScript.Shell")
path1=wshShell.ExpandEnvironmentStrings("%WINDIR%")
drv1=wshShell.ExpandEnvironmentStrings("%SYSTEMDRIVE%") & "\"

if objFSO.FileExists(drv1 & "Users\defaultuser0\AppData\Local\Microsoft\Windows\Shell\DefaultLayouts.xml")=false Then
WScript.Echo "Not Windows 10 OS"
WScript.Quit
end if

mconfig64 = path1 & "\Microsoft.NET\Framework64\v4.0.30319\Config\machine.config"
mconfig32 = path1 & "\Microsoft.NET\Framework\v4.0.30319\Config\machine.config"


if objFSO.FileExists(mconfig64)=true then 

WScript.Echo "Creating backup: " & mconfig64 & " ..."

objFSO.CopyFile mconfig64, mconfig64 & ".bak", true

WScript.Echo "Modifying " & mconfig64 & " ..."

Set objFile = objFSO.OpenTextFile(mconfig64, ForReading)

strText = objFile.ReadAll

objFile.Close

strNewText = Replace(strText, "<runtime />", "<runtime><NetFx40_LegacySecurityPolicy enabled=""true""/></runtime>")

Set objFile = objFSO.OpenTextFile(mconfig64, ForWriting)

objFile.WriteLine strNewText
objFile.Close

WScript.Echo "Running caspol ..."

cmdx=path1 & "\Microsoft.NET\Framework64\v4.0.30319\caspol.exe -polchgprompt off"
wshShell.Run cmdx,0,true

cmdx=path1 & "\Microsoft.NET\Framework64\v4.0.30319\caspol.exe -all -reset"
wshShell.Run cmdx,0,true

cmdx=path1 & "\Microsoft.NET\Framework64\v4.0.30319\caspol.exe -polchgprompt on"
wshShell.Run cmdx,0,true

WScript.Echo "Copying generated security.config file ..."

wshShell.Run "xcopy.exe """ & path1 & "\Microsoft.NET\Framework64\v4.0.30319\Config\security.config" & """ """ & path1 & "\Microsoft.NET\Framework64\v2.0.50727\Config\security.config" & """ /R /Y",0,True
wshShell.Run "xcopy.exe """ & path1 & "\Microsoft.NET\Framework64\v4.0.30319\Config\security.config.cch" & """ """ & path1 & "\Microsoft.NET\Framework64\v2.0.50727\Config\security.config.cch" & """ /R /Y",0,True

wshShell.Run "xcopy.exe """ & path1 & "\Microsoft.NET\Framework64\v4.0.30319\Config\security.config" & """ """ & path1 & "\Microsoft.NET\Framework\v2.0.50727\Config\security.config" & """ /R /Y",0,True
wshShell.Run "xcopy.exe """ & path1 & "\Microsoft.NET\Framework64\v4.0.30319\Config\security.config.cch" & """ """ & path1 & "\Microsoft.NET\Framework\v2.0.50727\Config\security.config.cch" & """ /R /Y",0,True

WScript.Echo "PROCESS COMPLETE!"

elseif objFSO.FileExists(mconfig32)=true then 

 WScript.Echo "Creating backup: " & mconfig32 & " ..."

objFSO.CopyFile mconfig32, mconfig32 & ".bak", true

WScript.Echo "Modifying " & mconfig32 & " ..."

Set objFile = objFSO.OpenTextFile(mconfig32, ForReading)

strText = objFile.ReadAll

objFile.Close

strNewText = Replace(strText, "<runtime />", "<runtime><NetFx40_LegacySecurityPolicy enabled=""true""/></runtime>")

Set objFile = objFSO.OpenTextFile(mconfig64, ForWriting)

objFile.WriteLine strNewText
objFile.Close

WScript.Echo "Running caspol ..."

cmdx=path1 & "\Microsoft.NET\Framework\v4.0.30319\caspol.exe -polchgprompt off"
wshShell.Run cmdx,0,true

cmdx=path1 & "\Microsoft.NET\Framework\v4.0.30319\caspol.exe -all -reset"
wshShell.Run cmdx,0,true

cmdx=path1 & "\Microsoft.NET\Framework\v4.0.30319\caspol.exe -polchgprompt on"
wshShell.Run cmdx,0,true

WScript.Echo "Copying generated security.config file ..."

wshShell.Run "xcopy.exe """ & path1 & "\Microsoft.NET\Framework\v4.0.30319\Config\security.config" & """ """ & path1 & "\Microsoft.NET\Framework\v2.0.50727\Config\security.config" & """ /R /Y",0,True
wshShell.Run "xcopy.exe """ & path1 & "\Microsoft.NET\Framework\v4.0.30319\Config\security.config.cch" & """ """ & path1 & "\Microsoft.NET\Framework\v2.0.50727\Config\security.config.cch" & """ /R /Y",0,True

WScript.Echo "PROCESS COMPLETE!"

end if
