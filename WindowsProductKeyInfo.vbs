'* Script: Windows Key Information
'* http://www.proy.info            *
'* Modified On 02/01/2017          *
'***********************************

Dim strComputer, objWMIService, objItem, Caption, colItems, ProductData, OSVersion, InstallDate, RegisteredUser, ProductID, systemOsType
Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")

strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem",,48)
For Each objItem in colItems
    Caption = objItem.Caption
	OSVersion = Trim(objItem.Caption) & " (Build " & objItem.Version & ")"
	
	dtmConvertedDate.Value = objItem.InstallDate
	InstallDate = dtmConvertedDate.GetVarDate

	RegisteredUser = objItem.RegisteredUser
	ProductID = Trim(objItem.SerialNumber)

Next


'Find Processor architecture
Set WshShell = CreateObject("WScript.Shell")
OsType = WshShell.RegRead("HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\PROCESSOR_ARCHITECTURE")
 
If OsType = "x86" then
systemOsType = "32-bit Operating System"
elseif OsType = "AMD64" then
systemOsType = "64-bit Operating System"
end if


Function WindowsKey

	Set WshShell = CreateObject("WScript.Shell")
	Key = WshShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\DigitalProductId")

    Const KeyOffset = 52
    Dim isWin8, Maps, i, j, Current, KeyOutput, Last, keypart1, insert
    'Check if OS is Windows 8
    isWin8 = (Key(66) \ 6) And 1
    Key(66) = (Key(66) And &HF7) Or ((isWin8 And 2) * 4)
    i = 24
    Maps = "BCDFGHJKMPQRTVWXY2346789"
    Do
       	Current= 0
        j = 14
        Do
           Current = Current* 256
           Current = Key(j + KeyOffset) + Current
           Key(j + KeyOffset) = (Current \ 24)
           Current=Current Mod 24
            j = j -1
        Loop While j >= 0
        i = i -1
        KeyOutput = Mid(Maps,Current+ 1, 1) & KeyOutput
        Last = Current
    Loop While i >= 0
    keypart1 = Mid(KeyOutput, 2, Last)
    insert = "N"
    KeyOutput = Replace(KeyOutput, keypart1, keypart1 & insert, 2, 1, 0)
    If Last = 0 Then KeyOutput = insert & KeyOutput
    WindowsKey = Mid(KeyOutput, 1, 5) & "-" & Mid(KeyOutput, 6, 5) & "-" & Mid(KeyOutput, 11, 5) & "-" & Mid(KeyOutput, 16, 5) & "-" & Mid(KeyOutput, 21, 5)


End Function

ProductData = OSVersion & vbNewLine & systemOsType & vbNewLine & vbNewLine & "Install Date: " & InstallDate & vbNewLine & "Registered To: " & RegisteredUser & vbNewLine & "Windows PID: " & ProductID & vbNewLine & "Windows Key: " & WindowsKey

'Show messbox if save to a file
If vbYes = MsgBox(ProductData  & vblf & vblf & "Click Yes to save these information to a file?", vbYesNo + vbQuestion, "BackUp Windows Key Information") then
    Save ProductData
End If

'Save data to a file
Function Save(Data)
    Dim fso, fName, txt,objshell,UserName
    Set objshell = CreateObject("wscript.shell")
    'Get current user name
    UserName = objshell.ExpandEnvironmentStrings("%UserName%")
    'Create a text file on desktop
    fName = "C:\Users\" & UserName & "\Desktop\WindowsKeyInfo.txt"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txt = fso.CreateTextFile(fName)
    txt.Writeline Data
    txt.Close
End Function
