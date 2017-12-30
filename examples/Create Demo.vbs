
'//----------------------------------------------------------------------------
'! Create Demo
'! @version: 1.0 - 03 Mar 2013
'! @author Eduardo Mozart de Oliveira
'//
'// This script is provided "AS IS" with no warranties, confers no rights and 
'// is not supported by the authors or Deployment Artist.
'//
'//----------------------------------------------------------------------------

Dim CR, iRet, iRet2, iRet3, iRet4, s1, sVal, T, A1, A2, A3, i2

MsgBox "This demo creates keys and values.", 64, "Reg Class"

Include "..\RegClass.vbs"  '-- load the class.

Set CR = New CWMIReg


iRet = CR.CreateKey("HKCU\software\microsoft\blah1\blah2\blah3\")
  MsgBox "CreateKey returned: " & iRet
  
   If iRet <> 0 Then
      Set CR = Nothing
      WScript.Quit
   End If   
   
		 ' The script can set REG_BINARY keys as long as they are in the format used by a regedit.exe export or a binary Array.
         A1 = Array(34, 23, 1, 0, 0, 255, 32, 100) ' A1 = "hex:22,17,01,00,00,ff,20,64"
		 ' "hex:34,23,01,00,00,255,32,100" are not the same! Array is written in byte values.
      iRet = CR.SetValue("HKCU\software\microsoft\blah1\blah2\blah3\BinVal", A1, "REG_BINARY")
      iRet2 = CR.SetValue("HKCU\software\microsoft\blah1\blah2\blah3\StrVal", "Some string value.", "REG_SZ")
	  iRet3 = CR.SetValue("HKCU\software\microsoft\blah1\blah2\blah3\XStrVal","%APPDATA%","REG_EXPAND_SZ")
      iRet4 = CR.SetValue("HKCU\software\microsoft\blah1\blah2\blah3\NumVal", 60, "REG_DWORD")
        A2 = Array("first multi string", "second multi string", "third multi string")
      iRet5 = CR.SetValue("HKCU\software\microsoft\blah1\blah2\blah3\MultiVal", A2, "REG_MULTI_SZ")
	  iRet6 = CR.SetValue("HKCU\software\microsoft\blah1\blah2\blah3\NumVal64", 60, "REG_QWORD")

       
        MsgBox "Attempt to set 6 values, binary, string, expand string, dword, multi-string and qword. Return codes are:" & vbCrLf & iRet & vbCrLf & iRet2 & vbCrLf & iRet3 & vbCrLf & iRet4 & vbCrLf & iRet5 & vbCrLf & iRet6


          
 Set CR = Nothing
 
 Sub Include(FileName)
  Dim sPath, FSO2, TS2, Pt1, s2
        On Error Resume Next
    sPath = WScript.ScriptFullName
    Pt1 = InStrRev(sPath, "\")
   sPath = Left(sPath, Pt1) & FileName
     Set FSO2 = CreateObject("Scripting.FileSystemObject") 
        Set TS2 = FSO2.OpenTextFile(sPath, 1)
           s2 = TS2.ReadAll
           TS2.Close
        Set TS2 = Nothing
     Set FSO2 = Nothing
    ExecuteGlobal s2   
End Sub
