
'//----------------------------------------------------------------------------
'! Create-Delete Demo
'! @version: 1.0 - 03 Mar 2013
'! @author Eduardo Mozart de Oliveira
'//
'// This script is provided "AS IS" with no warranties, confers no rights and 
'// is not supported by the authors or Deployment Artist.
'//
'//----------------------------------------------------------------------------

Dim CR, iRet, iRet2, iRet3, iRet4, s1, sVal, T, A1, A2, A3, i2

'MsgBox "This demo creates keys and values, then deletes them.", 64, "Reg Class"

Include "..\RegClass.vbs"  '-- load the class.

Set CR = New CWMIReg


MsgBox "After you run this script and confirm it worked, click OK to test the deletion lines.", 64, "Reg Class"
     
MsgBox "Delete one of the values. Click OK to continue.", 64, "Reg Class"
    iRet = CR.Delete("HKCU\software\microsoft\blah1\blah2\blah3\StrVal")
		MsgBox "Delete value returns: " & iRet & "."

MsgBox "Delete HKCU\software\microsoft\blah1 key. Click OK to continue.", 64, "Reg Class"
	iRet2 = CR.Delete("HKCU\software\microsoft\blah1")
		MsgBox "Delete key returns: " & iRet & "."

       
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
