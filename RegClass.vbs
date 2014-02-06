
'--  WMI-derived Registry class for VBScript. The variables in this class all have "_" appended.
'--  Unfortunately, that makes the code a bit more difficult to read. It was done in order to
'--  avoid possible conflicts with variable names in code in scripts that use the class.

' Function Exists(Path)    (Returns type of data if value, "K" if key,  or "" if not found.)
' Function CreateKey(Path)    Returns 0 on success. Path can have "\" at end or not.
' Function EnumKeys(Path, ArrayOut)  Returns number of subkeys. ArrayOut contains subkey names.
' Function EnumVals(Path, AValsOut, ATypesOut)    AVals is array of value names. ATypesOut is array of value types.
' Function GetValue(Path)  - Returns value data for value or key. Returns data if found.
' Function SetValue(Path, ValData, Type) - set value data.
' Function Delete(Path) ' delete key or value.

'-- private functions:
 '  Function EnumKeysAll(Path, ArrayOut) Return list of all subkeys in a key. Function returns number of subkeys.
'          ArrayOut returns key paths.
'  EnumKeysAll has been made public, in case it might be useful, but it was really written for use in deleting keys.
' Function DeleteKey deletes all subkeys in path by first calling EnumKeysAll. It then deletes parent key.

'-- ########################## BEGIN CLASS #####################################
'-- All variables in this class have "_" appended. 
'-- That makes the code harder to read, but prevents possible clashes With global variables in the "parent" script.


Class CWMIReg
   Private HKCR_, HKCU_, HKLM_, HKU_
   Private Loc_, Provider_, Processor_

 Public Function Exists(Path_)
    Dim i2_, i3_, AVals_, ATypes_, s1_, Pt1_, sName_, Path1_, IsKey_
      Exists = ""
         On Error Resume Next
    s1_ = Path_
       If Right(s1_, 1) = "\" Then  '-- key.
           s1_ = Left(s1_, (len(s1_) - 1))
           IsKey_ = True
       Else
           IsKey_ = False     
       End If     
            Pt1_ = InStrRev(s1_, "\")
                 If Pt1_ = 0 Then Exit Function 
            sName_ = Right(s1_, (len(s1_) - Pt1_))
            Path1_ = Left(s1_, (Pt1_ - 1))
    
   
    If (IsKey_ = True) Then
        i2_ = EnumKeys(Path1_, AVals_)
    Else
        i2_ = EnumVals(Path1_, AVals_, ATypes_)
    End If      
       If (i2_ < 1) Then Exit Function '-- if i2_ is 0 or neg. then doesn't exist.
     
   For i3_ = 0 to i2_  - 1
      If UCase(AVals_(i3_)) = UCase(sName_) Then
         If (IsKey_ = False) Then
            Exists = ATypes_(i3_)
         Else
            Exists = "K"
         End If
           Exit For
       End If
   Next    
End Function

'--------------------------------------------- GetValue -----------------------------------------------------
Public Function GetValue(Path_)
  Dim Path1_, sKey_, LKey_, iRet_, Val_, Pt1_, ValName_, Typ_
  Dim Ctx_, Svc_, Reg1_, Inparams_, Outparams_
  
     On Error Resume Next
      Typ_ = Exists(Path_)
      If Len(Typ_) = 0 Then Exit Function

   Pt1_ = InStr(1, Path_, "\")
 If (Pt1_ > 0) Then
    sKey_ = Left(Path_, (Pt1_ - 1)) 
    Path1_ = Right(Path_, (len(Path_) - Pt1_))
    LKey_ = GetHKey(sKey_)
 Else
   LKey_ = GetHKey(Path_)
   Path1_ = ""
 End If  
 
   GetValue = -3 '-- os arch mismatch.
   	 If IsEmpty(Provider_) Then Exit Function
 
 
 Set Ctx_ = CreateObject("WbemScripting.SWbemNamedValueSet")
 Ctx_.Add "__ProviderArchitecture", Provider_
 Set Svc_ = Loc_.ConnectServer("","root\default","","",,,,Ctx_)
 Set Reg1_ = Svc_.Get("StdRegProv") 
 
 If (Len(Path1_) = 0) Then
      'iRet_ = Reg1_.GetStringValue(LKey_, Path1_, "", Val_)
      'GetValue = Val_
      Set Inparams_ = Reg1_.Methods_("GetStringValue").Inparameters
	  Inparams_.Hdefkey = LKey_
	  Inparams_.Ssubkeyname = Path1_ 
	  Inparams_.sValueName = ""
			
	  Set Outparams_ = Reg1_.ExecMethod_("GetStringValue", Inparams_,,Ctx_) 
	  iRet_ = Outparams_.ReturnValue
	  GetValue = Outparams_.sValue
  Else
     If (Typ_ = "K") Then
        Path1_ = Left(Path1_, (len(Path1_) - 1))
       'iRet_ = Reg1_.GetStringValue(LKey_, Path1_, "", Val_)
       'GetValue = Val_
		Set Inparams_ = Reg1_.Methods_("GetStringValue").Inparameters
		Inparams_.Hdefkey = LKey_
		Inparams_.Ssubkeyname = Path1_ 
		Inparams_.sValueName = ""
		
		Set Outparams_ = Reg1_.ExecMethod_("GetStringValue", Inparams_,,Ctx_) 
		iRet_ = Outparams_.ReturnValue
        GetValue = Outparams_.sValue
     Else
       Pt1_ = InStrRev(Path1_, "\")
       ValName_ = Right(Path1_, (len(Path1_) - Pt1_))
       Path1_ = Left(Path1_, (Pt1_ - 1))    
          Select Case Typ_
            Case "REG_SZ"
              'iRet_ = Reg1_.GetStringValue(LKey_, Path1_, ValName_, Val_)
               	Set Inparams_ = Reg1_.Methods_("GetStringValue").Inparameters
				Inparams_.Hdefkey = LKey_
				Inparams_.Ssubkeyname = Path1_ 
				Inparams_.sValueName = ValName_
				
				Set Outparams_ = Reg1_.ExecMethod_("GetStringValue", Inparams_,,Ctx_) 
				iRet_ = Outparams_.ReturnValue
				GetValue = Outparams_.sValue
            Case "REG_EXPAND_SZ"
              'iRet_ = Reg1_.GetExpandedStringValue(LKey_, Path1_, ValName_, Val_)
                Set Inparams_ = Reg1_.Methods_("GetExpandedStringValue").Inparameters
				Inparams_.Hdefkey = LKey_
				Inparams_.Ssubkeyname = Path1_ 
				Inparams_.sValueName = ValName_
				
				Set Outparams_ = Reg1_.ExecMethod_("GetExpandedStringValue", Inparams_,,Ctx_) 
				iRet_ = Outparams_.ReturnValue
				GetValue = Outparams_.sValue
            Case "REG_BINARY"
              'iRet_ = Reg1_.GetBinaryValue(LKey_, Path1_, ValName_, Val_)
                Set Inparams_ = Reg1_.Methods_("GetBinaryValue").Inparameters
				Inparams_.Hdefkey = LKey_
				Inparams_.Ssubkeyname = Path1_ 
				Inparams_.sValueName = ValName_
				
				Set Outparams_ = Reg1_.ExecMethod_("GetBinaryValue", Inparams_,,Ctx_) 
				iRet_ = Outparams_.ReturnValue
				GetValue = Outparams_.uValue
            Case "REG_DWORD"
              'iRet_ = Reg1_.GetDWORDValue(LKey_, Path1_, ValName_, Val_)
                Set Inparams_ = Reg1_.Methods_("GetDWORDValue").Inparameters
				Inparams_.Hdefkey = LKey_
				Inparams_.Ssubkeyname = Path1_ 
				Inparams_.sValueName = ValName_
				
				Set Outparams_ = Reg1_.ExecMethod_("GetDWORDValue", Inparams_,,Ctx_) 
				iRet_ = Outparams_.ReturnValue
				GetValue = Outparams_.uValue
            Case "REG_MULTI_SZ"  
              'iRet_ = Reg1_.GetMultiStringValue(LKey_, Path1_, ValName_, Val_)
                Set Inparams_ = Reg1_.Methods_("GetMultiStringValue").Inparameters
				Inparams_.Hdefkey = LKey_
				Inparams_.Ssubkeyname = Path1_ 
				Inparams_.sValueName = ValName_
				
				Set Outparams_ = Reg1_.ExecMethod_("GetMultiStringValue", Inparams_,,Ctx_) 
				iRet_ = Outparams_.ReturnValue
				GetValue = Outparams_.sValue
			Case "REG_QWORD"
			    Set Inparams_ = Reg1_.Methods_("GetQWORDValue").Inparameters
				Inparams_.Hdefkey = LKey_
				Inparams_.Ssubkeyname = Path1_ 
				Inparams_.sValueName = ValName_
				
				Set Outparams_ = Reg1_.ExecMethod_("GetQWORDValue", Inparams_,,Ctx_) 
				iRet_ = Outparams_.ReturnValue
				GetValue = Outparams_.uValue
            Case Else
               Exit Function  
          End Select
        
     End If   
 End If  
End Function

'------------------------ Enum Keys -----------------------------------

Public Function EnumKeys(Path_, AKeys_)
  Dim iRet_, sKey_, LKey_, Pt1_, Pt2, Path1_
  Dim Ctx_, Svc_, Reg1_, Inparams_, Outparams_
  
    On Error Resume Next
  EnumKeys = -1 '-- invalid Path
    Pt1_ = InStr(1, Path_, "\")
      If (Pt1_ = 0) Then 
         sKey_ = Path_
      Else 
         sKey_ = left(Path_, (Pt1_ - 1))
      End If   
  LKey_ = GetHKey(sKey_)
  EnumKeys = -2 '-- invalid hkey.
     If (LKey_ = 0) Then Exit Function
  EnumKeys = -3 '-- os arch mismatch.
  	 If IsEmpty(Provider_) Then Exit Function
     
    If (sKey_ = Path_) Then
       Path1_ = ""
    Else
       Path1_ = Right(Path_, (len(Path_) - Pt1_))
       If Right(Path1_, 1) = "\" Then Path1_ = Left(Path1_, (len(Path1_) - 1))  
    End If  
    
   'iRet_ = Reg1_.EnumKey(LKey_, Path1_, AKeys_)
    Set Ctx_ = CreateObject("WbemScripting.SWbemNamedValueSet")
	Ctx_.Add "__ProviderArchitecture", Provider_
	Set Svc_ = Loc_.ConnectServer("","root\default","","",,,,Ctx_)
	Set Reg1_ = Svc_.Get("StdRegProv") 
		
	Set Inparams_ = Reg1_.Methods_("EnumKey").Inparameters
	Inparams_.hDefKey = LKey_
	Inparams_.sSubKeyName = Path1_ 
	
	Set Outparams_ = Reg1_.ExecMethod_("EnumKey", Inparams_,,Ctx_) 
	iRet_ = Outparams_.ReturnValue
	AKeys_ = Outparams_.snames
	

   Select Case iRet_
     Case 0
        If (isArray(AKeys_) = False) Then 
           EnumKeys = 0
        Else  
           EnumKeys = UBound(AKeys_) + 1
        End If  
     Case 2 '-- invalid key Path
         EnumKeys = -3  
     Case -2147217405   '--  access denied  H80041003
         EnumKeys = -4
     Case Else
         EnumKeys = iRet_  '-- some other error.  
   End Select
End Function

'---------------------------------------------------- Enum Values -----------------------------
Public Function EnumVals(Path_, Vals_, Types_)
  Dim sKey_, Pt1_, Pt2_, LKey_, iRet_, Path1_, iCnt_, Val_
  Dim Ctx_, Svc_, Reg1_, Inparams_, Outparams_
  
    On Error Resume Next
     EnumVals = -1 '-- invalid Path.
     Pt1_ = InStr(1, Path_, "\")
    If (Pt1_ = 0) Then 
        sKey_ = Path_
        Path1_ = ""
    Else     
       sKey_ = Left(Path_, (Pt1_ - 1))
       Path1_ = Right(Path_, (len(Path_) - Pt1_))
    End If   
       LKey_ = GetHKey(sKey_)
     EnumVals = -2 '-- invalid hkey.
       If (LKey_ = 0) Then Exit Function
     EnumVals = -3 '-- os arch mismatch.
       If IsEmpty(Provider_) Then Exit Function
       
    
      If Right(Path1_, 1) = "\" Then Path1_ = Left(Path1_, (len(Path1_) - 1))  
    'iRet_ = Reg1_.EnumValues(LKey_, Path1_, Vals_, Types_)
    Set Ctx_ = CreateObject("WbemScripting.SWbemNamedValueSet")
	Ctx_.Add "__ProviderArchitecture", Provider_
	Set Svc_ = Loc_.ConnectServer("","root\default","","",,,,Ctx_)
	Set Reg1_ = Svc_.Get("StdRegProv") 
		
	Set Inparams_ = Reg1_.Methods_("SetStringValue").Inparameters
	Inparams_.Hdefkey = LKey_
	Inparams_.Ssubkeyname = Path1_ 
	
	Set Outparams_ = Reg1_.ExecMethod_("EnumValues", Inparams_,,Ctx_) 
	iRet_ = Outparams_.ReturnValue
	Vals_ = Outparams_.snames
	Types_ = Outparams_.Types

   Select Case iRet_
     Case 0
        If (isArray(Vals_) = False) Then 
           EnumVals = 0  '-- no values in key.
        Else  '-- values found. convert types from numeric to letters.
           EnumVals = UBound(Vals_) + 1
            For iCnt_ = 0 to UBound(Types_)
               Val_ = Types_(iCnt_)
               Types_(iCnt_) = ConvertType(Val_)
            Next   
        End If  
     Case 2 '-- invalid key Path
        EnumVals = -4  
     Case -2147217405   '--  access denied  H80041003
        EnumVals = -5
     Case Else
        EnumVals = iRet_  '-- some other error.  
   End Select
End Function

Public Function SetValue(Path_, ValData_, TypeIn_)
   Dim Path1_, sKey_, LKey_, iRet_, Pt1_, Pt2_, ValName_, Typ_
   Dim Ctx_, Svc_, Reg1_, Inparams_, Outparams_
   
   '-- Typ_
   Dim vbArrayInteger, vbArrayString
   Dim RegEx_, REsult_ '-- reg_expand_sz
   
   Dim arrHexValues_, arrDecValues_, n '-- reg_binary
   
 
   
  On Error Resume Next
   SetValue = -1  '-- defaults to invalid path error.
   If Len(TypeIn_) = 0 Then 
     Typ_ = Exists(Path_)
       If Len(Typ_) = 0 Then
       		' vbArray = 8192
       		' vbInteger = 2
       		' vbString = 8
       		vbArrayInteger = 8194 ' 8192 + 2
   			vbArrayString = 8200 ' 8192 + 8
       		
       		If VarType(ValData_) = vbArrayInteger or LCase(Left(ValData_, Len("hex:"))) = "hex:" Then
				Typ_ = "REG_BINARY"
			ElseIf VarType(ValData_) = vbString Then
				If InStr(ValData_, "%") Then
					' http://regexr.com
					' String should start with % and end with %.
					' It can not contain < > | & ^ (http://www.microsoft.com/resources/documentation/windows/xp/all/proddocs/en-us/set.mspx?mfr=true)
					Set RegEx_ = New RegExp
					RegEx_.Global = True
					RegEx_.Pattern = "\B%([^\<\>\|\&\^]{1,})%\B"
					Set REsult_ = RegEx_.Execute(ValData_)
					
					If REsult_.Count > 0 Then
						'MsgBox Result.Count & vbCrLf & Result.Item(0).Value
						Typ_ = "REG_EXPAND_SZ"
					Else
						Typ_ = "REG_SZ"
					End If	
				Else
					Typ_ = "REG_SZ"
				End If
			ElseIf VarType(ValData_) = vbInteger or VarType(ValData_) = vbLong or VarType(ValData_) = vbBoolean then 
				Typ_ = "REG_DWORD"
			ElseIf VarType(ValData_) = vbArrayString Then
				Typ_ = "REG_MULTI_SZ"
			ElseIf VarType(ValData_) = vbCurrency Then
				Typ_ = "REG_QWORD"
			End If 
	   End If
   Else
      If IsNumeric(TypeIn_) Or Len(TypeIn_) = 1 Then 
         Typ_ = ConvertType(TypeIn_)
      Else
         Typ_ = UCase(TypeIn_)
      End If    
   End If    
   
     Pt1_ = InStr(1, Path_, "\")
    If (Pt1_ = 0) Then 
       sKey_ = Path_
       Path1_ = ""
    Else
       sKey_ = Left(Path_, (Pt1_ - 1))
       Path1_ = Right(Path_, (len(Path_) - Pt1_))
    End If   
        LKey_ = GetHKey(sKey_)
        If (LKey_ = 0) Then 
           SetValue = -2  '-- invalid hKey.
           Exit Function
        ElseIf IsEmpty(Provider_) Or Processor_ = "x86" And Typ_ = "REG_QWORD" Then
           SetValue = -3 '-- os arch mismatch
           Exit Function
        End If   
  
   '-- create a key if it does not exist ------------
  iRet_ = EnumKeys(sKey_ & "\" & Left(Path1_, InStrRev(Path1_, "\")-1), AKeys)
  If iRet_ = -3 Then 
  	iRet_ = CreateKey(sKey_ & "\" & Left(Path1_, InStrRev(Path1_, "\")-1))
  	If iRet_ <> 0 Then
  		SetValue = iRet_
  		Exit Function
  	End If
  End If
  
  Set Ctx_ = CreateObject("WbemScripting.SWbemNamedValueSet")
  Ctx_.Add "__ProviderArchitecture", Provider_
  Set Svc_ = Loc_.ConnectServer("","root\default","","",,,,Ctx_)
  Set Reg1_ = Svc_.Get("StdRegProv") 
  
  If (Typ_ = "K") Or (Right(Path1_, 1) = "\") Then
      If Right(Path1_, 1) = "\" Then Path1_ = Left(Path1_, (len(Path1_) - 1))
        'iRet_ = Reg1_.SetStringValue(LKey_, Path1_, "", ValData_)
		Set Inparams_ = Reg1_.Methods_("SetStringValue").Inparameters
		Inparams_.Hdefkey = LKey_
		Inparams_.Ssubkeyname = Path1_ 
		Inparams_.sValueName = ""
		Inparams_.sValue = ValData_
		
		Set Outparams_ = Reg1_.ExecMethod_("SetStringValue", Inparams_,,Ctx_) 
		iRet_ = Outparams_.ReturnValue
  Else
    Err.clear
    On Error Resume Next
     Pt1_ = InStrRev(Path1_, "\")
     ValName_ = Right(Path1_, (len(Path1_) - Pt1_))
     Path1_ = Left(Path1_, (Pt1_ - 1))     
        Select Case Typ_
          Case "REG_SZ"
             'iRet_ = Reg1_.SetStringValue(LKey_, Path1_, ValName_, ValData_)
             Set Inparams_ = Reg1_.Methods_("SetStringValue").Inparameters
			 Inparams_.Hdefkey = LKey_
			 Inparams_.Ssubkeyname = Path1_ 
			 Inparams_.sValueName = ValName_
			 Inparams_.sValue = ValData_
				
			 Set Outparams_ = Reg1_.ExecMethod_("SetStringValue", Inparams_,,Ctx_)
			 iRet_ = Outparams_.ReturnValue 
          Case "REG_EXPAND_SZ"
             'iRet_ = Reg1_.SetExpandedStringValue(LKey_, Path1_, ValName_, ValData_)
             Set Inparams_ = Reg1_.Methods_("SetExpandedStringValue").Inparameters
			 Inparams_.Hdefkey = LKey_
			 Inparams_.Ssubkeyname = Path1_ 
			 Inparams_.sValueName = ValName_
			 Inparams_.sValue = ValData_
				
			 Set Outparams_ = Reg1_.ExecMethod_("SetExpandedStringValue", Inparams_,,Ctx_)
			 iRet_ = Outparams_.ReturnValue 
          Case "REG_BINARY"
             'iRet_ = Reg1_.SetBinaryValue(LKey_, Path1_, ValName_, ValData_)
             Set Inparams_ = Reg1_.Methods_("SetBinaryValue").Inparameters
			 Inparams_.Hdefkey = LKey_
			 Inparams_.Ssubkeyname = Path1_ 
			 Inparams_.sValueName = ValName_
			 
			 If IsArray(ValData_) Then
			 	Inparams_.uValue = ValData_
			 ElseIf LCase(Left(ValData_, Len("hex:"))) = "hex:" Then
			 	'Example:   ValData_ = "hex:23,00,41,00,43,00,42,00,6c,00"
				arrHexValues_ = Split(Replace(ValData_, "hex:", ""), ",")
				arrDecValues_ = DecimalNumbers(arrHexValues_)
				Inparams_.uValue = arrDecValues_
			 Else
			 	SetValue = -6 '-- type mismatch. incoming value not valid.
			 	Exit Function
			 End If
				
			 Set Outparams_ = Reg1_.ExecMethod_("SetBinaryValue", Inparams_,,Ctx_)
			 iRet_ = Outparams_.ReturnValue 
          Case "REG_DWORD"
             'iRet_ = Reg1_.SetDWORDValue(LKey_, Path1_, ValName_, ValData_)
             Set Inparams_ = Reg1_.Methods_("SetDWORDValue").Inparameters
			 Inparams_.Hdefkey = LKey_
			 Inparams_.Ssubkeyname = Path1_ 
			 Inparams_.sValueName = ValName_
			 Inparams_.uValue = ValData_
				
			 Set Outparams_ = Reg1_.ExecMethod_("SetDWORDValue", Inparams_,,Ctx_)
			 iRet_ = Outparams_.ReturnValue 
          Case "REG_MULTI_SZ"  
             'iRet_ = Reg1_.SetMultiStringValue(LKey_, Path1_, ValName_, ValData_)
             Set Inparams_ = Reg1_.Methods_("SetMultiStringValue").Inparameters
			 Inparams_.Hdefkey = LKey_
			 Inparams_.Ssubkeyname = Path1_ 
			 Inparams_.sValueName = ValName_
			 Inparams_.sValue = ValData_
				
			 Set Outparams_ = Reg1_.ExecMethod_("SetMultiStringValue", Inparams_,,Ctx_)
			 iRet_ = Outparams_.ReturnValue 
          Case Else
             SetValue = -7
             Exit Function 
        End Select
  End If   
   
    If (Err.number = -2147217403) Then 
       SetValue = -6  '-- type mismatch. incoming value not valid.
       Exit Function
    End If
   Select Case iRet_
     Case 0
          SetValue = 0   'success.
     'Case 2 '-- invalid key path 
     '     SetValue = -4  
     Case -2147217405   '--  access denied  H80041003
          SetValue = -5
     Case Else
          SetValue = iRet_  '-- some other error.  
    End Select
End Function

 '-- create a key. Path can be with "\" at end or not, but must not have "\" if path is an HKey like "HKLM". ------------
Public Function CreateKey(Path_)
   Dim sKey_, LKey_, Path1_, iRet_, Pt1_
   Dim Ctx_, Svc_, Reg1_, Inparams_, Outparams_
  
   On Error Resume Next
    CreateKey = -1
   Pt1_ = InStr(1, Path_, "\") 
   If (Pt1_ = 0) Then
      sKey_ = Path_
      Path1_ = ""
   Else   
     sKey_ = Left(Path_, Pt1_ - 1)
     Path1_ = Right(Path_, (len(Path_) - Pt1_))
   End If  
    CreateKey = -2
  LKey_ = GetHKey(sKey_)
     If (LKey_ = 0) Then Exit Function
    CreateKey = -3
     If IsEmpty(Provider_) Then Exit Function
     
  If Right(Path1_, 1) = "\" Then Path1_ = Left(Path1_, (len(Path1_) - 1))
  'iRet_ = Reg1_.CreateKey(LKey_, Path1_)
  Set Ctx_ = CreateObject("WbemScripting.SWbemNamedValueSet")
  Ctx_.Add "__ProviderArchitecture", Provider_
  Set Svc_ = Loc_.ConnectServer("","root\default","","",,,,Ctx_)
  Set Reg1_ = Svc_.Get("StdRegProv") 
		
  Set Inparams_ = Reg1_.Methods_("CreateKey").Inparameters
  Inparams_.Hdefkey = LKey_
  Inparams_.Ssubkeyname = Path1_ 
	
  Set Outparams_ = Reg1_.ExecMethod_("CreateKey", Inparams_,,Ctx_) 
  iRet_ = Outparams_.ReturnValue
  
   Select Case iRet_
        Case 0			   	  '-- OK.
           CreateKey = 0  
        Case 2             	  '-- invalid key
           CreateKey = -4
        Case -2147217405      '--  access denied  H80041003
           CreateKey = -5
        Case Else
           CreateKey = iRet_  '-- some other error.  
    End Select

End Function

 '-------------- Delete ---- delete a key or value. Add "\" for keys. -------------------
Public Function Delete(Path_)
  Dim sKey_, LKey_, Path1_, Pt1_, ValName_, iRet_
  Dim Ctx_, Svc_, Reg1_, Inparams_, Outparams_

      On Error Resume Next
      Delete = -1 'invalid path.
	    Pt1_ = InStr(1, Path_, "\")      
	      If Pt1_ = 0 Then Exit Function
  sKey_ = Left(Path_, (Pt1_ - 1))     
  Path1_ = Right(Path_, (len(Path_) - Pt1_))
     Delete = -2  ' invalid hkey.
  LKey_ = GetHKey(sKey_)
    If (LKey_ = 0) Then Exit Function
     Delete = -3
    If IsEmpty(Provider_) Then Exit Function
    
    If Right(Path1_, 1) = "\" Then
       Path1_ = Left(Path1_, (len(Path1_) - 1))
       iRet_ = DeleteKey(Path_)  
    Else
       Pt1_ = InStrRev(Path1_, "\")
       ValName_ = Right(Path1_, (len(Path1_) - Pt1_))
       Path1_ = Left(Path1_, (Pt1_ - 1))
       'iRet_ = Reg1_.DeleteValue(LKey_, Path1_, ValName_)
       Set Ctx_ = CreateObject("WbemScripting.SWbemNamedValueSet")
	   Ctx_.Add "__ProviderArchitecture", Provider_
	   Set Svc_ = Loc_.ConnectServer("","root\default","","",,,,Ctx_)
	   Set Reg1_ = Svc_.Get("StdRegProv") 
	   
	   Set Inparams_ = Reg1_.Methods_("DeleteValue").Inparameters
	   Inparams_.Hdefkey = LKey_
	   Inparams_.Ssubkeyname = Path1_ 
	   Inparams_.sValueName = ValName_
		
	   Set Outparams_ = Reg1_.ExecMethod_("DeleteValue", Inparams_,,Ctx_) 
	   iRet_ = Outparams_.ReturnValue
    End If
     
   Select Case iRet_
        Case 0
          Delete = 0
        Case -1  ' returned from DeleteKey
          Delete = -1
        Case -2  ' returned from DeleteKey
          Delete = -2
        Case 2
          Delete = -4 ' value does not exists.
        Case -2147217405   '--  access denied  H80041003
          Delete = -5
       Case Else
          Delete = iRet_  '-- some other error.  
    End Select
  
End Function

Public Function TestPath(PathIn)
  ' HKCU\ControlPanel\Desktop = 3
  ' HKLM\Software\Microsoft\Windows NT\CurrentVersion = 5
  
  Dim sKey_, Path_, sHKUSubKey_
  Dim PathIn_, Cnt_, Pt1_, Pt2_
      On Error Resume Next
    If InStr(PathIn, "\") = 0 Then TestPath = 0 : Exit Function
      
    sKey_ = Left(PathIn, InStr(PathIn, "\")-1) : sKey_ = UCase(sKey_)
    Path_ = Mid(PathIn, InStr(PathIn, "\")+1, Len(PathIn)) : Path_ = UCase(Path_)
    	If Path_ = "" Then TestPath = 0 : Exit Function
    
    ' Consider HKEY_USERS\S-1-5-21-1501084202-2593169170-243912787-500 as HKCU
    ' Consider HKEY_USERS\S-1-5-21-1501084202-2593169170-243912787-500_Classes as HKCR
	If Left(sKey_,Len("HKU\")) = "HKU\" Or _
		Left(sKey_, Len("HKEY_USERS\")) = "HKEY_USERS\" Then
		
		' S-1-5-21-1501084202-2593169170-243912787-500 
		' S-1-5-21-1501084202-2593169170-243912787-500_Classes
		sHKUSubKey_ = Left(Path_, InStr(Path_, "\")-1) : sHKUSubKey_ = UCase(sHKUSubKey_)
		
		If Right(sHKUSubKey, Len("_CLASSES")) = "_CLASSES" Then
			sKey = "HKCR"
		Else
			sKey = "HKCU"
		End If 
	End If
	
	PathIn_ = sKey_ & "\" & Path_
	Cnt_ = 0
    Pt1_ = 1
      Do
        Pt2_ = InStr(Pt1_, PathIn_, "\")
          If (Pt2_ = 0) Then
              Exit Do
          Else
             Pt1_ = Pt2_ + 1
             Cnt_ = Cnt_ + 1
             If (Pt1_ > Len(PathIn_)) Then Exit Do
          End If
      Loop   
        TestPath = Cnt_
End Function

 '--------------------------------------- Private Functions ----------------------------

Private Sub Class_Initialize()
     On Error Resume Next
  HKCR_ = &H80000000
  HKCU_ = &H80000001
  HKLM_ = &H80000002
  HKU_ = &H80000003

  'Set Reg1_ = GetObject("winMgMts:root\default:StdRegProv")
  ' If (Err.number <> 0) Then
  '   Err.Raise 1, "WMIReg Class", "Failed to access WMI StdRegProv object. Class cannot function."
  ' End If
  
  Set Loc_ = CreateObject("Wbemscripting.SWbemLocator")
  ' http://csi-windows.com/toolkit/csi-getosbits
  Processor_ = GetValue("HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\PROCESSOR_ARCHITECTURE")
End Sub 

Private Sub Class_Terminate()
  'Set Reg1_ = Nothing
  Set Loc_ = Nothing
End Sub

Private Function GetHKey(sKey1_) 
  If Right(sKey1_, 2) = "64" Then
  	If Processor_ = "AMD64" Then
  		Provider_ = 64
  	Else
  		Provider_ = Empty
  	End If
  Else
  	Provider_ = 32
  End If

  Select Case UCase(Replace(sKey1_, "64", ""))
	   Case "HKLM" 
		GetHKey = HKLM_
	   Case "HKCU" 
		GetHKey = HKCU_
	   Case "HKCR" 
		GetHKey = HKCR_
	   Case "HKU"  
		GetHKey = HKU_ 
	      
	   Case "HKEY_LOCAL_MACHINE"  
		GetHKey = HKLM_
	   Case "HKEY_CURRENT_USER"   
		GetHKey = HKCU_
	   Case "HKEY_CLASSES_ROOT"   
		GetHKey = HKCR_
	   Case "HKEY_USERS"          
		GetHKey = HKU_ 

	   Case Else 
		GetHKey = 0
  End Select      
End Function

Private Function ConvertType(TypeIn)
    On Error Resume Next
  Select Case TypeIn
     Case 1
        ConvertType = "REG_SZ"
     Case 2
        ConvertType = "REG_EXPAND_SZ"
     Case 3
        ConvertType = "REG_BINARY"
     Case 4
        ConvertType = "REG_DWORD"
     Case 7
        ConvertType = "REG_MULTI_SZ"
     Case 11
     	ConvertType = "REG_QWORD"
     
     ' Compatibility Layer 
     Case "S"
     	ConvertType = "REG_SZ"
     Case "X"
     	ConvertType = "REG_EXPAND_SZ"
     Case "B"
     	ConvertType = "REG_BINARY"
     Case "D"
     	ConvertType = "REG_DWORD"
     Case "M"
     	ConvertType = "REG_MULTI_SZ"
     Case "Q"
     	ConvertType = "REG_QWORD"
     
     Case Else
       ConvertType = ""
  End Select        
End Function

Public Function EnumKeysAll(Path_, AKeys_)
  Dim sList_, s2_, AK1_, AK3_, AK2_(), iRet1_, i2_, UB_, iRet2_, i3_, Path1_
      Path1_ = Path_
         On Error Resume Next  
      If Right(Path1_, 1) <> "\" Then Path1_ = Path1_ & "\"
     iRet1_ = EnumKeys(Path_, AK1_)  
        If iRet1_ > 0 Then
          ReDim Preserve AK2_(iRet1_ - 1)
            For i2_ = 0 to iRet1_ - 1 
               AK2_(i2_) = Path1_ & AK1_(i2_) & "\"              
               iRet2_ = EnumKeysAll(AK2_(i2_), AK3_)
                 If (iRet2_ > 0) Then
                     UB_ = UBound(AK2_)
                     ReDim Preserve AK2_(UB_ + iRet2_)
                       For i3_ = 1 to iRet2_
                          AK2_(UB_ + i3_) = AK3_(i3_ - 1)
                      Next      
                 End If        
            Next
        End If 
        AKeys_ = AK2_
    EnumKeysAll = 0  
    EnumKeysAll = UBound(AKeys_) + 1
End Function

Private Function DeleteKey(Path_)
  Dim i3_, i4_, A1_, iRet_, Pt1_, s2_, hK_, sK_
  Dim Ctx_, Svc_, Reg1_, Inparams_, Outparams_
  
   On Error Resume Next
     i3_ = TestPath(Path_)
     If (i3_ < 3) Then 
        DeleteKey = -1  '-- invalid path. Key is top level. function fails on the premise that attempted deletion was a mistake.
        Exit Function
     End If
     
  Pt1_ = InStr(Path_, "\")
  sK_ = Left(Path_, (Pt1_ - 1))
  hK_ = GetHKey(sK_)
    If (hK_ = 0) Then
       DeleteKey = -2  '-- invalid hkey.
       Exit Function
    End If   
  
  
  Set Ctx_ = CreateObject("WbemScripting.SWbemNamedValueSet")
  Ctx_.Add "__ProviderArchitecture", Provider_
  Set Svc_ = Loc_.ConnectServer("","root\default","","",,,,Ctx_)
  Set Reg1_ = Svc_.Get("StdRegProv") 

  iRet_ = EnumKeysAll(Path_, A1_)
    If (iRet_ > 0) Then
       For i3_ = UBound(A1_) to 0 Step -1
          s2_ = A1_(i3_)
          s2_ = Right(s2_, (len(s2_) - Pt1_))  '-- remove hkey string from path.
          If Right(s2_, 1) = "\" Then s2_ = Left(s2_, (len(s2_) - 1))
          'i4_ = Reg1_.DeleteKey(hK_, s2_)
          Set Inparams_ = Reg1_.Methods_("DeleteKey").Inparameters
  		  Inparams_.Hdefkey = hK_
          Inparams_.Ssubkeyname = s2_ 
          
          Set Outparams_ = Reg1_.ExecMethod_("DeleteKey", Inparams_,,Ctx_) 
  		  i4_ = Outparams_.ReturnValue
       Next
    End If
   
   s2_ = Right(Path_, (len(Path_) - Pt1_))   
     If Right(s2_, 1) = "\" Then s2_ = Left(s2_, (len(s2_) - 1))
   'DeleteKey = Reg1_.DeleteKey(hK_, s2_)
   Set Inparams_ = Reg1_.Methods_("DeleteKey").Inparameters
   Inparams_.Hdefkey = hK_
   Inparams_.Ssubkeyname = s2_  
   
   Set Outparams_ = Reg1_.ExecMethod_("DeleteKey", Inparams_,,Ctx_) 
   DeleteKey = Outparams_.ReturnValue
End Function

Private Function DecimalNumbers(arrHex_)
    On Error Resume Next
    
    ' from: http://www.petri.co.il/forums/showthread.php?t=46158
    Dim i, strDecValues_
    For i = 0 to Ubound(arrHex_)
        If isEmpty(strDecValues_) Then
            strDecValues_ = CLng("&H" & arrHex_(i))
        Else
            strDecValues_ = strDecValues_ & "," & CLng("&H" & arrHex_(i))
        End If
    Next
      
    DecimalNumbers = split(strDecValues_, ",")
End Function

End Class