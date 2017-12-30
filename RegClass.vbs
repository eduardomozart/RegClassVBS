'-----------------------------------------------------------------------------
' Class: CWMIReg
' 
' WMI-derived Registry class for VBScript. 
'
' All variables in this class have "_" appended.
' Unfortunately, that makes the code a bit more difficult to read. It was done in order to
' avoid possible conflicts with variable names in code in scripts that use the class.
'
' Version:
'
' 	3.0.2
'
' Author:
'
' 	* Joe Priestley (2004-2008)
'	* Eduardo Mozart de Oliveira (2014-2017)
'
' Data Type:
'
' 	REG_SZ        - A string value.
' 	REG_EXPAND_SZ - An expanded string data value. The environment variable specified in the string must exist for the string to be expanded when you call GetValue.
' 	REG_BINARY    - An array of binary data values. The script can set REG_BINARY keys as long as they are in the format used by a regedit.exe export or a bytes Array. See *examples\Create Demo.vbs*.
' 	REG_DWORD     - A numeric data value.
' 	REG_MULTI_SZ  - A list of strings. The SetValue method accepts an array of strings as the parameter that determines the values of the entry. Note that if you use the SetValue method to append to an existing multistring-valued entry rather than create a new one, you have to first use the GetValue method to retrieve the existing list of strings. This is because SetValue overwrites any existing value. 
' 	REG_QWORD     - A QWORD data value for the named value.
'
' Error Codes:
'
' Most of the functions return error codes. There are standard error codes (negative numbers) that mean the same for all functions.
'
' -1 - Invalid Path
' -2 - Invalid HKey
' -3 - Invalid Key Path
' -4 - Permission denied
' -5 - OS arch mismatch (Example: Assigning a QWord into a 32-bit OS) 
' < 0 - Other error codes specific to the functions. See <https://msdn.microsoft.com/en-us/library/aa393978(v=vs.85).aspx> (WbemErrorEnum)
'
' See Also:
' 
' 	<WMI Reference, StdRegProv Class: http://msdn.microsoft.com/en-us/library/aa393664.aspx>
'	<Managing Windows Registry with Scripting (Part 2): http://www.serverwatch.com/tutorials/article.php/1476861>
'
Class CWMIReg
	Private HKCR_ ' HKEY_CLASSES_ROOT constant (StdRegProv).
	Private HKCU_ ' HKEY_CURRENT_USER constant (StdRegProv).
	Private HKLM_ ' HKEY_LOCAL_MACHINE constant (StdRegProv).
	Private HKU_ ' HKEY_USERS constant (StdRegProv).
   
	Private Loc_ ' WMI Locator (SWbemLocator) object (64-bit support for StdRegProv).
	Private Provider_ ' StdRegProv key target (32-bit or 64-bit).
	Private Processor_ ' Windows OS bitness (32-bit or 64-bit).
	
	Private debugEnabled ' Enable or disable debug logging.
	
	' ValueType
	Private vtNone_ ' No value type.
	Private vtString_ ' Nul terminated string.
	Private vtExpandString_ ' Nul terminated string (with environment variable references).
	Private vtBinary_ ' Free form binary.
	Private vtDWord_ ' 32-bit number.
	Private vtDWordBigEndian_ ' 32-bit number. In big-endian format, the most significant byte of a word is the low-order byte.
	Private vtLink_ ' Symbolic Link (unicode).
	Private vtMultiString_ ' Multiple strings.
	Private vtResourceList_ ' Resource list in the resource map.
	Private vtFullResourceDescriptor_ ' Resource list in the hardware description.
	Private vtResourceRequirementsList_ ' Resource list in the hardware description.
	Private vtQWord_ ' 64-bit number.
	' ValueType - end
	
	' Enable or disable debug logging. If enabled, debug messages are 
	' logged to the enabled facilities. Otherwise debug messages are 
	' silently discarded. This property is disabled by default.
	Public Property Get Debug
		Debug = debugEnabled
	End Property

	Public Property Let Debug(ByVal enable)
		debugEnabled = CBool(enable)
	End Property


'-----------------------------------------------------------------------------
' Function: Exists
'
' The class's Exists function uses key or value enumeration to check whether a key or value exists, and also returns the data type for existing values.
'
' Parameters:
'
' 	Path_  - Add "\" at end for keys.
'
' Returns:
'
' Returns data type if value, "K" if key, or "" if not found. (Note that WMI seems to often, perhaps always, return "REG_EXPAND_SZ" for plain string values, but it generally doesn't matter. A value that WMI says is "REG_EXPAND_SZ" can still be read or written as "REG_SZ", and when calling GetValue the Type parameter can be sent as "", leaving the GetValue function to handle the ambiguity.) 
'
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

'-----------------------------------------------------------------------------
' Function: GetValue
'
' GetValue returns the the value data, and also works to test the existence of a value.
'
' Parameters:
'
' 	Path_  - The source array.
'
' Returns:
'
' On success function returns value data. The method returns a string value that is "" (empty) if value does not exist. Otherwise the return can be a standard error code from -1 to -5 (see *Error Codes*). If the function fails, the return value is a nonzero error code.
'
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
		End If ' If Typ_ = "K"
	End If
End Function

'-----------------------------------------------------------------------------
' Function: EnumKeys
'
' Returns list of sub keys in a key.
'
' Parameters:
'
'	Path_ - The key to be enumerated. Path may have "\" at end or not.
'	AKeys_ - An array of key names.
'
' Returns:
'
' Function returns number of sub keys. Greater than 0 is the count of sub keys. Zero indicates no sub keys. Otherwise the return can be a standard error code from -1 to -5 (see *Error Codes*). If the function fails, the return value is a nonzero error code. AKeys contains sub key names.
'
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
  
	EnumKeys = -5 '-- os arch mismatch.
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
		Case -2147217405, 5   '--  access denied  H80041003
			EnumKeys = -4
		Case Else
			EnumKeys = iRet_  '-- some other error.  
	End Select
End Function

'-----------------------------------------------------------------------------
' Function: EnumVals
'
' Return value names and type from a given key.
'
' Parameters:
'
' 	Path_      - The source key.
'	AValsOut_  - An array of value names. The value array will include the key's default value only if it is not blank (WMI only returns it if there is content). Also, the default value is not necessarily found in array(0). The array seems to return values in the order that they were created.
'	ATypesOut_ - An array of data types.
'
' Returns:
'
' If return is > 0 it represents the number of values in the key. A return of 0 indicates no values present and no data saved in the default value. The value count is the same as UBound(AValsOut) + 1. For example, if a given key has 3 values written, EnumVals will return 3 and the empty default value will be ignored. If you then assign a string to the default value in that key, EnumVals will return 4. -1 to -5 are the standard error codes (see *Error Codes*). If the function fails, the return value is a nonzero error code.
'
Public Function EnumVals(Path_, AValsOut_, ATypesOut_)
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
    
	EnumVals = -5 '-- os arch mismatch.
		If IsEmpty(Provider_) Then Exit Function
       
    
    If Right(Path1_, 1) = "\" Then Path1_ = Left(Path1_, (len(Path1_) - 1))  
    'iRet_ = Reg1_.EnumValues(LKey_, Path1_, AValsOut_, ATypesOut_)
    Set Ctx_ = CreateObject("WbemScripting.SWbemNamedValueSet")
	Ctx_.Add "__ProviderArchitecture", Provider_
	Set Svc_ = Loc_.ConnectServer("","root\default","","",,,,Ctx_)
	Set Reg1_ = Svc_.Get("StdRegProv") 
		
	Set Inparams_ = Reg1_.Methods_("EnumValues").Inparameters
	Inparams_.Hdefkey = LKey_
	Inparams_.Ssubkeyname = Path1_ 
	
	Set Outparams_ = Reg1_.ExecMethod_("EnumValues", Inparams_,,Ctx_) 
	iRet_ = Outparams_.ReturnValue
	AValsOut_ = Outparams_.snames
	ATypesOut_ = Outparams_.Types

	Select Case iRet_
		Case 0
			If (IsArray(AValsOut_) = False) Then 
				EnumVals = 0  '-- no values in key.
			Else  '-- values found.
				EnumVals = UBound(AValsOut_) + 1
				For iCnt_ = 0 to UBound(ATypesOut_)
				   Val_ = ATypesOut_(iCnt_)
				   ATypesOut_(iCnt_) = ConvertType(Val_)
				Next   
			End If  
		Case 2 '-- invalid key Path
			EnumVals = -3
		Case -2147217405   '--  access denied  H80041003
			EnumVals = -4
		Case Else
			EnumVals = iRet_  '-- some other error.  
	End Select
End Function

'-----------------------------------------------------------------------------
' Function: SetValue
'
' Set value data.
'
' Parameters:
'
' 	Path_      - A key that contains the named value to be set. If key does not exist the key path will be created. You can specify an existing named value (update) or a new named value (create). Specify an empty string to set the data value for the default named value.
'	AValsOut_  - ValData is value data. Data type of the ValData parameter varies by type. For String or XString values a string must be sent. DWord values must be numeric. Binary values must be sent as an array of byte values (numbers from 0 to 255). A MultiString value must be sent as an array of strings.
'	TypeIn_    - (Optional) A data type. Use "" to SetValue function detect data type to use. If no type is sent then function will find type, but if value does not already exist when no type is sent it will default to string. Therefore, the type should always be sent when available. 
'
' Returns:
'
' The method returns a int value that is 0 (zero) if successful.  -1 to -5 are standard errors (see *Error Codes*). -5 = Incoming value not valid (data type of value data not coercible. Example: Assigning a string to ValData for a binary setting). -6 = Invalid data type value sent (for example, sending "A" as Type would be invalid). If the function fails, the return value is a nonzero error code.
'
Public Function SetValue(Path_, ValData_, TypeIn_)
	Dim Path1_, sKey_, LKey_, iRet_, Pt1_, Pt2_, ValName_ 
	Dim Ctx_, Svc_, Reg1_, Inparams_, Outparams_
   
	'-- Typ_
	Dim Typ_ : Typ_ = TypeIn_
	Dim vbArrayInteger, vbArrayString
	Dim RegEx_, REsult_ '-- reg_expand_sz
   
   
 
   
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
					' <http://regexr.com>
					' String should start with % and end with %.
					' It can not contain < > | & ^ <http://www.microsoft.com/resources/documentation/windows/xp/all/proddocs/en-us/set.mspx?mfr=true>
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
		End If ' If Typ_ = vtNone_
	Else ' If Len(TypeIn_) = 0
		If IsNumeric(TypeIn_) Then Typ_ = ConvertType(TypeIn_)
	End If    
   
	If Len(Typ_) = 0 Then
		' Exit Function
		Typ_ = "REG_SZ"
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
           SetValue = -5 '-- os arch mismatch
           Exit Function
        End If   

	'-- Create a key if it does not exist ------------
	iRet_ = EnumKeys(sKey_ & "\" & Left(Path1_, InStrRev(Path1_, "\")-1), AKeys)
	If iRet_ <> 0 Then 
		If iRet_ <> -4 Then
			iRet_ = CreateKey(sKey_ & "\" & Left(Path1_, InStrRev(Path1_, "\")-1))
			If iRet_ <> 0 Then
				If debugEnabled Then WScript.Echo "SetValue - Cannot create " & sKey_ & "\" & Left(Path1_, InStrRev(Path1_, "\")-1) & " key."
				SetValue = iRet_
				Exit Function
			End If
		Else
			If debugEnabled Then WScript.Echo "SetValue - Access denied to " & sKey_ & "\" & Left(Path1_, InStrRev(Path1_, "\")-1) & " key."
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
					Inparams_.uValue = DecimalNumbers(ValData_)
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
			Case "REG_QWORD"
				Set Inparams_ = Reg1_.Methods_("SetQWORDValue").Inparameters
				Inparams_.Hdefkey = LKey_
				Inparams_.Ssubkeyname = Path1_ 
				Inparams_.sValueName = ValName_
				Inparams_.uValue = ValData_
				
				Set Outparams_ = Reg1_.ExecMethod_("SetQWORDValue", Inparams_,,Ctx_)
				iRet_ = Outparams_.ReturnValue 
			Case Else
				SetValue = -6
				Exit Function 
        End Select
	End If   
   
    If (Err.number = -2147217403) Then 
       SetValue = -5  '-- type mismatch. incoming value not valid.
       Exit Function
    End If
	
	Select Case iRet_
		Case 0
			SetValue = 0   'success.
		Case 2 '-- invalid key path 
			SetValue = -3
		Case -2147217405   '--  access denied  H80041003
			SetValue = -4
		Case Else
			SetValue = iRet_  '-- some other error.  
    End Select
End Function

'-----------------------------------------------------------------------------
' Function: CreateKey
'
' Create a key or value. Path can have "\" at end or not (since the function is unambiguous), but must not have "\" if path is an HKey like "HKLM".
'
' Parameters:
'
' 	Path_ - The key to be created. The CreateKey method creates all sub keys specified in the path that do not exist. For example, if MyKey and MySubKey do not exist in the following path, the CreateKey method creates both keys: HKEY_CURRENT_USER\SOFTWARE\MyKey\MySubKey.
'
' Returns:
'
' The method returns a int value that is 0 (zero) if successful. If the function fails, the return value is a nonzero error code.
'
Public Function CreateKey(Path_)
	Dim sKey_, LKey_, Path1_, iRet_, Pt1_
	Dim Ctx_, Svc_, Reg1_, Inparams_, Outparams_
  
	On Error Resume Next
    CreateKey = -1 ' invalid path.
	Pt1_ = InStr(1, Path_, "\") 
	If (Pt1_ = 0) Then
		sKey_ = Path_
		Path1_ = ""
	Else   
		sKey_ = Left(Path_, Pt1_ - 1)
		Path1_ = Right(Path_, (len(Path_) - Pt1_))
	End If  
    
	CreateKey = -2 ' invalid hKey.
	LKey_ = GetHKey(sKey_)
		If (LKey_ = 0) Then Exit Function
    
	CreateKey = -5 ' os arch mismatch.
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
			CreateKey = -3
        Case -2147217405      '--  access denied  H80041003
			CreateKey = -4
        Case Else
			CreateKey = iRet_  '-- some other error.  
    End Select
	
End Function

'-----------------------------------------------------------------------------
' Function: Delete
'
' Delete a key (using DeleteKey) or value. Add "\" for keys.
'
' Parameters:
'
' 	Path_ - The key or value to be deleted. The function will enumerate sub keys when a key is being deleted and will delete the sub keys in reverse order to allow deletion of the key specified in Path. For example, if MyKey and MySubKey does exist in the following path, the Delete method deletes both keys: HKEY_CURRENT_USER\SOFTWARE\MyKey\MySubKey.
'
' Returns:
'
' The method returns a int value that is 0 (zero) if successful. If the function fails, the return value is a standard error code (see *Error Codes*) or a nonzero error code.
'
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
    
	Delete = -5 ' os arch mismatch
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
			Delete = -1 ' invalid path.
        Case -2  ' returned from DeleteKey
			Delete = -2 ' invalid hKey.
		Case 2 '-- invalid Key path
			Delete = -3
        Case -2147217405   '--  access denied  H80041003
			Delete = -4
       Case Else
			Delete = iRet_  '-- some other error.  
    End Select
	
End Function

'-----------------------------------------------------------------------------
' Function: TestPath
'
' Returns key level. The Delete function, since it is capable of deleting a key "tree", has a built-in safety feature. It will return error code -1 (Invalid Path) if the PathIn_ parameter does not contain at least 3 backslashes. That prevents the accidental deletion of the main hive keys and their direct sub keys. In other words, HKCU cannot be deleted. Nor can HKCU\Software\.
'
' Parameters:
'
' 	PathIn_ - The key to be tested.
'
' Returns:
'
' The method returns a int value > 0 (zero) if successful. If the function fails, the return value is a 0 (zero) error code.
'
' - HKCU\Software\Classes\ = HKCR\ = -1 (Invalid path)
' - HKLM\Software\ = -1 (Invalid Path)
' - HKCU\ControlPanel\Desktop = 3
' - HKLM\Software\Microsoft\Windows NT\CurrentVersion = 5
'
' See Also:
'
'	<DeleteKey>
'
Public Function TestPath(PathIn_)
  ' HKCU\ControlPanel\Desktop = 3
  ' HKLM\Software\Microsoft\Windows NT\CurrentVersion = 5
  
	Dim sKey_, Path_, sHKUSubKey_
	Dim Cnt_, Pt1_, Pt2_
	
	On Error Resume Next
		If InStr(PathIn_, "\") = 0 Then TestPath = 0 : Exit Function
      
    sKey_ = Left(PathIn_, InStr(PathIn_, "\")-1) : sKey_ = UCase(sKey_)
    Path_ = Mid(PathIn_, InStr(PathIn_, "\")+1, Len(PathIn_)) : Path_ = UCase(Path_)
    	If Path_ = "" Then TestPath = 0 : Exit Function
    
    ' Consider HKEY_USERS\S-1-5-21-1501084202-2593169170-243912787-500 as HKCU
    ' Consider HKEY_USERS\S-1-5-21-1501084202-2593169170-243912787-500_Classes as HKCR
	If Left(sKey_,Len("HKU\")) = "HKU\" Or _
		Left(sKey_, Len("HKEY_USERS\")) = "HKEY_USERS\" Then
		
		' S-1-5-21-1501084202-2593169170-243912787-500 
		' S-1-5-21-1501084202-2593169170-243912787-500_Classes
		sHKUSubKey_ = Left(Path_, InStr(Path_, "\")-1) : sHKUSubKey_ = UCase(sHKUSubKey_)
		
		If Right(sHKUSubKey_, Len("_CLASSES")) = "_CLASSES" Then
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
 
 ' - Constructor/Destructor ---------------------------------------------------

'-----------------------------------------------------------------------------
' Constructor: Class_Initialize
'
' - Set Locator (SWbemLocator) object and it's Constants.
' - Detect the bitness (32-bit or 64-bit) of the Windows OS.
' - Initialize logger objects with default values, i.e. disable debug.
'
Private Sub Class_Initialize()
	On Error Resume Next
	HKCR_ = &H80000000
	HKCU_ = &H80000001
	HKLM_ = &H80000002
	HKU_ = &H80000003
  
	'ValueType - begin
	vtNone_ = &H0 ' No value type
	vtString_ = &H1 ' Nul terminated string
	vtExpandString_ = &H2 ' Nul terminated string (with environment variable references)
	vtBinary_ = &H3 ' Free form binary
	vtDWord_ = &H4 ' 32-bit number
	vtDWordBigEndian_ = &H5 ' 32-bit number. In big-endian format, the most significant byte of a word is the low-order byte.
	vtLink_ = &H6 ' Symbolic Link (unicode)
	vtMultiString_ = &H7 ' Multiple strings
	vtResourceList_ = &H8 ' Resource list in the resource map
	vtFullResourceDescriptor_ = &H9 ' Resource list in the hardware description
	vtResourceRequirementsList_ = &HA ' Resource list in the hardware description
	vtQWord_ = &HB ' 64-bit number
	'ValueType - end
	
	debugEnabled = False

	'Set Re'g1_ = GetObject("winMgMts:root\default:StdRegProv")
	' If (Err.number <> 0) Then
	'   Err.Raise 1, "WMIReg Class", "Failed to access WMI StdRegProv object. Class cannot function."
	' End If
  
	Set Loc_ = CreateObject("Wbemscripting.SWbemLocator")
	If (Err.number <> 0) Then
		Err.Raise 1, "WMIReg Class", "Failed to access WMI SWbemLocator object. Class cannot function."
	End If
  
	' <http://csi-windows.com/toolkit/csi-getosbits>
	Processor_ = GetValue("HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\PROCESSOR_ARCHITECTURE")
End Sub 

 ' - Constructor/Destructor ---------------------------------------------------

'-----------------------------------------------------------------------------
' Function: Class_Terminate
'
' - Set Locator (SWbemLocator) object to Nothing.
'
Private Sub Class_Terminate()
	'Set Reg1_ = Nothing
	Set Loc_ = Nothing
End Sub

'-----------------------------------------------------------------------------
' Function: GetHKey
'
' Assign Registry Root variable with its WMI hex equivalent. 
'
' Parameters:
' 
'	sKey1_ - The key to be enumerated.
'
' Returns: 
'
' Assign Registry Root variable with its WMI hex equivalent, or 0 if the operation failed.
'
Public Function GetHKey(sKey1_) 
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
		Case "HKLM", "HKEY_LOCAL_MACHINE"  
			GetHKey = HKLM_
		Case "HKCU", "HKEY_CURRENT_USER"   
			GetHKey = HKCU_
		Case "HKCR", "HKEY_CLASSES_ROOT"   
			GetHKey = HKCR_
		Case "HKU", "HKEY_USERS"  
			GetHKey = HKU_ 
	      
		Case Else 
			GetHKey = 0
	End Select      
End Function

'-----------------------------------------------------------------------------
' Function: ConvertType
'
' Assign Type variable with its WMI hex equivalent and vice-versa.
'
' Parameters:
' 
'	TypeIn_ - The Type to be enumerated.
'
' Returns: 
'
' Assign Type variable with its WMI hex equivalent and vice-versa. Returns "" for invalid data type. 
'
Public Function ConvertType(TypeIn_)
	On Error Resume Next
	Select Case TypeIn_
		Case vtString_
			ConvertType = "REG_SZ"
		Case vtExpandString_
			ConvertType = "REG_EXPAND_SZ"
		Case vtBinary_
			ConvertType = "REG_BINARY"
		Case vtDWord_
			ConvertType = "REG_DWORD"
		Case vtMultiString_
			ConvertType = "REG_MULTI_SZ"
		Case vtQWord_
			ConvertType = "REG_QWORD"
			
		Case "REG_SZ"
			ConvertType = vtString_
		Case "REG_EXPAND_SZ"
			ConvertType = vtExpandString_
		Case "REG_BINARY"
			ConvertType = vtBinary_
		Case "REG_DWORD"
			ConvertType = vtDWord_
		Case "REG_MULTI_SZ"
			ConvertType = vtMultiString_
		Case "REG_QWORD"
			ConvertType = vtQWord_
		
		Case Else
			ConvertType = ""
	End Select        
End Function

'-----------------------------------------------------------------------------
' Function: EnumKeysAll
'
' Return list of all sub keys in a key. EnumKeysAll has been made public, in case it might be useful, but it was really written for use in deleting keys.
'
' Parameters:
' 
'	Path_  - The key to be enumerated.
'	AKeys_ -
'
' Returns: 
'
' Function returns number of sub keys. AKeys returns key paths.
'
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

'-----------------------------------------------------------------------------
' Function: DeleteKey
'
' Deletes all sub keys in path by first calling EnumKeysAll. It then deletes parent key.
'
' Parameters:
' 
'	Path_ - The key to be deleted.
'
' Returns: 
'
' The method returns a int value that is 0 (zero) if successful. If the function fails, the return value is a nonzero error code.
'
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

'-----------------------------------------------------------------------------
' Function: DecimalNumbers
'
' Convert hex string to binary Array.
'
' Author: Rems <http://www.petri.co.il/forums/showthread.php?t=46158>
'
' Parameters:
' 
'	strHex_ - The value to be converted. Example: "hex:23,00,41,00,43,00,42,00,6c,00"
'
' Returns: 
'
' The method returns a decimal binary Array if successful.
'
' See Also:
'
'	<SetValue>
'
Private Function DecimalNumbers(strHex_)
	Dim arrHex_ : arrHex_ = Split(Replace(strHex_, "hex:", ""), ",") 
	
    On Error Resume Next
    
    ' from: <http://www.petri.co.il/forums/showthread.php?t=46158>
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
