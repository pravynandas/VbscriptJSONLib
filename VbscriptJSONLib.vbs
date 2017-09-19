' ============================================================================================
' This is a VBScrit adaptation of the VB6 implementation by Michael Glaser (vbjson@ediy.co.nz)
' Added support for "ArrayList" since Collection is not directly supported in VBS
' ArrayList based StringBuilder class used instead of CopyMemory based (vbAccelerator) cStringBuilder.
' Edits by Praveen Nandagiri (pravynandas@gmail.com)
'
' Pre-Note:
' VBJSON is a VB6 adaptation of the VBA JSON project at http://code.google.com/p/vba-json/
' Some bugs fixed, speed improvements added for VB6 by Michael Glaser (vbjson@ediy.co.nz)
' BSD Licensed
' ============================================================================================
' ==========================================================================================
' Support Class: StringBuilder
' Author: Stan Schwartz
' Visit: http://www.nullskull.com/q/10167834/vbscript-stringbuilder-class.aspx
' ==========================================================================================
Option Explicit

Class VbscriptJSONLib
	Private INVALID_JSON
	Private INVALID_OBJECT
	Private INVALID_ARRAY
	Private INVALID_BOOLEAN
	Private INVALID_NULL
	Private INVALID_KEY
	Private INVALID_RPC_CALL

	Private psErrors 

	Private Sub Class_Initialize()
		psErrors = ""
		
		'Class constants
		INVALID_JSON = 1
		INVALID_OBJECT = 2
		INVALID_ARRAY = 3
		INVALID_BOOLEAN = 4
		INVALID_NULL = 5
		INVALID_KEY = 6
		INVALID_RPC_CALL = 7
	End Sub
	
	Public Function GetParserErrors() 
		GetParserErrors = psErrors
	End Function

	Public Function ClearParserErrors() 
		psErrors = ""
	End Function


	'
	'   parse string and create JSON object
	'
	Public Function parse(ByRef str ) 

		Dim index 
		index = 1
		psErrors = ""
		'On Error Resume Next
		Call skipChar(str, index)
		Select Case Mid(str, index, 1)
		Case "{"
			Set parse = parseObject(str, index)
		Case "["
			Set parse = parseArray(str, index)
		Case Else
			psErrors = "Invalid JSON"
		End Select


	End Function

	'
	'   parse collection of key/value
	'
	Private Function parseObject(ByRef str , ByRef index ) 

		Set parseObject = CreateObject("Scripting.Dictionary")
		Dim sKey 

		' "{"
		Call skipChar(str, index)
		If Mid(str, index, 1) <> "{" Then
			psErrors = psErrors & "Invalid Object at position " & index & " : " & Mid(str, index) & vbCrLf
			Exit Function
		End If

		index = index + 1

		Do
			Call skipChar(str, index)
			If "}" = Mid(str, index, 1) Then
				index = index + 1
				Exit Do
			ElseIf "," = Mid(str, index, 1) Then
				index = index + 1
				Call skipChar(str, index)
			ElseIf index > Len(str) Then
				psErrors = psErrors & "Missing '}': " & Right(str, 20) & vbCrLf
				Exit Do
			End If


			' add key/value pair
			sKey = parseKey(str, index)
			'On Error Resume Next

			parseObject.Add sKey, parseValue(str, index)
			If Err.Number <> 0 Then
				psErrors = psErrors & Err.Description & ": " & sKey & vbCrLf
				Exit Do
			End If
		Loop
	'eh:

	End Function

	'
	'   parse list
	'
	Private Function parseArray(ByRef str , ByRef index ) 

		Set parseArray = CreateObject("System.Collections.ArrayList")

		' "["
		Call skipChar(str, index)
		If Mid(str, index, 1) <> "[" Then
			psErrors = psErrors & "Invalid Array at position " & index & " : " + Mid(str, index, 20) & vbCrLf
			Exit Function
		End If

		index = index + 1

		Do

			Call skipChar(str, index)
			If "]" = Mid(str, index, 1) Then
				index = index + 1
				Exit Do
			ElseIf "," = Mid(str, index, 1) Then
				index = index + 1
				Call skipChar(str, index)
			ElseIf index > Len(str) Then
				psErrors = psErrors & "Missing ']': " & Right(str, 20) & vbCrLf
				Exit Do
			End If

			' add value
			'On Error Resume Next
			parseArray.Add parseValue(str, index)
			If Err.Number <> 0 Then
				psErrors = psErrors & Err.Description & ": " & Mid(str, index, 20) & vbCrLf
				Exit Do
			End If
		Loop

	End Function

	'
	'   parse string / number / object / array / true / false / null
	'
	Private Function parseValue(ByRef str , ByRef index )

		Call skipChar(str, index)

		Select Case Mid(str, index, 1)
		Case "{"
			Set parseValue = parseObject(str, index)
		Case "["
			Set parseValue = parseArray(str, index)
		Case """", "'"
			parseValue = parseString(str, index)
		Case "t", "f"
			parseValue = parseBoolean(str, index)
		Case "n"
			parseValue = parseNull(str, index)
		Case Else
			parseValue = parseNumber(str, index)
		End Select

	End Function

	'
	'   parse string
	'
	Private Function parseString(ByRef str , ByRef index ) 

		Dim quote 
		Dim Char 
		Dim Code 

		Dim SB : Set SB = New StringBuilder

		Call skipChar(str, index)
		quote = Mid(str, index, 1)
		index = index + 1

		Do While index > 0 And index <= Len(str)
			Char = Mid(str, index, 1)
			Select Case (Char)
			Case "\"
				index = index + 1
				Char = Mid(str, index, 1)
				Select Case (Char)
				Case """", "\", "/", "'"
					SB.Append Char
					index = index + 1
				Case "b"
					SB.Append vbBack
					index = index + 1
				Case "f"
					SB.Append vbFormFeed
					index = index + 1
				Case "n"
					SB.Append vbLf
					index = index + 1
				Case "r"
					SB.Append vbCr
					index = index + 1
				Case "t"
					SB.Append vbTab
					index = index + 1
				Case "u"
					index = index + 1
					Code = Mid(str, index, 4)
					'SB.Append ChrW(Val("&h" + Code))
					SB.Append Chr("&H" + Code)
					index = index + 4
				End Select
			Case quote
				index = index + 1

				parseString = SB.toString
				Set SB = Nothing

				Exit Function

			Case Else
				SB.Append Char
				index = index + 1
			End Select
		Loop

		parseString = SB.toString
		Set SB = Nothing

	End Function

	'
	'   parse number
	'
	Private Function parseNumber(ByRef str , ByRef index )

		Dim Value 
		Dim Char 

		Call skipChar(str, index)
		Do While index > 0 And index <= Len(str)
			Char = Mid(str, index, 1)
			If InStr("+-0123456789.eE", Char) Then
				Value = Value & Char
				index = index + 1
			Else
				parseNumber = CDbl(Value)
				Exit Function
			End If
		Loop
	End Function

	'
	'   parse true / false
	'
	Private Function parseBoolean(ByRef str , ByRef index ) 

		Call skipChar(str, index)
		If Mid(str, index, 4) = "true" Then
			parseBoolean = True
			index = index + 4
		ElseIf Mid(str, index, 5) = "false" Then
			parseBoolean = False
			index = index + 5
		Else
			psErrors = psErrors & "Invalid Boolean at position " & index & " : " & Mid(str, index) & vbCrLf
		End If

	End Function

	'
	'   parse null
	'
	Private Function parseNull(ByRef str , ByRef index )

		Call skipChar(str, index)
		If Mid(str, index, 4) = "null" Then
			parseNull = Null
			index = index + 4
		Else
			psErrors = psErrors & "Invalid null value at position " & index & " : " & Mid(str, index) & vbCrLf
		End If

	End Function

	Private Function parseKey(ByRef str , ByRef index ) 

		Dim dquote 
		Dim squote 
		Dim Char 

		Call skipChar(str, index)
		Do While index > 0 And index <= Len(str)
			Char = Mid(str, index, 1)
			Select Case (Char)
			Case """"
				dquote = Not dquote
				index = index + 1
				If Not dquote Then
					Call skipChar(str, index)
					If Mid(str, index, 1) <> ":" Then
						psErrors = psErrors & "Invalid Key at position " & index & " : " & parseKey & vbCrLf
						Exit Do
					End If
				End If
			Case "'"
				squote = Not squote
				index = index + 1
				If Not squote Then
					Call skipChar(str, index)
					If Mid(str, index, 1) <> ":" Then
						psErrors = psErrors & "Invalid Key at position " & index & " : " & parseKey & vbCrLf
						Exit Do
					End If
				End If
			Case ":"
				index = index + 1
				If Not dquote And Not squote Then
					Exit Do
				Else
					parseKey = parseKey & Char
				End If
			Case Else
				If InStr(vbCrLf & vbCr & vbLf & vbTab & " ", Char) Then
				Else
					parseKey = parseKey & Char
				End If
				index = index + 1
			End Select
		Loop

	End Function

	'
	'   skip special character
	'
	Private Sub skipChar(ByRef str , ByRef index )
		Dim bComment 
		Dim bStartComment 
		Dim bLongComment 
		Do While index > 0 And index <= Len(str)
			Select Case Mid(str, index, 1)
			Case vbCr, vbLf
				If Not bLongComment Then
					bStartComment = False
					bComment = False
				End If

			Case vbTab, " ", "(", ")"

			Case "/"
				If Not bLongComment Then
					If bStartComment Then
						bStartComment = False
						bComment = True
					Else
						bStartComment = True
						bComment = False
						bLongComment = False
					End If
				Else
					If bStartComment Then
						bLongComment = False
						bStartComment = False
						bComment = False
					End If
				End If

			Case "*"
				If bStartComment Then
					bStartComment = False
					bComment = True
					bLongComment = True
				Else
					bStartComment = True
				End If

			Case Else
				If Not bComment Then
					Exit Do
				End If
			End Select

			index = index + 1
		Loop

	End Sub

	Public Function toString(ByRef obj ) 
		Dim SB : Set SB = New StringBuilder
		Select Case VarType(obj)
		Case vbNull
			SB.Append "null"
		Case vbDate
			SB.Append """" & CStr(obj) & """"
		Case vbString
			SB.Append """" & Encode(obj) & """"
		Case vbObject

			Dim bFI 
			Dim i 

			bFI = True
			If TypeName(obj) = "Dictionary" Then

				SB.Append "{"
				Dim keys
				keys = obj.keys
				For i = 0 To obj.Count - 1
					If bFI Then bFI = False Else SB.Append ","
					Dim key
					key = keys(i)
					SB.Append """" & key & """:" & toString(obj.Item(key))
				Next 'i
				SB.Append "}"

			ElseIf TypeName(obj) = "Collection" OR TypeName(obj) = "ArrayList" Then

				SB.Append "["
				Dim Value
				For Each Value In obj
					If bFI Then bFI = False Else SB.Append ","
					SB.Append toString(Value)
				Next 'Value
				SB.Append "]"

			End If
		Case vbBoolean
			If obj Then SB.Append "true" Else SB.Append "false"
		Case vbVariant, vbArray, vbArray + vbVariant
			Dim sEB
			SB.Append multiArray(obj, 1, "", sEB)
		Case Else
			SB.Append Replace(obj, ",", ".")
		End Select

		toString = SB.toString
		Set SB = Nothing

	End Function

	Private Function Encode(str) 

		Dim SB : Set SB = New StringBuilder
		Dim i 
		Dim j 
		Dim aL1 
		Dim aL2 
		Dim c 
		Dim p 

		aL1 = Array(&H22, &H5C, &H2F, &H8, &HC, &HA, &HD, &H9)
		aL2 = Array(&H22, &H5C, &H2F, &H62, &H66, &H6E, &H72, &H74)
		For i = 1 To Len(str)
			p = True
			c = Mid(str, i, 1)
			For j = 0 To 7
				If c = Chr(aL1(j)) Then
					SB.Append "\" & Chr(aL2(j))
					p = False
					Exit For
				End If
			Next

			If p Then
				Dim a
				a = AscW(c)
				If a > 31 And a < 127 Then
					SB.Append c
				ElseIf a > -1 Or a < 65535 Then
					SB.Append "\u" & String(4 - Len(Hex(a)), "0") & Hex(a)
				End If
			End If
		Next

		Encode = SB.toString
		Set SB = Nothing

	End Function

	Private Function multiArray(aBD, iBC, sPS, ByRef sPT)   ' Array BoDy, Integer BaseCount, String PoSition

		Dim iDU 
		Dim iDL 
		Dim i 

		On Error Resume Next
		iDL = LBound(aBD, iBC)
		iDU = UBound(aBD, iBC)

		Dim SB : Set SB = New StringBuilder

		Dim sPB1, sPB2  ' String PointBuffer1, String PointBuffer2
		If Err.Number = 9 Then
			sPB1 = sPT & sPS
			For i = 1 To Len(sPB1)
				If i <> 1 Then sPB2 = sPB2 & ","
				sPB2 = sPB2 & Mid(sPB1, i, 1)
			Next
			'        multiArray = multiArray & toString(Eval("aBD(" & sPB2 & ")"))
			SB.Append toString(aBD(sPB2))
		Else
			sPT = sPT & sPS
			SB.Append "["
			For i = iDL To iDU
				SB.Append multiArray(aBD, iBC + 1, i, sPT)
				If i < iDU Then SB.Append ","
			Next
			SB.Append "]"
			sPT = Left(sPT, iBC - 2)
		End If
		Err.Clear
		multiArray = SB.toString

		Set SB = Nothing
	End Function

	' Miscellaneous JSON functions

	Public Function StringToJSON(st ) 

		Dim FIELD_SEP : FIELD_SEP = "~"
		Dim RECORD_SEP : RECORD_SEP = "|"

		Dim sFlds 
		Dim sRecs : Set sRecs = New StringBuilder
		Dim lRecCnt 
		Dim lFld 
		Dim fld 
		Dim rows 

		lRecCnt = 0
		If st = "" Then
			StringToJSON = "null"
		Else
			rows = Split(st, RECORD_SEP)
			For lRecCnt = LBound(rows) To UBound(rows)
				sFlds = ""
				fld = Split(rows(lRecCnt), FIELD_SEP)
				For lFld = LBound(fld) To UBound(fld) Step 2
					sFlds = (sFlds & IIf(sFlds <> "", ",", "") & """" & fld(lFld) & """:""" & toUnicode(fld(lFld + 1) & "") & """")
				Next    'fld
				sRecs.Append IIf((Trim(sRecs.toString) <> ""), "," & vbCrLf, "") & "{" & sFlds & "}"
			Next    'rec
			StringToJSON = ("( {""Records"": [" & vbCrLf & sRecs.toString & vbCrLf & "], " & """RecordCount"":""" & lRecCnt & """ } )")
		End If
	End Function


	Public Function RStoJSON(rs ) 
		'On Error GoTo errHandler
		Dim sFlds 
		Dim sRecs : Set sRecs = New StringBuilder
		Dim lRecCnt 
		Dim fld 

		lRecCnt = 0
		If rs.State = adStateClosed Then
			RStoJSON = "null"
		Else
			If rs.EOF Or rs.BOF Then
				RStoJSON = "null"
			Else
				Do While Not rs.EOF And Not rs.BOF
					lRecCnt = lRecCnt + 1
					sFlds = ""
					For Each fld In rs.Fields
						sFlds = (sFlds & IIf(sFlds <> "", ",", "") & """" & fld.Name & """:""" & toUnicode(fld.Value & "") & """")
					Next    'fld
					sRecs.Append IIf((Trim(sRecs.toString) <> ""), "," & vbCrLf, "") & "{" & sFlds & "}"
					rs.MoveNext
				Loop
				RStoJSON = ("( {""Records"": [" & vbCrLf & sRecs.toString & vbCrLf & "], " & """RecordCount"":""" & lRecCnt & """ } )")
			End If
		End If

		Exit Function
	'errHandler:

	End Function

	'Public Function JsonRpcCall(url , methName , args(), Optional user , Optional pwd ) 
	'    Dim r 
	'    Dim cli 
	'    Dim pText 
	'    Static reqId 
	'
	'    reqId = reqId + 1
	'
	'    Set r = CreateObject("Scripting.Dictionary")
	'    r("jsonrpc") = "2.0"
	'    r("method") = methName
	'    r("params") = args
	'    r("id") = reqId
	'
	'    pText = toString(r)
	'
	'    Set cli = CreateObject("MSXML2.XMLHTTP.6.0")
	'   ' Set cli = New MSXML2.XMLHTTP60
	'    If Len(user) > 0 Then   ' If Not IsMissing(user) Then
	'        cli.Open "POST", url, False, user, pwd
	'    Else
	'        cli.Open "POST", url, False
	'    End If
	'    cli.setRequestHeader "Content-Type", "application/json"
	'    cli.Send pText
	'
	'    If cli.Status <> 200 Then
	'        Err.Raise vbObjectError + INVALID_RPC_CALL + cli.Status, , cli.statusText
	'    End If
	'
	'    Set r = parse(cli.responseText)
	'    Set cli = Nothing
	'
	'    If r("id") <> reqId Then Err.Raise vbObjectError + INVALID_RPC_CALL, , "Bad Response id"
	'
	'    If r.Exists("error") Or Not r.Exists("result") Then
	'        Err.Raise vbObjectError + INVALID_RPC_CALL, , "Json-Rpc Response error: " & r("error")("message")
	'    End If
	'
	'    If Not r.Exists("result") Then Err.Raise vbObjectError + INVALID_RPC_CALL, , "Bad Response, missing result"
	'
	'    Set JsonRpcCall = r("result")
	'End Function




	Public Function toUnicode(str ) 

		Dim x 
		Dim uStr : Set uStr = New StringBuilder
		Dim uChrCode 

		For x = 1 To Len(str)
			uChrCode = Asc(Mid(str, x, 1))
			

				Select Case uChrCode
				Case 8:    ' backspace
					uStr.Append "\b"
				Case 9:    ' tab
					uStr.Append "\t"
				Case 10:    ' line feed
					uStr.Append "\n"
				Case 12:    ' formfeed
					uStr.Append "\f"
				Case 13:    ' carriage return
					uStr.Append "\r"
				Case 34:    ' quote
					uStr.Append "\"""
				Case 39:    ' apostrophe
					uStr.Append "\'"
				Case 92:    ' backslash
					uStr.Append "\\"
				Case 123, 125:    ' "{" and "}"
					uStr.Append ("\u" & Right("0000" & Hex(uChrCode), 4))
				'Below condition migrated to below IF condition in VBS
				'Case Is < 32, Is > 127:    ' non-ascii characters
				'    uStr.Append ("\u" & Right("0000" & Hex(uChrCode), 4))
				Case Else
					If uChrCode < 32 or uChrCode > 127 Then    ' non-ascii characters
						uStr.Append ("\u" & Right("0000" & Hex(uChrCode), 4))
					Else
						uStr.Append Chr(uChrCode)	'everything else
					End If
				End Select
		Next
		toUnicode = uStr.toString
		Exit Function

	End Function
End Class

Class StringBuilder
     
    Dim stringArray
     
    Private Sub Class_Initialize()
        Set stringArray = CreateObject("System.Collections.ArrayList")
    End Sub
     
    Public Sub Append(ByVal strValue)
        stringArray.Add strValue
    End Sub
     
    Public Sub PrePend(ByVal strValue)
        stringArray.Insert 0, strValue
    End Sub
     
    Public Function toString()
        toString = Join(stringArray.ToArray(), "")
    End Function
     
    Public Function Count()
        Count = stringArray.Count()
    End Function
     
    Public Sub Reset()
        stringArray.Clear()
        Class_Initialize
    End Sub
     
    Public Function Contains(ByVal strString)
        Contains = False
        If stringArray.Contains(strString) Then Contains = True
    End Function
     
End Class