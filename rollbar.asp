<%
Dim strRollbarAccessToken
Dim strRollbarPersonUserId
Dim strRollbarPersonUserName
Dim strRollbarEnvironment

strRollbarAccessToken    = ""
strRollbarPersonUserId   = ""
strRollbarPersonUserName = ""
strRollbarEnvironment    = "production"

Public Sub RollbarASPError()
    Dim objError
    Set objError = Server.GetLastError()
    Call Rollbar("error", "", "", objError)
    Set objError = Nothing
End Sub

Public Sub RollbarError(strMessage, strExtraPayload)
    Call Rollbar("error", strMessage, strExtraPayload, NULL)
End Sub

Public Sub RollbarWarning(strMessage, strExtraPayload)
    Call Rollbar("warning", strMessage, strExtraPayload, NULL)
End Sub

Public Sub RollbarInfo(strMessage, strExtraPayload)
    Call Rollbar("info", strMessage, strExtraPayload, NULL)
End Sub

Public Sub RollbarDebug(strMessage, strExtraPayload)
    Call Rollbar("debug", strMessage, strExtraPayload, NULL)
End Sub

Private Sub Rollbar(strLevel, strMessage, strExtraPayload, objError)
    Dim strPayload, strURL

    If strRollbarAccessToken = "" Then
        Exit Sub
    End If

    On Error Resume Next
    If strLevel = "error" Then
        If IsObject(objError) Then
            strMessage = objError.Description
        End If
    End If

    If Request.ServerVariables("HTTPS") = "ON" Then
        strURL = "https://"
    Else
        strURL = "http://"
    End If
    strURL = strURL & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL")

    strPayload = "{"
    strPayload = strPayload & """access_token"": """&strRollbarAccessToken&""","
    strPayload = strPayload & """data"": "
    strPayload = strPayload & "{"
    strPayload = strPayload & "    ""environment"": """&strRollbarEnvironment&""","
    strPayload = strPayload & "    ""level"": """&strLevel&""", "
    strPayload = strPayload & "    ""body"": "
    strPayload = strPayload & "    { "
    strPayload = strPayload & "        ""message"": "
    strPayload = strPayload & "        {"
    strPayload = strPayload & "            ""body"": """&PrepareForRollbar(strMessage)&""","
    strPayload = strPayload & "            ""session"": """&PrepareForRollbar(GetSessionAsString())&""""
    If strExtraPayload <> "" Then
        strPayload = strPayload & ","
        strPayload = strPayload & strExtraPayload
    End If
    If strLevel = "error" AND IsObject(objError) Then
        strPayload = strPayload & ","
        strPayload = strPayload & "       ""ASPCode"": """&PrepareForRollbar(objError.ASPCode)&""","
        strPayload = strPayload & "       ""ASPDescription"": """&PrepareForRollbar(objError.ASPDescription)&""","
        strPayload = strPayload & "       ""Category"": """&PrepareForRollbar(objError.Category)&""","
        strPayload = strPayload & "       ""Column"": """&PrepareForRollbar(objError.Column)&""","
        strPayload = strPayload & "       ""Description"": """&PrepareForRollbar(objError.Description)&""","
        strPayload = strPayload & "       ""File"": """&PrepareForRollbar(objError.File)&""","
        strPayload = strPayload & "       ""Line"": """&PrepareForRollbar(objError.Line)&""","
        strPayload = strPayload & "       ""Number"": """&PrepareForRollbar(objError.Number)&""","
        strPayload = strPayload & "       ""Source"": """&PrepareForRollbar(objError.Source)&""""
    End If
    strPayload = strPayload & "        }"
    strPayload = strPayload & "    },"
    If strRollbarPersonUserId <> "" OR strRollbarPersonUserName <> "" Then
        strPayload = strPayload & "    ""person"": "
        strPayload = strPayload & "    { "
        strPayload = strPayload & "        ""id"": """&PrepareForRollbar(strRollbarPersonUserId)&""","
        strPayload = strPayload & "        ""username"": """&PrepareForRollbar(strRollbarPersonUserName)&""""
        strPayload = strPayload & "    },"
    End If
    strPayload = strPayload & "    ""request"": "
    strPayload = strPayload & "    { "
    strPayload = strPayload & "        ""url"": """&PrepareForRollbar(strUrl)&""","
    strPayload = strPayload & "        ""method"": """&PrepareForRollbar(Request.ServerVariables("HTTP_METHOD"))&""","
    strPayload = strPayload & "        ""query_string"": """&PrepareForRollbar(Request.QueryString)&""","
    strPayload = strPayload & "        ""body"": """&PrepareForRollbar(Request.Form)&""","
    strPayload = strPayload & "        ""user_ip"": """&PrepareForRollbar(Request.ServerVariables("REMOTE_ADDR"))&""""
    strPayload = strPayload & "    }"
    strPayload = strPayload & "}"
    strPayload = strPayload & "}"

    Call GetURLPostJSON("https://api.rollbar.com/api/1/item/", 1, strPayload)
    On Error Goto 0

    If strLevel = "error" Then
        Set objError = Nothing
    End If
End Sub

Function PrepareForRollbar(strData)
    strData = EnsureIsTrimmedString(strData)
    strData = Replace(strData, """", "")
    strData = Replace(strData, "\", "\\")
    strData = Replace(strData, VbCrLf, "\n")
    PrepareForRollbar = strData
End Function

Function EnsureIsTrimmedString(ByVal strString)
    strString = strString & ""
    If NOT IsNull(strString) Then
        strString = Trim(strString)
    End If
    EnsureIsTrimmedString = strString 
End Function

Function GetSessionAsString()
    On Error Resume Next
    Dim sessionItem, strSession
    strSession = ""
    For Each sessionItem in Session.Contents
        If IsArray(Session(sessionItem)) Then
            strSession = strSession & sessionItem & "=" & PrintArray(Session(sessionItem)) & VbCrLf
        Else
            strSession = strSession & sessionItem & "=" & Session(sessionItem) & VbCrLf
        End If
    Next
    GetSessionAsString = strSession
End Function

Function PrintArray(aryArray)
	On Error Resume Next
	Dim i, j, k, strOut, strElement, aryDimensions(10), strDimensions
	i=0
	strDimensions = ""
	For Each strElement in aryArray
		i = i + 1
	Next
	j = 0
	k=1
	Do While j >= 0 AND k < 10
		j = UBound(aryArray, k)
		If j > 0 Then
			strDimensions = strDimensions & j & ","
			aryDimensions(k-1) = j
		End If
		j = 0
		k = k + 1
	Loop
	If strDimensions <> "" Then
		strDimensions = Left(strDimensions, Len(strDimensions)-1)
	End If
	strOut = "Array ("&strDimensions&"): " & VbCrLf
	strOut = strOut & "----------" & VbCrLf
	j=0
	For Each strElement in aryArray
		strOut = strOut & strElement
		j = j + 1
		If j = aryDimensions(0)+1 Then
			strOut = strOut & VbCrLf
			j=0
		Else
			strOut = strOut & ","
		End If
	Next	
    strOut = strOut & "----------"
	PrintArray = strOut
End Function

Function GetURLPostJSON(strUrl, lTimeout, strData)
    Dim objHttp, GotResponse, intSecondsWait
    On Error Resume Next
    GetURLPostJSON = ""    
        
    Set objHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
    If lTotal = 0 Then
        objHttp.open strMethod, strUrl, True
    Else
        objHttp.setTimeouts lTotal*1000/4, lTotal*1000/4, lTotal*1000/4, lTotal*1000
        objHttp.open strMethod, strUrl, False 
    End if
    objHttp.setRequestHeader "Content-Type", "application/json"
    objHttp.send strData
    If lTotal = 0 Then
        Set objHttp = Nothing
        Exit Function
    End If
    
    intSecondsWait  = 0
    GotResponse     = False
    Do While objHttp.readyState <> 4
        If Err.Number <> 0 Then
            Exit Do
        End If
        
        objHttp.waitForResponse 1
        intSecondsWait = intSecondsWait + 1
        
        If objHttp.readyState = 4 Then
            GotResponse = True
            Exit Do
        End If
        If intSecondsWait > lTotal Then
            GotResponse = False
            Exit Do
        End If
    Loop
    If objHttp.readyState = 4 Then
        GotResponse = True
    End If
    
    If GotResponse AND Err.Number = 0 Then
        If objHttp.status >= 200 AND objHttp.status <= 299 Then
	    GetURLPostJSON = objHTTP.ResponseText 
        End If
        
    ElseIf Err.Number <> 0 Then
        Err.Clear
    End If
    
    Set objHttp = Nothing
    
    On Error Goto 0
End Function
%>
