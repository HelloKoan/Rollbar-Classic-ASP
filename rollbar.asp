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
    strPayload = strPayload & "            ""body"": """&strMessage&""""
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

    response.write strPayload
    Call GetURLPostJSON("https://api.rollbar.com/api/1/item/", 1, strPayload, "", "")
    On Error Goto 0

    If strLevel = "error" Then
        Set objError = Nothing
    End If
End Sub

Function PrepareForRollbar(strData)
    strData = EnsureIsTrimmedString(strData)
    strData = Replace(strData, """", "")
    strData = Replace(strData, VbCrLf, " ")
    PrepareForRollbar = strData
End Function

Function EnsureIsTrimmedString(ByVal strString)
    strString = strString & ""
    If NOT IsNull(strString) Then
        strString = Trim(strString)
    End If
    EnsureIsTrimmedString = strString 
End Function
%>