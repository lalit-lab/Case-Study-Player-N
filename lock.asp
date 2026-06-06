<%
' =====================================================================
'  SINGLE-DEVICE LOCK helpers (no secrets — safe to commit)
'  Each successful login mints a token, stored server-side per account.
'  A later login overwrites it, so the earlier session stops validating
'  and gets bounced to the login screen. Only one device stays active.
' =====================================================================

Function NewToken()
    Dim s, i
    Randomize
    s = ""
    For i = 1 To 24
        s = s & Hex(Int(Rnd() * 16))
    Next
    NewToken = s & "-" & Hex(CLng(Timer * 1000))
End Function

Function LockPath(user)
    Dim safe, i, c
    safe = ""
    For i = 1 To Len(user)
        c = Mid(user, i, 1)
        If (c >= "a" And c <= "z") Or (c >= "A" And c <= "Z") Or (c >= "0" And c <= "9") Then safe = safe & c
    Next
    If safe = "" Then safe = "x"
    LockPath = Server.MapPath(".lock_" & safe & ".txt")
End Function

Sub SetActiveToken(user, tok)
    Dim fso, f
    On Error Resume Next
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    Set f = fso.CreateTextFile(LockPath(user), True)
    f.Write tok
    f.Close
    Set fso = Nothing
End Sub

Function GetActiveToken(user)
    Dim fso, f, v
    v = ""
    On Error Resume Next
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(LockPath(user)) Then
        Set f = fso.OpenTextFile(LockPath(user), 1)
        If Not f.AtEndOfStream Then v = f.ReadAll
        f.Close
    End If
    Set fso = Nothing
    GetActiveToken = v
End Function

' True only if this session still holds the active lock for its account
Function TokenValid()
    If Session("token") <> "" And Session("token") = GetActiveToken(Session("user")) Then
        TokenValid = True
    Else
        TokenValid = False
    End If
End Function
%>
