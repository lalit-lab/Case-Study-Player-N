<%
Response.ContentType = "application/json"
Response.AddHeader "Access-Control-Allow-Origin", "*"
Response.AddHeader "Cache-Control", "no-cache"

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim fso2, fp2
    fp2 = Server.MapPath("events.jsonl")
    Set fso2 = Server.CreateObject("Scripting.FileSystemObject")
    If fso2.FileExists(fp2) Then
        fso2.DeleteFile fp2
    End If
    Dim fc
    Set fc = fso2.CreateTextFile(fp2, True)
    fc.Close
    Set fso2 = Nothing
    Response.Write "[]"
    Response.End
End If

Dim fso, filePath, f, content, lines, valid, i, line
filePath = Server.MapPath("events.jsonl")
Set fso = Server.CreateObject("Scripting.FileSystemObject")

If Not fso.FileExists(filePath) Then
    Response.Write "[]"
    Response.End
End If

Set f = fso.OpenTextFile(filePath, 1)
If f.AtEndOfStream Then
    f.Close
    Set fso = Nothing
    Response.Write "[]"
    Response.End
End If
content = f.ReadAll
f.Close
Set fso = Nothing

lines = Split(content, Chr(10))
valid = ""
For i = 0 To UBound(lines)
    line = Trim(Replace(lines(i), Chr(13), ""))
    If Left(line, 1) = "{" Then
        If Len(valid) > 0 Then valid = valid & ","
        valid = valid & line
    End If
Next

Response.Write "[" & valid & "]"
%>
