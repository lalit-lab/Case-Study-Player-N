<%
Response.ContentType = "text/plain"
Response.AddHeader "Access-Control-Allow-Origin", "*"
Response.AddHeader "Access-Control-Allow-Methods", "POST, OPTIONS"
Response.AddHeader "Access-Control-Allow-Headers", "Content-Type"

If Request.ServerVariables("REQUEST_METHOD") = "OPTIONS" Then
    Response.Status = "200 OK"
    Response.End
End If

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    On Error Resume Next

    Dim bytes, stream, body
    bytes = Request.BinaryRead(Request.TotalBytes)

    Set stream = Server.CreateObject("ADODB.Stream")
    stream.Open
    stream.Type = 1
    stream.Write bytes
    stream.Position = 0
    stream.Type = 2
    stream.CharSet = "UTF-8"
    body = stream.ReadText
    stream.Close
    Set stream = Nothing

    Dim ts, mm, dd, hh, mi, ss
    ts = Now()
    mm = Right("0" & Month(ts), 2)
    dd = Right("0" & Day(ts), 2)
    hh = Right("0" & Hour(ts), 2)
    mi = Right("0" & Minute(ts), 2)
    ss = Right("0" & Second(ts), 2)
    Dim tsStr
    tsStr = Year(ts) & "-" & mm & "-" & dd & "T" & hh & ":" & mi & ":" & ss & "Z"

    Dim jsonLine
    If Len(body) > 1 Then
        jsonLine = "{""ts"":""" & tsStr & """," & Mid(body, 2)
    Else
        jsonLine = "{""ts"":""" & tsStr & """}"
    End If

    Dim fso, filePath, f
    filePath = Server.MapPath("events.jsonl")
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(filePath, 8, True)
    f.WriteLine jsonLine
    f.Close
    Set fso = Nothing

    Response.Write "ok"
End If
%>
