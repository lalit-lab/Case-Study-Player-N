<%@ Language="JScript" %>
<%
Response.ContentType = "application/json";
Response.AddHeader("Access-Control-Allow-Origin", "*");
Response.AddHeader("Cache-Control", "no-cache");

if (Request.ServerVariables("REQUEST_METHOD") == "POST") {
    var fso = Server.CreateObject("Scripting.FileSystemObject");
    var fp = Server.MapPath("events.jsonl");
    if (fso.FileExists(fp)) {
        fso.DeleteFile(fp);
        fso.CreateTextFile(fp, true);
    }
    Response.Write("[]");
    Response.End();
}

try {
    var fso = Server.CreateObject("Scripting.FileSystemObject");
    var filePath = Server.MapPath("events.jsonl");
    if (!fso.FileExists(filePath)) {
        Response.Write("[]");
        Response.End();
    }
    var f = fso.OpenTextFile(filePath, 1, false);
    var content = f.AtEndOfStream ? "" : f.ReadAll();
    f.Close();
    var lines = content.split("\n");
    var valid = [];
    for (var i = 0; i < lines.length; i++) {
        var line = lines[i].replace(/\r/g, "").trim();
        if (line && line.charAt(0) === "{") valid.push(line);
    }
    Response.Write("[" + valid.join(",") + "]");
} catch(e) {
    Response.Write("[]");
}
%>
