<%@ Language="JScript" %>
<%
Response.ContentType = "text/plain";
Response.AddHeader("Access-Control-Allow-Origin", "*");
Response.AddHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
Response.AddHeader("Access-Control-Allow-Headers", "Content-Type");

if (Request.ServerVariables("REQUEST_METHOD") == "OPTIONS") {
    Response.Status = "200 OK";
    Response.End();
}

if (Request.ServerVariables("REQUEST_METHOD") == "POST") {
    try {
        var bytes = Request.BinaryRead(Request.TotalBytes);
        var stream = Server.CreateObject("ADODB.Stream");
        stream.Open();
        stream.Type = 1;
        stream.Write(bytes);
        stream.Position = 0;
        stream.Type = 2;
        stream.CharSet = "UTF-8";
        var body = stream.ReadText();
        stream.Close();

        var ts = new Date().toISOString();
        var enriched = '{"ts":"' + ts + '",' + body.substring(1);

        var fso = Server.CreateObject("Scripting.FileSystemObject");
        var filePath = Server.MapPath("events.jsonl");
        var f = fso.OpenTextFile(filePath, 8, true);
        f.WriteLine(enriched);
        f.Close();

        Response.Write("ok");
    } catch(e) {
        Response.Write("err:" + e.message);
    }
}
%>
