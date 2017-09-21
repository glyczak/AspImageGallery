<%@ Language="VBScript" %>
<%
    Dim directory, file
    directory = Request.Querystring("d")
    file = Request.Querystring("f")
    If StrComp(directory, "") = 0 Then
        directory = "D:\\inetpub\\Gallery\\content\\"
    Else
        directory = "D:\\inetpub\\Gallery\\content\\" & Request.Querystring("d") & "\\"
    End If
    Response.Buffer = False
    Dim objStream
    Set objStream = Server.CreateObject("ADODB.Stream")
    objStream.Type = 1 'adTypeBinary
    objStream.Open
    objStream.LoadFromFile(directory & file)
    Response.ContentType = "application/x-unknown"
    Response.Addheader "Content-Disposition", "attachment; filename=" & file
    Response.BinaryWrite objStream.Read
    objStream.Close
    Set objStream = Nothing
%>