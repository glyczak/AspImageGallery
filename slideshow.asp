<%@ Language="VBScript" %>
<%
    Dim fso, directory, files
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set directory = fso.GetFolder("D:\\inetpub\\Gallery\\content\\")
    Set files = directory.files
%>
<!DOCTYPE html>
<html lang="en">
    <head>
        <meta charset="utf-8" />
        <title></title>
    </head>
    <body>
        
    </body>
</html>
