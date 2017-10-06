<%@ Language="VBScript" %>
<%
    Dim fso, directory, subfolders
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set directory = fso.GetFolder("D:\\inetpub\\Gallery\\content\\")
    Set subfolders = directory.subfolders
%>
<!DOCTYPE html>
<html lang="en">
    <head>
        <meta charset="utf-8" />
        <title></title>
    </head>
    <body>
        <h1>Content Gallery</h1>
        <ul>
            <% For Each subfolder in subfolders %>
                <li> <%= subfolder.name %> - <a href="slideshow.asp?d=<%= subfolder.name %>">Start Slideshow</a></li>
            <% Next %>
        </ul>
    </body>
</html>
