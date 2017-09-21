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
        <a href="slideshow.asp">Start Slideshow</a>
        <ul>
            <%
                For Each subfolder in subfolders 
                    Response.write("<li>" & subfolder.name & "</li>") 
                Next 
            %>
        </ul>
        <img src="ajax.asp?d=cats&f=cat1.jpg"></img>
    </body>
</html>
