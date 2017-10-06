<%@ Language="VBScript" %>
<%
    Randomize Timer

    ' declare all variables
    Dim objFSO, objFolder, objFiles, objFile
    Dim strFiles, strImages, strPhysical, strFile, strImage

    strPath = "content/" & Request.Querystring("d")

    ' this constant has the names of valid image file name
    ' extensions and can be modified for more image types
    Const strValid = ".gif.jpg.jpeg.png"

    ' make sure we have a trailing slash in the path
    If Right(strPath,1) <> Chr(47) Then strPath = strPath & Chr(47)
    ' get the physical path of the folder
    strPhysical = Server.MapPath(strPath)
    ' get a File System Object
    Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
    ' create a folder object
    Set objFolder = objFSO.GetFolder(strPhysical)
    ' get the files collection
    Set objFiles = objFolder.Files

    ' enumerate the files collection looking for images
    For Each objFile in objFiles
      strFile = LCase(objFile.Name)
      If Instr(strValid,Right(strFile,4)) Then
        ' add vaild images to a string of image names
        strFiles = strFiles & strFile & vbTab
      End If
    Next

    ' split the image names into an array
    strImages = Split(strFiles,vbTab)
%>
<!DOCTYPE html>
<html lang="en">
    <head>
        <meta charset="utf-8" />
        <title>Slideshow</title>
    </head>
    <body>
        <div id="images">
            <%
                For Each strImage In strImages
                If strImage <> "" Then
            %>
                <img class="slide" width="100%" src="content/<%= Request.Querystring("d") %>/<%= strImage %>"></img>
            <%
                End If
                Next
            %>
        </div>
        <script>
            var slideIndex = 0;
            showSlides();

            function showSlides() {
                var i;
                var slides = document.getElementsByClassName("slide");
                for (i = 0; i < slides.length; i++) {
                    slides[i].style.display = "none"; 
                }
                slideIndex++;
                if (slideIndex > slides.length) {slideIndex = 1} 
                slides[slideIndex-1].style.display = "block"; 
                setTimeout(showSlides, 3000);
            }
        </script>
    </body>
</html>
