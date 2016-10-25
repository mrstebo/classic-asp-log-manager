<%
Class FileLogger
  Private pName
  Private pFilename

  Public Default Function construct(name, filename)
    pName = name
    pFilename = filename
    Set construct = Me
  End Function

  Public Sub Log(logObject)
    Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
    Dim f : Set f = fso.OpenTextFile(pFilename, 8, True, -1)
    Dim content : content = logObject.ToJSON()

    f.WriteLine(content)
    f.Close()
    
    Set f = Nothing
    Set fso = Nothing
  End Sub
End Class
%>
