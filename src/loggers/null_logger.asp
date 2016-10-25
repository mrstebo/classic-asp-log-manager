<%
' So we don't end up with null reference-type exceptions :)
Class NullLogger
  Private pName

  Public Default Function construct(name)
    pName = name
    Set construct = Me
  End Function

  Public Sub Log(logObject)
  End Sub
End Class
%>
