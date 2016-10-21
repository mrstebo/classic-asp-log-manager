<%
Class JsonUtils_
	Public Function IsJson(value)
		IsJson = False

		If Len(value&"") > 0 Then
			If Mid(value, 1, 1) = "{" And Mid(value, Len(value), 1) = "}" Then
				IsJson = True
			ElseIf Mid(value, 1, 1) = "[" And Mid(value, Len(value), 1) = "]" Then
				IsJson = True
			End If
		End If
	End Function

	Public Function MakeJsonSafe(value)
		Dim result : result = value&""

		result = Replace(result, "\", "\\")
		result = Replace(result, """", "\""")
		result = Replace(result, "'", "\'")

		MakeJsonSafe = result
	End Function
End Class

Dim JsonUtils : Set JsonUtils = New JsonUtils_
%>
