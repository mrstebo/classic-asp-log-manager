<%
Class DetailedLogObject
	Public LogLevel
	Public Message
	Public Tags
	Public Meta

	Private Sub Class_Initialize
		Set Tags = New ArrayList
		Set Meta = CreateObject("Scripting.Dictionary")
	End Sub

	Private Sub Class_Terminate
		Set Meta = Nothing
		Set Tags = Nothing
	End Sub

	Public Function ToJSON()
		Dim result : result = ""

		result = result & "{"
		result = result & """level"": "&LogLevel&","
		result = result & """message"": """&JsonUtils.MakeJsonSafe(Message)&""","
		result = result & """tags"": ["
		result = result & Join(TagsToArray(), ",")
		result = result & "],"
		result = result & """meta"": ["
		result = result & Join(MetaToArray(), ",")
		result = result & "]"
		result = result & "}"

		ToJson = result
	End Function

	Private Function TagsToArray()
		Dim arr : Set arr = New ArrayList

		For Each tag in Tags.ToArray()
			Call arr.AddItem(""""&JsonUtils.MakeJsonSafe(tag)&"""")
		Next

		TagsToArray = arr.ToArray()
	End Function

	Private Function MetaToArray()
		Dim arr : Set arr = New ArrayList

		For Each key in Meta.Keys
			Call arr.AddItem("{""key"": """&JsonUtils.MakeJsonSafe(key)&""", ""value"": """&JsonUtils.MakeJsonSafe(Meta(key))&"""}")
		Next

		MetaToArray = arr.ToArray()
	End Function
End Class
%>
