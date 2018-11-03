Sub dictFunc()

	‘ENSURE MICROSOFT SCRIPTING RUNTIME IS ENABLED, instead of Key,Item, can just list dict.Add “test”,6
	Dim dict As New Scripting.Dictionary
	Set dict = New Scripting.Dictionary
	dict.Add Key:="test_key_1", Item:=6
	dict.Add Key:="test_key_2", Item:=12
	
	If dict.Exists("test_key_1") Then
		Debug.print(dict("test_key_1"))
	End If

	' Iterate over whole dictionary
	Dim k As Variant
	For Each k in dict.Keys
		Debug.print k, dict(k)
	Next


End Sub