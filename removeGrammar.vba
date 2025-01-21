Function removeGrammar(Text) As String

    Dim gVar, grammarArr As Variant
    
    grammarArr = Array(".", ",", "!", "?", ":", ";", "-", "(", ")", "'")
    
    For Each gVar In grammarArr
        Text = Replace(Text, gVar, " ")
    Next gVar
    
    removeGrammar = Text

End Function