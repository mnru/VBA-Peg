Attribute VB_Name = "G"
Function ParseState(inputs As String, pos As Long, nodes) As ParseState
    Set ParseState = New ParseState
    Call ParseState.init(inputs, pos, nodes)
End Function
