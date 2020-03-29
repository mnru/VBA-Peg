Attribute VB_Name = "G"
Function Token(str As String) As Token
    Set Token = New Token
    Token.init (str)
End Function

Function Char(str As String) As Char
    Set Char = New Char
    Char.init (str)
End Function

Function RegEx(str As String) As RegEx
    Set RegEx = New RegEx
    RegEx.init (str)
End Function

Function Seq(ParamArray arg()) As Seq
    Set Seq = New Seq
    prm = arg
    Seq.init (prm)
End Function

Function Choice(ParamArray arg()) As Choice
    Set Choice = New Choice
    prm = arg
    Choice.init (prm)
End Function

Function Rep0or1(ParamArray arg()) As Rep0or1
    Set Rep0or1 = New Rep0or1
    prm = arg
    Rep0or1.init (prm)
End Function

Function Rep0orMore(ParamArray arg()) As Rep0orMore
    Set Rep0orMore = New Rep0orMore
    prm = arg
    Rep0orMore.init (prm)
End Function

Function Rep1orMore(ParamArray arg()) As Rep1orMore
    Set Rep1orMore = New Rep1orMore
    prm = arg
    Rep1orMore.init (prm)
End Function

Function T(ParamArray arg()) As T
    Set T = New T
    prm = arg
    T.init (prm)
End Function

Function F(ParamArray arg()) As F
    Set F = New F
    prm = arg
    F.init (prm)
End Function

Function ParseState(inputs As String, pos As Long, nodes) As ParseState
    Set ParseState = New ParseState
    Call ParseState.init(inputs, pos, nodes)
End Function
