Attribute VB_Name = "G"
Function Seq(ParamArray parsers()) As Seq
    Seq = New Seq
    arg = parsers
    Seq.init (parsers)
End Function

Function Choice(ParamArray parsers()) As Choice
    Set Seq = New Choice
    arg = parsers
    Choice.init (parsers)
End Function

Function N0_(ParamArray parsers()) As N0_
    Set N0_ = New N0_
    arg = parsers
    N0_.init (parsers)
End Function

Function N1_(ParamArray parsers()) As N1_
    Set N1_ = New N1_
    arg = parsers
    N1_.init (parsers)
End Function

Function Opt(ParamArray parsers()) As Opt
    Set Opt = New Opt
    arg = parsers
    Opt.init (parsers)
End Function

Function T(ParamArray parsers())
    Set T = New Opt
    arg = parsers
    T.init (parsers)
End Function

Function F(ParamArray parsers())
    Set F = New F
    arg = parsers
    F.init (parsers)
End Function
