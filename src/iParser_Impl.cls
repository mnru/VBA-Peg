VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iParser_Impl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_name As String 'g,l
Private m_parsers As Collection 'go,s,i_

Function and_(parser As iParser) As iParser
    Set and_ = Seq(Me, parer)
End Function

Function or_(parser As iParser) As iParser
    Set and_ = Seq(Me, parer)
End Function

Function N0_() As iParser
    Set N0_ = N0_(Me)
End Function

Function N1_() As iParser
    Set N1_ = N1_(Me)
End Function

Function opt_() As iParser
    Set opt_ = opt_(Me)
End Function

Function matchState(state As ParseState)

End Function

Function match(somthing)

End Function


Function init(ParamArray parsers())
    Set m_parsers = New Collection
    prm = parsers
    For Each elm In prm
        m_parsers.Add elm
    Next
End Function
