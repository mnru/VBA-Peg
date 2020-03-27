VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iParser_Impl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
MultiUse = -1 'True
End
Private m_name As String 'g,l
Private m_parsers As Collection 'go,s,i_

Function and_(parser As iParser) As iParser
    Set and_ = Seq(Me, parer)
End Function

Function or_(parser As iParser) As iParser
    Set and_ = Seq(Me, parer)
End Function

Function r0() As iParser
    Set r0 = Rep0orMore(Me)
End Function

Function r1() As iParser
    Set r1 = Rep1orMore(Me)
End Function

Function r_() As iParser
    Set r_ = Rep0or1(Me)
End Function

Function matchState(state As ParseState)
End Function

Function match(somthing)
    If TypeName(something) = "String" Then
        Set state = Paresestate(something, 0)
    End Function

    Function init(ParamArray parsers())
        Set m_parsers = New Collection
        prm = parsers
        For Each elm In prm
            m_parsers.Add elm
        Next
    End Function
