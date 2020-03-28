VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iParser_Impl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_name As String 'i_,g,l
Private m_parsers As Collection 'is,igo_,go,s_

Function and_(parser As iParser) As iParser
    Set and_ = Seq(Me, parser)
End Function

Function or_(parser As iParser) As iParser
    Set or_ = Choice(Me, parser)
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
    End If
End Function

Function init(prm)
    Set m_parsers = New Collection
    For Each elm In prm
        Dim x As iParser
        Set x = elm
        m_parsers.Add x
    Next
End Function
