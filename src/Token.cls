VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Token"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements iParser
Private m_name As String 'i_,g,l
Private m_Parsers As Collection 'is,igo_,go,s_

Private Property Get name() As String
    name = m_name
End Property

Private Property Let name(ByVal name_ As String)
    m_name = name_
End Property

Private Property Get Parsers() As Collection
    Set Parsers = m_Parsers
End Property

Public Property Set Parsers(Parsers_ As Collection)
    Set m_Parsers = Parsers_
End Property

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

Function internalMatch(state As ParseState)
End Function

Function match(somthing)
    If TypeName(something) = "String" Then
        Set state = PareseState(something, 0)
    Else
        Set state = internalMatch(something)
    End If
    Set match = state
End Function

Function init(prm)
    Set m_Parsers = New Collection
    For Each elm In prm
        Dim x As iParser
        Set x = elm
        m_Parsers.Add x
    Next
End Function
