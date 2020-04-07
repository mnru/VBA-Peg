Attribute VB_Name = "iParser_Impl"
Private m_name As String 'i_,g,l
Private m_Parsers As Collection 'is,igo_,go,s_

Function and_(parser As iParser) As iParser
    Set and_ = seq(Me, parser)
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
