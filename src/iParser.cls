VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public name As String

Private Property Set parsers(parsers_ As Collection)
End Property

Public Property Get parsers() As Collection
End Property

Function and_(parser As iParser) As iParser
End Function

Function or_(parser As iParser) As iParser
End Function

Function r0() As iParser
End Function

Function r1() As iParser
End Function

Function r_() As iParser
End Function

Function matchState(state As ParseState)
End Function

Function match(somthing)
End Function

Function init(ParamArray parsers())
End Function
