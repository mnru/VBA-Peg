VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Node"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_begin As Long 'g,l
Private m_label As String 'g,l
Private m_inputs As String 'g,l

Private Property Get begin() As Long
    begin = m_begin
End Property

Private Property Let begin(ByVal begin_ As Long)
    m_begin = begin_
End Property

Private Property Get label() As String
    label = m_label
End Property

Private Property Let label(ByVal label_ As String)
    m_label = label_
End Property

Private Property Get inputs() As String
    inputs = m_inputs
End Property

Private Property Let inputs(ByVal inputs_ As String)
    m_inputs = inputs_
End Property
