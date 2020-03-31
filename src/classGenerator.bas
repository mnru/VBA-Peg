Attribute VB_Name = "classGenerator"
'R_ "Rep0or1",R0 "Rep0orMore",R1 "Rep1orMore"
'
'"AnyChar","ASTNode","Char","CharRange","CharSet","Choice","classGenerator","classUtil","Delay","F","G","iParser","iParser_Impl","N1_","Opt","Parser_Impl","ParseState","RegEx","Rule","Seq","T","Token"

Sub genrateConstructor()
    clsns1 = Array("Seq", "Choice", "Rep0or1", "Rep0orMore", "Rep1orMore", "T", "F")
    clsns2 = Array("Token", "Char", "RegEx")
    Call mkInterFace("iParser", "iParser_impl")
    Call mkSubClass("iParser", "iParser_impl", clsns1)
    Call mkSubClass("iParser", "iParser_impl", clsns2)
    Call mkCst("Parser_Parsers", "G", "classGenerator", clsns1)
    Call mkCst("Parser_String", "G", "classGenerator", clsns2)
    Call mkCstPrm("G", "ParseState", "inputs As String, pos As Long, nodes")
    Call mkCstPrm("G", "Node", "begin As long,label As String,inputs As String")
End Sub

Sub initializeClass()
    clsns = Array("iParser_Impl", "classUtil", "classGenerator", "testCode", "ParseState", "Node")
    delComponentExcept (clsns)
End Sub

Function Cst_Parser_Parsers()
    'Function $0(ParamArray arg()) As $0
    '  Set $0 = New $0
    '  prm = arg
    '  $0.init (prm)
    'End Function
End Function

Function Cst_Parser_String()
    'Function $0(str As String) As $0
    '  Set $0 = New $0
    '  $0.init (str)
    'End Function
End Function

Function Cst_Parser_Prm()
    'Function $0($1) As $0
    '  Set $0 = New $0
    '  call $0.init($2)
    'End Function
End Function
