Attribute VB_Name = "classGenerator"
'R_ "Rep0or1",R0 "Rep0orMore",R1 "Rep1orMore"
'
'"AnyChar","ASTNode","Char","CharRange","CharSet","Choice","classGenerator","classUtil","Delay","F","G","iParser","iParser_Impl","N1_","Opt","Parser_Impl","ParseState","RegEx","Rule","Seq","T","Token"

Sub genrateConstructor()
    clsns = Array("Seq", "Choice", "Rep0or1", "Rep0orMore", "Rep1orMore", "T", "F")
    Call mkInterFace("iParser", "iParser_impl")
    Call mkSubClass("iParser", "iParser_impl", clsns)
    Call mkCst("iParser", "G", "classGenerator", clsns)
End Sub

Sub initializeClass()
    clsns = Array("iParser_Impl", "classUtil", "classGenerator", "testCode", "ParseState")
    delModComponentExcept (clsns)
End Sub

Function Cst_iParser()
    'Function ?(ParamArray arg()) As ?
    '  Set ? = New ?
    '  prm = arg
    '  ?.init (prm)
    'End Function
End Function
