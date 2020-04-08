Attribute VB_Name = "classGenerator"
'R_ "Rep0or1",R0 "Rep0orMore",R1 "Rep1orMore"
'
'"AnyChar","ASTNode","Char","CharRange","CharSet","Choice","classGenerator","classUtil","Delay","F","G","iParser","iParser_Impl","N1_","Opt","Parser_Impl","ParseState","RegEx","Rule","Seq","T","Token"

Sub genrateConstructor()
    Dim tmplmod As String
    tmplmod = "classUtilTmpl"
    ovrmod = "classGenerator"
    clsns1 = Array("Seq", "Choice", "Rep0or1", "Rep0orMore", "Rep1orMore", "T", "F")
    clsns2 = Array("Token", "Char", "RegEx")
    Call mkInterFace("iParser", "iParser_impl")
    Call mapA("mkCstInitByPAry", clsns1, "G", "iParser_impl", "iParser", "Parsers")
    Call mapA("mkCstInitByDclPrm", clsns2, "G", "str;String", "iParser_impl", "iParser")
    Call mkCstPrm("ParseState", "G", "inputs;String, pos;Long,nodes;")
    Call mkCstPrm("Node", "G", "begin;long,label;String,inputs;String")
End Sub

Sub initializeClass()
    clsns = Array("iParser_Impl", "classUtil", "classUtilCmp", "classUtilPrp", "classUtilTmpl", "classGenerator", "testCode", "ParseState", "Node", "FunctionalArrayMin")
    delComponentExcept (clsns)
End Sub
