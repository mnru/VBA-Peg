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
    Call mkSubClass("iParser", "iParser_impl", clsns1)
    Call mkSubClass("iParser", "iParser_impl", clsns2)
    Call mkCst("G", "tmpl_Cst_ParamArray", tmplmod, clsns1)
    Dim dclprm2 As String
    dclprm2 = "str;String"
    For Each clsn In clsns2
        Call mkCstPrm("G", CStr(clsn), dclprm2)
        Call addInitByDclPrm(CStr(clsn), dclprm2)
    Next
    Call mkCstPrm("G", "ParseState", "inputs;String, pos;Long,nodes;")
    Call mkCstPrm("G", "Node", "begin;long,label;String,inputs;String")
End Sub

Sub initializeClass()
    clsns = Array("iParser_Impl", "classUtil", "classUtilCmp", "classUtilPrp", "classUtilTmpl", "classGenerator", "testCode", "ParseState", "Node", "FunctionalArrayMin")
    delComponentExcept (clsns)
End Sub
