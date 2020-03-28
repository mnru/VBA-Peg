Attribute VB_Name = "testCode"

Sub test0()
    strA = Join(Split("g,f,e,a,b", ","), vbCrLf)
    Call disposeProc("replace", "Logwriter", "myoutput", , strA)
End Sub

Sub testCode0()
    Call mkInterFace("", "iParser_Impl")
End Sub

Sub testcode1()
    Call mkSubClass("SpecialLogWriter", , "LogWriter")
End Sub

Sub testprp()
    Debug.Print mkPrpStatement("abc", "Long", "g")(1)
    Debug.Print mkPrpStatement("abc", "Long", "g_")(1)
    Debug.Print mkPrpStatement("abc", "Long", "s_")(1)
    Debug.Print mkPrpStatement("abc", "Long", "s")(1)
    Debug.Print mkPrpStatement("abc", "Long", "sov")(1)
    Debug.Print mkPrpStatement("abc", "Long", "l")(1)
    Debug.Print mkPrpStatement("abc", "Long", "il")(1)
    Debug.Print mkPrpStatement("abc", "Long", "ig_")(1)
    Debug.Print mkPrpStatement("abc", "Long", "is_")(1)
    Debug.Print mkPrpStatement("abc", "Long", "i")(1)
    Debug.Print mkPrpStatement("abc", "Long", "i_")(1)
    Debug.Print mkPrpStatement("abc", "", "g")(1)
    Debug.Print mkPrpStatement("abc", "", "g_")(1)
    Debug.Print mkPrpStatement("abc", "", "s_")(1)
    Debug.Print mkPrpStatement("abc", "", "s")(1)
    Debug.Print mkPrpStatement("abc", "", "sov")(1)
    Debug.Print mkPrpStatement("abc", "", "l")(1)
    Debug.Print mkPrpStatement("abc", "", "il")(1)
    Debug.Print mkPrpStatement("abc", "", "ig_")(1)
    Debug.Print mkPrpStatement("abc", "", "is_")(1)
    Debug.Print mkPrpStatement("abc", "", "i")(1)
    Debug.Print mkPrpStatement("abc", "", "i_")(1)
End Sub

Sub testDel()
    clsns = Array("classGenerator", "classUtil", "G", "iParser_Impl")
    Call delModComponent("N0_")
    Call delModComponentExcept(clsns)
End Sub
'
'"AnyChar","ASTNode","Char","CharRange","CharSet","Choice","classGenerator","classUtil","Delay","F","G","iParser","iParser_Impl","N1_","Opt","Parser_Impl","ParseState","RegEx","Rule","Seq","T","Token"

Sub testPart()
    x = Split("i,j,k,gt,ilt")
    y = partSymbol(x)
    Stop
End Sub

