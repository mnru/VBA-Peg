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

Function testcode2(src As String, prms)
    Dim tmp
    Dim arg
    Dim n, i, j
    arg = prms
    tmp = Split(src, vbCrLf)
    ReDim ret(LBound(tmp) To UBound(tmp))
    For i = LBound(tmp) To UBound(tmp)
        sLine = Trim(tmp(i))
        If Len(sLine) > 0 And Left(sLine, 1) = "'" Then sLine = Right(sLine, Len(sLine) - 1)
        For j = 0 To UBound(arg)
            v = "$" & j
            sLine = Replace(sLine, v, arg(j))
        Next j
        tmp(i) = sLine
    Next i
    writePrmsToTmpl = Join(tmp, vbCrLf)
End Function

Function testcode3(src As String, prms)
End Function

Sub testcode4()
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
    Call delComponent("N0_")
    Call delComponentExcept(clsns)
End Sub

Sub testPart()
    x = Split("i,j,k,gt,ilt")
    y = partSymbol(x)
    Stop
End Sub

Sub testTmpl()
    Dim arg
    Dim tmpl As String
    Dim dclPrms As String
    Dim prms As String
    fromMod = "classGenerator"
    tmpln = "Parser_Prm"
    dclPrms = "a as string,b as long"
    prms = delTypeInDcl(dclPrms)
    arg = Array("Node", dclPrms, prms)
    '   Set cmp = mkComponent(toMod, "std")
    '    With cmp.CodeModule
    tmpl = disposeProc("get", fromMod, "Cst_" & tmpln)(1)
    ret = writePrmsToTmpl(tmpl, arg)
    Debug.Print ret
    '   End With
End Sub

Sub testTmpl1()
    str0 = mkCstPrmLines("ParseState", "inputs As String, pos As Long, nodes")
    Debug.Print str0
End Sub

Sub testo()
    Set cmp = mkComponent("testCode", "std")
    With cmp.CodeModule
        Call .InsertLines(.ProcBodyLine("testtmpl1", 0), "")
    End With
End Sub

Sub testOverride()
    Call overRide("testcode2", 0, "testCode", "classUtil")
    Call overRide("testcode3", 0, "testCode", "classUtil")
End Sub
