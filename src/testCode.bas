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
'//SubClass
'Private m_inputs As String
'Private m_ pos As Long
'Private m_ nodes As Collection
'
'Sub init(inputs_ As String, pos_ As Long, nodes_ As Collection)
'  m_inputs = inputs_
'  m_ pos = pos_
'  m_ nodes = nodes_
'End Sub
'//Constructor
'Function Node(a As String, b As Long) As Node
' Set Node = New Node
' Call Node.init(a, b)
'End Function

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

Sub testDcl()
    clsns = Array("classGenerator", "classUtil", "G", "iParser_Impl")
    Call delComponent("N0_")
    Call delComponentExcept(clsns)
End Sub

Sub testPart()
    x = Split("i,j,k,gt,ilt")
    y = partSymbol(x)
    Stop
End Sub

Sub testInit()
    x = dclPrmToInit("inputs;String, pos;Long;_, nodes;;o")
    Debug.Print x
End Sub

Sub testTmpl1()
    str0 = dclPrmToCst("ParseState", "inputs;String, pos;Long, nodes")
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

Sub testTmpl()
    Dim arg
    Dim tmpl As String
    Dim dclPrms As String
    Dim prms As String
    fromMod = "classUtilTmpl"
    tmpln = "tmpl_cst_prms"
    prms = "a;string,b;long"
    dcl = expandDcl(prms, "dcl", "", "", "")
    asn = expandDcl(prms, "asn")
    arg = Array("Node", dcl, asn)
    '  Set cmp = mkComponent(toMod, "std")
    '  With cmp.CodeModule
    tmpl = disposeProc("get", fromMod, tmpln)(1)
    ret = tmplToCode0(tmpl, arg)
    Debug.Print ret
    '  End With
End Sub

Sub testdclprm()
    Dim dclprm As String
    dclprm = "inputs;String, pos;Long, nodes"
    Debug.Print "//dclPrmToDcl"
    Debug.Print dclPrmToDcl(dclprm)
    Debug.Print "//dclPrmToInit"
    Debug.Print dclPrmToInit(dclprm)
    Debug.Print "//dclPrmToCst"
    Debug.Print dclPrmToCst("Node", dclprm)
End Sub
