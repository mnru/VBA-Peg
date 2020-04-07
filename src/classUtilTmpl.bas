Attribute VB_Name = "classUtilTmpl"
Sub mkCst(toMod As String, tmpln As String, fromMod As String, clsns)
    Dim arg
    Dim tmpl As String
    Dim sLines As String
    arg = clsns
    Set cmp = mkComponent(toMod, "std")
    With cmp.CodeModule
        tmpl = disposeProc("get", fromMod, tmpln)(1)
        For i = UBound(arg) To LBound(arg) Step -1
            sLines = tmplToCode(tmpl, CStr(arg(i)))
            .AddFromString (vbCrLf & sLines)
        Next i
    End With
End Sub

Sub mkCstPrm(clsn As String, toMod As String, dclPrms As String, _
    Optional tmpln As String = "tmpl_Cst_Prms", Optional fromMod As String = "classUtilTmpl")
    Dim sLines As String
    Dim cmp
    sLines = dclPrmToCst(clsn, dclPrms, tmpln, fromMod)
    Set cmp = mkComponent(toMod, "std")
    With cmp.CodeModule
        .AddFromString (vbCrLf & sLines)
    End With
End Sub

Function tmplToCode(tmpl As String, ParamArray prms())
    args = prms
    tmplToCode = tmplToCode0(tmpl, args)
End Function

Function tmplToCode0(tmpl As String, args)
    Dim ret, sLine
    Dim i As Long, j As Long
    ret = Split(tmpl, vbCrLf)
    For i = LBound(ret) To UBound(ret)
        sLine = LTrim(ret(i))
        If Len(sLine) > 0 And Left(sLine, 1) = "'" Then sLine = Right(sLine, Len(sLine) - 1)
        For j = 0 To lenAry(args) - 1
            sLine = Replace(sLine, "$" & j, getAryAt(args, j, 0))
        Next j
        ret(i) = sLine
    Next i
    tmplToCode0 = Join(ret, vbCrLf)
End Function

Function dclPrmToCst(clsn As String, dclPrms As String, _
    Optional tmpln As String = "tmpl_Cst_Prms", Optional fromMod As String = "classUtilTmpl")
    Dim arg
    Dim tmpl As String
    Dim prms As String
    tmp = dclPrmToAry(dclPrms)
    arg1 = Join(mapA("mkDcl", tmp, "", "", ""), ",")
    arg2 = Join(mapA("getAryAt", tmp, 1), ",")
    arg = Array(clsn, arg1, arg2)
    tmpl = disposeProc("get", fromMod, tmpln)(1)
    dclPrmToCst = tmplToCode0(tmpl, arg)
End Function

Function dclPrmToDcl(dclPrms As String)
    Dim ret
    tmp = dclPrmToAry(dclPrms)
    tmp0 = mapA("mkdcl", tmp, "auto", "m_", "")
    ret = Join(tmp0, vbCrLf) & vbCrLf & vbCrLf
    dclPrmToDcl = ret
End Function

Function dclPrmToInit(dclPrms As String) As String
    Dim ret As String
    tmp = dclPrmToAry(dclPrms)
    tmp1 = mapA("mkdcl", tmp, "", "", "_")
    tmp2 = mapA("mkasn", tmp)
    tmp3 = mapA("addstr", tmp2, "  ")
    ret = "Sub init(" & Join(tmp1, ",") & ")" & vbCrLf & Join(tmp3, vbCrLf) & vbCrLf & "End Sub"
    dclPrmToInit = ret
End Function

Sub addInitByDclPrm(clsn As String, dclprm As String)
    str1 = dclPrmToDcl(dclprm)
    str2 = dclPrmToInit(dclprm)
    Set cmp = getComponent(clsn)
    With cmp.CodeModule
        .AddFromString (str1 & vbCrLf & str2)
    End With
End Sub

Sub overRide(fnc As String, knd As Long, toMod As String, tmpln As String, fromMod As String)
    Dim sLines, cmp
    tmpln = Join(Array("ovr", toMod, fnc), "_")
    tmpl = disposeProc("get", fromMod, tmpln)
    sCode = tmplToCode(tmpl)
    Call disposeProc("replace", toMod, fnc, knd, sCode)
End Sub

Function tmpl_Cst_ParamArray()
    'Function $0(ParamArray arg()) As $0
    ' Set $0 = New $0
    ' prm = arg
    ' $0.init (prm)
    'End Function
End Function

Function tmpl_Cst_String()
    'Function $0(str As String) As $0
    ' Set $0 = New $0
    ' $0.init (str)
    'End Function
End Function

Function tmpl_Cst_Prms()
    'Function $0($1) As $0
    ' Set $0 = New $0
    ' call $0.init($2)
    'End Function
End Function

Function tmpl_init_ParamAry()
    'Function init(prm)
    ' Set m_$0 = New Collection
    ' For Each elm In prm
    '  Dim x As $1
    '  Set x = elm
    '  m_$0.Add x
    ' Next
    'End Function
End Function

Function dclPrmToAry(dclprm As String)
    tmp = Split(dclprm, ",")
    ret = mapA("mcsplit", tmp, ";")
    For i = LBound(ret) To UBound(ret)
        ret(i) = mapA("mcTrim", ret(i))
    Next i
    dclPrmToAry = ret
End Function

Function mkAsn(ary)
    vr = getAryAt(ary, 1)
    ret = "m_" & vr & " = " & vr & "_"
    If lenAry(ary) >= 3 Then
        x = LCase(getAryAt(ary, 3))
        If InStr(x, "o") > 0 Then ret = "Set " & ret
    End If
    mkAsn = ret
End Function

Function mkDcl(ary, Optional scp = "auto", Optional pre = "m_", Optional suf = "_")
    vr = getAryAt(ary, 1)
    tp = ""
    If lenAry(ary) >= 2 Then tp = getAryAt(ary, 2)
    If LCase(scp) = "auto" Then
        scp = "Private "
        If lenAry(ary) >= 3 Then
            x = getAryAt(ary, 3)
            If InStr(x, "_") > 0 Then
                scp = "Public "
            End If
        End If
    End If
    ret = scp & pre & vr & suf
    If tp <> "" Then ret = ret & " As " & tp
    mkDcl = ret
End Function

Function mcSplit(str, dlm)
    mcSplit = Split(str, dlm)
End Function

Function mcTrim(str)
    mcTrim = Trim(str)
End Function

Function addStr(x, Optional a = "", Optional b = "")
    addStr = a & x & b
End Function
