Attribute VB_Name = "classUtilTmpl"
Sub mkCstPAry(clsn As String, toMod As String, tmpln As String, fromMod As String)
    Dim tmpl As String
    Dim sLines As String
    sLines = tmplnToCode(tmpln, fromMod, clsn)
    Call writeToComponent(toMod, vbCrLf & sLines)
End Sub

Sub addInitByPAry(clsn As String, tmpln As String, fromMod As String, clcn As String)
    Dim sLines As String
    sLines = tmplnToCode(tmpln, fromMod, clcn, clsn)
    Call writeToComponent(clsn, vbCrLf & sLines)
End Sub

Sub mkCstInitByPAry(clsn As String, toMod As String, _
    Optional impln As String = "", Optional itfn As String = "", Optional clcn As String = "Parsers", _
    Optional cstn As String = "tmpl_Cst_PAry", Optional initn As String = "tmpl_Init_PAry", Optional fromMod As String = "classUtilTmpl")
    Call mkCstPAry(clsn, toMod, cstn, fromMod)
    If impln = "" Then
        Call mkComponent(clsn, "cls")
    Else
        Call mkSubClass(clsn, impln, itfn)
    End If
    Call addInitByPAry(clsn, initn, fromMod, clcn)
End Sub

Sub mkCstPrm(clsn As String, toMod As String, dclPrm As String, _
    Optional tmpln As String = "tmpl_Cst_Prms", Optional fromMod As String = "classUtilTmpl")
    Dim sLines As String
    sLines = dclPrmToCst(clsn, dclPrm, tmpln, fromMod)
    Call writeToComponent(toMod, vbCrLf & sLines)
End Sub

Sub mkCstInitByDclPrm(clsn As String, toMod As String, dclPrm As String, _
    Optional impln As String = "", Optional itfn As String = "", _
    Optional tmpln As String = "tmpl_Cst_Prms", Optional fromMod As String = "classUtilTmpl")
    Call mkCstPrm(clsn, toMod, dclPrm)
    If impln = "" Then
        Call mkComponent(clsn, "cls")
    Else
        Call mkSubClass(clsn, impln, itfn)
    End If
    Call addInitByDclPrm(clsn, dclPrm)
End Sub

Function tmplnToCode(tmpln As String, modn As String, ParamArray prms())
    Dim tmpl As String
    args = prms
    tmpl = disposeProc("get", modn, tmpln, 0)(1)
    tmplnToCode = tmplToCode(tmpl, args)
End Function

Function tmplToCode(tmpl As String, args)
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
    tmplToCode = Join(ret, vbCrLf)
End Function

Function dclPrmToCst(clsn As String, dclPrm As String, _
    Optional tmpln As String = "tmpl_Cst_Prms", Optional fromMod As String = "classUtilTmpl")
    Dim arg
    Dim tmpl As String
    Dim prms As String
    tmp = dclPrmToAry(dclPrm)
    arg1 = Join(mapA("mkDcl", tmp, "", "", ""), ",")
    arg2 = Join(mapA("getAryAt", tmp, 1), ",")
    arg = Array(clsn, arg1, arg2)
    tmpl = disposeProc("get", fromMod, tmpln)(1)
    dclPrmToCst = tmplToCode(tmpl, arg)
End Function

Function dclPrmToDcl(dclPrm As String)
    Dim ret
    tmp = dclPrmToAry(dclPrm)
    tmp0 = mapA("mkdcl", tmp, "auto", "m_", "")
    ret = Join(tmp0, vbCrLf) & vbCrLf & vbCrLf
    dclPrmToDcl = ret
End Function

Function dclPrmToInit(dclPrm As String) As String
    Dim ret As String
    tmp = dclPrmToAry(dclPrm)
    tmp1 = mapA("mkdcl", tmp, "", "", "_")
    tmp2 = mapA("mkasn", tmp)
    tmp3 = mapA("addstr", tmp2, " ")
    ret = "Sub init(" & Join(tmp1, ",") & ")" & vbCrLf & Join(tmp3, vbCrLf) & vbCrLf & "End Sub"
    dclPrmToInit = ret
End Function

Sub addInitByDclPrm(clsn As String, dclPrm As String)
    str1 = dclPrmToDcl(dclPrm)
    str2 = dclPrmToInit(dclPrm)
    Set cmp = getComponent(clsn)
    With cmp.CodeModule
        .AddFromString (str1 & vbCrLf & str2)
    End With
End Sub

Sub overRide(fnc As String, knd As Long, toMod As String, tmpln As String, fromMod As String)
    Dim tmpln As String
    tmpln = Join(Array("ovr", toMod, fnc), "_")
    sCode = tmplnToCode(tmpln, fromMod)
    Call disposeProc("replace", toMod, fnc, knd, sCode)
End Sub

Function tmpl_Cst_PAry()
    'Function $0(ParamArray arg()) As $0
    ' Set $0 = New $0
    ' prm = arg
    ' $0.init (prm)
    'End Function
End Function

Function tmpl_Cst_Prms()
    'Function $0($1) As $0
    ' Set $0 = New $0
    ' call $0.init($2)
    'End Function
End Function

Function tmpl_init_PAry()
    'Sub init(prm)
    ' Set m_$0 = New Collection
    ' For Each elm In prm
    ' Dim x As $1
    ' Set x = elm
    ' m_$0.Add x
    ' Next
    'End Sub
End Function

Function tmpl_init_Prm()
    'Sub init($0)
    '$1
    'End Sub
End Function

Function dclPrmToAry(dclPrm As String)
    tmp = Split(dclPrm, ",")
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
