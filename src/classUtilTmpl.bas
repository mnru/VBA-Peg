Attribute VB_Name = "classUtilTmpl"
Public interfaceDics

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

Sub mkCstPrm(toMod As String, clsn As String, dclprms As String, _
    Optional tmpln As String = "tmpl_Cst_Prms", Optional fromMod As String = "classUtilTmpl")
    Dim sLines As String
    Dim cmp
    sLines = mkCstPrmLines(clsn, dclprms, tmpln, fromMod)
    Set cmp = mkComponent(toMod, "std")
    With cmp.CodeModule
        .AddFromString (vbCrLf & sLines)
    End With
End Sub

Function mkCstPrmLines(clsn As String, dclprms As String, _
    Optional tmpln As String = "tmpl_Cst_Prms", Optional fromMod As String = "classUtilTmpl")
    Dim arg
    Dim tmpl As String
    Dim prms As String
    dcl = expandDcl(dclprms)
    asn = expandDcl(dclprms, "asn")
    arg = Array(clsn, dcl, asn)
    tmpl = disposeProc("get", fromMod, tmpln)(1)
    mkCstPrmLines = tmplToCode0(tmpl, arg)
End Function

Function tmplToCode(tmpl As String, ParamArray prms())
    args = prms
    tmplToCode = tmplToCode0(tmpl, args)
End Function

Function tmplToCode0(tmpl As String, args)
    Dim ret, sLine
    Dim n0, i, j
    n0 = LBound(args)
    ret = Split(tmpl, vbCrLf)
    For i = LBound(ret) To UBound(ret)
        sLine = LTrim(ret(i))
        If Len(sLine) > 0 And Left(sLine, 1) = "'" Then sLine = Right(sLine, Len(sLine) - 1)
        For j = 0 To lenAry(args) - 1
            sLine = Replace(sLine, "$" & j, args(n0 + j))
        Next j
        ret(i) = sLine
    Next i
    tmplToCode0 = Join(ret, vbCrLf)
End Function

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
    '  Set m_$0 = New Collection
    '  For Each elm In prm
    '    Dim x As $1
    '    Set x = elm
    '    m_$0.Add x
    '  Next
    'End Function
End Function

Function expandDcl(elms As String, Optional tp = "dcl")
    prms = Split(elms, ",")
    n = UBound(prms)
    ReDim ret(0 To n)
    For i = 0 To n
        tmp = Split(prms(i), ";")
        For j = 0 To UBound(tmp)
            tmp(j) = Trim(tmp(j))
        Next
        Select Case tp
            Case "asn"
                ret(i) = tmp(0)
            Case "dcl"
                If tmp(1) = "" Then
                    ret(i) = tmp(0)
                Else
                    ret(i) = tmp(0) & " As " & tmp(1)
                End If
            Case Else
        End Select
    Next i
    expandDcl = Join(ret, ",")
End Function

Function dclToAry(dclPrm As String)
    tmp = Split(dclPrm, ",")
    ReDim ret(LBound(tmp) To UBound(tmp))
    For i = LBound(tmp) To UBound(tmp)
        sLine = tmp(i)
        stmp = Split(sLine, ";")
        ret(i) = stmp
    Next i
    dclToAry = ret
End Function

Function mkInitByDclPrms(dclprms As String)
    dcls = expandDcl(dclprms)
    asns = expandDcl(dclprms, "asn")
End Function

Function mkasn(vr)
    ret = "m_" & vr & " = " & vr & "_"
    mkasn = ret
End Function

Function mkdcl(vr, tp, Optional scp = "Private")
    If scp <> "" Then scp = scp & " "
    ret = scp & "m_" & vr & " As " & tp
    mkdcl = ret
End Function

Function initDcl(dclprms As String)
    ret = ""
    tmp = dclToAry(dclprms)
    For Each elm In tmp
        ret = ret & vbCrLf & mkdcl(elm(0), elm(1))
    Next elm
    ReDim tmp1(LBound(tmp) To UBound(tmp))
    For i = LBound(tmp) To UBound(tmp)
        tmp1(i) = tmp(i)(0) & "_ As " & tmp(i)(1)
    Next
    ret = ret & vbCrLf & vbCrLf & "Sub init(" & Join(tmp1, ",") & ")"
    ReDim tmp2(LBound(tmp) To UBound(tmp))
    For i = LBound(tmp) To UBound(tmp)
        tmp2(i) = mkasn(tmp(i)(0))
        If lenAry(tmp(i)) = 3 Then
            If tmp(i)(2) = "o" Then tmp2(i) = "Set " & tmp2(i)
        End If
        tmp2(i) = space(4) & tmp2(i)
    Next i
    ret = ret & vbCrLf & Join(tmp2, vbCrLf) & vbCrLf & "End Sub"
    initDcl = ret
End Function
