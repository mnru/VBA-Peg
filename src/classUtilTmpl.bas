Attribute VB_Name = "classUtilTmpl"
Sub mkCst(tmpln As String, toMod As String, fromMod As String, clsns)
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

Sub mkCstPrm(toMod As String, clsn As String, dclPrms As String, _
    Optional tmpln As String = "tmpl_Cst_Prms", Optional fromMod As String = "classUtilTmpl")
    Dim sLines As String
    Dim cmp
    sLines = mkCstPrmLines(clsn, dclPrms, tmpln, fromMod)
    Set cmp = mkComponent(toMod, "std")
    With cmp.CodeModule
        .AddFromString (vbCrLf & sLines)
    End With
End Sub

Function commentToDcl(elms As String, Optional tp = "declare")
    prms = Split(elms, ",")
    n = UBound(prms)
    ReDim ret(0 To n)
    For i = 0 To n
        tmp = Split(prms(i), ";")
        For j = 0 To UBound(tmp)
            tmp(j) = Trim(tmp(j))
        Next
        Select Case tp
            Case "assign"
                ret(i) = tmp(0)
            Case "declare"
                If tmp(1) = "" Then
                    ret(i) = tmp(0)
                Else
                    ret(i) = tmp(0) & " As " & tmp(1)
                End If
            Case Else
        End Select
    Next i
    commentToDcl = Join(ret, ",")
End Function

Function mkCstPrmLines(clsn As String, dclPrms As String, _
    Optional tmpln As String = "tmpl_Cst_Prms", Optional fromMod As String = "classUtilTmpl")
    Dim arg
    Dim tmpl As String
    Dim prms As String
    dcl = commentToDcl(dclPrms)
    asn = commentToDcl(dclPrms, "assign")
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

Sub overRideByTmpl(fnc As String, knd As Long, tmpln As String, toMod As String, fromMod As String)
    Dim sLines, cmp
    tmpl = disposeProc("get", fromMod, tmpln, 0)(1)
    Set cmp = getComponent(toMod)
    With cmp.CodeModule
        lnum = .ProcStartLine(fnc, knd)
        Call .DeleteLines(lnum, .ProcCountLines(fnc, knd))
        sCode = tmplToCode(tmpl)
        Call .InsertLines(lnum, sCode)
    End With
End Sub

Sub overRide(fnc As String, knd As Long, toMod As String, fromMod As String)
    Dim sLines, cmp
    sLines = disposeProc("get", fromMod, fnc, knd)
    Set cmp = getComponent(toMod)
    With cmp.CodeModule
        lnum = .ProcStartLine(fnc, knd)
        Call .DeleteLines(lnum, .ProcCountLines(fnc, knd))
        Call .InsertLines(lnum, Join(sLines, vbCrLf))
    End With
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
