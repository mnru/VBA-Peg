Attribute VB_Name = "classUtil"
Option Base 0

Function disposeProc(tp, modn, procName, Optional knd = 0, Optional sCode = "")
    'tp "get","del","replace"
    Dim cmp
    Dim lineStart, lineDef, lineContent, lineEnd, defcnt, linecnt, i
    Dim xdef, xcnt, xend
    disposeProc = ""
    tp = LCase(tp)
    If tp = "del" Then sCode = ""
    Set cmp = Application.VBE.ActiveVBProject.VBComponents(modn)
    With cmp.CodeModule
        linecnt = .ProcCountLines(procName, knd)
        lineStart = .ProcStartLine(procName, knd)
        lineDef = .ProcBodyLine(procName, knd)
        lineEnd = lineStart + linecnt - 1
        For i = 1 To linecnt - 1
            str0 = Trim(.Lines(lineEnd, 1))
            If str0 = "End Function" Or str0 = "End Sub" Or str0 = "End Property" Then
                Exit For
            Else
                lineEnd = lineEnd - 1
            End If
        Next i
        defcnt = 0
        Do While lineDef + defcnt < lineEnd
            strLine = Trim(.Lines(lineDef + defcnt, 1))
            defcnt = defcnt + 1
            If Not strLine Like "* _" Then
                Exit Do
            End If
        Loop
        lineContent = lineDef + defcnt
        If tp = "get" Then
            xdef = .Lines(lineDef, lineContent - lineDef)
            xcnt = ""
            If lineEnd > lineContent Then
                xcnt = .Lines(lineContent, lineEnd - lineContent)
            End If
            xend = .Lines(lineEnd, 1)
            disposeProc = Array(xdef, xcnt, xend)
            Exit Function
        End If
        On Error Resume Next
        If tp = "del" Or tp = "replace" Then
            Call .DeleteLines(lineContent, lineEnd - lineContent)
            If tp = "replace" Then
                Call .InsertLines(lineContent, sCode)
            End If
        End If
        On Error GoTo 0
    End With
    Set cmp = Nothing
End Function

Function getNoOptionLine(modn)
    Dim ret, i
    ret = 0
    Set cmp = ActiveWorkbook.VBProject.VBComponents(modn)
    With cmp.CodeModule
        dcln = .CountOfDeclarationLines
        For i = 1 To decln
            sLine = Trim(.Lines(i, 1))
            If sLine Like "Option " Then
                ret = ret + 1
            Else
                Exit For
            End If
        Next i
    End With
    getNoOptionLine = ret
End Function

Function getComponent(modn As String)
    'get  module  of which name of "modn"
    Set cmps = Application.VBE.ActiveVBProject.VBComponents
    For Each cmp In cmps
        If LCase(cmp.name) = LCase(modn) Then
            Set getComponent = cmp
            Exit Function
        End If
    Next cmp
    Debug.Print "doesn't exists module " & modn
End Function

Function mkComponent(modn As String, tp As String)
    'get  module  of which name of "modn"  if exist,or create module of type "tp"
    'tp mod,cls,frm
    Set cmps = Application.VBE.ActiveVBProject.VBComponents
    For Each cmp In cmps
        '    Debug.Print cmp.name
        If LCase(cmp.name) = LCase(modn) Then
            Debug.Print "Already Exists Component " & modn
            Set mkComponent = cmp
            Exit Function
        End If
    Next cmp
    Set mkComponent = cmps.Add(typeNum(tp))
    mkComponent.name = modn
End Function

Sub delProcs(Optional modn As String = "", Optional bolPrP = False, Optional bolFnc = False)
    'delete procedures in module "modn"
    If modn = "" Then modn = Application.VBE.SelectedVBComponent.name
    Set cmp = Application.VBE.ActiveVBProject.VBComponents(modn)
    fncs = getModProcDics(modn)
    With cmp.CodeModule
        If bolFnc Then
            For Each fnc In fncs(0).keys
                Call .DeleteLines(.ProcStartLine(fnc, 0), .ProcCountLines(fnc, 0))
            Next fnc
        End If
        If bolPrP Then
            For Each prp In fncs(1).keys
                For Each knd In fncs(1)(prp)
                    Call .DeleteLines(.ProcStartLine(prp, knd), .ProcCountLines(prp, knd))
                Next knd
            Next prp
        End If
    End With
End Sub

Sub delComponent(modn As String)
    'delete module component
    Set cmps = Application.VBE.ActiveVBProject.VBComponents
    For Each cmp In cmps
        '   Debug.Print cmp.name
        If cmp.name = modn Then
            cmps.Remove cmp
            'MsgBox  "Delete Component " & modn
            Debug.Print "Delete Component " & modn
            bol = True
            Exit Sub
        End If
    Next cmp
    If Not bol Then
        Debug.Print "Doesn't Exists Component " & modn
    End If
End Sub

Sub delComponentExcept(modns)
    Set cmps = Application.VBE.ActiveVBProject.VBComponents
    For Each cmp In cmps
        modn = cmp.name
        Debug.Print modn
        bol = True
        If cmp.Type <> 1 And cmp.Type <> 2 And cmp.Type <> 3 Then bol = False
        For Each elm In modns
            If elm = modn Then bol = False
        Next elm
        If bol Then
            cmps.Remove cmp
            Debug.Print "Delete Component " & modn
        End If
    Next cmp
End Sub

Sub printComponents()
    Set cmps = Application.VBE.ActiveVBProject.VBComponents
    For Each cmp In cmps
        Debug.Print cmp.name
    Next cmp
End Sub

Function typeNum(tp As String)
    Dim ret
    Select Case tp
        Case "std"
            ret = 1
        Case "cls"
            ret = 2
        Case "frm"
            ret = 3
        Case Else
    End Select
    typeNum = ret
End Function

Sub cpCode(smodn, tmodn, Optional part = "all")
    Dim sCmp, tCmp
    Set sCmp = Application.VBE.ActiveVBProject.VBComponents(smodn)
    With sCmp.CodeModule
        Select Case LCase(part)
            Case "all"
                sCode = .Lines(1, .CountOfLines)
            Case "dcl"
                sCode = .Lines(1, .CountOfDeclarationLines)
            Case "prc"
                sCode = .Lines(.CountOfDeclarationLines + 1, .CountOfLines)
            Case Else
        End Select
    End With
    Set tCmp = Application.VBE.ActiveVBProject.VBComponents(tmodn)
    If sCode <> "" Then
        tCmp.CodeModule.AddFromString sCode
    End If
    Set sCmp = Nothing
    Set tCmp = Nothing
End Sub

Function getModProcDics(modn As String)
    Dim cmp
    Dim procName
    Dim procLineNum As Long
    Dim linecnt  As Long
    Dim fncDic
    Dim prpDic
    Set fncDic = CreateObject("Scripting.Dictionary")
    Set prpDic = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Set cmp = Application.VBE.ActiveVBProject.VBComponents(modn)
    With cmp.CodeModule
        If .CountOfLines > 0 Then
            procName = ""
            For linecnt = .CountOfDeclarationLines + 1 To .CountOfLines
                If procName <> .ProcOfLine(linecnt, 0) Then
                    procName = .ProcOfLine(linecnt, 0)
                    procLineNum = tryToGetProcLineNum(cmp, procName, 0)
                    If procLineNum <> 0 Then
                        Call fncDic.Add(procName, 0)
                    Else
                        If Not prpDic.Exists(procName) Then
                            Call prpDic.Add(procName, New Collection)
                            For knd = 1 To 3
                                procLineNum = tryToGetProcLineNum(cmp, procName, knd)
                                If procLineNum <> 0 Then
                                    prpDic(procName).Add knd
                                End If
                            Next knd
                        End If
                    End If
                End If
            Next linecnt
        End If
    End With
    getModProcDics = Array(fncDic, prpDic)
    Set cmp = Nothing
End Function

Private Function tryToGetProcLineNum(cmp, procName, Optional knd = 0)
    On Error Resume Next
    ret = 0
    ret = cmp.CodeModule.ProcCountLines(procName, knd)
    tryToGetProcLineNum = ret
End Function

Private Function lenAry(ary)
    lenAry = UBound(ary) - LBound(ary) + 1
End Function

Private Function conArys(ParamArray argArys())
    Dim num As Long, i As Long
    Dim arys, ret, elm, ary
    arys = argArys
    num = 0
    For Each ary In arys
        If IsArray(ary) Then
            num = num + lenAry(ary)
        Else
            num = num + 1
        End If
    Next ary
    ReDim ret(0 To num - 1)
    i = 0
    For Each ary In arys
        If IsArray(ary) Then
            For Each elm In ary
                ret(i) = elm
                i = i + 1
            Next elm
        Else
            ret(i) = ary
            i = i + 1
        End If
    Next ary
    conArys = ret
End Function

Private Function defaultInterfaceName(clsn As String)
    n = InStr(clsn, "_")
    If n > 0 Then
        ifc = Left(clsn, n - 1)
    Else
        ifc = "I" & clsn
    End If
    defaultInterfaceName = ifc
End Function

Function isProc(sLine, pos, n)
    Dim n0, c1, c2
    Dim ret
    n0 = Len(sLine)
    ret = True
    pos2 = pos + n - 1
    If pos <= 0 Or pos2 > n0 Then ret = False
    If pos > 1 Then
        c1 = Mid(sLine, pos - 1, 1)
        If c1 <> " " And c1 <> "(" Then
            ret = False
        End If
    End If
    If pos2 < n0 Then
        c2 = Mid(sLine, pos2 + 1, 1)
        If c2 <> " " And c2 <> "," And c2 <> "(" And c2 <> ")" Then
            ret = False
        End If
    End If
    isProc = ret
End Function

Function isPrefix(sLine, pos, n)
    Dim n0, c1, c2
    Dim ret
    n0 = Len(sLine)
    ret = True
    pos2 = pos + n - 1
    If pos <= 0 Or pos2 > n0 Then ret = False
    If pos > 1 Then
        c1 = Mid(sLine, pos - 1, 1)
        If c1 <> " " And c1 <> "(" Then
            ret = False
        End If
    End If
    If pos2 < n0 Then
        c2 = Mid(sLine, pos2 + 1, 1)
        If c2 <> "_" Then
            ret = False
        End If
    End If
    isPrefix = ret
End Function

Sub addPrefix(ifsn As String, clsn As String)
    Dim tmp
    fncs = getModProcDics(ifsn)
    Dim sLine
    Dim cmp 'As VBComponent
    Set cmp = mkComponent(clsn, "cls")
    With cmp.CodeModule
        For i = .CountOfDeclarationLines To .CountOfLines
            tmp = .Lines(i, 1)
            For j = 0 To 1
                For Each s In fncs(j).keys
                    n = Len(s)
                    pos = Len(tmp)
                    Do While pos > 0
                        pos = InStrRev(tmp, s, pos)
                        If pos > 0 Then
                            If isProc(tmp, pos, n) Then
                                If pos = 1 Then
                                    tmp = ifsn & "_" & tmp
                                Else
                                    tmp = Left(tmp, pos - 1) & ifsn & "_" & Right(tmp, Len(tmp) - pos + 1)
                                End If
                            End If
                        End If
                        pos = pos - 1
                    Loop
                Next s
            Next j
            If .Lines(i, 1) <> tmp Then Call .ReplaceLine(i, tmp)
        Next i
    End With
End Sub

Sub delPrefix(ifsn As String, clsn As String)
    Dim tmp, t1, t2
    fncs = getModProcDics(ifsn)
    Dim sLine
    Dim cmp 'As VBComponent
    Set cmp = mkComponent(clsn, "cls")
    With cmp.CodeModule
        For i = .CountOfDeclarationLines To .CountOfLines
            tmp = .Lines(i, 1)
            For j = 0 To 1
                For Each s In fncs(j).keys
                    n = Len(ifsn)
                    pos = Len(tmp)
                    Do While pos > 0
                        pos = InStrRev(tmp, ifsn, pos)
                        If pos > 0 Then
                            If isPrefix(tmp, pos, n) Then
                                t1 = ""
                                If pos > 0 Then t1 = Left(tmp, pos - 1)
                                If Len(tmp) - (pos + n) > 0 Then t2 = Right(tmp, Len(tmp) - (pos + n))
                                tmp = t1 & t2
                            End If
                        End If
                        pos = pos - 1
                    Loop
                Next s
            Next j
            If .Lines(i, 1) <> tmp Then Call .ReplaceLine(i, tmp)
        Next i
    End With
End Sub

Private Function mkPrpStatement(x, tp, symbol)
    symbol = LCase(symbol)
    Dim ibol
    Dim st
    Dim sts(1 To 3)
    Dim sc, gls, o, v
    sc = IIf(symbol Like "*_", "Public", "Private")
    tpp = IIf(tp = "", "", " As " & tp)
    ibol = Left(symbol, 1) = "i"
    If ibol Then
        If symbol = "i" Or symbol = "i_" Then
            st = sc & " " & x & tpp
            mkPrpStatement = Array(ibol, st)
            Exit Function
        Else
            symbol = Right(symbol, Len(symbol) - 1)
        End If
    End If
    sc = IIf(symbol Like "*_", "Public", "Private")
    gls = UCase(Left(symbol, 1)) & "et"
    o = IIf(gls = "Set" Or InStr(symbol, "o") > 0, "Set ", "")
    v = IIf(gls = "Let" Or (gls = "Set" And InStr(symbol, "v") > 0), "ByVal ", "")
    tmp = sc & " Property " & gls & " " & x
    If gls = "Get" Then
        tmp = tmp & "()" & tpp
    Else
        tmp = tmp & "(" & v & x & "_" & tpp & ")"
    End If
    sts(1) = tmp
    tmp = o
    If gls = "Get" Then
        tmp = tmp & x & " = m_" & x
    Else
        tmp = tmp & "m_" & x & " = " & x & "_"
    End If
    sts(2) = tmp
    sts(3) = "End Property"
    If ibol Then
        mkPrpStatement = Array(ibol, sts(1) & vbCrLf & sts(3))
    Else
        mkPrpStatement = Array(ibol, Join(sts, vbCrLf))
    End If
End Function

Sub mkPrp(Optional ifcn As String = "", Optional impln As String = "", Optional mkI = False, Optional mkNotI = True)
    Dim cmp
    Dim sLine As String
    Dim i
    Dim aryLine, aryDcl, arySymbol
    If impln = "" Then
        Set cmp = Application.VBE.SelectedVBComponent
    Else
        Set cmp = Application.VBE.ActiveVBProject.VBComponents(impln)
    End If
    Debug.Print cmp.name
    If ifcn = "" Then ifcn = defaultInterfaceName(cmp.name)
    With cmp.CodeModule
        For i = .CountOfDeclarationLines To 1 Step -1
            sLine = .Lines(i, 1)
            aryLine = Split(sLine, "'")
            If lenAry(aryLine) <> 2 Then GoTo endfor
            aryDcl = partDcl(aryLine(0))
            If aryDcl(0) Then
                arySymbol = partSymbol(aryLine(1))
                Set symbolI = arySymbol(0)
                Set symbolNotI = arySymbol(1)
                If mkI Then
                    Set cmp0 = mkComponent(ifcn, "cls")
                    For j = symbolI.Count To 1 Step -1
                        s = symbolI(j)
                        s1 = aryDcl(1)
                        s2 = aryDcl(2)
                        sts = mkPrpStatement(s1, s2, s)
                        Call cmp0.CodeModule.AddFromString(vbCrLf & sts(1))
                    Next j
                End If
                If mkNotI Then
                    For j = symbolNotI.Count To 1 Step -1
                        s = symbolNotI(j)
                        s1 = aryDcl(1)
                        s2 = aryDcl(2)
                        sts = mkPrpStatement(s1, s2, s)
                        Call .AddFromString(vbCrLf & sts(1))
                    Next j
                End If
            End If
endfor:
        Next i
    End With
End Sub

Private Function partDcl(str)
    bol = True
    s2 = ""
    s4 = ""
    ary1 = Split(Trim(str))
    If lenAry(ary1) <> 2 And lenAry(ary1) <> 4 Then bol = False
    If lenAry(ary1) = 4 And Trim(ary1(2)) <> "As" Then bol = False
    If bol Then
        s1 = Trim(ary1(0))
        s2 = Trim(ary1(1))
        If s1 <> "Dim" And s1 = "Private" <> s1 = "Public" Then
            bol = False
        Else
            If lenAry(ary1) = 2 Then
                s4 = ""
            Else
                s4 = Trim(ary1(3))
            End If
            If Left(s2, 2) = "m_" Then s2 = Right(s2, Len(s2) - 2)
        End If
    End If
    partDcl = Array(bol, s2, s4)
End Function

Private Function partSymbol(str)
    Dim clc1, clc2, ary, elm
    Set clc1 = New Collection
    Set clc2 = New Collection
    ary = Split(str, ",")
    For Each elm In ary
        elm = LCase(Trim(elm))
        If elm <> "" Then
            If Left(elm, 1) = "i" Then
                clc1.Add elm
            Else
                clc2.Add elm
            End If
        End If
    Next elm
    partSymbol = Array(clc1, clc2)
End Function

Sub mkCst(tmpln As String, toMod As String, fromMod As String, clsns)
    Dim arg
    Dim tmpl As String
    Dim sLines As String
    arg = clsns
    Set cmp = mkComponent(toMod, "std")
    With cmp.CodeModule
        tmpl = disposeProc("get", fromMod, "Cst_" & tmpln)(1)
        For i = UBound(arg) To LBound(arg) Step -1
            sLines = writeToTmpl(tmpl, CStr(arg(i)))
            .AddFromString (vbCrLf & sLines)
        Next i
    End With
End Sub

Sub mkCstPrm(toMod As String, clsn As String, dclPrms As String, _
    Optional tmpln As String = "Cst_Parser_Prm", Optional fromMod As String = "classGenerator")
    Dim sLines As String
    Dim cmp
    sLines = mkCstPrmLines(clsn, dclPrms)
    Set cmp = mkComponent(toMod, "std")
    With cmp.CodeModule
        .AddFromString (vbCrLf & sLines)
    End With
End Sub

Function delTypeInDcl(elms As String)
    ret = Split(elms, ",")
    For i = LBound(ret) To UBound(ret)
        tmp = Trim(ret(i))
        pos = InStr(tmp, " ")
        If pos > 1 Then tmp = Left(tmp, pos - 1)
        ret(i) = tmp
    Next i
    delTypeInDcl = Join(ret, ",")
End Function

Function mkCstPrmLines(clsn As String, dclPrms As String, _
    Optional tmpln As String = "Cst_Parser_Prm", Optional fromMod As String = "classGenerator")
    Dim arg
    Dim tmpl As String
    Dim prms As String
    prms = delTypeInDcl(dclPrms)
    arg = Array(clsn, dclPrms, prms)
    tmpl = disposeProc("get", fromMod, tmpln)(1)
    mkCstPrmLines = writePrmsToTmpl(tmpl, arg)
End Function

Function writeToTmpl(src As String, nm As String)
    Dim tmp
    Dim i
    tmp = Split(src, vbCrLf)
    For i = LBound(tmp) To UBound(tmp)
        sLine = Trim(tmp(i))
        If Len(sLine) > 0 And Left(sLine, 1) = "'" Then sLine = Right(sLine, Len(sLine) - 1)
        sLine = Replace(sLine, "$", nm)
        tmp(i) = sLine
    Next i
    writeToTmpl = Join(tmp, vbCrLf)
End Function

Function writePrmsToTmpl(src As String, prms)
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

Sub override(fnc, knd, toMod, fromMod)
End Sub

Sub mkInterFace(ifcn As String, impln As String, ParamArray ArgClsns())
    Dim i
    Dim clsns
    Dim sCmp, tCmp
    Dim fncs, fnckeys, fnc
    clsns = ArgClsns
    If impln = "" Then
        Set cmp = Application.VBE.SelectedVBComponent
    Else
        Set cmp = Application.VBE.ActiveVBProject.VBComponents(impln)
    End If
    If ifcn = "" Then ifcn = defaultInterfaceName(CStr(impln))
    Set sCmp = Application.VBE.ActiveVBProject.VBComponents(impln)
    Set tCmp = mkComponent(ifcn, "cls")
    fncs = getModProcDics(impln)
    With tCmp.CodeModule
        fnckeys = fncs(0).keys
        For i = UBound(fnckeys) To LBound(fnckeys) Step -1
            fnc = fnckeys(i)
            codes = disposeProc("get", impln, fnc)
            code = codes(0) & vbCrLf & codes(2)
            Call .AddFromString(vbCrLf & code)
            'Call .InsertLines(.CountOfLines, code)
        Next i
    End With
    Call mkPrp(ifcn, impln, True, False)
End Sub

Sub mkSubClass(ifcn As String, impln As String, sclsns)
    If ifcn = "" Then ifcn = defaultInterfaceName(CStr(impln))
    Set cmps = Application.VBE.SelectedVBComponent
    For Each sclsn In sclsns
        Set tCmp = mkComponent(CStr(sclsn), "cls")
        Call cpCode(impln, sclsn, "all")
        With tCmp.CodeModule
            sLine = "implements " & ifcn
            Call .InsertLines(1, sLine)
        End With
    Next sclsn
End Sub
