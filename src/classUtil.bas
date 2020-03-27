Attribute VB_Name = "classUtil"
Option Base 0

Function disposeProc(tp, cmpn, procName, Optional knd = 0, Optional sCode = "")
    'tp "get","del","replace"
    Dim cmp
    Dim lineStart, lineDef, lineContent, lineEnd, defcnt, linecnt, i
    Dim xdef, xcnt, xend
    disposeProc = ""
    tp = LCase(tp)
    If tp = "del" Then sCode = ""
    Set cmp = ActiveWorkbook.VBProject.VBComponents(cmpn)
    With cmp.CodeModule
        linecnt = .ProcCountLines(procName, knd)
        lineStart = .ProcStartLine(procName, knd)
        lineDef = .procBodyLine(procName, knd)
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
            xcnt = .Lines(lineContent, lineEnd - lineContent)
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

Sub mkModComponent(modn As String, tp As String)
    'tp mod,cls,frm
    Set cmps = Application.VBE.ActiveVBProject.VBComponents
    For Each cmp In cmps
        If cmp.name = modn Then
            'MsgBox  "Already Exists Component " & modn
            Debug.Print "Already Exists Component " & modn
            Exit Sub
        End If
    Next cmp
    Set cmp0 = cmps.Add(typeNum(tp))
    cmp0.name = modn
End Sub

Sub delModComponent(modn As String)
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

Sub delModComponentExcept(modns)
    Set cmps = Application.VBE.ActiveVBProject.VBComponents
    For Each cmp In cmps
        modn = cmp.name
        Debug.Print modn
        bol = True
        If cmp.Type <> 1 And cmp.Type <> 2 Then bol = False
        
        For Each elm In modns
            If elm = modn Then bol = False
        Next elm
        If bol Then
            cmps.Remove cmp
            Debug.Print "Delete Component " & modn
        End If
    Next cmp
End Sub

Sub printModComponents()
    Set cmps = Application.VBE.ActiveVBProject.VBComponents
    For Each cmp In cmps
    Debug.Print cmp.name
    Next cmp
End Sub

Function typeNum(tp As String)
    Dim ret
    Select Case tp
        Case "mod"
            ret = 1
        Case "cls"
            ret = 2
        Case "frm"
            ret = 3
        Case Else
    End Select
    typeNum = ret
End Function

Sub mkInterFace(ifcn As String, ParamArray ArgClsns())
    clsns = ArgClsns
    clsn = clsns(0)
    If ifcn = "" Then ifcn = defaultInterfaceName(CStr(clsn))
    Set cmps = ActiveWorkbook.VBProject.VBComponents
    Set sCmp = cmps(clsn)
    Set tCmp = cmps.Add(2)
    tCmp.name = ifcn
    Call cpCode(clsn, ifcn, "all")
    fncs = getModProcs(ifcn)
    For Each fnc In fncs(0).keys
        Call disposeProc("del", ifcn, fnc)
    Next
    For Each prp In fncs(1).keys
        For Each knd In fncs(1)(prp)
            Call disposeProc("del", ifcn, prp, knd)
        Next
    Next
End Sub

Sub mkSubClass(sclsn As String, ifcn As String, ParamArray ArgClsns())
    clsns = ArgClsns
    clsn = clsns(0)
    If ifcn = "" Then ifcn = defaultInterfaceName(CStr(clsn))
    Set cmps = ActiveWorkbook.VBProject.VBComponents
    Set sCmp = cmps(clsn)
    Set tCmp = cmps.Add(2)
    tCmp.name = sclsn
    Call cpCode(clsn, sclsn, "all")
    fncs = getModProcs(ifcn)
End Sub

Sub cpCode(smodn, tmodn, Optional part = "all")
    Set sCmp = ActiveWorkbook.VBProject.VBComponents(smodn)
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
    Set tCmp = ActiveWorkbook.VBProject.VBComponents(tmodn)
    If sCode <> "" Then
        tCmp.CodeModule.AddFromString sCode
    End If
    Set sCmp = Nothing
    Set tCmp = Nothing
End Sub

Function getModProcs(modn As String)
    bn = ActiveWorkbook.name
    Dim procName
    Dim procLineNum As Long
    Dim linecnt  As Long
    Dim fncDic
    Dim prpDic
    Set fncDic = CreateObject("Scripting.Dictionary")
    Set prpDic = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Set cmp = ActiveWorkbook.VBProject.VBComponents(modn)
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
                        If Not prpDic.exists(procName) Then
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
    getModProcs = Array(fncDic, prpDic)
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

Private Function defaultInterfaceName(clsn As String)
    n = InStr(clsn, "_")
    If n > 0 Then
        ifc = Left(clsn, n - 1)
    Else
        ifc = "I" & clsn
    End If
    defaultInterfaceName = ifc
End Function

Function isLexicallyProc(sLine, pos, n)
    Dim N0, c1, c2
    Dim ret
    N0 = Len(sLine)
    ret = True
    If pos + n > N0 Then ret = False
    If n > 1 Then
        c1 = Mid(lStr, n - 1, 1)
        If c1 <> " " And c1 <> "(" Then
            ret = False
        End If
    End If
    pos2 = pos + n - 1
    If pos2 < N0 Then
        c2 = Mid(lStr, pos2, 1)
        If c2 <> " " And c2 <> "," And c2 <> "(" And c2 <> ")" Then
            ret = False
        End If
    End If
    isLexicallyProc = ret
End Function

Function addIfcPreFix(sLine, ifc, pos)
    Dim ret
    If pos = 1 Then
        ret = ifc & "_" & sLine
    Else
        ret = Left(sLine, pos - 1) & ifc & "_" & Right(sLine, Len(sLine) - pos + 1)
    End If
    addIfcPreFix = ret
End Function

Function delIfcPreFix(sLine, ifc)
    Dim ret
    ret = Replace(sLine, ifc & "_", "")
    ret = ifc & "_" & sLine
    delIfcPreFix = ret
End Function

Function mkPrpStatement(x, tp, symbol)
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

Sub mkPrp(Optional ifcn As String = "")
    Dim cmp
    Dim sLine
    Dim i, n
    Set cmp = Application.VBE.SelectedVBComponent
    Debug.Print cmp.name
    If ifcn = "" Then ifcn = defaultInterfaceName(cmp.name)
    With cmp.CodeModule
        For i = .CountOfDeclarationLines To 1 Step -1
            sLine = .Lines(i, 1)
            n = InStr(sLine, "'")
            If n = 0 Then GoTo endfor
            ary1 = Split(Trim(Left(sLine, n - 1)))
            ary2 = Split(Trim(Right(sLine, Len(sLine) - n)), ",")
            If lenAry(ary1) <> 2 And lenAry(ary1) <> 4 Then GoTo endfor
            If lenAry(ary1) = 4 And Trim(ary1(2)) <> "As" Then GoTo endfor
            s1 = Trim(ary1(0))
            s2 = Trim(ary1(1))
            If s1 <> "Dim" And s1 = "Private" <> s1 = "Public" Then GoTo endfor
            If lenAry(ary1) = 2 Then
                s4 = ""
            Else
                s4 = Trim(ary1(3))
            End If
            If Left(s2, 2) = "m_" Then s2 = Right(s2, Len(s2) - 2)
            For j = UBound(ary2) To LBound(ary2) Step -1
                s = Trim(ary2(j))
                If s <> "" Then
                    sts = mkPrpStatement(s2, s4, s)
                    If Not sts(0) Then
                        .AddFromString (vbCrLf & sts(1))
                    Else
                        Set cmp0 = Application.VBE.ActiveVBProject.VBComponents(ifcn)
                        cmp0.CodeModule.AddFromString (vbCrLf & sts(1))
                    End If
                End If
            Next j
endfor:
        Next i
    End With
End Sub

Function mkConstructorStatement(clsn As String, Optional arg = "ParamArray arg()") As String
    Dim ret(1 To 5)
    ret(1) = "Function " & clsn & "(" & arg & " ) As " & clsn
    ret(2) = "  Set " & clsn & " = New " & clsn
    ret(3) = "  prm = arg"
    ret(4) = "  " & clsn & ".init(prm)"
    ret(5) = "End Function"
    mkConstructorStatement = Join(ret, vbCrLf)
End Function

Sub mkCst(toMod As String, clsns() As String)
    arg = clsns
    Set cmp = Application.VBE.ActiveVBProject.VBComponents(toMod)
    With cmp.CodeModule
        For i = UBound(arg) To LBound(arg) Step -1
            str0 = mkConstructorStatement(CStr(arg(i)))
            .AddFromString (vbCrLf & str0)
        Next i
    End With
End Sub
