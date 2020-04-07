Attribute VB_Name = "classUtil"
Option Base 0

Function disposeProc(tp, modn, procName, Optional knd = 0, Optional sCode = "")
    'tp "get","del","replace"
    Dim cmp
    Dim lineStart, lineDef, lineContent, lineEnd, defcnt, linecnt, i
    Dim xdef, xcnt, xend
    Dim insertCode
    disposeProc = ""
    tp = LCase(tp)
    If tp = "del" Then sCode = ""
    If modn = "" Then modn = Application.VBE.SelectedVBComponent.name
    Set cmp = Application.VBE.ActiveVBProject.VBComponents(modn)
    With cmp.CodeModule
        linecnt = .ProcCountLines(procName, knd)
        lineStart = .ProcStartLine(procName, knd)
        lineDef = .ProcBodyLine(procName, knd)
        lineEnd = lineStart + linecnt - 1
        'determine lineEnd
        For i = 1 To linecnt - 1
            str0 = Trim(.Lines(lineEnd, 1))
            If str0 = "End Function" Or str0 = "End Sub" Or str0 = "End Property" Then
                Exit For
            Else
                lineEnd = lineEnd - 1
            End If
        Next i
        defcnt = 0
        'determine lineContent
        Do While lineDef + defcnt < lineEnd
            strLine = Trim(.Lines(lineDef + defcnt, 1))
            defcnt = defcnt + 1
            If Not strLine Like "* _" Then
                Exit Do
            End If
        Loop
        lineContent = lineDef + defcnt
        'each type dispose start
        If tp = "del" Then
            On Error Resume Next
            Call .DeleteLines(lineStart, lineEnd - linecnt)
            Exit Function
        End If
        'determine xdef,xcnt,xend
        xdef = .Lines(lineDef, lineContent - lineDef)
        xcnt = ""
        If lineEnd > lineContent Then
            xcnt = .Lines(lineContent, lineEnd - lineContent)
        End If
        xend = .Lines(lineEnd, 1)
        If tp = "get" Then
            disposeProc = Array(xdef, xcnt, xend)
            Exit Function
        End If
        On Error Resume Next
        If tp = "replace" Then
            If sCode = "" Then
                insertCode = xdef & vbCrLf & xend
            Else
                insertCode = xdef & vbCrLf & sCode & vbCrLf & xend
            End If
            Call .DeleteLines(lineDef, lineEnd - lineDef + 1)
            Call .InsertLines(lineDef, insertCode)
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

Sub delProcs(Optional modn As String = "", Optional bolPrP = False, Optional bolFnc = False)
    'delete procedures in module "modn"
    If modn = "" Then modn = Application.VBE.SelectedVBComponent.name
    fncs = getModProcDics(modn)
    If bolFnc Then
        For Each fnc In fncs(0).keys
            Call disposeProc("del", modn, fnc, 0)
        Next fnc
    End If
    If bolPrP Then
        For Each prp In fncs(1).keys
            For Each knd In fncs(1)(prp)
                Call disposeProc("del", modn, fnc, knd)
            Next knd
        Next prp
    End If
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
    Dim linecnt As Long
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
    getModProcDics = Array(fncDic, prpDic)
    Set cmp = Nothing
End Function

Private Function tryToGetProcLineNum(cmp, procName, Optional knd = 0)
    On Error Resume Next
    ret = 0
    ret = cmp.CodeModule.ProcCountLines(procName, knd)
    tryToGetProcLineNum = ret
End Function

Function defaultInterfaceName(clsn As String)
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

Sub mkSubClass(sclsn As String, ifcn As String, impln As String)
    If ifcn = "" Then ifcn = defaultInterfaceName(CStr(impln))
    Set cmps = Application.VBE.SelectedVBComponent
    Set tCmp = mkComponent(CStr(sclsn), "cls")
    Call cpCode(impln, sclsn, "all")
    With tCmp.CodeModule
        sLine = "implements " & ifcn
        Call .InsertLines(1, sLine)
    End With
End Sub

Sub mkSubClasses(sclsns, ifcn As String, impln As String)
    For Each sclsn In sclsns
        Call mkSubClass(CStr(sclsn), ifcn, impln)
    Next sclsn
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
    tmplToCode0 = Join(tmp, vbCrLf)
End Function

Function testcode3(src As String, prms)
End Function
