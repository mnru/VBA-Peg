Attribute VB_Name = "classUtilPrp"
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
