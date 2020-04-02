Attribute VB_Name = "classUtilCmp"
Function getComponent(modn As String)
    'get module of which name of "modn"
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
    'get module of which name of "modn" if exist,or create module of type "tp"
    'tp mod,cls,frm
    Set cmps = Application.VBE.ActiveVBProject.VBComponents
    For Each cmp In cmps
        '  Debug.Print cmp.name
        If LCase(cmp.name) = LCase(modn) Then
            Debug.Print "Already Exists Component " & modn
            Set mkComponent = cmp
            Exit Function
        End If
    Next cmp
    Set mkComponent = cmps.Add(typeNum(tp))
    mkComponent.name = modn
End Function

Sub delComponent(modn As String)
    'delete module component
    Set cmps = Application.VBE.ActiveVBProject.VBComponents
    For Each cmp In cmps
        '  Debug.Print cmp.name
        If cmp.name = modn Then
            cmps.Remove cmp
            'MsgBox "Delete Component " & modn
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
