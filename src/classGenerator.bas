Attribute VB_Name = "classGenerator"
Sub genrateConstructor()
    clsns = Array("Seq", "Choice", "Rep0or1", "Rep0orMore", "Rep1orMore", "T", "F")
    Call mkCst("G", clsns)
End Sub
