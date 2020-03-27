Attribute VB_Name = "classGenerator"
Sub genrateConstructor()
    clsns = Array("Seq", "Choice", "Rep0or1", "Rep0orMore", "Rep1orMore", "T", "F")
    Call mkCst("G", clsns)
End Sub

Sub test()
    clsns = Array("classGenerator", "classUtil", "G", "iParser_Impl", "")
    Call delModComponent("N0_")
   Call delModComponentExcept(clsns)
End Sub
'
'"AnyChar","ASTNode","Char","CharRange","CharSet","Choice","classGenerator","classUtil","Delay","F","G","iParser","iParser_Impl","N1_","Opt","Parser_Impl","ParseState","RegEx","Rule","Seq","T","Token"

