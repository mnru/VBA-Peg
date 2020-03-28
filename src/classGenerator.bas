Attribute VB_Name = "classGenerator"
Sub genrateConstructor()
    clsns = Array("Seq", "Choice", "Rep0or1", "Rep0orMore", "Rep1orMore", "T", "F")
    Call mkInterFace("iParser", "iParser_impl")
      Call mkCst("G", clsns)
End Sub
'"AnyChar","ASTNode","Char","CharRange","CharSet","Choice","classGenerator","classUtil","Delay","F","G","iParser","iParser_Impl","N1_","Opt","Parser_Impl","ParseState","RegEx","Rule","Seq","T","Token"
