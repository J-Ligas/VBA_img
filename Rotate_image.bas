Sub Flip()

    Dim WIA_imgFile As WIA.ImageFile
    Dim WIA_imgP As WIA.ImageProcess
    Dim MyFile As String

    'Seteando objetos
    Set WIA_imgFile = New WIA.ImageFile
    Set WIA_imgP = New WIA.ImageProcess
    Ruta_img = "PATH"
    Ruta_img_v = "PATH of new image"

    'Añadir filtro
    WIA_imgP.Filters.Add WIA_imgP.FilterInfos("RotateFlip").FilterID
    WIA_imgP.Filters(1).Properties("FlipHorizontal") = True
    
    'Cargar archivo
    WIA_imgFile.LoadFile Ruta_img

    'Procesar imagen usando el método apply
    Set WIA_imgFile = WIA_imgP.Apply(WIA_imgFile)

    On Error Resume Next
    VBA.Kill Ruta_img_v
    WIA_imgFile.SaveFile Ruta_img_v
    On Error GoTo 0

    Set WIA_imgP = Nothing
    Set WIA_imgFile = Nothing

End Sub