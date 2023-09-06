Sub WIA_Crop_Image()

    'You need to active references: 
    'WI

    Dim WIA_ImgFile As WIA.ImageFile
    Dim WIA_ImgP As WIA.ImageProcess
    Dim Myfile As String
    
    Set WIA_ImgFile = New WIA.ImageFile
    Set WIA_Img = New WIA.ImageProcess
    
    'Add the path of your image
    Myfile = "PATH"

    'Apply filter
    WIA_ImgP.Filters.Add WIA_ImgP.FilterInfos("Crop").FilterID

    'Change de dimensions "crop"
    WIA_ImgP.Filters(1).Properties("Left") = 50
    WIA_ImgP.Filters(1).Properties("Right") = 50
    WIA_ImgP.Filters(1).Properties("Top") = 50
    WIA_ImgP.Filters(1).Properties("Bottom") = 50

    'Load the file
    WIA_ImgFile.LoadFile Myfile

    'Process the file using apply method
    Set WIA_ImgFile = WIA_ImgP.Apply(WIA_ImgFile)
    
    On Error Resume Next
    VBA.Kill "PATH of new image"
    WIA_ImgFile.SaveFile "PATH of new image"
    On Error GoTo 0

    Set WIA_ImgP = Nothing
    Set WIA_ImgFile = Nothing

End Sub


