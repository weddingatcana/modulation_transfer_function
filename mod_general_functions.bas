Attribute VB_Name = "mod_general_functions"
Option Explicit

Function create_device_independent_bitmap(ByVal image_input_path$) As LongPtr

    #If Win64 Then
    
        Dim inputGDI As t_gdiplus, _
            tokenGDI As LongPtr, _
            bmpGDI As LongPtr, _
            hBmpGDI As LongPtr

    #Else
    
        Dim inputGDI As t_gdiplus, _
            tokenGDI&, _
            bmpGDI&, _
            hBmpGDI&
    
    #End If
    
            inputGDI.gdiplus_version = 1
            
            If GdiplusStartup(tokenGDI, inputGDI) = 0 Then
                If GdipCreateBitmapFromFile(StrPtr(image_input_path), bmpGDI) = 0 Then
                    GdipCreateHBITMAPFromBitmap bmpGDI, hBmpGDI, 0
                    GdipDisposeImage bmpGDI
                End If
            End If
            
            GdiplusShutdown tokenGDI
            create_device_independent_bitmap = hBmpGDI
    
End Function

Function choose_image_file$()

    Dim FDO As FileDialog, _
        SelectionChosen&
    
        Set FDO = Application.FileDialog(msoFileDialogFilePicker)
        SelectionChosen = -1
        
        With FDO
            .InitialFileName = "C:\"
            .Title = "Choose Image File"
            .AllowMultiSelect = False
            .Filters.Clear
            .Filters.Add "Allowed Image Extensions", "*.bmp; *.gif; *.jpg; *.jpeg; *.png; *.tiff; *.tif; *.dib; *.wmf; *.emf"
            
            If .Show = SelectionChosen Then
                choose_image_file = .SelectedItems(1)
            Else
            End If
            
        End With
    
        Set FDO = Nothing

End Function

Function choose_image_file_extension$(ByVal SelectedItem$)

    Dim strExtension$, _
        strDelimiter$, _
        DelimiterPosition&, _
        MaxExtensionLength&

        MaxExtensionLength = 5
        strDelimiter = "."
        strExtension = SelectedItem
        
        DelimiterPosition = InStr(Len(strExtension) - MaxExtensionLength, strExtension, strDelimiter)
        
        strExtension = Right(strExtension, Len(strExtension) - DelimiterPosition + 1)
        choose_image_file_extension = strExtension

End Function

Function save_image(ByRef image_description As t_description, _
                    ByRef guid As t_guid, _
                    ByRef image_interface As IPicture, _
           Optional ByVal strOutputImage$ = "", _
           Optional ByVal strExtension$ = "")

    Dim check&, _
        completion_status&
        
        completion_status = 1
        check = mod_api.OleCreatePictureIndirect(image_description, guid, completion_status, image_interface)
        
        If strOutputImage <> "" And strExtension <> "" Then
            stdole.SavePicture image_interface, strOutputImage & strExtension
        End If
        
        Set image_interface = Nothing
    
End Function

Function image_input_dimensions(ByVal image_input_path$) As t_dimensions

    Dim wia As Object
    
        Set wia = CreateObject("WIA.ImageFile")
    
        If wia Is Nothing Then
            Exit Function
        End If
    
        wia.LoadFile image_input_path
        image_input_dimensions.width = wia.width
        image_input_dimensions.height = wia.height
        Set wia = Nothing
        
End Function

Public Function write_csv(ByRef A&(), _
                          ByVal csvFilename$, _
                 Optional ByVal csvDirectory$ = "C:\Users\qp\Desktop\") As Boolean

    Dim FSO As Object, _
        txtFile As Object, _
        rawMaxRowA&, _
        rawMaxColA&, _
        concatString$, _
        delimiter$, _
        i&, j&
        
        Set FSO = CreateObject("Scripting.FileSystemObject")
        Set txtFile = FSO.CreateTextFile(csvDirectory & csvFilename)
        delimiter = ","
        
        rawMaxRowA = UBound(A, 1)
        rawMaxColA = UBound(A, 2)
       
        For i = 1 To rawMaxRowA
            For j = 1 To rawMaxColA
            
                concatString = concatString & A(i, j) & delimiter
                
            Next j
            
            txtFile.write concatString & vbCrLf
            concatString = ""
            
        Next i
        
        txtFile.Close
        write_csv = True
        Set FSO = Nothing
        Set txtFile = Nothing

End Function

Public Function write_csv_double(ByRef A#(), _
                                 ByVal csvFilename$, _
                        Optional ByVal csvDirectory$ = "C:\Users\qp\Desktop\mtf\") As Boolean

    Dim FSO As Object, _
        txtFile As Object, _
        rawMaxRowA&, _
        rawMaxColA&, _
        concatString$, _
        delimiter$, _
        i&, j&
        
        Set FSO = CreateObject("Scripting.FileSystemObject")
        Set txtFile = FSO.CreateTextFile(csvDirectory & csvFilename)
        delimiter = ","
        
        rawMaxRowA = UBound(A, 1)
        rawMaxColA = UBound(A, 2)
       
        For i = 1 To rawMaxRowA
            For j = 1 To rawMaxColA
            
                concatString = concatString & A(i, j) & delimiter
                
            Next j
            
            txtFile.write concatString & vbCrLf
            concatString = ""
            
        Next i
        
        txtFile.Close
        write_csv_double = True
        Set FSO = Nothing
        Set txtFile = Nothing

End Function
