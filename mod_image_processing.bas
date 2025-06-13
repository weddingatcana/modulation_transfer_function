Attribute VB_Name = "mod_image_processing"
Option Explicit

Const PI = 3.14159265
Const LINEPAIR = 36

Public Type t_center
    x As Long
    y As Long
End Type

Public Type t_max_data
    pixel_position As Long
    magnitude As Long
End Type

Public Type t_circle
    rgb_data() As Byte
    analysis_data() As Double
End Type
    
Public Enum t_write
    write_file = 0
    no_file = 1
End Enum

Function greyscale&(ByVal R&, _
                    ByVal G&, _
                    ByVal B&, _
           Optional ByVal greyscale_calculation As greyType = greyType_AVERAGE)

    Dim k&(), _
        max&, _
        min&, _
        temp&, _
        i&, j&
        
        If greyscale_calculation = greyType_AVERAGE Then
        
            greyscale = (R + G + B) / 3
             
        ElseIf greyscale_calculation = greyType_LUMINANCE Then
        
            greyscale = R * 0.3 + G * 0.59 + B * 0.11
            
        ElseIf greyscale_calculation = greyType_DESATURATION Then
        
            ReDim k(1 To 3)
            k(1) = R: k(2) = G: k(3) = B
           
           ' Sorts array from largest to smallest. Quick & dirty bubblesort.
            For i = LBound(k) To UBound(k)
                For j = i + 1 To UBound(k)
                   If k(i) < k(j) Then
                    
                        temp = k(j)
                        k(j) = k(i)
                        k(i) = temp
                        
                    End If
                Next j
            Next i
            
            max = k(1)
            min = k(3)
            greyscale = (max + min) / 2
            
        ElseIf greyscale_calculation = greyType_RED Then
        
            greyscale = R
            
        ElseIf greyscale_calculation = greyType_GREEN Then
        
            greyscale = G
            
        ElseIf greyscale_calculation = greyType_BLUE Then
        
            greyscale = B
            
        End If

End Function

Function raw_image_to_grey(ByRef rawData() As Byte, _
                  Optional ByVal greyscale_calculation As greyType = greyType_AVERAGE) As Long()

    Dim rawCopy&(), _
        rawMaxRow&, _
        rawMaxCol&, _
        rawMaxColor&, _
        grey&(), _
        R&, G&, B&, _
        i&, j&, k&
        
        rawMaxRow = UBound(rawData, 3)
        rawMaxCol = UBound(rawData, 2)
        rawMaxColor = UBound(rawData, 1)
        
        ReDim rawCopy(1 To rawMaxColor, 1 To rawMaxCol, 1 To rawMaxRow)
        
        For k = 1 To rawMaxColor
            For j = 1 To rawMaxCol
                For i = 1 To rawMaxRow
                
                    rawCopy(k, j, i) = CLng(rawData(k, j, i))
                
                Next i
            Next j
        Next k

        ReDim grey(1 To rawMaxCol, 1 To rawMaxRow)
        
        For j = 1 To rawMaxCol
            For i = 1 To rawMaxRow
            
                R = rawCopy(1, j, i)
                G = rawCopy(2, j, i)
                B = rawCopy(3, j, i)
                'Alpha = rawCopy(4, j, i)
                
                grey(j, i) = greyscale(R, G, B, greyscale_calculation)
                
            Next i
        Next j
        
        raw_image_to_grey = grey

End Function

Public Function dimension_2D_to_3D(ByRef src2D&()) As Byte()

    Const BitsPixel& = &HC '12
    Const hWnd As LongPtr = &H0 '0
    Const base& = 256

    Dim src3D() As Byte, _
        srcMaxRow&, _
        srcMaxCol&, _
        srcMaxColor&, _
        srcMaxColorCorrection&, _
        check&, _
        hDC As LongPtr, _
        i&, j&, k&

        'Obtain handle to desktop window, find if we're on a 24bit/32bit color architecture.
        hDC = GetDC(hWnd)
        srcMaxColor = GetDeviceCaps(hDC, BitsPixel) \ 8
        check = ReleaseDC(hWnd, hDC)
        
        If check = 1 Then
            'released
        ElseIf check = 0 Then
            'not released
        End If
        
        srcMaxRow = UBound(src2D, 2)
        srcMaxCol = UBound(src2D, 1)
        
        'We still want to include the alpha channel in our array, if on a 32bit system. Hence, dimensioning to srcMaxColor. _
        SetBitmapBits will need that extra alpha channel if on a 32bit system.
        ReDim src3D(1 To srcMaxColor, 1 To srcMaxCol, 1 To srcMaxRow)
        
        'But, as following from the comment above, we don't want to copy src2D data into the alpha channel (#4).
        If srcMaxColor > 3 Then
            srcMaxColorCorrection = 3
        Else
            srcMaxColorCorrection = srcMaxColor
        End If
        
        For k = 1 To srcMaxColorCorrection
            For j = 1 To srcMaxCol
                For i = 1 To srcMaxRow
                
                    src3D(k, j, i) = CByte(Abs(src2D(j, i) Mod base))
                                
                Next i
            Next j
        Next k
        
        dimension_2D_to_3D = src3D

End Function

Public Function draw_center(ByRef rgb() As Byte, _
                            ByVal x&, _
                            ByVal y&) As Byte()

    rgb(1, x, y) = 0
    rgb(1, x, y) = 255
    rgb(1, x, y) = 0
    
    draw_center = rgb
    
End Function

Public Function clean_data(ByRef raw_3D#(), _
                  Optional ByVal start_radius& = 10, _
                  Optional ByVal dr& = 10, _
                  Optional ByRef write_data As t_write) As Double()

    Dim bool As Boolean, _
        rawMaxCircle&, _
        rawMaxRow&, _
        rawMaxCol&, _
        new_3D#(), _
        new_2D#(), _
        data_theta#(), _
        data_pixel#(), _
        min_coeffs#(), _
        min_poly#(), _
        min_avg#, _
        max_coeffs#(), _
        max_poly#(), _
        max_avg#, _
        final#(), _
        theta#, _
        radius#, _
        circumference#, _
        pixel&, _
        i&, j&, k&
        
        rawMaxCircle = UBound(raw_3D, 1)
        rawMaxRow = UBound(raw_3D, 2)
        rawMaxCol = UBound(raw_3D, 3)
        
        ReDim final(1 To rawMaxCircle, 1 To 4)
        
        j = 0
        For k = 1 To rawMaxCircle
            For i = 1 To rawMaxRow
            
                theta = raw_3D(k, i, 1)
                pixel = raw_3D(k, i, 2)
                
                If theta <> 0 And pixel <> 0 Then
                    j = j + 1
                End If
            
            Next i
        
            ReDim new_3D(1 To rawMaxCircle, 1 To j, 1 To rawMaxCol)
            
            j = 0
            For i = 1 To rawMaxRow
            
                theta = raw_3D(k, i, 1)
                pixel = raw_3D(k, i, 2)
                
                If theta <> 0 And pixel <> 0 Then
                    j = j + 1
                    new_3D(k, j, 1) = theta
                    new_3D(k, j, 2) = pixel
                End If
            
            Next i
            
            ReDim new_2D(1 To j, 1 To rawMaxCol)
            
            For i = 1 To j
            
                new_2D(i, 1) = new_3D(k, i, 1)
                new_2D(i, 2) = new_3D(k, i, 2)
            
            Next i
            
            data_theta = mod_matrix.matVec(new_2D, 1)
            data_pixel = mod_matrix.matVec(new_2D, 2)
            
            min_coeffs = mod_optimization.weighted_least_squares_min(data_theta, data_pixel, 5, , 20)
            max_coeffs = mod_optimization.weighted_least_squares_max(data_theta, data_pixel, 5, , 20)
            
            min_poly = mod_optimization.poly_fit_seperate_coeff(data_theta, min_coeffs)
            max_poly = mod_optimization.poly_fit_seperate_coeff(data_theta, max_coeffs)
            
            min_avg = mod_optimization.average(mod_matrix.matVec(min_poly, 2))
            max_avg = mod_optimization.average(mod_matrix.matVec(max_poly, 2))
            
            radius = start_radius + ((k - 1) * dr)
            circumference = 2 * PI * radius
            
            final(k, 1) = max_avg
            final(k, 2) = min_avg
            final(k, 3) = LINEPAIR / circumference
            final(k, 4) = (max_avg - min_avg) / (max_avg + min_avg)
            
            If write_data = write_file Then
            
                bool = write_csv_double(new_2D, "circle_" & k & ".csv")
                bool = write_csv_double(min_poly, "circle_" & k & "_min" & ".csv")
                bool = write_csv_double(max_poly, "circle_" & k & "_max" & ".csv")
                
            End If
            
            j = 0
            
        Next k
        
        clean_data = final

End Function

Public Function draw_circles(ByRef grey_data&(), _
                             ByRef circle_center As t_center, _
                    Optional ByVal start_radius& = 10, _
                    Optional ByVal num_circles& = 10, _
                    Optional ByVal dr& = 10) As t_circle

    Dim rgb() As Byte, _
        analysis_data#(), _
        greyMaxRow&, _
        greyMaxCol&, _
        theta_max#, _
        max_size#, _
        t#, dt#, _
        k&, ra#, _
        x#, y#, _
        xr&, yr&, _
        xc&, yc&, _
        p&
        
        greyMaxRow = UBound(grey_data, 1)
        greyMaxCol = UBound(grey_data, 2)
        
        theta_max = 2 * PI
        dt = 1 / (100 * theta_max)
        max_size = Round(theta_max / dt)
        
        ReDim analysis_data(1 To num_circles, 1 To max_size, 1 To 2)
        rgb = mod_image_processing.dimension_2D_to_3D(grey_data)
                
        ra = start_radius
        k = 0
        Do
        
            If k = num_circles Then
                draw_circles.rgb_data = rgb
                draw_circles.analysis_data = analysis_data
                Exit Function
            End If
            
            t = 0
            p = 1
            Do
            
                x = ra * Cos(t) + circle_center.x
                y = ra * Sin(t) + circle_center.y
                xr = Round(x)
                yr = Round(y)
                
                If xr = 0 Then
                    xr = 1
                ElseIf yr = 0 Then
                    yr = 1
                End If
                
                If t = 0 Then
                
                    xc = xr
                    yc = yr
                    
                    rgb(1, xr, yr) = 0
                    rgb(2, xr, yr) = 255
                    rgb(3, xr, yr) = 0
                    
                    analysis_data(k + 1, p, 1) = t
                    analysis_data(k + 1, p, 2) = grey_data(xr, yr)
                    
                End If
                
                If xc <> xr Or yc <> yr Then
                    
                    rgb(1, xr, yr) = 0      'B
                    rgb(2, xr, yr) = 255    'G
                    rgb(3, xr, yr) = 0      'R
                    
                    analysis_data(k + 1, p, 1) = t
                    analysis_data(k + 1, p, 2) = grey_data(xr, yr)
                    
                    xc = xr
                    yc = yr
                
                End If
                
                t = t + dt
                p = p + 1
                
                If t > theta_max Then
                    Exit Do
                End If
                
            Loop

        ra = ra + dr
        k = k + 1
        
        Loop

End Function

Public Function find_center(ByRef grey_data&(), _
                            ByVal write_data As t_write, _
                            ByVal threshold&) As t_center

    Dim x_data&(), _
        y_data&(), _
        x_full&(), _
        y_full&(), _
        x_prune&(), _
        y_prune&(), _
        greyMaxRow&, _
        greyMaxCol&, _
        bool As Boolean, _
        xp As t_max_data, _
        yp As t_max_data, _
        sum&, _
        i&, j&
        
        greyMaxRow = UBound(grey_data, 1)
        greyMaxCol = UBound(grey_data, 2)

        ReDim x_data(1 To greyMaxCol)
        ReDim x_full(1 To greyMaxRow, 1 To 2)
        ReDim y_data(1 To greyMaxRow)
        ReDim y_full(1 To greyMaxCol, 1 To 2)
        
        'x data
        For i = 1 To greyMaxRow
            For j = 1 To greyMaxCol
            
                x_data(j) = grey_data(i, j)
            
            Next j
            
            sum = find_sum(x_data)
            x_full(i, 1) = i
            x_full(i, 2) = sum
            
        Next i
        
        'y data
        For j = 1 To greyMaxCol
            For i = 1 To greyMaxRow
            
                y_data(i) = grey_data(i, j)
            
            Next i
            
            sum = find_sum(y_data)
            y_full(j, 1) = j
            y_full(j, 2) = sum
            
        Next j
        
        'write data
        If write_data = write_file Then
            bool = mod_general_functions.write_csv(x_full, "x_full.csv")
            bool = mod_general_functions.write_csv(y_full, "y_full.csv")
        End If
        
        'prune if necessary
        x_prune = prune_data(x_full, threshold)
        y_prune = prune_data(y_full, threshold)
        
        'find center
        xp = find_min(x_prune, threshold)
        yp = find_min(y_prune, threshold)
        find_center.x = xp.pixel_position
        find_center.y = yp.pixel_position

End Function

Public Function find_sum&(ByRef raw_data&())

    Dim rawMaxRow&, _
        sum&, _
        i&
        
        rawMaxRow = UBound(raw_data, 1)
        
        sum = 0
        For i = 1 To rawMaxRow
            sum = sum + raw_data(i)
        Next i
        
        find_sum = sum
        
End Function

Public Function find_min(ByRef raw_data&(), _
                         ByVal threshold&) As t_max_data

    Dim rawMaxRow&, _
        rawMaxCol&, _
        current&, _
        min&, _
        i&
        
        rawMaxRow = UBound(raw_data, 1)
        rawMaxCol = UBound(raw_data, 2)
        
        find_min.magnitude = raw_data(threshold, 2)
        find_min.pixel_position = 1
        
        min = raw_data(threshold, 2)
        For i = (threshold + 1) To (rawMaxRow - threshold)
        
            current = raw_data(i, 2)
            
            If current < min Then
                find_min.magnitude = raw_data(i, 2)
                find_min.pixel_position = i
                min = current
            End If
        
        Next i

End Function

Public Function prune_data(ByRef raw_data&(), _
                           ByVal threshold&) As Long()

    Dim prune&(), _
        start&, _
        finish&, _
        rawMaxRow&, _
        rawMaxCol&, _
        i&, j&
        
        If threshold < 1 Then
            threshold = 1
        End If
        
        rawMaxRow = UBound(raw_data, 1)
        rawMaxCol = UBound(raw_data, 2)
        start = threshold
        finish = rawMaxRow - threshold
        
        ReDim prune(start To finish, 1 To rawMaxCol)
        
        For i = start To finish
            For j = 1 To rawMaxCol
            
                prune(i, j) = raw_data(i, j)
            
            Next j
        Next i
        
        prune_data = prune

End Function
