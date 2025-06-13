Attribute VB_Name = "mod_optimization"
Option Explicit

Public Type t_circle_stats
    avg_max_contrast As Double
    avg_min_contrast As Double
    frequency As Double
End Type

Public Function weighted_least_squares_min(ByRef Ax#(), _
                                           ByRef Ay#(), _
                                  Optional ByVal polyOrder# = 5, _
                                  Optional ByVal p# = 0.01, _
                                  Optional ByVal iter& = 5) As Double()
                                                 
    Dim rawMaxRowA&, _
        i&, j&, k&, _
        W#(), B1#(), _
        B2#(), B3#(), _
        B4#(), B5#(), _
        B6#(), Bf#(), _
        Xv#(), WeightedPoly#(), _
        y#, z#, _
        residual#
        
        rawMaxRowA = UBound(Ax, 1)
        ReDim Bf(1 To polyOrder, 1 To 1)
        
        W = mod_matrix.matSpy(rawMaxRowA)
        Xv = mod_optimization.vandermonde(Ax, polyOrder)
        
        k = 1
        Do
        
            B1 = mod_matrix.matTra(Xv)
            B2 = mod_matrix.matMul(B1, W)
            B3 = mod_matrix.matMul(B2, Xv)
            B4 = mod_matrix.matInv(B3)
            B5 = mod_matrix.matMul(B4, B1)
            B6 = mod_matrix.matMul(B5, W)
            Bf = mod_matrix.matMul(B6, Ay)
            
            WeightedPoly = mod_optimization.poly_fit_seperate_coeff(Ax, Bf)
        
            For i = 1 To rawMaxRowA
                
                y = Ay(i, 1)
                z = WeightedPoly(i, 2)
                residual = y - z
                
                If residual > 0 Then
                    W(i, i) = p
                Else
                    W(i, i) = 1 - p
                End If
                
            Next i
        
            k = k + 1
        
            If k > iter Then
                Exit Do
            End If
        
        Loop
    
        weighted_least_squares_min = Bf

End Function

Public Function weighted_least_squares_max(ByRef Ax#(), _
                                           ByRef Ay#(), _
                                  Optional ByVal polyOrder# = 5, _
                                  Optional ByVal p# = 0.01, _
                                  Optional ByVal iter& = 5) As Double()
                                                 
    Dim rawMaxRowA&, _
        i&, j&, k&, _
        W#(), B1#(), _
        B2#(), B3#(), _
        B4#(), B5#(), _
        B6#(), Bf#(), _
        Xv#(), WeightedPoly#(), _
        y#, z#, _
        residual#
        
        rawMaxRowA = UBound(Ax, 1)
        ReDim Bf(1 To polyOrder, 1 To 1)
        
        W = mod_matrix.matSpy(rawMaxRowA)
        Xv = mod_optimization.vandermonde(Ax, polyOrder)
        
        k = 1
        Do
        
            B1 = mod_matrix.matTra(Xv)
            B2 = mod_matrix.matMul(B1, W)
            B3 = mod_matrix.matMul(B2, Xv)
            B4 = mod_matrix.matInv(B3)
            B5 = mod_matrix.matMul(B4, B1)
            B6 = mod_matrix.matMul(B5, W)
            Bf = mod_matrix.matMul(B6, Ay)
            
            WeightedPoly = mod_optimization.poly_fit_seperate_coeff(Ax, Bf)
        
            For i = 1 To rawMaxRowA
                
                y = Ay(i, 1)
                z = WeightedPoly(i, 2)
                residual = y - z
                
                If residual > 0 Then
                    W(i, i) = 1 - p
                Else
                    W(i, i) = p
                End If
                
            Next i
        
            k = k + 1
        
            If k > iter Then
                Exit Do
            End If
        
        Loop
    
        weighted_least_squares_max = Bf

End Function

Public Function poly_fit_seperate_coeff(ByRef A#(), _
                                        ByRef coeff#()) As Double()

    Dim rawMaxRowA&, _
        rawMaxColA&, _
        rawMaxRowCoeff&, _
        i&, k&, _
        sum#, _
        C#()
        
        rawMaxRowA = UBound(A, 1)
        rawMaxColA = UBound(A, 2)
        rawMaxRowCoeff = UBound(coeff, 1)
        
        ReDim C(1 To rawMaxRowA, 1 To (rawMaxColA + 1))
        
        For i = 1 To rawMaxRowA
        
            sum = 0
            For k = 1 To rawMaxRowCoeff
            
                sum = sum + (coeff(k, 1) * (A(i, 1) ^ (k - 1)))
                    
            Next k
            
            C(i, 1) = A(i, 1)
            C(i, 2) = sum
            
        Next i
        
        poly_fit_seperate_coeff = C

End Function

Public Function pixel_distribution(ByRef pixel_data#()) As Long()

    Dim pixelMaxRow&, _
        byte_limit&, _
        pixel&, _
        dist&(), _
        i&, j&, k&
        
        pixelMaxRow = UBound(pixel_data, 1)
        byte_limit = 255
        
        ReDim dist(0 To byte_limit, 1 To 2)
        
        k = 0
        For j = 0 To byte_limit
            For i = 1 To pixelMaxRow
            
                pixel = pixel_data(i)
                
                If pixel = j Then
                    k = k + 1
                End If
            
            Next i
            
            dist(j, 1) = j
            dist(j, 2) = k
            k = 0
            
        Next j

        pixel_distribution = dist
        
End Function

Public Function average#(ByRef A#())

    Dim rawMaxRowA&, _
        avg#, _
        sum#, _
        i&
        
        rawMaxRowA = UBound(A, 1)
        
        sum = 0
        For i = 1 To rawMaxRowA
            sum = sum + A(i, 1)
        Next i
        
        avg = sum / rawMaxRowA
        average = avg

End Function

Public Function standard_deviation#(ByRef A#(), _
                                    ByVal avg#)

    Dim rawMaxRowA&, _
        sum#, _
        dev#, _
        i&
        
        rawMaxRowA = UBound(A, 1)
        
        sum = 0
        dev = 0
        For i = 1 To rawMaxRowA
            sum = (A(i, 1) - avg) ^ 2
            dev = dev + sum
        Next i
        
        dev = Sqr(dev / rawMaxRowA)
        standard_deviation = dev
                        
End Function

Public Function vandermonde(ByRef rawX#(), _
                            ByVal polyOrder&) As Double()

    Dim rawMaxRowX&, _
        i&, j&, _
        final#()
        
        rawMaxRowX = UBound(rawX, 1)
        
        ReDim final(1 To rawMaxRowX, 1 To (polyOrder + 1))
        
        For j = 0 To polyOrder
            For i = 1 To rawMaxRowX
            
                final(i, (j + 1)) = rawX(i, 1) ^ (j)
            
            Next i
        Next j

        vandermonde = final

End Function
