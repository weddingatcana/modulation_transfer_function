Attribute VB_Name = "prg_mtf"
Option Explicit

    #If Win64 Then
    
        Public Type t_gdiplus
            gdiplus_version             As Long
            debug_event_callback        As LongPtr
            suppress_background_thread  As Long
            suppress_external_codecs    As Long
        End Type
        
        Public Type t_bitmap
            type                        As Long
            width                       As Long
            height                      As Long
            width_bytes                 As Long
            planes                      As Integer
            bits_pixel                  As Integer
            bits                        As LongPtr
        End Type
        
        Public Type t_guid
            data1                       As Long
            data2                       As Integer
            data3                       As Integer
            data4(7)                    As Byte
        End Type
        
        Public Type t_description
            size                        As Long
            type                        As Long
            hBmp                        As LongPtr
            hPal                        As LongPtr
        End Type

    #Else
        
        Public Type t_gdiplus
            gdiplus_version             As Long
            debug_event_callback        As Long
            suppress_background_thread  As Long
            suppress_external_codecs    As Long
        End Type
        
        Public Type t_bitmap
            type                        As Long
            width                       As Long
            height                      As Long
            width_bytes                 As Long
            planes                      As Integer
            bits_pixel                  As Integer
            bits                        As Long
        End Type
        
        Public Type t_guid
            data1                       As Long
            data2                       As Integer
            data3                       As Integer
            data4(7)                    As Byte
        End Type
        
        Public Type t_description
            size                        As Long
            type                        As Long
            hBmp                        As Long
            hPal                        As Long
        End Type
        
    #End If
    
    Public Type t_dimensions
        width As Long
        height As Long
    End Type
    
    Public Enum greyType
        greyType_AVERAGE = 0
        greyType_LUMINANCE = 1
        greyType_DESATURATION = 2
        greyType_RED = 3
        greyType_GREEN = 4
        greyType_BLUE = 5
    End Enum
    
    Public global_grey&()
    Public global_rgb() As Byte
    Public global_final#()
    
    Const CF_BITMAP& = &H1

Sub prg_mtf()

    #If Win64 Then
    
       Dim hBmp As LongPtr, _
           image_width&, _
           image_width_bytes&, _
           image_height&, _
           image_color_channels&, _
 _
           image_description As t_description, _
           image_interface As IPicture, _
           guid As t_guid, _
           bmp_info As t_bitmap, _
 _
           image_input_path$, _
           image_input_extension$, _
           image_output_path$, _
           image_output_extension$, _
 _
           image_center As t_center, _
           image_threshold&, _
           image_mask() As Byte, _
           circle_data As t_circle, _
           circle_start_radius&, _
           circle_spacing&, _
           circle_number&, _
           bool As Boolean
    
    #Else
    
       Dim hBmp&, _
           image_width&, _
           image_width_bytes&, _
           image_height&, _
           image_color_channels&, _
 _
           image_description As t_description, _
           image_interface As IPicture, _
           guid As t_guid, _
           bmp_info As t_bitmap, _
 _
           image_input_path$, _
           image_input_extension$, _
           image_output_path$, _
           image_output_extension$, _
 _
           image_center As t_center, _
           image_threshold&, _
           image_mask() As Byte, _
           circle_data As t_circle, _
           circle_start_radius&, _
           circle_spacing&, _
           circle_number, _
           bool As Boolean

    #End If

           guid.data1 = &H7BF80980
           guid.data2 = &HBF32
           guid.data3 = &H101A
           guid.data4(0) = &H8B
           guid.data4(1) = &HBB
           guid.data4(2) = &H0
           guid.data4(3) = &HAA
           guid.data4(4) = &H0
           guid.data4(5) = &H30
           guid.data4(6) = &HC
           guid.data4(7) = &HAB
                        
           image_input_path = mod_general_functions.choose_image_file
           
           If Len(image_input_path) = 0 Then
               Exit Sub
           End If
           
           image_input_extension = mod_general_functions.choose_image_file_extension(image_input_path)
           image_output_path = "C:\Users\qp\Desktop\mtf\test"
           image_output_extension = image_input_extension
            
           hBmp = mod_general_functions.create_device_independent_bitmap(image_input_path)
           
           image_description.size = LenB(image_description)
           image_description.type = CF_BITMAP
           image_description.hBmp = hBmp
           image_description.hPal = 0&

           mod_api.GetObject hBmp, LenB(bmp_info), bmp_info
           
           image_width = bmp_info.width
           image_width_bytes = bmp_info.width_bytes
           image_height = bmp_info.height
           image_color_channels = bmp_info.bits_pixel \ 8
            
           ReDim global_rgb(1 To image_color_channels, 1 To image_width, 1 To image_height)
           mod_api.GetBitmapBits hBmp, (image_width_bytes * image_height), global_rgb(1, 1, 1)
            
           ReDim global_grey(1 To image_width, 1 To image_height)
           global_grey = mod_image_processing.raw_image_to_grey(global_rgb, greyType_LUMINANCE)
            
           image_threshold = 100
           image_center = mod_image_processing.find_center(global_grey, write_file, image_threshold)
            
           circle_start_radius = 10
           circle_spacing = 10
           circle_number = 10
           circle_data = mod_image_processing.draw_circles(global_grey, image_center, circle_start_radius, circle_number, circle_spacing)
           
           image_mask = mod_image_processing.draw_center(circle_data.rgb_data, image_center.x, image_center.y)
            
           mod_api.SetBitmapBits hBmp, (image_width_bytes * image_height), image_mask(1, 1, 1)
           mod_general_functions.save_image image_description, guid, image_interface, image_output_path, image_output_extension
            
           global_final = mod_image_processing.clean_data(circle_data.analysis_data, circle_start_radius, circle_spacing, no_file)
           bool = mod_general_functions.write_csv_double(global_final, "final.csv")
            
End Sub
